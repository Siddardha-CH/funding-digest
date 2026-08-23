"""
Daily Startup Funding Digest
-----------------------------
Pulls recent startup-funding news from free RSS feeds, filters for
funding-related keywords, and emails you a clean HTML digest.

SETUP:
1. pip install feedparser
2. Set these environment variables (or GitHub Secrets if using Actions):
     GMAIL_ADDRESS   = your gmail address
     GMAIL_APP_PASSWORD = the 16-char app password from Google Account settings
     TO_EMAIL        = where you want the digest sent (can be same as GMAIL_ADDRESS)
3. Run: python funding_digest.py

NOTE ON CONTACTS: This script only aggregates public funding *news*.
It intentionally does NOT scrape LinkedIn or generate founder emails —
that data is unreliable when scraped and LinkedIn's terms prohibit
automated scraping. For companies you're actually interested in,
look up the founder manually (LinkedIn search) and try a free tool
like Hunter.io (25 free lookups/month) for a likely email pattern.
"""

import os
import re
import smtplib
import ssl
from datetime import datetime, timedelta, timezone
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import feedparser
from openpyxl import Workbook
from openpyxl.styles import Font

# --- Free RSS sources that cover startup funding news ---
FEEDS = {
    "TechCrunch": "https://techcrunch.com/category/venture/feed/",
    "VentureBeat": "https://venturebeat.com/category/venture/feed/",
    "Crunchbase News": "https://news.crunchbase.com/feed/",
    "Forbes": "https://www.forbes.com/business/feed/",
    "Business Insider": "https://www.businessinsider.com/rss",
    "Fortune": "https://fortune.com/feed/",
    "Axios": "https://api.axios.com/feed/",
    "Inc42 (India)": "https://inc42.com/feed/",
    "YourStory (India)": "https://yourstory.com/feed",
    "Entrackr (India)": "https://entrackr.com/feed",
    "Economic Times Startups": "https://economictimes.indiatimes.com/tech/startups/rssfeeds/13358288.cms",
    "TechInAsia": "https://www.techinasia.com/feed",
}

# Keywords that signal an actual funding announcement
FUNDING_KEYWORDS = re.compile(
    r"\b(raises?|raised|funding|fundraise|seed round|pre-seed|series [a-e]\b|"
    r"closes.*round|secures.*(million|funding|crore)|"
    r"valued at|investment (round|of)|backed by|led by \w+ (capital|ventures|partners))\b",
    re.IGNORECASE,
)

# Optional: extra keywords to prioritize companies likely to be hiring
# engineers (mentions of "hiring", "expand", "scale engineering", etc.)
HIRING_SIGNAL = re.compile(
    r"\b(hire|hiring|expand|scale|grow the team|engineering team)\b",
    re.IGNORECASE,
)

LOOKBACK_HOURS = 48  # widened so a slow news day doesn't come up empty

# --- Best-effort extraction patterns ---
# News headlines are inconsistent, so these are "best guess" patterns, not
# guaranteed extraction. Blank cells mean the article didn't state it plainly
# enough for a simple pattern to catch — check the link for the full story.

# "Acme raises $12M" / "Acme raised $3.5 million" -> company + amount
COMPANY_AMOUNT = re.compile(
    r"^([A-Z][\w&.\- ]{1,50}?)\s+(?:raises?|raised|secures?|closes?|lands?|nets?)\s+"
    r"(\$[\d.,]+\s?(?:[MmBbKk]illion|[MmBb]|[Kk])?)",
)

# "led by Sequoia Capital" / "backed by Andreessen Horowitz"
INVESTORS = re.compile(
    r"\b(?:led by|backed by|from|with participation from)\s+"
    r"([A-Z][\w&.\- ]{2,60}?)(?:[.,]|\s+to\s|\s+for\s|$)",
)

# "plans to use the funds to expand its engineering team" / "will use the
# money to hire" — captures the goal clause when explicitly stated
GOALS = re.compile(
    r"(?:plans? to|will|to)\s+use\s+the\s+(?:funds?|money|capital)\s+to\s+"
    r"([^.]{5,120})",
    re.IGNORECASE,
)


def extract_company_and_amount(title: str):
    m = COMPANY_AMOUNT.search(title)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return "", ""


def extract_investors(text: str) -> str:
    m = INVESTORS.search(text)
    return m.group(1).strip() if m else ""


def extract_goals(text: str) -> str:
    m = GOALS.search(text)
    return m.group(1).strip() if m else ""


def entry_is_recent(entry) -> bool:
    for field in ("published_parsed", "updated_parsed"):
        t = getattr(entry, field, None)
        if t:
            published = datetime(*t[:6], tzinfo=timezone.utc)
            return published > datetime.now(timezone.utc) - timedelta(hours=LOOKBACK_HOURS)
    return True  # if no date, include it rather than risk missing it


def collect_funding_news():
    items = []
    for source, url in FEEDS.items():
        try:
            parsed = feedparser.parse(url)
        except Exception as e:
            print(f"Could not fetch {source}: {e}")
            continue

        for entry in parsed.entries:
            title = getattr(entry, "title", "")
            summary = getattr(entry, "summary", "")
            if not FUNDING_KEYWORDS.search(title + " " + summary):
                continue
            if not entry_is_recent(entry):
                continue

            clean_summary = re.sub("<[^<]+?>", "", summary)[:280]
            combined_text = title + " " + clean_summary
            company, amount = extract_company_and_amount(title)

            items.append(
                {
                    "source": source,
                    "title": title,
                    "link": getattr(entry, "link", ""),
                    "summary": clean_summary,  # doubles as "what they do" description
                    "company": company,
                    "amount": amount,
                    "investors": extract_investors(combined_text),
                    "goals": extract_goals(combined_text),
                    "hiring_signal": bool(HIRING_SIGNAL.search(combined_text)),
                }
            )
    return items


def build_html(items):
    if not items:
        return "<p>No new funding news matched today. Try widening the feed list.</p>"

    # Put likely-hiring companies first
    items.sort(key=lambda x: not x["hiring_signal"])

    rows = []
    for it in items:
        tag = "🟢 hiring signal" if it["hiring_signal"] else ""
        rows.append(
            f"""
            <div style="margin-bottom:18px;padding:12px;border:1px solid #ddd;border-radius:8px;">
              <div style="font-size:12px;color:#888;">{it['source']} {tag}</div>
              <a href="{it['link']}" style="font-size:16px;font-weight:bold;color:#1a0dab;text-decoration:none;">
                {it['title']}
              </a>
              <p style="font-size:14px;color:#444;">{it['summary']}...</p>
            </div>
            """
        )
    return "<h2>Startup Funding Digest — {}</h2>{}".format(
        datetime.now().strftime("%B %d, %Y"), "".join(rows)
    )


def build_excel(items, path="funding_digest.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Funding Digest"

    headers = [
        "Company Name",
        "What They Do",
        "Amount Raised",
        "Investors",
        "Stated Goals",
        "Hiring Signal",
        "Source",
        "Link",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Hiring-signal companies first, same as the email
    items_sorted = sorted(items, key=lambda x: not x["hiring_signal"])
    for it in items_sorted:
        ws.append(
            [
                it["company"] or it["title"],  # fall back to headline if name wasn't parsed
                it["summary"],
                it["amount"],
                it["investors"],
                it["goals"],
                "Yes" if it["hiring_signal"] else "",
                it["source"],
                it["link"],
            ]
        )

    # Reasonable column widths so it's readable without manual resizing
    widths = [24, 55, 14, 28, 40, 12, 14, 45]
    for col, width in zip("ABCDEFGH", widths):
        ws.column_dimensions[col].width = width

    wb.save(path)
    return path


def send_email(html_body: str, excel_path: str | None = None):
    gmail_address = os.environ["GMAIL_ADDRESS"]
    gmail_password = os.environ["GMAIL_APP_PASSWORD"]
    to_email = os.environ.get("TO_EMAIL", gmail_address)

    msg = MIMEMultipart("mixed")
    msg["Subject"] = f"Funding Digest — {datetime.now().strftime('%b %d')}"
    msg["From"] = gmail_address
    msg["To"] = to_email

    alt_part = MIMEMultipart("alternative")
    alt_part.attach(MIMEText(html_body, "html"))
    msg.attach(alt_part)

    if excel_path and os.path.exists(excel_path):
        with open(excel_path, "rb") as f:
            attachment = MIMEApplication(f.read(), _subtype="xlsx")
            attachment.add_header(
                "Content-Disposition", "attachment", filename=os.path.basename(excel_path)
            )
            msg.attach(attachment)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(gmail_address, gmail_password)
        server.sendmail(gmail_address, to_email, msg.as_string())


if __name__ == "__main__":
    news = collect_funding_news()
    html = build_html(news)
    excel_file = build_excel(news) if news else None
    send_email(html, excel_file)
    print(f"Sent digest with {len(news)} items.")
