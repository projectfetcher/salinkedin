"""
LinkedIn Job Scraper — Python v1
Converted from Google Apps Script v5

Requirements:
    pip install requests beautifulsoup4 openpyxl

Usage:
    python linkedin_job_scraper.py

Output:
    jobs_output.xlsx  (same columns as the original Google Sheet)
"""

import re
import time
import base64
import json
import logging
from datetime import datetime, timedelta
from urllib.parse import urljoin, urlparse, unquote
import requests
from bs4 import BeautifulSoup
import openpyxl

# =============================================================================
#  CONFIG
# =============================================================================

SHEET_NAME  = "Sheet1"
PAGES       = 1          # 1 page ≈ 25 results
JOB_LIMIT   = 45
DELAY_S     = 2.5        # seconds between requests
FETCH_CHAR_LIMIT = 120_000

SEARCH_BASE = "https://www.linkedin.com/jobs/search?keywords=&location=Saudi-Arabia&start="

OUTPUT_FILE = "jobs_output.xlsx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}

# WordPress credentials (optional — set to empty strings to skip logo upload)
WP_URL      = "https://mauritius.mimusjobs.com/wp-json/wp/v2/"
WP_USER     = "calolina"
WP_PASSWORD = "st8a 6mWY wqgV 0syR mB3i y5FQ"

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

# ANSI colours (auto-disabled if terminal doesn't support them)
import sys
_USE_COLOUR = sys.stdout.isatty()

def _c(code, text):
    return f"\033[{code}m{text}\033[0m" if _USE_COLOUR else text

C_HEADER  = lambda t: _c("1;36",  t)   # bold cyan
C_LABEL   = lambda t: _c("1;33",  t)   # bold yellow
C_VALUE   = lambda t: _c("97",    t)   # bright white
C_DIM     = lambda t: _c("2",     t)   # dim
C_GREEN   = lambda t: _c("1;32",  t)   # bold green
C_RED     = lambda t: _c("1;31",  t)   # bold red
C_BLUE    = lambda t: _c("1;34",  t)   # bold blue
C_DIVIDER = lambda: _c("2", "─" * 72)


def print_job_verbose(job: dict, index: int, total: int):
    """Print a full human-readable summary of a scraped job to stdout."""

    desc = job.get("jobDescription", "")
    # Show first 400 chars of description, then ellipsis
    desc_preview = (desc[:400] + " [...]") if len(desc) > 400 else desc
    # Indent description lines
    desc_indented = "\n".join("   " + line for line in desc_preview.splitlines() if line.strip())

    apply = job.get("application", "")
    apply_display = apply if apply else C_DIM("— not found —")

    logo = job.get("companyLogo", "")
    logo_display = logo if logo else C_DIM("— none —")

    print()
    print(C_DIVIDER())
    print(C_HEADER(f"  JOB {index}/{total}"))
    print(C_DIVIDER())

    # ── Core job info ─────────────────────────────────────────────────────────
    print(f"  {C_LABEL('Title')}          : {C_VALUE(job.get('jobTitle', ''))}")
    print(f"  {C_LABEL('Job Type')}       : {job.get('jobType', '')}")
    print(f"  {C_LABEL('Field')}          : {job.get('jobField', '') or C_DIM('—')}")
    print(f"  {C_LABEL('Location')}       : {job.get('jobLocation', '') or C_DIM('—')}")
    print(f"  {C_LABEL('Seniority')}      : {job.get('jobExperience', '') or C_DIM('—')}")
    print(f"  {C_LABEL('Qualifications')} : {job.get('jobQualifications', '') or C_DIM('—')}")
    print(f"  {C_LABEL('Salary')}         : {job.get('salaryRange', '') or C_DIM('—')}")
    print(f"  {C_LABEL('Date Posted')}    : {job.get('datePosted', '') or C_DIM('—')}")
    print(f"  {C_LABEL('Deadline')}       : {job.get('deadline', '') or C_DIM('—')}")
    print(f"  {C_LABEL('Est. Deadline')}  : {job.get('estimatedDeadline', '') or C_DIM('—')}")

    # ── Application ───────────────────────────────────────────────────────────
    print(f"  {C_LABEL('Apply Link')}     : {C_GREEN(apply) if apply else apply_display}")

    # ── Company ───────────────────────────────────────────────────────────────
    print()
    print(f"  {C_BLUE('── COMPANY ──────────────────────────────────────────────')}")
    print(f"  {C_LABEL('Name')}           : {C_VALUE(job.get('companyName', '') or C_DIM('—'))}")
    print(f"  {C_LABEL('Industry')}       : {job.get('companyIndustry', '') or C_DIM('—')}")
    print(f"  {C_LABEL('Type')}           : {job.get('companyType', '') or C_DIM('—')}")
    print(f"  {C_LABEL('Founded')}        : {job.get('companyFounded', '') or C_DIM('—')}")
    print(f"  {C_LABEL('Headquarters')}   : {job.get('companyAddress', '') or C_DIM('—')}")
    print(f"  {C_LABEL('Website')}        : {job.get('companyWebsite', '') or C_DIM('—')}")
    print(f"  {C_LABEL('LinkedIn URL')}   : {job.get('companyUrl', '') or C_DIM('—')}")
    print(f"  {C_LABEL('Logo')}           : {logo_display}")

    # ── Company about ─────────────────────────────────────────────────────────
    about = job.get("companyDetails", "")
    if about:
        about_preview = (about[:200] + " [...]") if len(about) > 200 else about
        print(f"  {C_LABEL('About')}          : {about_preview}")

    # ── Description preview ───────────────────────────────────────────────────
    print()
    print(f"  {C_BLUE('── DESCRIPTION PREVIEW ──────────────────────────────────')}")
    if desc_indented:
        print(desc_indented)
    else:
        print(C_DIM("   — no description —"))

    # ── Raw URL ───────────────────────────────────────────────────────────────
    print()
    print(f"  {C_LABEL('Job URL')}        : {C_DIM(job.get('jobUrl', ''))}")
    print(C_DIVIDER())

# =============================================================================
#  SKIP-CRAWL DOMAINS
# =============================================================================

SKIP_CRAWL_DOMAINS = [
    "dhl.com", "fedex.com", "ups.com",
    "amazon.com", "amazon.jobs",
    "google.com", "microsoft.com", "apple.com",
    "meta.com", "ibm.com", "oracle.com", "sap.com",
    "accenture.com", "deloitte.com", "pwc.com", "kpmg.com", "ey.com",
    "mckinsey.com", "bcg.com", "bain.com",
    "citibank.com", "hsbc.com", "barclays.com", "bnpparibas.com",
    "airbus.com", "boeing.com", "siemens.com", "ge.com",
    "unilever.com", "nestle.com", "pg.com", "shell.com", "bp.com",
]

BAD_DOMAINS = [
    "linkedin.com", "google.com", "youtube.com", "facebook.com",
    "twitter.com", "x.com", "instagram.com", "t.co", "example.com",
    "w3.org", "sentry.io", "schema.org",
]

NOISE_EMAIL_DOMAINS = [
    "example.com", "sentry.io", "google.com", "w3.org",
    "schema.org", "wixpress.com", "squarespace.com",
]

FAKE_LOCAL_RE  = re.compile(
    r"^(name|user|email|mail|yourname|your[-_.]?email|sample|test|info|hello"
    r"|noreply|no[-_.]?reply|admin|webmaster|support|contact|example)$", re.I)
FAKE_DOMAIN_RE = re.compile(
    r"^(domain|example|yoursite|yourdomain|yourbrand|company|mycompany"
    r"|website|yourcompany|mysite|placeholder|site)\.[a-z]{2,}$", re.I)

MONTH_MAP = {
    "jan": 0, "feb": 1, "mar": 2, "apr": 3, "may": 4, "jun": 5,
    "jul": 6, "aug": 7, "sep": 8, "oct": 9, "nov": 10, "dec": 11,
}

# =============================================================================
#  HELPERS
# =============================================================================

def should_skip_crawl(url: str) -> bool:
    if not url:
        return True
    lower = url.lower()
    return any(d in lower for d in SKIP_CRAWL_DOMAINS)


def is_bad_url(url: str) -> bool:
    if not url or not url.startswith("http"):
        return True
    lower = url.lower()
    return any(d in lower for d in BAD_DOMAINS)


def is_career_url(url: str) -> bool:
    l = url.lower()
    return any(k in l for k in ["career", "jobs", "apply", "vacanci", "recruit", "opening", "hiring", "work-with"])


def is_contact_url(url: str) -> bool:
    l = url.lower()
    return any(k in l for k in ["contact", "about", "reach", "get-in", "enquir", "support"])


def make_absolute(href: str, root_url: str) -> str:
    if not href:
        return ""
    href = href.strip()
    if href.startswith("http"):
        return href
    if href.startswith("//"):
        return "https:" + href
    if href.startswith("/"):
        return root_url.rstrip("/") + href
    return ""


def decode_html_entities(s: str) -> str:
    if not s:
        return ""
    replacements = [
        ("&amp;",  "&"), ("&lt;",   "<"), ("&gt;",   ">"),
        ("&quot;", '"'), ("&#39;",  "'"), ("\\u0026", "&"),
        ("\\u003D","="), ("\\u003A",":"), ("\\u002F","/"),
    ]
    for old, new in replacements:
        s = s.replace(old, new)
    return s

# =============================================================================
#  HTTP
# =============================================================================

def fetch_page(url: str, follow_redirects: bool = True) -> str | None:
    try:
        time.sleep(0.3)
        r = requests.get(url, headers=HEADERS, allow_redirects=follow_redirects, timeout=15)
        if r.status_code != 200:
            return None
        text = r.text
        if len(text) > FETCH_CHAR_LIMIT:
            text = text[:FETCH_CHAR_LIMIT]
        return text
    except Exception as e:
        log.warning(f"fetchPage failed ({url}): {e}")
        return None

# =============================================================================
#  NORMALISE DATE TEXT
# =============================================================================

def normalise_date_text(text: str) -> str:
    if not text:
        return ""
    fr_map = {
        "heure":"hour","heures":"hours","jour":"day","jours":"days",
        "semaine":"week","semaines":"weeks","mois":"month","an":"year","ans":"years",
    }
    m = re.search(r"il\s+y\s+a\s+(\d+)\s+([a-zéè]+)", text, re.I)
    if m:
        unit = fr_map.get(m.group(2).lower())
        if unit:
            return f"{m.group(1)} {unit} ago"
    if re.match(r"^hier$", text.strip(), re.I):
        return "1 day ago"
    if re.search(r"aujourd|today", text, re.I):
        return "0 days ago"
    return text

# =============================================================================
#  DESCRIPTION CLEANER
# =============================================================================

def clean_description(raw: str) -> str:
    if not raw:
        return ""
    text = raw
    text = text.replace("\u00a0", " ").replace("\u200b", "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    text = re.sub(r"([a-z])([A-Z])", r"\1 \2", text)
    text = re.sub(r"(\d)([A-Za-z])", r"\1 \2", text)
    text = re.sub(r"([A-Za-z])(\d)", r"\1 \2", text)
    text = re.sub(r"([.,:;!?])([A-Za-z0-9])", r"\1 \2", text)
    text = re.sub(r"\s*[•·▪◦]\s*", "\n• ", text)
    text = re.sub(r"\n\s*[-–—]\s+", "\n• ", text)

    paragraphs = re.split(r"\n{2,}", text)
    cleaned_paras = []
    for para in paragraphs:
        lines = para.split("\n")
        cleaned_lines = []
        for line in lines:
            line = line.strip()
            if not line:
                continue
            ends_with_punct = bool(re.search(r"[.!?:;,]$", line))
            looks_like_heading = bool(re.match(r"^[A-Z\s]{3,30}$", line)) and len(line.split()) <= 5
            if not ends_with_punct and not looks_like_heading and len(line) > 8:
                if not re.match(r"^[•\-–]", line) and not re.match(r"^\w+:$", line):
                    line = line + "."
            cleaned_lines.append(line)
        cleaned_paras.append("\n".join(cleaned_lines))

    text = "\n\n".join(p for p in cleaned_paras if p.strip())
    text = re.sub(r" {2,}", " ", text)
    return text.strip()

# =============================================================================
#  EMAIL CLEANER
# =============================================================================

def clean_email(raw: str) -> str:
    if not raw:
        return ""
    em = raw
    em = re.sub(r"^mailto:", "", em, flags=re.I)
    em = re.sub(r"\?.*$", "", em)
    for pattern, replacement in [
        (r"\\u003[Ee]", ""), (r"\\u003[Cc]", ""), (r"\\u0040", "@"),
        (r"\\u002[Ee]", "."), (r"\\u0026", ""), (r"u003[Ee]", ""),
        (r"u003[Cc]", ""), (r"u0040", "@"), (r"&amp;", ""),
        (r"&lt;", ""), (r"&gt;", ""), (r"&#64;", "@"), (r"&#46;", "."),
        (r"&nbsp;", ""), (r"%40", "@"), (r"%2[Ee]", "."), (r"%20", ""),
        (r"[>]+$", ""), (r"[<]+$", ""),
    ]:
        em = re.sub(pattern, replacement, em, flags=re.I)
    em = em.strip().lower()

    if not em or "@" not in em or "." not in em:
        return ""
    if not re.match(r"^[a-zA-Z0-9]", em):
        return ""

    at_index = em.rfind("@")
    if at_index == -1:
        return ""
    local  = em[:at_index]
    domain = em[at_index + 1:]

    if ".mu" in domain:
        domain = re.sub(r"\.mu.*", ".mu", domain, flags=re.I)
    elif ".uk" in domain:
        domain = re.sub(r"\.uk.*", ".uk", domain, flags=re.I)
    else:
        domain = re.sub(r"(\.[a-z]{2,6})[a-z0-9\-_/?#+]*$", r"\1", domain, flags=re.I)

    em = local + "@" + domain
    if not em or "@" not in em or "." not in em:
        return ""
    if not re.match(r"^[a-zA-Z0-9]", em):
        return ""
    return em


def clean_application_link(raw: str) -> str:
    if not raw:
        return ""
    raw = raw.strip()
    is_email = "@" in raw and not raw.startswith("http")
    if is_email:
        return clean_email(raw)
    if raw.startswith("http"):
        url = raw
        if ".mu" in url.lower():
            def mu_replace(m):
                tld, path = m.group(1), m.group(2) or ""
                if path and re.match(r"^/[a-z0-9\-/]+$", path, re.I):
                    return tld + path
                return tld
            url = re.sub(r"(\.mu)(\/[^\s]*)?$", mu_replace, url, flags=re.I)
        url = re.sub(r"#.*$", "", url)
        url = re.sub(r"(subject|applysubject|refno|applyref|applyhere|clickhere|applynow)(\?.*)?$", "", url, flags=re.I)
        url = re.sub(r"[.,;:!?)]+$", "", url)
        return url.strip()
    return raw


def clean_logo_url(raw: str) -> str:
    if not raw:
        return ""
    raw = decode_html_entities(raw).strip()
    if not raw.startswith("http"):
        return ""
    raw = re.sub(r"[\"')\s]+$", "", raw)
    return raw

# =============================================================================
#  EMAIL EXTRACTOR
# =============================================================================

def extract_email_from_text(text: str) -> str:
    if not text:
        return ""
    emails = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text)
    for raw_em in emails:
        em = clean_email(raw_em)
        if not em or "@" not in em:
            continue
        parts = em.split("@")
        if len(parts) != 2:
            continue
        if any(em.find(d) != -1 for d in NOISE_EMAIL_DOMAINS):
            continue
        if FAKE_LOCAL_RE.match(parts[0]):
            continue
        if FAKE_DOMAIN_RE.match(parts[1]):
            continue
        return em
    return ""

# =============================================================================
#  DEADLINE PARSER
# =============================================================================

def try_parse_date(s: str) -> datetime | None:
    if not s:
        return None
    try:
        d = datetime.fromisoformat(s)
        return d
    except Exception:
        pass
    try:
        return datetime.strptime(s, "%B %d, %Y")
    except Exception:
        pass
    try:
        return datetime.strptime(s, "%d %B %Y")
    except Exception:
        pass
    m = re.match(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", s)
    if m:
        mon = MONTH_MAP.get(m.group(2)[:3].lower())
        if mon is not None:
            return datetime(int(m.group(3)), mon + 1, int(m.group(1)))
    return None


def parse_deadline(soup: BeautifulSoup) -> str:
    full_text = soup.get_text()
    patterns = [
        r"closes?\s+on\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"closes?\s+on\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
        r"apply\s+by\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"apply\s+by\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
        r"applications?\s+close[sd]?\s*(?:on)?\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"deadline[:\s]+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"deadline[:\s]+(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})",
        r"closing\s+date[:\s]+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"closing\s+date[:\s]+(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
    ]
    now = datetime.now()
    for pattern in patterns:
        m = re.search(pattern, full_text, re.I)
        if m:
            d = try_parse_date(m.group(1))
            if d and d > now:
                return d.strftime("%Y-%m-%d")
    return ""


def estimate_deadline_from_posted(posted_text: str) -> str:
    if not posted_text:
        return ""
    text = normalise_date_text(posted_text)
    base = datetime.now()
    m = re.search(r"(\d+)\s*(hour|day|week|month)", text, re.I)
    if m:
        n, unit = int(m.group(1)), m.group(2).lower()
        if "hour"  in unit: base -= timedelta(hours=n)
        elif "day"  in unit: base -= timedelta(days=n)
        elif "week" in unit: base -= timedelta(weeks=n)
        elif "month" in unit:
            month = base.month - n
            year  = base.year + month // 12
            month = month % 12 or 12
            base  = base.replace(year=year, month=month)
    # Add 3 months for estimated close
    month = base.month + 3
    year  = base.year + (month - 1) // 12
    month = (month - 1) % 12 + 1
    base  = base.replace(year=year, month=month)
    return base.strftime("%Y-%m-%d")


def resolve_posted_date(raw: str) -> str:
    if not raw:
        return ""
    text = normalise_date_text(raw)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", text.strip()):
        return text.strip()
    try:
        d = datetime.fromisoformat(text)
        return d.strftime("%Y-%m-%d")
    except Exception:
        pass
    base = datetime.now()
    m = re.search(r"(\d+)\s*(hour|day|week|month|year)", text, re.I)
    if m:
        n, unit = int(m.group(1)), m.group(2).lower()
        if "hour"  in unit: base -= timedelta(hours=n)
        elif "day"  in unit: base -= timedelta(days=n)
        elif "week" in unit: base -= timedelta(weeks=n)
        elif "month" in unit:
            month = base.month - n
            year  = base.year + month // 12
            month = month % 12 or 12
            base  = base.replace(year=year, month=month)
        elif "year" in unit:
            base = base.replace(year=base.year - n)
        return base.strftime("%Y-%m-%d")
    if re.search(r"just\s*now|today", text, re.I):
        return datetime.now().strftime("%Y-%m-%d")
    return datetime.now().strftime("%Y-%m-%d")

# =============================================================================
#  JOB FIELD INFERENCE
# =============================================================================

FIELD_KEYWORD_MAP = [
    ("Information Technology",
     ["software engineer","developer","devops","frontend","backend","full stack","fullstack",
      "sysadmin","cloud","cybersecurity","data engineer","machine learning","artificial intelligence",
      "ai/ml","it support","network engineer","database","kubernetes","docker","aws","azure",
      "react","node.js","python developer","java developer"],
     ["programming","coding","api","agile","scrum","git","linux","server","infrastructure","software"]),

    ("Finance & Accounting",
     ["accountant","auditor","finance manager","financial analyst","cfo","treasurer","tax",
      "bookkeeper","payroll","budget analyst","credit analyst","investment","portfolio manager",
      "risk analyst","forex","actuary","acca","cfa","cpa"],
     ["financial","accounting","balance sheet","p&l","reconciliation","ifrs","gaap","ledger","invoicing"]),

    ("Sales & Business Development",
     ["sales executive","sales manager","business development","account manager",
      "sales representative","bd manager","regional sales","key account","sales director",
      "commercial manager","sales officer"],
     ["revenue","pipeline","crm","leads","prospects","quota","target","upsell","cross-sell","b2b","b2c"]),

    ("Marketing & Communications",
     ["marketing manager","digital marketing","seo","sem","content marketer","social media manager",
      "brand manager","marketing executive","communications manager","pr manager","copywriter",
      "growth hacker","email marketing","campaign manager"],
     ["marketing","branding","advertising","social media","content","campaign","analytics",
      "google ads","facebook ads","influencer"]),

    ("Human Resources",
     ["hr manager","human resources","recruiter","talent acquisition","hr business partner",
      "hrbp","hr officer","compensation","benefits manager","organisational development",
      "learning and development","l&d","hr generalist","payroll manager"],
     ["recruitment","onboarding","performance management","employee relations","hr","workforce",
      "headhunting","staffing"]),

    ("Engineering",
     ["mechanical engineer","civil engineer","electrical engineer","structural engineer",
      "process engineer","project engineer","maintenance engineer","production engineer",
      "quality engineer","safety engineer","site engineer","design engineer"],
     ["engineering","cad","autocad","solidworks","manufacturing","plant","machinery",
      "commissioning","maintenance"]),

    ("Healthcare & Medicine",
     ["doctor","physician","nurse","pharmacist","medical officer","surgeon","anaesthetist",
      "physiotherapist","radiographer","lab technician","clinical","healthcare manager",
      "occupational therapist","dentist","midwife"],
     ["hospital","clinic","patient","medical","health","pharmaceutical","diagnosis","treatment","ward"]),

    ("Education & Training",
     ["teacher","lecturer","professor","trainer","educator","tutor","school principal",
      "academic","curriculum","e-learning","instructional designer","teaching assistant"],
     ["school","university","college","classroom","students","pedagogy","curriculum","education","training"]),

    ("Hospitality & Tourism",
     ["hotel manager","front desk","housekeeping","chef","sous chef","food and beverage",
      "f&b manager","restaurant manager","bartender","waiter","concierge","tour guide",
      "travel agent","events coordinator","catering"],
     ["hospitality","hotel","resort","tourism","guest","accommodation","restaurant","kitchen","culinary"]),

    ("Logistics & Supply Chain",
     ["supply chain manager","logistics coordinator","warehouse manager","fleet manager",
      "procurement manager","purchasing manager","import export","freight","shipping coordinator",
      "inventory manager","demand planner"],
     ["logistics","supply chain","warehouse","inventory","freight","procurement","sourcing",
      "distribution","customs"]),

    ("Legal",
     ["lawyer","attorney","legal counsel","paralegal","compliance officer","legal advisor",
      "solicitor","barrister","corporate counsel","legal manager","contract manager"],
     ["legal","law","contracts","litigation","regulatory","compliance","gdpr","intellectual property"]),

    ("Administration & Operations",
     ["office manager","executive assistant","administrative officer","operations manager",
      "pa","personal assistant","receptionist","data entry","office administrator",
      "company secretary","business analyst"],
     ["administration","operations","office","coordination","scheduling","reporting","clerical","filing"]),

    ("Customer Service",
     ["customer service","call centre","customer success","customer support","help desk",
      "service advisor","client relations","customer experience","contact centre"],
     ["customer","support","helpdesk","tickets","escalation","satisfaction","service level",
      "inbound","outbound"]),

    ("Construction & Real Estate",
     ["quantity surveyor","site supervisor","project manager construction","architect",
      "draughtsman","property manager","estate agent","real estate","building inspector",
      "land surveyor","construction manager"],
     ["construction","building","property","real estate","site","contractor","tender","bof","drawings"]),

    ("Manufacturing & Production",
     ["production manager","quality control","quality assurance","qa","qc","factory manager",
      "plant manager","production supervisor","assembly","cnc operator","technician"],
     ["production","manufacturing","factory","assembly","quality","lean","six sigma","ohs","safety","line"]),

    ("Design & Creative",
     ["graphic designer","ui/ux","product designer","art director","creative director",
      "animator","illustrator","photographer","videographer","motion designer","web designer"],
     ["design","creative","adobe","figma","photoshop","illustrator","indesign","sketch","branding","visual"]),

    ("Research & Science",
     ["research scientist","data scientist","lab researcher","research analyst",
      "clinical researcher","environmental scientist","chemist","biologist",
      "statistician","epidemiologist"],
     ["research","analysis","data","laboratory","science","experiment","findings","methodology","survey"]),

    ("Security",
     ["security officer","security guard","security manager","cctv","loss prevention",
      "risk manager","health and safety","hse officer","osh","fire safety"],
     ["security","safety","risk","surveillance","patrol","access control","emergency","incident"]),

    ("Media & Journalism",
     ["journalist","editor","reporter","broadcast","news anchor","content creator",
      "media manager","radio","television","producer","scriptwriter"],
     ["media","journalism","broadcast","news","editorial","publishing","press","interview"]),

    ("Non-Profit & Social Work",
     ["social worker","ngo","charity","programme coordinator","community development",
      "welfare officer","case manager","development officer","fundraiser","volunteer coordinator"],
     ["social","ngo","community","welfare","beneficiary","donor","impact","charity","development"]),
]

def infer_job_field(title: str, description: str) -> str:
    if not title and not description:
        return ""
    combined = ((title or "") + " " + (description or "")).lower()
    best_field, best_score = "", 0
    for label, high_keys, supporting in FIELD_KEYWORD_MAP:
        score = sum(3 for k in high_keys if k in combined)
        score += sum(1 for k in supporting if k in combined)
        if score > best_score:
            best_score, best_field = score, label
    if best_score >= 3:
        log.info(f"Inferred job field: {best_field} (score {best_score})")
        return best_field
    return ""

# =============================================================================
#  QUALIFICATION EXTRACTOR
# =============================================================================

QUALIFICATION_TIERS = [
    ("PhD / Doctorate",
     ["phd","ph.d","doctorate","doctoral","doctor of philosophy"]),
    ("Master's Degree",
     ["master","msc","m.sc","ma ","m.a ","mba","m.b.a","meng","m.eng","mphil",
      "postgraduate","post-graduate","post graduate"]),
    ("Bachelor's Degree",
     ["bachelor","bsc","b.sc","ba ","b.a ","beng","b.eng","bcom","b.com","bba",
      "llb","degree in","undergraduate degree","honours degree","hons"]),
    ("Higher National Diploma",
     ["hnd","hnc","higher national diploma","higher national certificate",
      "higher diploma","advanced diploma"]),
    ("Diploma",
     ["diploma","dip ","dip.","associate degree","foundation degree"]),
    ("Professional Certification",
     ["acca","cpa","cfa","cima","pmp","prince2","cissp","aws certified","comptia",
      "cisco","ccna","ccnp","shrm","cipd","chartered","certified public",
      "certified financial","certified project","professional certification",
      "professional certificate"]),
    ("A-Levels / HSC",
     ["a-level","a level","hsc","higher school certificate","ib diploma",
      "international baccalaureate","gce advanced"]),
    ("O-Levels / School Certificate",
     ["o-level","o level","igcse","gcse","school certificate","sc ","cpe",
      "certificate of primary"]),
    ("No Formal Qualification Required",
     ["no qualification","no degree","no formal","school leaver","entry level",
      "no experience required","training provided","will train"]),
]

def extract_qualification(text: str) -> str:
    if not text:
        return ""
    if re.search(r"nursery|primary years|ib pyp|aged between|boys and girls", text, re.I):
        return ""
    lower = text.lower()
    for label, keywords in QUALIFICATION_TIERS:
        if any(k in lower for k in keywords):
            log.info(f"Qualification matched: {label}")
            return label
    return ""

# =============================================================================
#  EXPERIENCE EXTRACTOR
# =============================================================================

NO_EXP_KEYWORDS = [
    "no experience","no prior experience","fresh graduate","freshers",
    "entry level","entry-level","0 years","zero experience",
    "training provided","will train","no experience required",
]
LESS_THAN_1_KEYWORDS = [
    "less than 1 year","under 1 year","6 months","less than a year",
    "some experience","minimal experience",
]

def years_to_band(n: int) -> str:
    if n <= 0:  return "No Experience Required"
    if n <= 2:  return "1 - 2 Years"
    if n <= 5:  return "3 - 5 Years"
    if n <= 10: return "6 - 10 Years"
    return "10+ Years"

def extract_experience(text: str) -> str:
    if not text:
        return ""
    if re.search(r"aged?\s+between|boys\s+and\s+girls|nursery|primary\s+years|IB\s+PYP", text, re.I):
        return ""
    lower = text.lower()
    if any(k in lower for k in NO_EXP_KEYWORDS):
        return "No Experience Required"
    if any(k in lower for k in LESS_THAN_1_KEYWORDS):
        return "Less than 1 Year"
    patterns = [
        r"(\d+)\s*[-–to]+\s*(\d+)\s*\+?\s*years?",
        r"(\d+)\s*\+\s*years?\s*(?:of\s+)?(?:experience)?",
        r"(?:minimum|at\s+least|over|more\s+than)\s+(\d+)\s*\+?\s*years?",
        r"(\d+)\s*years?\s*(?:of\s+)?(?:relevant\s+)?(?:work\s+)?experience",
    ]
    for pattern in patterns:
        m = re.search(pattern, text, re.I)
        if m:
            raw = int(m.group(1))
            if raw > 20:
                continue
            band = years_to_band(raw)
            log.info(f"Experience extracted: {raw} yrs → {band}")
            return band
    return ""

# =============================================================================
#  LINKEDIN APPLY URL DECODER
# =============================================================================

def decode_linkedin_apply_url(raw: str) -> str:
    if not raw:
        return ""
    raw = decode_html_entities(raw)
    if raw.startswith("http") and "linkedin.com" not in raw:
        return raw

    m = re.search(r"[?&]url=([^&]+)", raw)
    if m:
        try:
            decoded = unquote(m.group(1))
            if "%" in decoded:
                decoded = unquote(decoded)
            if decoded.startswith("http") and "linkedin.com" not in decoded:
                return decoded
        except Exception:
            pass

    b64_m = re.search(r"[?&]offsiteApplyUrl=([^&]+)", raw)
    if b64_m:
        try:
            b64     = unquote(b64_m.group(1))
            decoded2 = base64.b64decode(b64).decode("utf-8")
            parsed  = json.loads(decoded2)
            if parsed and "url" in parsed:
                return parsed["url"]
        except Exception:
            pass

    return ""

# =============================================================================
#  FOLLOW LINKEDIN APPLY BUTTON
# =============================================================================

def follow_linkedin_apply_button(soup: BeautifulSoup, job_url: str) -> str:
    selectors = [
        {"attrs": {"data-tracking-control-name": re.compile(r"apply.link.offsite", re.I)}},
    ]
    raw = ""

    # Try common apply link patterns
    for tag in soup.find_all("a", href=True):
        href = tag.get("href", "")
        control = tag.get("data-tracking-control-name", "")
        if "offsite" in control.lower() or "apply" in control.lower():
            raw = href
            break

    if not raw:
        for tag in soup.find_all("a", href=True):
            href = tag.get("href", "")
            text = tag.get_text().lower()
            if ("apply" in text or "/apply" in href) and "linkedin" not in href:
                raw = href
                break

    if not raw:
        return ""

    resolved = decode_linkedin_apply_url(raw)
    if resolved and not is_bad_url(resolved):
        return resolved
    if raw.startswith("http") and not is_bad_url(raw):
        return raw
    return ""

# =============================================================================
#  PAGE EMAIL SCANNER
# =============================================================================

def scan_page_for_email(soup: BeautifulSoup, raw_html: str = "") -> str:
    for tag in soup.find_all("a", href=re.compile(r"^mailto:", re.I)):
        em = clean_email(tag.get("href", ""))
        if not em:
            continue
        if any(d in em for d in NOISE_EMAIL_DOMAINS):
            continue
        parts = em.split("@")
        if len(parts) == 2 and not FAKE_LOCAL_RE.match(parts[0]) and not FAKE_DOMAIN_RE.match(parts[1]):
            return em

    section_text = ""
    for sel in ["footer","#footer",".footer","#contact",".contact"]:
        for tag in soup.select(sel):
            section_text += " " + tag.get_text()
    found = extract_email_from_text(section_text)
    if found:
        return found

    body_email = extract_email_from_text(soup.get_text())
    if body_email:
        return body_email

    if raw_html:
        obfuscated = re.findall(
            r"[a-zA-Z0-9._%+\-]+\s*[\[\(]?\s*at\s*[\]\)]?\s*[a-zA-Z0-9.\-]+"
            r"\s*[\[\(]?\s*dot\s*[\]\)]?\s*[a-zA-Z]{2,}",
            raw_html, re.I
        )
        if obfuscated:
            norm = re.sub(r"\s*[\[\(]?\s*at\s*[\]\)]?\s*", "@", obfuscated[0], flags=re.I)
            norm = re.sub(r"\s*[\[\(]?\s*dot\s*[\]\)]?\s*", ".", norm, flags=re.I)
            norm = re.sub(r"\s+", "", norm).lower()
            if any(d in norm for d in NOISE_EMAIL_DOMAINS):
                pass
            elif "@" in norm and not FAKE_LOCAL_RE.match(norm.split("@")[0]):
                return norm
        html_email = extract_email_from_text(raw_html)
        if html_email:
            return html_email
    return ""

# =============================================================================
#  COMPANY WEBSITE CRAWLER
# =============================================================================

def crawl_company_website(website_url: str, job_title: str) -> dict:
    log.info(f"Crawling company site: {website_url}")
    if should_skip_crawl(website_url):
        return {"url": website_url, "email": "", "method": "fallback_website"}

    deadline = time.time() + 12
    root_url = website_url.rstrip("/")

    def get(url):
        if time.time() > deadline:
            return None
        try:
            time.sleep(0.5)
            r = requests.get(url, headers=HEADERS, timeout=10, allow_redirects=True)
            if r.status_code != 200:
                return None
            text = r.text
            if len(text) > FETCH_CHAR_LIMIT:
                text = text[:FETCH_CHAR_LIMIT]
            return text
        except Exception:
            return None

    home_html = get(root_url)
    if not home_html:
        return {"url": "", "email": "", "method": ""}

    soup_h = BeautifulSoup(home_html, "html.parser")
    home_email = scan_page_for_email(soup_h, home_html)
    if home_email:
        return {"url": "", "email": home_email, "method": "s7_homepage_email"}

    careers_url = contact_url = ""
    for tag in soup_h.find_all("a", href=True):
        href = make_absolute(tag.get("href", ""), root_url)
        link_text = tag.get_text().lower()
        if not href or is_bad_url(href) or href == root_url:
            continue
        if root_url not in href:
            continue
        if not careers_url and is_career_url(href):
            careers_url = href
        if not contact_url and (is_contact_url(href) or "contact" in link_text):
            contact_url = href
        if careers_url and contact_url:
            break

    if careers_url and time.time() < deadline:
        careers_html = get(careers_url)
        if careers_html:
            soup_c = BeautifulSoup(careers_html, "html.parser")
            email = scan_page_for_email(soup_c, careers_html)
            if email:
                return {"url": "", "email": email, "method": "s7_careers_email"}

    if contact_url and time.time() < deadline:
        contact_html = get(contact_url)
        if contact_html:
            soup_ct = BeautifulSoup(contact_html, "html.parser")
            email = scan_page_for_email(soup_ct, contact_html)
            if email:
                return {"url": "", "email": email, "method": "s7_contact_email"}

    if careers_url:
        return {"url": careers_url, "email": "", "method": "s7_careers_page"}
    return {"url": root_url, "email": "", "method": "fallback_website"}

# =============================================================================
#  APPLICATION DETAILS EXTRACTOR
# =============================================================================

def extract_application_details(job_url: str, soup: BeautifulSoup, company_website: str) -> dict:
    desc_text = ""
    for sel in [".show-more-less-html__markup", ".description__text"]:
        el = soup.select_one(sel)
        if el:
            desc_text = el.get_text()
            break

    job_title = ""
    el = soup.select_one(".top-card-layout__title")
    if el:
        job_title = el.get_text(strip=True)

    apply_btn = follow_linkedin_apply_button(soup, job_url)
    if apply_btn:
        log.info(f"S0 apply button: {apply_btn}")
        return {"url": apply_btn, "email": "", "method": "s0_apply_button"}

    # Script tag search
    apply_from_script = ""
    for script in soup.find_all("script"):
        if apply_from_script:
            break
        txt = script.string or ""
        for pattern in [
            r'"applyStartUrl"\s*:\s*"([^"]+)"',
            r'"applicationUrl"\s*:\s*"([^"]+)"',
        ]:
            m = re.search(pattern, txt)
            if m:
                candidate = decode_html_entities(m.group(1)).replace("\\", "")
                if candidate.startswith("http") and not is_bad_url(candidate):
                    apply_from_script = candidate
                    break
    if apply_from_script:
        return {"url": apply_from_script, "email": "", "method": "s1b_script_tag"}

    # Description links
    desc_el = soup.select_one(".show-more-less-html__markup") or soup.select_one(".description__text")
    if desc_el:
        for a in desc_el.find_all("a", href=True):
            h = a.get("href", "")
            if not is_bad_url(h):
                return {"url": h, "email": "", "method": "s3_desc_link"}

    url_matches = re.findall(r"https?://[^\s\"'<>)(,\]]+", desc_text)
    for u in url_matches:
        u = re.sub(r"[.,;:!?)]+$", "", u)
        if not is_bad_url(u):
            return {"url": u, "email": "", "method": "s4_desc_url"}

    email_in_desc = extract_email_from_text(desc_text)
    if email_in_desc:
        return {"url": "", "email": email_in_desc, "method": "s5_desc_email"}

    resolved_website = decode_linkedin_apply_url(company_website) or company_website
    if resolved_website and not is_bad_url(resolved_website):
        if should_skip_crawl(resolved_website):
            return {"url": resolved_website, "email": "", "method": "fallback_website"}
        result = crawl_company_website(resolved_website, job_title)
        if result.get("email") or result.get("url"):
            return result
        return {"url": resolved_website, "email": "", "method": "fallback_website"}

    return {"url": "", "email": "", "method": "not_found"}

# =============================================================================
#  JOB CRITERIA
# =============================================================================

def get_job_criteria(soup: BeautifulSoup, label: str) -> str:
    lower_label = label.lower()
    for li in soup.select(".description__job-criteria-list > li"):
        h3 = li.find("h3")
        if h3 and lower_label in h3.get_text().strip().lower():
            spans = li.select(".description__job-criteria-text, span")
            if spans:
                return spans[-1].get_text(strip=True)
    for chip in soup.select(".job-details-jobs-unified-top-card__job-insight, .jobs-unified-top-card__job-insight"):
        text = chip.get_text(strip=True).lower()
        if "employment" in lower_label or "type" in lower_label:
            if re.search(r"full[\-\s]?time|part[\-\s]?time|contract|temporary|internship|freelance", text, re.I):
                return chip.get_text(strip=True)
        elif "seniority" in lower_label:
            if re.search(r"entry|associate|mid[\-\s]?senior|senior|director|executive|intern", text, re.I):
                return chip.get_text(strip=True)
    meta_map = {
        "employment type": soup.find("meta", {"name": "employmentType"}),
        "seniority level": soup.find("meta", {"name": "seniorityLevel"}),
        "industries":      soup.find("meta", {"name": "industry"}),
    }
    meta_tag = meta_map.get(lower_label)
    if meta_tag:
        return meta_tag.get("content", "")
    return ""

def get_workplace_type(soup: BeautifulSoup) -> str:
    for sel in [
        ".topcard__workplace-type",
        ".job-details-jobs-unified-top-card__workplace-type",
        ".jobs-unified-top-card__workplace-type",
    ]:
        el = soup.select_one(sel)
        if el:
            return el.get_text(strip=True)
    for chip in soup.select(".job-details-jobs-unified-top-card__job-insight, .jobs-unified-top-card__job-insight"):
        t = chip.get_text(strip=True)
        if re.match(r"^(remote|on[\-\s]?site|hybrid)$", t, re.I):
            return t
    return ""

# =============================================================================
#  WORDPRESS LOGO UPLOAD
# =============================================================================

def upload_logo_to_wordpress(logo_url: str, company_name: str) -> str:
    if not logo_url or not logo_url.startswith("http") or not WP_USER:
        return ""
    try:
        r = requests.get(logo_url, headers={
            "User-Agent": HEADERS["User-Agent"],
            "Referer": "https://www.linkedin.com/",
        }, timeout=15)
        if r.status_code != 200:
            log.warning(f"Logo download failed ({r.status_code}): {logo_url}")
            return ""
        content_type = r.headers.get("Content-Type", "image/jpeg")
        ext = "png" if "png" in content_type else "jpg"
        file_name = re.sub(r"[^a-z0-9]", "-", company_name.lower()) + "-logo." + ext
        credentials = base64.b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()
        wp_r = requests.post(
            WP_URL + "media",
            headers={
                "Authorization": "Basic " + credentials,
                "Content-Disposition": f"attachment; filename={file_name}",
                "Content-Type": content_type,
            },
            data=r.content,
            timeout=20,
        )
        if wp_r.status_code in (200, 201):
            data = wp_r.json()
            hosted_url = data.get("source_url", "")
            log.info(f"Logo uploaded to WP: {hosted_url}")
            return hosted_url
        log.warning(f"WP media upload failed ({wp_r.status_code}): {wp_r.text[:200]}")
    except Exception as e:
        log.warning(f"uploadLogoToWordPress error: {e}")
    return ""

# =============================================================================
#  COMPANY DETAILS SCRAPER
# =============================================================================

def get_company_detail(soup: BeautifulSoup, label: str) -> str:
    lower = label.lower()
    for div in soup.select("section.core-section-container dl > div"):
        dt = div.find("dt")
        if dt and lower in dt.get_text().strip().lower():
            dd = div.find("dd")
            if dd:
                return dd.get_text(strip=True)
    return ""

def scrape_company_details(company_url: str) -> dict:
    if not company_url:
        return {}
    log.info(f"Scraping company: {company_url}")
    try:
        resp = requests.get(company_url, headers=HEADERS, timeout=15)
        if resp.status_code != 200:
            return {}
        html = resp.text
        soup = BeautifulSoup(html, "html.parser")

        logo_m = re.search(
            r'<meta[^>]+property=["\']og:image["\'][^>]+content=["\']([^"\']+)["\']', html, re.I
        ) or re.search(
            r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']og:image["\']', html, re.I
        )
        raw_logo = clean_logo_url(logo_m.group(1)) if logo_m else ""

        website_tag = soup.select_one("a[data-tracking-control-name='about_website']")
        raw_website = website_tag.get("href", "") if website_tag else get_company_detail(soup, "Website")
        website = decode_linkedin_apply_url(raw_website) or raw_website

        name = soup.find("h1")
        name = name.get_text(strip=True) if name else ""

        hosted_logo = upload_logo_to_wordpress(raw_logo, name) if raw_logo else ""

        return {
            "name":         name,
            "industry":     get_company_detail(soup, "Industry"),
            "size":         get_company_detail(soup, "Company size"),
            "headquarters": get_company_detail(soup, "Headquarters"),
            "type":         get_company_detail(soup, "Type"),
            "founded":      get_company_detail(soup, "Founded"),
            "specialties":  get_company_detail(soup, "Specialties"),
            "website":      website,
            "logo":         hosted_logo or raw_logo,
            "about":        (soup.select_one("section.about-us p") or
                             soup.select_one(".core-section-container__content p") or
                             type("_", (), {"get_text": lambda self, **k: ""})()
                            ).get_text(strip=True),
        }
    except Exception as e:
        log.warning(f"Company scrape failed: {e}")
        return {}

# =============================================================================
#  JOB DETAIL SCRAPER
# =============================================================================

def scrape_job_details(job_url: str) -> dict | None:
    log.info(f"Scraping job: {job_url}")
    try:
        resp = requests.get(job_url, headers=HEADERS, timeout=15)
        if resp.status_code != 200:
            return None
        html = resp.text
        soup = BeautifulSoup(html, "html.parser")
    except Exception as e:
        log.warning(f"Job fetch failed: {e}")
        return None

    def sel_text(*selectors):
        for s in selectors:
            el = soup.select_one(s)
            if el:
                t = el.get_text(strip=True)
                if t:
                    return t
        return ""

    title         = sel_text(".top-card-layout__title","h1.topcard__title",
                              ".job-details-jobs-unified-top-card__job-title","h1")
    company_name  = sel_text(".topcard__org-name-link",
                              ".job-details-jobs-unified-top-card__company-name",".topcard__flavor")
    company_url_el = (soup.select_one(".topcard__org-name-link") or
                      soup.select_one(".job-details-jobs-unified-top-card__company-name a"))
    company_url   = company_url_el.get("href", "") if company_url_el else ""
    location      = sel_text(".topcard__flavor--bullet",
                              ".job-details-jobs-unified-top-card__bullet")
    workplace_type = get_workplace_type(soup)

    time_el = soup.find("time")
    raw_posted = (time_el.get("datetime", "") if time_el else "") or \
                 sel_text(".posted-time-ago__text",
                          ".job-details-jobs-unified-top-card__posted-date")
    posted_date = resolve_posted_date(raw_posted)

    applicants = sel_text(".num-applicants__caption", ".jobs-unified-top-card__applicant-count")

    raw_desc = sel_text(".show-more-less-html__markup", ".description__text")
    description = clean_description(raw_desc)

    salary = ""
    for sel in [".compensation__salary",".salary","[class*='salary']","[class*='compensation']"]:
        el = soup.select_one(sel)
        if el:
            salary = el.get_text(strip=True)
            break
    if not salary:
        for chip in soup.select(".job-details-jobs-unified-top-card__job-insight"):
            t = chip.get_text(strip=True)
            if re.search(r"\$|MUR|Rs\.?|salary|/yr|/hour|per month", t, re.I):
                salary = t
                break

    raw_job_type  = get_job_criteria(soup, "Employment type") or workplace_type
    job_type      = raw_job_type or "Full-time"
    seniority     = get_job_criteria(soup, "Seniority level")

    linkedin_function = get_job_criteria(soup, "Job function")
    linkedin_industry = get_job_criteria(soup, "Industries")
    job_field     = linkedin_function or linkedin_industry or infer_job_field(title, description)
    industry      = linkedin_industry or get_job_criteria(soup, "Industries")

    real_deadline      = parse_deadline(soup)
    estimated_deadline = estimate_deadline_from_posted(posted_date) if not real_deadline else ""
    effective_deadline = real_deadline or estimated_deadline

    time.sleep(0.2)
    company = scrape_company_details(company_url)

    time.sleep(0.2)
    apply_data = extract_application_details(job_url, soup, company.get("website", ""))

    raw_apply = ""
    if apply_data.get("email"):
        raw_apply = clean_email(apply_data["email"])
    elif apply_data.get("url") and apply_data.get("method") != "not_found":
        raw_apply = apply_data["url"]
    apply_link = clean_application_link(raw_apply)

    qualifications = extract_qualification(description)
    experience     = extract_experience(description)

    return {
        "jobTitle":          title,
        "jobType":           job_type,
        "jobQualifications": qualifications,
        "jobExperience":     experience,
        "jobLocation":       location,
        "jobField":          job_field,
        "datePosted":        posted_date,
        "deadline":          effective_deadline,
        "jobDescription":    description,
        "application":       apply_link,
        "companyUrl":        company_url,
        "companyName":       company.get("name") or company_name,
        "companyLogo":       clean_logo_url(company.get("logo", "")),
        "companyIndustry":   company.get("industry") or industry,
        "companyFounded":    company.get("founded", ""),
        "companyType":       company.get("type", ""),
        "companyWebsite":    company.get("website", ""),
        "companyAddress":    company.get("headquarters", ""),
        "companyDetails":    company.get("about", ""),
        "jobUrl":            job_url,
        "estimatedDeadline": estimated_deadline,
        "salaryRange":       salary,
    }

# =============================================================================
#  MAIN CRAWL
# =============================================================================

def craw():
    start_time = time.time()
    log.info(f"SCRAPE STARTED: {datetime.now()}")

    result   = []
    seen_urls = set()

    for i in range(PAGES):
        list_url = SEARCH_BASE + str(i * 25)
        log.info(f"Fetching list page {i+1}: {list_url}")
        try:
            resp = requests.get(list_url, headers=HEADERS, timeout=15)
            soup = BeautifulSoup(resp.text, "html.parser")
            found = 0
            for a in soup.select("#main-content > section > ul > li > div > a, a.base-card__full-link"):
                href = a.get("href", "")
                if not href:
                    continue
                key = href.split("?")[0]
                if key not in seen_urls:
                    seen_urls.add(key)
                    result.append(href)
                    found += 1
            log.info(f"Page {i+1}: found {found} new URLs")
        except Exception as e:
            log.warning(f"List page error: {e}")
        time.sleep(DELAY_S)

    log.info(f"Total unique job URLs: {len(result)}")
    result = result[:JOB_LIMIT]
    log.info(f"Capped to: {len(result)} jobs")

    jobs, errors = [], 0
    for j, url in enumerate(result):
        print(f"\n{C_HEADER(f'>>> Scraping job {j+1}/{len(result)} ...')}")
        log.info(f"URL: {url}")
        try:
            job = scrape_job_details(url)
            if job and job.get("jobTitle"):
                jobs.append(job)
                print_job_verbose(job, j + 1, len(result))
            else:
                print(C_RED(f"  ✗  No title found — skipped"))
        except Exception as e:
            errors += 1
            print(C_RED(f"  ✗  ERROR: {e}"))
            log.warning(f"Job error: {e}")
        time.sleep(DELAY_S)

    # ── Final summary banner ──────────────────────────────────────────────────
    mins = round((time.time() - start_time) / 60, 1)
    print()
    print(C_HEADER("=" * 72))
    print(C_HEADER("  SCRAPE COMPLETE"))
    print(C_HEADER("=" * 72))
    print(f"  {C_LABEL('Total scraped')}  : {C_GREEN(str(len(jobs)))} jobs")
    print(f"  {C_LABEL('Errors')}         : {C_RED(str(errors)) if errors else '0'}")
    print(f"  {C_LABEL('Duration')}       : ~{mins} min")
    print(f"  {C_LABEL('Output file')}    : {OUTPUT_FILE}")

    # ── Field breakdown ───────────────────────────────────────────────────────
    if jobs:
        from collections import Counter
        fields = Counter(j.get("jobField") or "Unknown" for j in jobs)
        print(f"\n  {C_LABEL('Jobs by field:')}")
        for field, count in fields.most_common():
            bar = "█" * count
            print(f"    {field:<35} {C_GREEN(bar)} {count}")

        # ── Apply method breakdown ────────────────────────────────────────────
        with_apply = sum(1 for j in jobs if j.get("application"))
        with_email = sum(1 for j in jobs if "@" in (j.get("application") or ""))
        with_url   = with_apply - with_email
        no_apply   = len(jobs) - with_apply
        print(f"\n  {C_LABEL('Application links:')}")
        print(f"    URL found    : {with_url}")
        print(f"    Email found  : {with_email}")
        print(f"    Not found    : {no_apply}")

    print(C_HEADER("=" * 72))
    log.info(f"Scraped {len(jobs)} jobs, {errors} errors")

    # ── Write to Excel ────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    headers = [
        "Job Title", "Job Type", "Job Qualifications", "Job Experience",
        "Job Location", "Job Field", "Date Posted", "Deadline",
        "Job Description", "Application", "Company URL", "Company Name",
        "Company Logo", "Company Industry", "Company Founded", "Company Type",
        "Company Website", "Company Address", "Company Details", "Job URL",
        "Estimated Deadline", "Salary Range",
    ]
    ws.append(headers)

    for job in jobs:
        ws.append([
            job["jobTitle"], job["jobType"], job["jobQualifications"], job["jobExperience"],
            job["jobLocation"], job["jobField"], job["datePosted"], job["deadline"],
            job["jobDescription"], job["application"], job["companyUrl"], job["companyName"],
            job["companyLogo"], job["companyIndustry"], job["companyFounded"], job["companyType"],
            job["companyWebsite"], job["companyAddress"], job["companyDetails"], job["jobUrl"],
            job["estimatedDeadline"], job["salaryRange"],
        ])

    wb.save(OUTPUT_FILE)
    log.info(f"Written {len(jobs)} rows to {OUTPUT_FILE}")
    log.info(f"DONE in ~{mins} min")


if __name__ == "__main__":
    craw()
