"""
LinkedIn Job Scraper — v5 (Merged)
===================================
Built on v4, now incorporates the full website-crawling engine from the
Saudi Pipeline v5:

  NEW / UPGRADED in this version:
  ─ ATS detection (Greenhouse, Lever, Workday, Ashby, SmartRecruiters,
    BambooHR, iCIMS, Taleo, SuccessFactors, Recruitee, Workable,
    Teamtailor, PageUp, Wynt, EasyJobs, Zoho, BambooHR, Breezy …)
  ─ Multi-strategy career page finder:
      subdomain probe → path probe → link scan → sitemap → DuckDuckGo
  ─ ATS-native job listing APIs (Greenhouse JSON, Lever JSON, Ashby JSON,
      SmartRecruiters API, Workday JSON)
  ─ SuccessFactors table scraper with pagination
  ─ Deep job-detail page extractor:
      JSON-LD · headed-section parser · bold-label extractor ·
      salary/date/deadline regex · logo · apply-button detection
  ─ Playwright (headless Chromium) for JS-heavy pages, with plain-requests
    fallback first (faster, polite)
  ─ Saudi filter preserved (skips non-Saudi domains / Wikipedia noise)
  ─ Standardisation: field · experience band · qualification tier
  ─ Deduplication by (title, company, location) fingerprint
  ─ Checkpoint / resume support  →  pipeline_checkpoint.json
  ─ Saves:  jobs_output.xlsx  +  saudi_jobs.csv  +  saudi_companies_found.csv
  ─ All existing v4 LinkedIn guest-API pagination preserved as the
    URL-collection layer; website crawling is the detail/apply enrichment layer

Requirements:
    pip install requests beautifulsoup4 openpyxl playwright pandas nest_asyncio tqdm
    playwright install chromium

Usage:
    python linkedin_job_scraper_v5.py
"""

# ─────────────────────────────────────────────────────────────────────────────
#  BOOTSTRAP  (install deps if running inside a notebook / fresh env)
# ─────────────────────────────────────────────────────────────────────────────
import subprocess, sys, os

_SILENT = os.environ.get("LJS_NO_INSTALL")
if not _SILENT:
    try:
        subprocess.run(
            ["apt-get", "install", "-y",
             "libatk1.0-0", "libatk-bridge2.0-0", "libcups2", "libdrm2",
             "libxkbcommon0", "libxcomposite1", "libxdamage1", "libxfixes3",
             "libxrandr2", "libgbm1", "libasound2"],
            capture_output=True, check=False,
        )
        subprocess.run(
            [sys.executable, "-m", "pip", "install",
             "playwright", "pandas", "beautifulsoup4",
             "requests", "nest_asyncio", "tqdm", "openpyxl", "-q"],
            check=True,
        )
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            capture_output=True, check=False,
        )
    except Exception as _e:
        print(f"[bootstrap] warn: {_e}")

import importlib, site as _site
importlib.invalidate_caches()
for _sp in _site.getsitepackages():
    if _sp not in sys.path:
        sys.path.insert(0, _sp)

# ─────────────────────────────────────────────────────────────────────────────
#  STANDARD IMPORTS
# ─────────────────────────────────────────────────────────────────────────────
import asyncio, re, json, random, time, base64, logging
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urljoin, urlparse, unquote, quote_plus

import requests
import pandas as pd
import nest_asyncio
from bs4 import BeautifulSoup
import openpyxl

try:
    from playwright.async_api import async_playwright
    _PLAYWRIGHT_OK = True
except ImportError:
    _PLAYWRIGHT_OK = False
    print("[warn] playwright not available — JS rendering disabled")

try:
    from tqdm import tqdm
except ImportError:
    def tqdm(it, **_): return it  # type: ignore

nest_asyncio.apply()

# ═════════════════════════════════════════════════════════════════════════════
#  CONFIG
# ═════════════════════════════════════════════════════════════════════════════

SHEET_NAME       = "Sheet1"
DELAY_S          = 2.0
FETCH_CHAR_LIMIT = 120_000

MAX_PAGES        = 0   # 0 = unlimited LinkedIn pages per keyword
MAX_EMPTY_PAGES  = 5
JOB_LIMIT        = 0   # 0 = no cap on LinkedIn URLs collected

OUTPUT_XLSX      = "jobs_output.xlsx"
OUTPUT_CSV       = "saudi_jobs.csv"
COMPANIES_FILE   = "saudi_companies_found.csv"
CHECKPOINT_FILE  = "pipeline_checkpoint.json"

# WordPress (optional — leave blank to skip logo upload)
WP_URL      = "https://mauritius.mimusjobs.com/wp-json/wp/v2/"
WP_USER     = "calolina"
WP_PASSWORD = "st8a 6mWY wqgV 0syR mB3i y5FQ"

VERBOSE = True

# ─── LinkedIn search keywords ────────────────────────────────────────────────
SEARCH_KEYWORDS = [
    "",
    "engineer", "developer", "manager", "finance", "sales", "HR",
    "doctor", "construction", "logistics", "operations", "customer service",
    "teacher", "chef", "lawyer", "graphic designer", "production manager",
    "petroleum", "driver", "security", "researcher", "journalist",
    "banker", "retail", "renewable energy",
]

# ─── Logging / colour ────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

import sys as _sys
_USE_COLOUR = _sys.stdout.isatty()

def _c(code, text):  return f"\033[{code}m{text}\033[0m" if _USE_COLOUR else text
C_HEADER  = lambda t: _c("1;36",  t)
C_LABEL   = lambda t: _c("1;33",  t)
C_VALUE   = lambda t: _c("97",    t)
C_DIM     = lambda t: _c("2",     t)
C_GREEN   = lambda t: _c("1;32",  t)
C_RED     = lambda t: _c("1;31",  t)
C_BLUE    = lambda t: _c("1;34",  t)
C_DIVIDER = lambda: _c("2", "─" * 72)

# ═════════════════════════════════════════════════════════════════════════════
#  ROTATING USER-AGENTS
# ═════════════════════════════════════════════════════════════════════════════

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_4) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
]
_ua_idx = 0

def _next_headers() -> dict:
    global _ua_idx
    ua = USER_AGENTS[_ua_idx % len(USER_AGENTS)]
    _ua_idx += 1
    return {
        "User-Agent":       ua,
        "Accept":           "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language":  "en-US,en;q=0.9",
        "Accept-Encoding":  "gzip, deflate, br",
        "Cache-Control":    "no-cache",
        "X-Li-Lang":        "en_US",
        "X-Requested-With": "XMLHttpRequest",
    }

HEADERS = _next_headers()

# ═════════════════════════════════════════════════════════════════════════════
#  DOMAIN / URL LISTS
# ═════════════════════════════════════════════════════════════════════════════

SKIP_CRAWL_DOMAINS = [
    "dhl.com","fedex.com","ups.com","amazon.com","amazon.jobs",
    "google.com","microsoft.com","apple.com","meta.com","ibm.com",
    "oracle.com","sap.com","accenture.com","deloitte.com","pwc.com",
    "kpmg.com","ey.com","mckinsey.com","bcg.com","bain.com",
    "citibank.com","hsbc.com","barclays.com","bnpparibas.com",
    "airbus.com","boeing.com","siemens.com","ge.com",
    "unilever.com","nestle.com","pg.com","shell.com","bp.com",
]
BAD_DOMAINS = [
    "linkedin.com","google.com","youtube.com","facebook.com",
    "twitter.com","x.com","instagram.com","t.co","example.com",
    "w3.org","sentry.io","schema.org",
]
NOISE_EMAIL_DOMAINS = [
    "example.com","sentry.io","google.com","w3.org",
    "schema.org","wixpress.com","squarespace.com",
]
NON_SAUDI_DOMAINS = {
    "fortune.com","forbes.com","bloomberg.com","reuters.com",
    "arabnews.com","aljazeera.com","bbc.com","bbc.co.uk",
    "cnn.com","nytimes.com","wsj.com","ft.com",
    "federalreserve.gov","sec.gov","state.gov","treasury.gov",
    "londonstockexchange.com","lseg.com","nyse.com","nasdaq.com",
    "imf.org","worldbank.org","opec.org","un.org","unesco.org",
    "pepsico.com","linkedin.com","facebook.com","twitter.com","x.com",
    "instagram.com","tiktok.com","snapchat.com","amazon.com",
    "google.com","microsoft.com","apple.com","hsbc.com","barclays.com",
    "houseofsaud.com",
}
BLOCKED_CAREER_DOMAINS = {"linkedin.com", "www.linkedin.com"}
SKIP_DOMAINS = {
    "google","facebook","twitter","linkedin","wikipedia","youtube",
    "instagram","tiktok","snapchat","amazon","duckduck","bing",
    "yahoo","reddit","quora","trustpilot","glassdoor","indeed",
    "zawya","arabnews","saudigazette","bloomberg","reuters",
}

FAKE_LOCAL_RE  = re.compile(
    r"^(name|user|email|mail|yourname|your[-_.]?email|sample|test|info|hello"
    r"|noreply|no[-_.]?reply|admin|webmaster|support|contact|example)$", re.I)
FAKE_DOMAIN_RE = re.compile(
    r"^(domain|example|yoursite|yourdomain|yourbrand|company|mycompany"
    r"|website|yourcompany|mysite|placeholder|site)\.[a-z]{2,}$", re.I)

MONTH_MAP = {
    "jan":0,"feb":1,"mar":2,"apr":3,"may":4,"jun":5,
    "jul":6,"aug":7,"sep":8,"oct":9,"nov":10,"dec":11,
}

# ─── ATS detection tables ────────────────────────────────────────────────────
ATS_DOMAINS = {
    "greenhouse.io": "Greenhouse",
    "lever.co": "Lever",
    "myworkdayjobs.com": "Workday",
    "ashbyhq.com": "Ashby",
    "smartrecruiters.com": "SmartRecruiters",
    "bamboohr.com": "BambooHR",
    "icims.com": "iCIMS",
    "taleo.net": "Taleo",
    "successfactors.com": "SuccessFactors",
    "sapsf.com": "SuccessFactors",
    "oraclecloud.com": "Oracle",
    "recruitee.com": "Recruitee",
    "workable.com": "Workable",
    "jobvite.com": "Jobvite",
    "breezy.hr": "Breezy",
    "zohorecruit.com": "Zoho",
    "zoho.com/recruit": "Zoho",
    "bayt.com": "Bayt",
    "gulftalent.com": "GulfTalent",
    "teamtailor.com": "Teamtailor",
    "comeet.com": "Comeet",
    "apply.workable.com": "Workable",
    "jobs.lever.co": "Lever",
    "boards.eu.greenhouse.io": "Greenhouse",
    "rmkcdn.successfactors.com": "SuccessFactors",
    "successfactors": "SuccessFactors",
    "ats.sa": "ATS.sa",
    "pageuppeople.com": "PageUp",
    "jobsoid.com": "Jobsoid",
    "freshteam.com": "Freshteam",
    "cornerstone": "Cornerstone",
    "talentsoft.com": "TalentSoft",
    "hibob.com": "HiBob",
    "rippling.com": "Rippling",
    "personio.com": "Personio",
    "easy.jobs": "EasyJobs",
    "apply.wynt.ai": "Wynt",
    "wynt.ai": "Wynt",
    "pinpointhq.com": "Pinpoint",
    "peoplehr.net": "PeopleHR",
}
ATS_HTML_FINGERPRINTS = {
    "rmkcdn.successfactors.com": "SuccessFactors",
    "talentcommunity/apply": "SuccessFactors",
    "/go/Job-Search/": "SuccessFactors",
    "greenhouse-io": "Greenhouse",
    "myworkdayjobs": "Workday",
    "lever.co/v0/postings": "Lever",
    "ashbyhq.com": "Ashby",
    "icims.com": "iCIMS",
    "bamboohr.com": "BambooHR",
    "smartrecruiters.com": "SmartRecruiters",
    "teamtailor.com": "Teamtailor",
    "recruitee.com": "Recruitee",
    "pageuppeople.com": "PageUp",
    "jobvite.com": "Jobvite",
    "taleo.net": "Taleo",
    "wynt.ai": "Wynt",
    "easy.jobs": "EasyJobs",
}

# ─── Career-page discovery ───────────────────────────────────────────────────
CAREER_SUBDOMAINS = [
    "careers","jobs","career","job","work","hiring",
    "apply","talent","recruitment","hr","people",
    "vacancies","opportunities","join",
]
CAREER_PATHS = [
    "/careers","/jobs","/career","/job","/careers/","/jobs/",
    "/en/careers","/en/jobs","/ar/careers","/about/careers",
    "/about/jobs","/company/careers","/join-us","/join",
    "/work-with-us","/openings","/vacancies","/opportunities",
    "/employment","/hiring","/recruitment","/apply",
]
JOB_LISTING_SUFFIXES = [
    "/go/Job-Search/","/go/All-Jobs/",
    "/job-search-results","/job-search-results/",
    "/en/job-search-results","/search","/search/",
    "/job-search","/job-search/","/openings","/openings/",
    "/current-openings","/positions","/positions/",
    "/open-positions","/listings","/all-jobs","/all-jobs/",
    "/vacancies","/vacancies/","/opportunities","/jobs","/jobs/",
    "/apply","/join","/en/jobs","/en/careers","/en/search",
    "/en/openings","/en/positions","/ar/jobs","/ar/careers",
]
SF_LISTING_SUFFIXES = [
    "/go/Job-Search/","/go/All-Jobs/",
    "/job-search-results","/job-search-results/",
    "/en/job-search-results","/search","/search/",
    "/jobs","/jobs/","/openings","/openings/","/all-jobs","/all-jobs/",
]
SF_SEARCH_PARAMS = [
    "/?createNewAlert=false&q=&locationsearch=",
    "/search/?createNewAlert=false&q=",
    "/search/?q=&locationsearch=",
    "/?q=","search?q=",
]

CAREER_KEYWORDS = [
    "career","careers","job","jobs","vacancy","vacancies",
    "hiring","work with us","join us","join our team",
    "employment","opportunities","opening","openings",
    "وظائف","التوظيف","انضم إلينا","فرص عمل",
]
DEFINITE_PATTERNS = [
    r"\bcareers?\b",r"\bjobs?\b",r"join\s+us",r"work\s+with\s+us",
    r"we'?re?\s+hiring",r"open\s+positions?",r"current\s+openings?",
    r"وظائف",r"التوظيف",
]
JOB_PAGE_SIGNALS = [
    "apply now","apply for","job description","requirements",
    "qualifications","responsibilities","we are looking",
    "we are hiring","open position","full-time","part-time",
    "remote","salary","benefits","وظيفة","تقديم","المتطلبات",
    "job opportunities","job search","talentcommunity","create alert",
    "sort by title","sort by location",
]

LOCATION_PATTERN = re.compile(
    r"(Riyadh|Jeddah|Dammam|Khobar|Mecca|Medina|Saudi Arabia|KSA|Remote|"
    r"Dhahran|Jubail|Yanbu|Taif|Abha|Buraidah|Hail|Tabuk|"
    r"الرياض|جدة|الدمام|مكة|المدينة|Central Province|Eastern Province|"
    r"Western Province|Makkah|Madinah)", re.I,
)

HARD_BLOCKED_PATTERNS = re.compile(
    r"(/legal/|/user.agreement|/privacy|/terms|/cookie"
    r"|/uas/request.password|/forgot.password|/sign.in|/login"
    r"|/register|/signup|/sign.up|/new.cv|/myworkspace|/dashboard"
    r"|/comments/feed|\.svg$|\.png$|\.jpg$|\.jpeg$|\.gif$|\.ico$"
    r"|/sitemap|/feed\.xml|/rss|javascript:|mailto:|tel:"
    r"|/cdn.cgi/access|objectstorage.*\.png)", re.I,
)

JOB_PATH_CORE       = re.compile(r"/jobs?/", re.I)
JOB_URL_ATS_SPECIFIC = re.compile(
    r"(/go/[A-Za-z0-9%-]+/\d+"
    r"|/\d{6,}/?$"
    r"|/job-detail/\d+"
    r"|/req/\d+"
    r"|/posting/[A-Za-z0-9-]+"
    r"|/application/\d+"
    r"|/apply/\d+)", re.I,
)

_GARBAGE_FIELD_RE = re.compile(
    r"^(job\s*description|key\s*accountabilities?|requirements?|qualifications?"
    r"|responsibilities|overview|role\s*purpose|job\s*purpose|about\s*the\s*role"
    r"|what\s*you|we\s*are\s*looking|education\s*&|experience\s*&|skills\s*&"
    r"|others?|not\s*applicable|n/?a|—|-)$", re.I,
)

# ─── Regex extractors ────────────────────────────────────────────────────────
SALARY_RE = re.compile(
    r"(?:SAR|SR|USD|\$|€|£)\s?[\d,]+(?:\s?[-–]\s?[\d,]+)?(?:\s?[Kk])?|"
    r"[\d,]+(?:\s?[-–]\s?[\d,]+)?\s?(?:SAR|SR|USD|per\s+month|/month|monthly)",
    re.I,
)
DATE_POSTED_RE = re.compile(
    r"(?:posted|date posted|published)[:\s]*([^\n<]{1,40})|"
    r"(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})|"
    r"(\d+\s+(?:day|hour|week|month)s?\s+ago)|"
    r"(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4})",
    re.I,
)
DEADLINE_RE = re.compile(
    r"(?:deadline|closing date|apply by|applications? close)[:\s]*([^\n<]{1,40})|"
    r"(?:expires?|expiry)[:\s]*([^\n<]{1,40})", re.I,
)
EXPERIENCE_RE = re.compile(
    r"(\d+\+?\s*(?:–|-|to)?\s*\d*\s*years?\s*(?:of\s+)?experience|"
    r"experience[:\s]*(\d+[\+\s\-]*\d*\s*years?)|"
    r"(?:minimum|min\.?)\s+\d+\s+years?)", re.I,
)

# ═════════════════════════════════════════════════════════════════════════════
#  SHARED RUNTIME STATE
# ═════════════════════════════════════════════════════════════════════════════

discovered_domains: set = set()
all_jobs:           list = []
company_results:    list = []
_JOB_COUNTER            = [0]

_stats = {
    "companies_seen": 0, "companies_with_careers": 0,
    "companies_no_careers": 0, "companies_skipped_non_saudi": 0,
    "jobs_total": 0, "jobs_with_title": 0, "jobs_with_salary": 0,
    "jobs_with_description": 0, "detail_fetches": 0,
    "detail_failures": 0, "ats_hits": {}, "strategy_hits": {},
}

# ═════════════════════════════════════════════════════════════════════════════
#  BASIC HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def _ts():  return datetime.now().strftime("%H:%M:%S")
def vlog(msg, indent=0):   print(f"[{_ts()}] {'   '*indent}{msg}", flush=True)
def vprint(msg, indent=0):
    if VERBOSE: vlog(msg, indent)

def clean(text, max_len=300):
    return re.sub(r"\s+", " ", str(text or "")).strip()[:max_len]

def get_domain(url):
    try:    return urlparse(str(url)).netloc.lower().replace("www.","")
    except: return ""

def get_base(website):
    p = urlparse(str(website))
    return p.netloc.lower().replace("www.",""), p.scheme or "https"

def should_skip_crawl(url: str) -> bool:
    if not url: return True
    return any(d in url.lower() for d in SKIP_CRAWL_DOMAINS)

def is_bad_url(url: str) -> bool:
    if not url or not url.startswith("http"): return True
    return any(d in url.lower() for d in BAD_DOMAINS)

def _is_blocked_career_url(url: str) -> bool:
    return get_domain(url) in BLOCKED_CAREER_DOMAINS

def make_absolute(href: str, root_url: str) -> str:
    if not href: return ""
    href = href.strip()
    if href.startswith("http"):   return href
    if href.startswith("//"): return "https:" + href
    if href.startswith("/"):  return root_url.rstrip("/") + href
    return ""

def decode_html_entities(s: str) -> str:
    if not s: return ""
    for old, new in [("&amp;","&"),("&lt;","<"),("&gt;",">"),("&quot;",'"'),
                     ("&#39;","'"),("\\u0026","&"),("\\u003D","="),
                     ("\\u003A",":"),("\\u002F","/")]:
        s = s.replace(old, new)
    return s

def canonicalise_job_url(url: str) -> str:
    if not url: return ""
    m = re.search(r"/jobs/view/(\d+)", url)
    if m: return f"https://www.linkedin.com/jobs/view/{m.group(1)}/"
    return re.sub(r"[?#].*$", "", url)

def _strip_li_tracking(url: str) -> str:
    if not url: return ""
    url = decode_html_entities(url)
    m = re.search(r"[?&]url=([^&]+)", url)
    if m:
        try:
            decoded = unquote(m.group(1))
            if "%" in decoded: decoded = unquote(decoded)
            if decoded.startswith("http") and "linkedin.com" not in decoded:
                return decoded
        except Exception:
            pass
    return url

def _title_similarity(a: str, b: str) -> float:
    if not a or not b: return 0.0
    def tokens(s):
        return set(re.sub(r"[^a-z0-9 ]", " ", s.lower()).split())
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb: return 0.0
    return len(ta & tb) / max(len(ta), len(tb))

def _sanitize_field(value, max_len=300):
    if not value: return ""
    v = re.sub(r"\s+", " ", str(value)).strip()
    if _GARBAGE_FIELD_RE.match(v): return ""
    if len(v) < 2: return ""
    return v[:max_len]

def parse_domain(url):
    try:
        p = urlparse(url if url.startswith("http") else "https://"+url)
        netloc = p.netloc.lower().replace("www.","")
        if not netloc or any(s in netloc for s in SKIP_DOMAINS): return None, None
        return netloc, f"https://www.{netloc}"
    except:
        return None, None

def is_ats(url):
    url = str(url).lower()
    if "linkedin.com" in url: return None
    for pat, name in ATS_DOMAINS.items():
        if pat in url: return name
    return None

def is_ats_from_html(html):
    lower = html.lower()
    for fp, name in ATS_HTML_FINGERPRINTS.items():
        if fp.lower() in lower: return name
    return None

def is_career_link(href, text):
    combined = (str(href) + " " + str(text)).lower()
    if "linkedin.com" in combined: return False, None
    for pat in DEFINITE_PATTERNS:
        if re.search(pat, text.lower()): return True, "definite"
    for kw in CAREER_KEYWORDS:
        if kw in combined: return True, "keyword"
    return False, None

def _job_signal_count(html):
    lower = html.lower()
    return sum(1 for s in JOB_PAGE_SIGNALS if s in lower)

def page_has_jobs(html):
    return _job_signal_count(html) >= 2

def is_likely_job_url(url, career_base_domain=""):
    if not url or len(url) < 10: return False
    url_lower = url.lower()
    path = urlparse(url).path.lower()
    if HARD_BLOCKED_PATTERNS.search(url_lower): return False
    if re.search(r"\.(svg|png|jpg|jpeg|gif|ico|pdf|zip|xml|json)$", path, re.I): return False
    if JOB_URL_ATS_SPECIFIC.search(path):
        if re.search(r"/go/job.?search/\d+/?$", path, re.I): return False
        if re.search(r"/go/all.?jobs/\d+/?$", path, re.I):   return False
        return True
    if JOB_PATH_CORE.search(path): return True
    return False

def _is_detail_page_url(url):
    if not url or len(url) < 10: return False
    path = urlparse(url).path.rstrip("/").lower()
    if HARD_BLOCKED_PATTERNS.search(url.lower()): return False
    if re.search(r"\.(svg|png|jpg|jpeg|gif|ico|pdf|zip|xml)$", path, re.I): return False
    if re.search(r"^/(jobs|job.search|job.search.results|openings"
                 r"|vacancies|all.jobs|opportunities|positions|search)$", path, re.I): return False
    return is_likely_job_url(url)

# ═════════════════════════════════════════════════════════════════════════════
#  HTTP HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def simple_get(url, timeout=10):
    try:
        r = requests.get(str(url), headers=HEADERS, timeout=timeout, allow_redirects=True)
        if r.status_code == 200: return r
    except: pass
    return None

def fetch_page(url: str, follow_redirects: bool = True, retries: int = 3) -> str | None:
    for attempt in range(retries):
        try:
            time.sleep(0.3 + attempt * 1.5)
            r = requests.get(url, headers=_next_headers(),
                             allow_redirects=follow_redirects, timeout=20)
            if r.status_code == 429:
                wait = 30 + attempt * 30
                log.warning(f"Rate-limited (429) — sleeping {wait}s"); time.sleep(wait); continue
            if r.status_code in (403, 999):
                log.warning(f"Blocked ({r.status_code}): {url}"); return None
            if r.status_code != 200:
                log.warning(f"HTTP {r.status_code}: {url}"); return None
            text = r.text
            return text[:FETCH_CHAR_LIMIT] if len(text) > FETCH_CHAR_LIMIT else text
        except Exception as e:
            log.warning(f"fetch attempt {attempt+1} failed ({url}): {e}")
            time.sleep(2 + attempt * 2)
    return None

# ═════════════════════════════════════════════════════════════════════════════
#  DATE HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def normalise_date_text(text: str) -> str:
    if not text: return ""
    fr_map = {"heure":"hour","heures":"hours","jour":"day","jours":"days",
              "semaine":"week","semaines":"weeks","mois":"month","an":"year","ans":"years"}
    m = re.search(r"il\s+y\s+a\s+(\d+)\s+([a-zéè]+)", text, re.I)
    if m:
        unit = fr_map.get(m.group(2).lower())
        if unit: return f"{m.group(1)} {unit} ago"
    if re.match(r"^hier$", text.strip(), re.I): return "1 day ago"
    if re.search(r"aujourd|today", text, re.I): return "0 days ago"
    return text

def resolve_posted_date(raw: str) -> str:
    if not raw: return ""
    text = normalise_date_text(raw)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", text.strip()): return text.strip()
    try: return datetime.fromisoformat(text).strftime("%Y-%m-%d")
    except: pass
    base = datetime.now()
    m = re.search(r"(\d+)\s*(hour|day|week|month|year)", text, re.I)
    if m:
        n, unit = int(m.group(1)), m.group(2).lower()
        if "hour"  in unit: base -= timedelta(hours=n)
        elif "day"  in unit: base -= timedelta(days=n)
        elif "week" in unit: base -= timedelta(weeks=n)
        elif "month" in unit:
            mo = base.month - n; yr = base.year + mo // 12; mo = mo % 12 or 12
            base = base.replace(year=yr, month=mo)
        elif "year" in unit: base = base.replace(year=base.year - n)
        return base.strftime("%Y-%m-%d")
    if re.search(r"just\s*now|today", text, re.I): return datetime.now().strftime("%Y-%m-%d")
    return datetime.now().strftime("%Y-%m-%d")

def try_parse_date(s: str) -> datetime | None:
    if not s: return None
    for fmt in ("%Y-%m-%dT%H:%M:%S","%Y-%m-%d","%B %d, %Y","%d %B %Y"):
        try: return datetime.strptime(s.strip(), fmt)
        except: pass
    try: return datetime.fromisoformat(s)
    except: pass
    m = re.match(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", s)
    if m:
        mon = MONTH_MAP.get(m.group(2)[:3].lower())
        if mon is not None: return datetime(int(m.group(3)), mon+1, int(m.group(1)))
    return None

def parse_deadline(soup: BeautifulSoup) -> str:
    full_text = soup.get_text()
    patterns = [
        r"closes?\s+on\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"apply\s+by\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"applications?\s+close[sd]?\s*(?:on)?\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"deadline[:\s]+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"deadline[:\s]+(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})",
        r"closing\s+date[:\s]+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
    ]
    now = datetime.now()
    for p in patterns:
        m = re.search(p, full_text, re.I)
        if m:
            d = try_parse_date(m.group(1))
            if d and d > now: return d.strftime("%Y-%m-%d")
    return ""

def estimate_deadline_from_posted(posted_text: str) -> str:
    if not posted_text: return ""
    text = normalise_date_text(posted_text)
    base = datetime.now()
    m = re.search(r"(\d+)\s*(hour|day|week|month)", text, re.I)
    if m:
        n, unit = int(m.group(1)), m.group(2).lower()
        if "hour"  in unit: base -= timedelta(hours=n)
        elif "day"  in unit: base -= timedelta(days=n)
        elif "week" in unit: base -= timedelta(weeks=n)
        elif "month" in unit:
            mo = base.month - n; yr = base.year + mo // 12; mo = mo % 12 or 12
            base = base.replace(year=yr, month=mo)
    mo = base.month + 3; yr = base.year + (mo - 1) // 12; mo = (mo - 1) % 12 + 1
    return base.replace(year=yr, month=mo).strftime("%Y-%m-%d")

def _estimate_deadline(date_posted_str):
    if not date_posted_str: return ""
    try:
        for fmt in ("%Y-%m-%d","%d/%m/%Y","%m/%d/%Y","%B %d, %Y","%b %d, %Y"):
            try:
                dt = datetime.strptime(date_posted_str.strip(), fmt)
                return (dt + timedelta(days=30)).strftime("%Y-%m-%d")
            except: continue
        m = re.search(r"(\d+)\s+days?\s+ago", date_posted_str, re.I)
        if m:
            posted = datetime.now() - timedelta(days=int(m.group(1)))
            return (posted + timedelta(days=30)).strftime("%Y-%m-%d")
    except: pass
    return ""

# ═════════════════════════════════════════════════════════════════════════════
#  TEXT CLEANERS
# ═════════════════════════════════════════════════════════════════════════════

def clean_description(raw: str) -> str:
    if not raw: return ""
    text = raw.replace("\u00a0"," ").replace("\u200b","").replace("\r\n","\n").replace("\r","\n")
    text = re.sub(r"([a-z])([A-Z])", r"\1 \2", text)
    text = re.sub(r"(\d)([A-Za-z])", r"\1 \2", text)
    text = re.sub(r"([A-Za-z])(\d)", r"\1 \2", text)
    text = re.sub(r"([.,:;!?])([A-Za-z0-9])", r"\1 \2", text)
    text = re.sub(r"\s*[•·▪◦]\s*", "\n• ", text)
    text = re.sub(r"\n\s*[-–—]\s+", "\n• ", text)
    paragraphs = re.split(r"\n{2,}", text)
    cleaned = []
    for para in paragraphs:
        lines = [l.strip() for l in para.split("\n") if l.strip()]
        out = []
        for line in lines:
            if (not re.search(r"[.!?:;,]$", line)
                    and not re.match(r"^[A-Z\s]{3,30}$", line)
                    and len(line) > 8
                    and not re.match(r"^[•\-–]", line)
                    and not re.match(r"^\w+:$", line)):
                line += "."
            out.append(line)
        cleaned.append("\n".join(out))
    return re.sub(r" {2,}", " ", "\n\n".join(p for p in cleaned if p.strip())).strip()

def clean_email(raw: str) -> str:
    if not raw: return ""
    em = raw
    em = re.sub(r"^mailto:", "", em, flags=re.I)
    em = re.sub(r"\?.*$", "", em)
    for pat, rep in [(r"\\u003[Ee]",""),(r"\\u003[Cc]",""),(r"\\u0040","@"),
                     (r"\\u002[Ee]","."),(r"&amp;",""),(r"&lt;",""),(r"&gt;",""),
                     (r"&#64;","@"),(r"&#46;","."),(r"%40","@"),(r"%2[Ee]",".")]:
        em = re.sub(pat, rep, em, flags=re.I)
    em = em.strip().lower()
    if not em or "@" not in em or "." not in em: return ""
    if not re.match(r"^[a-zA-Z0-9]", em): return ""
    at = em.rfind("@")
    if at == -1: return ""
    local, domain = em[:at], em[at+1:]
    if ".mu" in domain:   domain = re.sub(r"\.mu.*", ".mu", domain, flags=re.I)
    elif ".uk" in domain: domain = re.sub(r"\.uk.*", ".uk", domain, flags=re.I)
    else:                 domain = re.sub(r"(\.[a-z]{2,6})[a-z0-9\-_/?#+]*$", r"\1", domain, flags=re.I)
    em = local + "@" + domain
    if "@" not in em or "." not in em: return ""
    return em

def clean_application_link(raw: str) -> str:
    if not raw: return ""
    raw = raw.strip()
    if "@" in raw and not raw.startswith("http"): return clean_email(raw)
    if raw.startswith("http"):
        url = re.sub(r"#.*$", "", raw)
        url = re.sub(r"[.,;:!?)]+$", "", url)
        return url.strip()
    return raw

def clean_logo_url(raw: str) -> str:
    if not raw: return ""
    raw = decode_html_entities(raw).strip()
    if not raw.startswith("http"): return ""
    return re.sub(r"[\"')\s]+$", "", raw)

# ═════════════════════════════════════════════════════════════════════════════
#  EMAIL HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def extract_email_from_text(text: str) -> str:
    if not text: return ""
    emails = re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text)
    for raw_em in emails:
        em = clean_email(raw_em)
        if not em or "@" not in em: continue
        parts = em.split("@")
        if len(parts) != 2: continue
        if any(em.find(d) != -1 for d in NOISE_EMAIL_DOMAINS): continue
        if FAKE_LOCAL_RE.match(parts[0]) or FAKE_DOMAIN_RE.match(parts[1]): continue
        return em
    return ""

def scan_page_for_email(soup: BeautifulSoup, raw_html: str = "") -> str:
    for tag in soup.find_all("a", href=re.compile(r"^mailto:", re.I)):
        em = clean_email(tag.get("href", ""))
        if not em: continue
        if any(d in em for d in NOISE_EMAIL_DOMAINS): continue
        parts = em.split("@")
        if len(parts) == 2 and not FAKE_LOCAL_RE.match(parts[0]) and not FAKE_DOMAIN_RE.match(parts[1]):
            return em
    for sel in ["footer","#footer",".footer","#contact",".contact"]:
        for tag in soup.select(sel):
            found = extract_email_from_text(tag.get_text())
            if found: return found
    body = extract_email_from_text(soup.get_text())
    if body: return body
    if raw_html:
        found = extract_email_from_text(raw_html)
        if found: return found
    return ""

# ═════════════════════════════════════════════════════════════════════════════
#  JSON-LD PARSER
# ═════════════════════════════════════════════════════════════════════════════

def _parse_jsonld(html: str) -> dict:
    result = {}
    for raw in re.findall(
            r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
            html, re.S | re.I):
        try:    data = json.loads(raw.strip())
        except: continue
        if isinstance(data, list):
            data = next((d for d in data if isinstance(d, dict)), {})
        if not isinstance(data, dict): continue
        schema_type = data.get("@type", "")
        if schema_type == "JobPosting":
            org = data.get("hiringOrganization", {}) or {}
            result.update({
                "job_title":       data.get("title", ""),
                "job_description": data.get("description", ""),
                "date_posted":     data.get("datePosted", ""),
                "valid_through":   data.get("validThrough", ""),
                "employment_type": data.get("employmentType", ""),
                "salary":          _extract_salary_jsonld(data.get("baseSalary", {})),
                "company_name":    org.get("name", ""),
                "company_logo":    clean_logo_url(
                    org.get("logo","") if isinstance(org.get("logo"),str)
                    else org.get("logo",{}).get("url","") if isinstance(org.get("logo"),dict) else ""),
                "company_url":     org.get("sameAs","") or org.get("url",""),
                "company_website": org.get("sameAs","") or org.get("url",""),
                "apply_url":       (data.get("url","") or
                                    (data.get("applicationContact",{}) or {}).get("url","")),
                "location":        _extract_location_jsonld(data.get("jobLocation",{})),
            })
        elif schema_type in ("Organization","Corporation","LocalBusiness"):
            result.update({
                "company_name":     data.get("name",""),
                "company_logo":     clean_logo_url(
                    data.get("logo","") if isinstance(data.get("logo"),str)
                    else data.get("logo",{}).get("url","") if isinstance(data.get("logo"),dict) else ""),
                "company_url":      data.get("sameAs","") or data.get("url",""),
                "company_website":  data.get("sameAs","") or data.get("url",""),
                "company_industry": data.get("industry",""),
                "company_founded":  str(data.get("foundingDate","") or ""),
                "company_address":  _extract_address_jsonld(data.get("address",{})),
                "company_about":    data.get("description",""),
            })
    return result

def _extract_salary_jsonld(obj) -> str:
    if not obj: return ""
    if isinstance(obj, str): return obj
    if isinstance(obj, dict):
        val = obj.get("value", {}); currency = obj.get("currency", "")
        if isinstance(val, dict):
            lo = val.get("minValue",""); hi = val.get("maxValue",""); unit = val.get("unitText","")
            parts = [str(x) for x in [lo, hi] if x]
            return f"{currency} {' - '.join(parts)} {unit}".strip()
        return f"{currency} {val}".strip()
    return ""

def _extract_location_jsonld(obj) -> str:
    if not obj: return ""
    if isinstance(obj, list): obj = obj[0] if obj else {}
    if not isinstance(obj, dict): return str(obj)
    addr = obj.get("address", {})
    if isinstance(addr, dict):
        return ", ".join(filter(None, [
            addr.get("addressLocality",""), addr.get("addressRegion",""), addr.get("addressCountry",""),
        ]))
    return str(addr)

def _extract_address_jsonld(obj) -> str:
    if not obj: return ""
    if isinstance(obj, str): return obj
    if isinstance(obj, dict):
        return ", ".join(filter(None, [
            obj.get("streetAddress",""), obj.get("addressLocality",""),
            obj.get("addressRegion",""), obj.get("postalCode",""), obj.get("addressCountry",""),
        ]))
    return ""

def _extract_json_ld(soup):
    data = {}
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            obj = json.loads(script.string or "{}")
            if "@graph" in obj:
                for item in obj["@graph"]: data.update(_parse_json_ld_job(item))
            else: data.update(_parse_json_ld_job(obj))
        except: pass
    return data

def _parse_json_ld_job(obj):
    out = {}
    t = obj.get("@type", "")
    if "JobPosting" not in str(t): return out
    out["title"]       = obj.get("title", "")
    out["description"] = obj.get("description", "")
    out["date_posted"] = obj.get("datePosted", "")
    out["deadline"]    = obj.get("validThrough", "")
    out["job_type"]    = obj.get("employmentType", "")
    out["salary_range"] = ""
    bs = obj.get("baseSalary", {})
    if isinstance(bs, dict):
        val = bs.get("value", {})
        if isinstance(val, dict):
            mn = val.get("minValue",""); mx = val.get("maxValue",""); unit = val.get("unitText","")
            cur = bs.get("currency","")
            if mn or mx: out["salary_range"] = f"{cur} {mn}–{mx} ({unit})".strip()
        elif val: out["salary_range"] = str(val)
    loc = obj.get("jobLocation", {})
    if isinstance(loc, dict):
        addr = loc.get("address", {})
        if isinstance(addr, dict):
            parts = [addr.get("addressLocality",""), addr.get("addressRegion",""), addr.get("addressCountry","")]
            out["location"] = ", ".join(p for p in parts if p)
        else: out["location"] = str(addr)
    out["qualifications"] = obj.get("qualifications","")
    out["experience"]     = obj.get("experienceRequirements","")
    org = obj.get("hiringOrganization", {})
    if isinstance(org, dict):
        out["company_name"]    = org.get("name","")
        out["company_logo"]    = org.get("logo","")
        out["company_website"] = org.get("sameAs","")
        out["company_type"]    = org.get("@type","")
    out["field"] = obj.get("occupationalCategory","") or obj.get("industry","")
    return out

# ═════════════════════════════════════════════════════════════════════════════
#  DOM / PAGE HELPERS
# ═════════════════════════════════════════════════════════════════════════════

def _extract_logo(soup, base_url):
    og = soup.find("meta", property="og:image")
    if og and og.get("content"): return og["content"]
    for sel in ["img.logo","img[class*='logo']","img[alt*='logo']",
                ".logo img","header img",".navbar-brand img"]:
        el = soup.select_one(sel)
        if el and el.get("src"):
            src = el["src"]
            return src if src.startswith("http") else urljoin(base_url, src)
    return ""

def _extract_meta(soup, key):
    for attr in ["name","property"]:
        tag = soup.find("meta", {attr: key})
        if tag and tag.get("content"): return tag["content"]
    return ""

def _find_text_near_label(soup, *labels):
    for label in labels:
        pat = re.compile(re.escape(label), re.I)
        for dt in soup.find_all(["dt","th","strong","label","span","td","b"]):
            if pat.search(dt.get_text()):
                sib = dt.find_next_sibling()
                if sib:
                    val = _sanitize_field(sib.get_text())
                    if val: return clean(val)
                nxt = dt.find_next("td")
                if nxt:
                    val = _sanitize_field(nxt.get_text())
                    if val: return clean(val)
    return ""

def _parse_headed_sections(soup):
    sections = {}
    for h in soup.find_all(["h1","h2","h3","h4"]):
        key = h.get_text(strip=True).lower()
        if not key or len(key) > 120: continue
        body_parts = []
        for sib in h.find_next_siblings():
            if sib.name in ("h1","h2","h3","h4"): break
            text = sib.get_text(separator=" ", strip=True)
            if text: body_parts.append(text)
        sections[key] = " ".join(body_parts).strip()
    return sections

def _pick_section(sections, *keywords):
    for kw in keywords:
        kw_lower = kw.lower()
        for key, val in sections.items():
            if kw_lower in key and val:
                cleaned = _sanitize_field(val[:300])
                if cleaned: return val
    return ""

def _extract_bold_field(soup, *labels):
    for label in labels:
        pat = re.compile(re.escape(label) + r"\s*:?", re.I)
        for el in soup.find_all(["strong","b","p","li"]):
            text = el.get_text()
            if pat.search(text):
                cleaned = pat.sub("", text).strip()
                if cleaned and len(cleaned) > 2:
                    val = _sanitize_field(cleaned, 300)
                    if val: return val
    return ""

def _collect_job_links(soup, base_url, career_domain):
    seen, results = set(), []
    for a in soup.find_all("a", href=True):
        href = a.get("href","")
        if not href or href.startswith(("#","mailto:","tel:","javascript")): continue
        full = href if href.startswith("http") else urljoin(base_url, href)
        if full in seen: continue
        seen.add(full)
        if not is_likely_job_url(full, career_domain): continue
        text = a.get_text(strip=True)
        if not text or len(text) < 3:
            parent = a.parent
            for _ in range(3):
                if parent is None: break
                heading = parent.find(["h1","h2","h3","h4","strong","b"])
                if heading: text = heading.get_text(strip=True); break
                parent = parent.parent
        if not text or len(text) < 3:
            parts = urlparse(full).path.rstrip("/").split("/")
            slug = parts[-2] if (len(parts) >= 2 and parts[-1].isdigit()) else parts[-1]
            text = re.sub(r"[-_]", " ", slug).strip()
            text = re.sub(r"\s+\d{5,}$", "", text).strip()
        if text and len(text) >= 3:
            results.append((full, text))
    return results

def _pagination(soup, base_url):
    urls = []
    for a in soup.find_all("a", href=True):
        text = a.get_text(strip=True).lower()
        href = a.get("href","")
        if not href: continue
        full = href if href.startswith("http") else urljoin(base_url, href)
        if any(kw in text for kw in ["next","›","»","load more","show more"]):
            urls.append(full)
        elif re.match(r"^\d+$", text) and int(text) <= 50:
            urls.append(full)
    return list(dict.fromkeys(urls))

def _scan_links(soup, website, base_domain):
    best_score, best_url, best_strat, best_ats = 0, None, None, None
    for a in soup.find_all("a", href=True):
        href = a.get("href",""); text = clean(a.get_text(), 80).lower()
        if not href or href.startswith(("#","mailto:","tel:")): continue
        full = href if href.startswith("http") else urljoin(website, href)
        if _is_blocked_career_url(full): continue
        ats_name = is_ats(full)
        if ats_name: return full, f"ats_link:{ats_name}", ats_name
        ok, reason = is_career_link(href, text)
        if ok:
            score = (10 if reason == "definite" else 5) + (3 if base_domain in full else 0)
            if score > best_score:
                best_score, best_url, best_strat = score, full, f"link:{reason}"
    return (best_url, best_strat, best_ats) if best_url and best_score >= 5 else None

# ═════════════════════════════════════════════════════════════════════════════
#  SF TABLE SCRAPER
# ═════════════════════════════════════════════════════════════════════════════

def _scrape_sf_listing_table(soup, listing_url):
    results = []
    table = soup.find("table")
    if not table: return results
    for row in table.find_all("tr"):
        cells = row.find_all("td")
        if len(cells) < 2: continue
        title_cell = cells[0]
        links = title_cell.find_all("a", href=True)
        if not links: continue
        seen_urls = set()
        for link in links:
            href = link.get("href","")
            if not href or href in seen_urls: continue
            seen_urls.add(href)
            full_url = href if href.startswith("http") else urljoin(listing_url, href)
            if not _is_detail_page_url(full_url): continue
            title       = link.get_text(strip=True)
            location    = clean(cells[1].get_text()) if len(cells) >= 2 else ""
            date_posted = clean(cells[2].get_text()) if len(cells) >= 3 else ""
            if title and len(title) > 2:
                results.append({"url": full_url, "title": title,
                                 "location": location, "date_posted": date_posted})
            break
    return results

# ═════════════════════════════════════════════════════════════════════════════
#  STANDARDISATION TABLES
# ═════════════════════════════════════════════════════════════════════════════

FIELD_KEYWORD_MAP = [
    ("Information Technology",
     ["software engineer","developer","devops","frontend","backend","full stack","fullstack",
      "sysadmin","cloud","cybersecurity","data engineer","machine learning","artificial intelligence",
      "ai/ml","it support","network engineer","database","kubernetes","docker","aws","azure",
      "react","node.js","python developer","java developer","it manager","systems analyst",
      "erp","sap","technology","tech lead","infrastructure engineer","qa engineer",
      "automation engineer","business intelligence","data analyst"],
     ["programming","coding","api","agile","scrum","git","linux","server","infrastructure",
      "software","digital","tech"]),
    ("Finance & Accounting",
     ["accountant","auditor","finance manager","financial analyst","cfo","treasurer","tax",
      "bookkeeper","payroll","budget analyst","credit analyst","investment","portfolio manager",
      "risk analyst","forex","actuary","acca","cfa","cpa","finance officer",
      "financial controller","internal audit","external audit","accounts payable","treasury"],
     ["financial","accounting","balance sheet","p&l","reconciliation","ifrs","gaap",
      "ledger","invoicing","fiscal","budget","revenue"]),
    ("Sales & Business Development",
     ["sales executive","sales manager","business development","account manager",
      "sales representative","bd manager","regional sales","key account","sales director",
      "commercial manager","sales officer","revenue manager","channel manager","pre-sales"],
     ["revenue","pipeline","crm","leads","prospects","quota","target","upsell","b2b","b2c"]),
    ("Marketing & Communications",
     ["marketing manager","digital marketing","seo","sem","content marketer",
      "social media manager","brand manager","marketing executive","communications manager",
      "pr manager","copywriter","growth hacker","email marketing","campaign manager",
      "marketing director","public relations"],
     ["marketing","branding","advertising","social media","content","campaign","analytics",
      "google ads","facebook ads","influencer","positioning"]),
    ("Human Resources",
     ["hr manager","human resources","recruiter","talent acquisition","hr business partner",
      "hrbp","hr officer","compensation","benefits manager","organisational development",
      "learning and development","l&d","hr generalist","payroll manager",
      "people operations","talent management","hr director"],
     ["recruitment","onboarding","performance management","employee relations","hr",
      "workforce","staffing","saudization","nitaqat"]),
    ("Engineering",
     ["mechanical engineer","civil engineer","electrical engineer","structural engineer",
      "process engineer","project engineer","maintenance engineer","production engineer",
      "quality engineer","safety engineer","site engineer","design engineer",
      "petroleum engineer","chemical engineer","industrial engineer","hvac engineer"],
     ["engineering","cad","autocad","solidworks","manufacturing","plant","machinery",
      "commissioning","maintenance","iso","asme"]),
    ("Healthcare & Medicine",
     ["doctor","physician","nurse","pharmacist","medical officer","surgeon","anaesthetist",
      "physiotherapist","radiographer","lab technician","clinical","healthcare manager",
      "occupational therapist","dentist","midwife","radiologist","oncologist","cardiologist",
      "icu","medical director"],
     ["hospital","clinic","patient","medical","health","pharmaceutical","diagnosis","treatment"]),
    ("Education & Training",
     ["teacher","lecturer","professor","trainer","educator","tutor","school principal",
      "academic","curriculum","e-learning","instructional designer","teaching assistant",
      "academic advisor","dean","faculty"],
     ["school","university","college","classroom","students","pedagogy","curriculum","education"]),
    ("Hospitality & Tourism",
     ["hotel manager","front desk","housekeeping","chef","sous chef","food and beverage",
      "f&b manager","restaurant manager","bartender","waiter","concierge","tour guide",
      "travel agent","events coordinator","catering","guest relations"],
     ["hospitality","hotel","resort","tourism","guest","accommodation","restaurant","kitchen"]),
    ("Logistics & Supply Chain",
     ["supply chain manager","logistics coordinator","warehouse manager","fleet manager",
      "procurement manager","purchasing manager","import export","freight",
      "shipping coordinator","inventory manager","demand planner","customs clearance",
      "logistics manager","distribution manager"],
     ["logistics","supply chain","warehouse","inventory","freight","procurement","sourcing",
      "distribution","customs","3pl"]),
    ("Legal",
     ["lawyer","attorney","legal counsel","paralegal","compliance officer","legal advisor",
      "solicitor","barrister","corporate counsel","legal manager","contract manager",
      "in-house counsel","data protection officer"],
     ["legal","law","contracts","litigation","regulatory","compliance","gdpr","arbitration"]),
    ("Administration & Operations",
     ["office manager","executive assistant","administrative officer","operations manager",
      "personal assistant","receptionist","data entry","office administrator",
      "company secretary","business analyst","operations officer","facility manager"],
     ["administration","operations","office","coordination","scheduling","reporting","clerical"]),
    ("Customer Service",
     ["customer service","call centre","customer success","customer support","help desk",
      "service advisor","client relations","customer experience","contact centre",
      "customer care","cx specialist","complaints officer"],
     ["customer","support","helpdesk","tickets","escalation","satisfaction","service level"]),
    ("Construction & Real Estate",
     ["quantity surveyor","site supervisor","project manager construction","architect",
      "draughtsman","property manager","estate agent","real estate","building inspector",
      "land surveyor","construction manager","project director","bim engineer","fit out"],
     ["construction","building","property","real estate","site","contractor","tender","neom"]),
    ("Manufacturing & Production",
     ["production manager","quality control","quality assurance","qa","qc","factory manager",
      "plant manager","production supervisor","assembly","cnc operator","technician"],
     ["production","manufacturing","factory","assembly","quality","lean","six sigma"]),
    ("Design & Creative",
     ["graphic designer","ui/ux","product designer","art director","creative director",
      "animator","illustrator","photographer","videographer","motion designer","web designer",
      "ux researcher","visual designer"],
     ["design","creative","adobe","figma","photoshop","illustrator","indesign","branding"]),
    ("Research & Science",
     ["research scientist","data scientist","lab researcher","research analyst",
      "clinical researcher","environmental scientist","chemist","biologist","statistician",
      "geologist","geophysicist","reservoir engineer","r&d"],
     ["research","analysis","data","laboratory","science","experiment","findings"]),
    ("Security",
     ["security officer","security guard","security manager","cctv","loss prevention",
      "risk manager","health and safety","hse officer","osh","fire safety","security analyst",
      "information security","cyber analyst","soc analyst"],
     ["security","safety","risk","surveillance","patrol","access control","emergency"]),
    ("Media & Journalism",
     ["journalist","editor","reporter","broadcast","news anchor","content creator",
      "media manager","radio","television","producer","scriptwriter","editorial producer"],
     ["media","journalism","broadcast","news","editorial","publishing","press"]),
    ("Oil & Gas",
     ["petroleum engineer","drilling engineer","reservoir engineer","production engineer oil",
      "subsurface","geoscientist","upstream","downstream","refinery","petrochemical",
      "gas plant","field operator","well engineer","hse oil"],
     ["oil","gas","petroleum","refinery","upstream","downstream","aramco","sabic",
      "petrochemical","drilling","reservoir","pipeline","lng","lpg"]),
    ("Non-Profit & Social Work",
     ["social worker","ngo","charity","programme coordinator","community development",
      "welfare officer","case manager","development officer","fundraiser"],
     ["social","ngo","community","welfare","beneficiary","donor","impact","charity"]),
]


def standardise_field(raw_field: str, title: str = "", description: str = "",
                       industry: str = "") -> str:
    combined = " ".join([
        (raw_field or ""), (title or ""), (description or "")[:800], (industry or ""),
    ]).lower()
    best_label, best_score = "", 0
    for label, high_kws, low_kws in FIELD_KEYWORD_MAP:
        score  = sum(3 for kw in high_kws if kw in combined)
        score += sum(1 for kw in low_kws  if kw in combined)
        if score > best_score:
            best_score, best_label = score, label
    if best_score >= 1: return best_label
    return ""

# ── experience ────────────────────────────────────────────────────────────────
_NO_EXP_KW = ["no experience","no prior experience","fresh graduate","freshers",
              "entry level","entry-level","0 years","zero experience",
              "training provided","will train","no experience required"]
_LT1_KW    = ["less than 1 year","under 1 year","6 months","less than a year",
              "some experience","minimal experience","up to 1 year"]
_EXP_RE    = re.compile(
    r"(?:minimum|min\.?|at\s+least|over|more\s+than)?\s*"
    r"(\d+)\s*(?:\+|plus)?\s*(?:[-–to]+\s*(\d+))?\s*"
    r"years?(?:\s+of)?(?:\s+(?:relevant\s+)?experience)?", re.I,
)

def _years_to_band(n: int) -> str:
    if n <= 0:  return "No Experience Required"
    if n <= 2:  return "1 - 2 Years"
    if n <= 5:  return "3 - 5 Years"
    if n <= 10: return "6 - 10 Years"
    return "10+ Years"

def standardise_experience(raw: str) -> str:
    if not raw: return ""
    text = raw.lower().strip()
    for kw in _NO_EXP_KW:
        if kw in text: return "No Experience Required"
    for kw in _LT1_KW:
        if kw in text: return "Less than 1 Year"
    matches = _EXP_RE.findall(text)
    if matches:
        nums = []
        for m in matches:
            for g in m:
                if g:
                    try: nums.append(int(g))
                    except: pass
        if nums: return _years_to_band(min(nums))
    m = re.search(r"(\d+)\s*\+?\s*years?", text, re.I)
    if m: return _years_to_band(int(m.group(1)))
    if re.search(r"\b(senior|sr\.?|lead|principal|head of|director|vp)\b", text, re.I): return "6 - 10 Years"
    if re.search(r"\b(mid.?level|intermediate|associate)\b", text, re.I): return "3 - 5 Years"
    if re.search(r"\b(junior|jr\.?|graduate|intern|trainee|fresh)\b", text, re.I): return "Less than 1 Year"
    return ""

# ── qualification ─────────────────────────────────────────────────────────────
QUALIFICATION_TIERS = [
    ("PhD / Doctorate",          ["phd","ph.d","doctorate","doctoral","doctor of philosophy"]),
    ("Master's Degree",          ["master","msc","m.sc","mba","m.b.a","meng","m.eng","mphil",
                                   "postgraduate","post-graduate"]),
    ("Bachelor's Degree",        ["bachelor","bsc","b.sc","beng","b.eng","bcom","b.com","bba",
                                   "llb","degree in","undergraduate degree","honours degree","hons","b.tech","btech"]),
    ("Higher National Diploma",  ["hnd","hnc","higher national diploma","higher national certificate",
                                   "higher diploma","advanced diploma"]),
    ("Diploma",                  ["diploma","associate degree","foundation degree"]),
    ("Professional Certification",["acca","cpa","cfa","cima","pmp","prince2","cissp","aws certified",
                                   "comptia","cisco","ccna","ccnp","shrm","cipd","chartered",
                                   "professional certification","professional certificate"]),
    ("A-Levels / HSC",           ["a-level","a level","hsc","higher school certificate",
                                   "ib diploma","international baccalaureate"]),
    ("O-Levels / School Certificate",["o-level","o level","igcse","gcse","school certificate"]),
    ("No Formal Qualification Required",["no qualification","no degree","no formal","school leaver",
                                         "no experience required","training provided","will train"]),
]

def standardise_qualification(raw: str, full_text: str = "") -> str:
    corpus = ((raw or "") + " " + (full_text or "")[:2000]).lower()
    for label, keywords in QUALIFICATION_TIERS:
        for kw in keywords:
            if kw in corpus: return label
    return ""

# ═════════════════════════════════════════════════════════════════════════════
#  MAKE JOB ROW (unified record builder)
# ═════════════════════════════════════════════════════════════════════════════

def make_job(company, website, industry, careers_url, source,
             title="", location="", job_type="", department="", apply_url="",
             qualifications="", experience="", field="",
             date_posted="", deadline="", description="",
             company_logo="", company_founded="", company_type="",
             company_address="", company_details="",
             estimated_deadline="", salary_range="",
             company_industry="", company_name_override=""):

    std_field = standardise_field(field or department, title, description, industry)
    std_exp   = standardise_experience(experience)
    std_qual  = standardise_qualification(qualifications, description)

    return {
        "Job Title":          clean(title),
        "Job Type":           _sanitize_field(job_type) or "Full-time",
        "Job Qualifications": std_qual,
        "Job Experience":     std_exp,
        "Job Location":       clean(location) or "Saudi Arabia",
        "Job Field":          std_field,
        "Date Posted":        _sanitize_field(date_posted),
        "Deadline":           _sanitize_field(deadline),
        "Job Description":    clean(description, 2000),
        "Application":        apply_url,
        "Company URL":        careers_url,
        "Company Name":       clean(company_name_override or company),
        "Company Logo":       company_logo,
        "Company Industry":   clean(company_industry or industry),
        "Company Founded":    _sanitize_field(company_founded),
        "Company Type":       _sanitize_field(company_type),
        "Company Website":    website,
        "Company Address":    clean(company_address),
        "Company Details":    clean(company_details, 1000),
        "Job URL":            apply_url,
        "Estimated Deadline": _sanitize_field(estimated_deadline),
        "Salary Range":       _sanitize_field(salary_range),
        "source":             source,
    }

# ═════════════════════════════════════════════════════════════════════════════
#  WORDPRESS LOGO UPLOAD
# ═════════════════════════════════════════════════════════════════════════════

def upload_logo_to_wordpress(logo_url: str, company_name: str) -> str:
    if not logo_url or not logo_url.startswith("http") or not WP_USER: return ""
    try:
        r = requests.get(logo_url, headers={"User-Agent": HEADERS["User-Agent"],
                                             "Referer": "https://www.linkedin.com/"}, timeout=15)
        if r.status_code != 200: return ""
        ct = r.headers.get("Content-Type", "image/jpeg")
        ext = "png" if "png" in ct else "jpg"
        fn = re.sub(r"[^a-z0-9]", "-", company_name.lower()) + "-logo." + ext
        creds = base64.b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()
        wr = requests.post(WP_URL + "media",
                           headers={"Authorization": "Basic " + creds,
                                    "Content-Disposition": f"attachment; filename={fn}",
                                    "Content-Type": ct},
                           data=r.content, timeout=20)
        if wr.status_code in (200, 201): return wr.json().get("source_url","")
        log.warning(f"WP upload failed ({wr.status_code})")
    except Exception as e:
        log.warning(f"uploadLogoToWordPress: {e}")
    return ""

# ═════════════════════════════════════════════════════════════════════════════
#  LINKEDIN-SPECIFIC HELPERS  (kept from v4)
# ═════════════════════════════════════════════════════════════════════════════

def decode_linkedin_apply_url(raw: str) -> str:
    if not raw: return ""
    raw = decode_html_entities(raw)
    if raw.startswith("http") and "linkedin.com" not in raw: return raw
    m = re.search(r"[?&]url=([^&]+)", raw)
    if m:
        try:
            d = unquote(m.group(1))
            if "%" in d: d = unquote(d)
            if d.startswith("http") and "linkedin.com" not in d: return d
        except: pass
    b64m = re.search(r"[?&]offsiteApplyUrl=([^&]+)", raw)
    if b64m:
        try:
            d2 = base64.b64decode(unquote(b64m.group(1))).decode("utf-8")
            p  = json.loads(d2)
            if p and "url" in p: return p["url"]
        except: pass
    return ""

def follow_linkedin_apply_button(soup: BeautifulSoup, job_url: str) -> str:
    for tag in soup.find_all("a", href=True):
        ctrl = tag.get("data-tracking-control-name","")
        if "offsite" in ctrl.lower() or "apply" in ctrl.lower():
            r = decode_linkedin_apply_url(tag["href"])
            if r and not is_bad_url(r): return r
    for tag in soup.find_all("a", href=True):
        href = tag["href"]; text = tag.get_text().lower()
        if ("apply" in text or "/apply" in href) and "linkedin.com" not in href:
            if href.startswith("http") and not is_bad_url(href): return href
    return ""

def extract_company_from_job_page(html: str, soup: BeautifulSoup) -> dict:
    result = {}
    ld = _parse_jsonld(html)
    if ld: result.update({k: v for k, v in ld.items() if v})
    def _meta(name_or_prop):
        tag = (soup.find("meta", attrs={"property": name_or_prop}) or
               soup.find("meta", attrs={"name": name_or_prop}))
        return (tag.get("content","") if tag else "").strip()
    og_image = _meta("og:image")
    if og_image and not result.get("company_logo"):
        result["company_logo"] = clean_logo_url(og_image)
    def _sel(*selectors):
        for s in selectors:
            el = soup.select_one(s)
            if el:
                t = el.get_text(strip=True)
                if t: return t
        return ""
    if not result.get("company_name"):
        result["company_name"] = _sel(
            ".topcard__org-name-link",
            ".job-details-jobs-unified-top-card__company-name",
            ".topcard__flavor",
        )
    if not result.get("company_logo"):
        for img_sel in [".artdeco-entity-image","img.company-logo",
                        ".jobs-unified-top-card__company-logo img",".topcard__logo img"]:
            img = soup.select_one(img_sel)
            if img:
                src = img.get("src","") or img.get("data-delayed-url","") or img.get("data-ghost-url","")
                src = clean_logo_url(src)
                if src and "ghost" not in src.lower() and "placeholder" not in src.lower():
                    result["company_logo"] = src; break
    for script in soup.find_all("script"):
        txt = script.string or ""
        if not txt.strip(): continue
        if not result.get("company_website"):
            for pat in [r'"companyPageUrl"\s*:\s*"([^"]+)"',
                        r'"companyUrl"\s*:\s*"([^"]+)"',
                        r'"websiteUrl"\s*:\s*"([^"]+)"']:
                m = re.search(pat, txt)
                if m:
                    url = _strip_li_tracking(decode_html_entities(m.group(1)))
                    if url.startswith("http") and "linkedin.com" not in url:
                        result["company_website"] = url; break
        if not result.get("company_logo"):
            for pat in [r'"logoUrl"\s*:\s*"([^"]+)"',
                        r'"companyLogo"\s*:\s*"([^"]+)"',
                        r'"logo"\s*:\s*"([^"]+)"']:
                m = re.search(pat, txt)
                if m:
                    logo = clean_logo_url(decode_html_entities(m.group(1)))
                    if logo: result["company_logo"] = logo; break
    return result

def get_job_criteria(soup: BeautifulSoup, label: str) -> str:
    lower = label.lower()
    for li in soup.select(".description__job-criteria-list > li"):
        h3 = li.find("h3")
        if h3 and lower in h3.get_text().strip().lower():
            spans = li.select(".description__job-criteria-text, span")
            if spans: return spans[-1].get_text(strip=True)
    for chip in soup.select(
        ".job-details-jobs-unified-top-card__job-insight,"
        ".jobs-unified-top-card__job-insight"):
        text = chip.get_text(strip=True).lower()
        if "employment" in lower or "type" in lower:
            if re.search(r"full[\-\s]?time|part[\-\s]?time|contract|temporary|internship|freelance", text, re.I):
                return chip.get_text(strip=True)
    return ""

def get_workplace_type(soup: BeautifulSoup) -> str:
    for s in [".topcard__workplace-type",
              ".job-details-jobs-unified-top-card__workplace-type",
              ".jobs-unified-top-card__workplace-type"]:
        el = soup.select_one(s)
        if el: return el.get_text(strip=True)
    return ""

# ═════════════════════════════════════════════════════════════════════════════
#  LINKEDIN COMPANY PAGE SCRAPER  (v4 unchanged)
# ═════════════════════════════════════════════════════════════════════════════

def scrape_company_details(company_url: str) -> dict:
    empty = {"name":"","industry":"","size":"","headquarters":"","type":"",
             "founded":"","specialties":"","website":"","logo":"","about":""}
    if not company_url: return empty
    log.info(f"Scraping LinkedIn company page: {company_url}")
    base_url = re.sub(r"\?.*$", "", company_url.rstrip("/"))
    html = None
    guest_url = base_url.replace(
        "https://www.linkedin.com/company/",
        "https://www.linkedin.com/company-guest/",
    )
    if guest_url != base_url: html = fetch_page(guest_url)
    if not html:
        for attempt in range(3):
            try:
                time.sleep(1.5 + attempt * 2)
                r = requests.get(base_url, headers=_next_headers(), allow_redirects=True, timeout=20)
                if r.status_code == 429:
                    log.warning("Company page rate-limited — sleeping 60s"); time.sleep(60); continue
                if r.status_code == 200:
                    text = r.text
                    html = text[:FETCH_CHAR_LIMIT] if len(text) > FETCH_CHAR_LIMIT else text
                    break
            except Exception as e:
                log.warning(f"Company page fetch error (attempt {attempt+1}): {e}")
                time.sleep(2 + attempt * 2)
    if not html: return empty
    soup = BeautifulSoup(html, "html.parser")
    ld   = _parse_jsonld(html)
    def _sel(*selectors):
        for s in selectors:
            el = soup.select_one(s)
            if el:
                t = el.get_text(strip=True)
                if t: return t
        return ""
    def _get_detail(label):
        lower = label.lower()
        for div in soup.select("section.core-section-container dl > div"):
            dt = div.find("dt")
            if dt and lower in dt.get_text().strip().lower():
                dd = div.find("dd")
                if dd: return dd.get_text(strip=True)
        for row in soup.select(".org-page-details__definition-list dt, .about-us__basicInfo dt"):
            if lower in row.get_text().strip().lower():
                dd = row.find_next_sibling("dd")
                if dd: return dd.get_text(strip=True)
        return ""
    og_img_tag = (soup.find("meta", property="og:image") or
                  soup.find("meta", attrs={"name": "og:image"}))
    raw_logo = (og_img_tag.get("content","") if og_img_tag else "") or ld.get("company_logo","")
    if not raw_logo:
        for img in soup.select("img.org-top-card-primary-content__logo, img.artdeco-entity-image"):
            src = img.get("src","") or img.get("data-delayed-url","")
            if src and "ghost" not in src.lower():
                raw_logo = src; break
    logo = clean_logo_url(raw_logo)
    ws_tag = soup.select_one("a[data-tracking-control-name='about_website']")
    raw_ws = (ws_tag.get("href","") if ws_tag else "") or _get_detail("Website") or ld.get("company_website","")
    website = decode_linkedin_apply_url(raw_ws) or raw_ws
    name = (ld.get("company_name","") or _sel("h1.org-top-card-summary__title","h1","title") or "")
    if " | LinkedIn" in name: name = name.split(" | ")[0].strip()
    about = (ld.get("company_about","") or
             _sel("section.about-us p",".core-section-container__content p",
                  ".org-about-us-organization-description__text",
                  ".org-about-module__description") or "")
    hosted_logo = upload_logo_to_wordpress(logo, name) if logo else ""
    return {
        "name":         name,
        "industry":     _get_detail("Industry") or ld.get("company_industry",""),
        "size":         _get_detail("Company size"),
        "headquarters": _get_detail("Headquarters") or ld.get("company_address",""),
        "type":         _get_detail("Type"),
        "founded":      _get_detail("Founded") or ld.get("company_founded",""),
        "specialties":  _get_detail("Specialties"),
        "website":      website,
        "logo":         hosted_logo or logo,
        "about":        about,
    }

# ═════════════════════════════════════════════════════════════════════════════
#  PLAYWRIGHT HELPERS
# ═════════════════════════════════════════════════════════════════════════════

async def _resolve_iframe_ats(page, html, base_url):
    soup = BeautifulSoup(html, "html.parser")
    for iframe in soup.find_all("iframe"):
        src = iframe.get("src","") or iframe.get("data-src","")
        if not src: continue
        full = src if src.startswith("http") else urljoin(base_url, src)
        if is_ats(full):
            print(f"       🖼  Iframe ATS detected: {full}")
            return full
    try:
        for frame in page.frames:
            furl = frame.url
            if furl and furl != base_url and furl != "about:blank":
                if is_ats(furl):
                    print(f"       🖼  Live iframe ATS frame: {furl}")
                    return furl
    except: pass
    m = re.search(r'<iframe[^>]+src=["\']?(https?://[^"\'>\s]+)["\']?', html, re.I)
    if m:
        src = m.group(1)
        if is_ats(src):
            print(f"       🖼  Regex iframe ATS: {src}")
            return src
    return None

# ─── suffix prober ───────────────────────────────────────────────────────────
async def _probe_suffixes(page, career_root, root_score=0):
    root = career_root.rstrip("/")
    best_url, best_score = None, root_score
    for suffix in JOB_LISTING_SUFFIXES:
        candidate = root + suffix
        if candidate.rstrip("/") == root: continue
        html = None; final_url = candidate
        r = simple_get(candidate)
        if r and r.url.rstrip("/") != root:
            html = r.text; final_url = r.url
        else:
            try:
                resp = await page.goto(candidate, timeout=15000, wait_until="domcontentloaded")
                if resp and resp.status < 400:
                    final_url = page.url
                    if final_url.rstrip("/") != root:
                        await page.wait_for_timeout(1200)
                        html = await page.content()
            except: pass
        if not html: continue
        score = _job_signal_count(html)
        if score > best_score:
            best_score = score; best_url = final_url
            print(f"       Better jobs URL via suffix '{suffix}': score={score} → {final_url}")
    return best_url, best_score

# ─── crawl career links ──────────────────────────────────────────────────────
def _crawl_career_links(html, career_root):
    career_domain = get_domain(career_root)
    soup = BeautifulSoup(html, "html.parser")
    candidates = {}
    for a in soup.find_all("a", href=True):
        href = a.get("href",""); text = a.get_text(strip=True).lower()
        if not href or href.startswith(("#","mailto","tel")): continue
        full = href if href.startswith("http") else urljoin(career_root, href)
        full_lower = full.lower()
        if "linkedin.com" in full_lower: continue
        if career_domain not in full_lower and not is_ats(full): continue
        score = 0
        if re.search(r"/(job.search.results?|job.search|jobs.search|career.search"
                     r"|search|openings?|positions?|listings?|vacancies|all.jobs)", full_lower): score += 8
        if re.search(r"/(job|career|role|vacanc|posting)s?[/_-]", full_lower): score += 5
        if re.search(r"/(en|ar)/(job|career|vacanc|opening|position|search)", full_lower): score += 7
        if re.search(r"\b(view|see|all|browse|search|explore)\b.*\b(job|position|opening|vacanc|career)", text): score += 6
        if re.search(r"\b(job|position|opening|vacanc|career)s?\b", text): score += 3
        if re.search(r"/go/[A-Za-z0-9%-]+/\d+", full_lower): score += 10
        if score >= 3: candidates[full] = max(candidates.get(full, 0), score)
    for url, _ in sorted(candidates.items(), key=lambda x: -x[1])[:5]:
        return url
    return None

# ─── resolve to jobs listing URL ─────────────────────────────────────────────
async def _resolve_to_jobs_url(page, career_root, strategy, ats_name):
    if _is_blocked_career_url(career_root):
        return None, strategy + "+blocked", None
    if ats_name:
        return career_root, strategy, ats_name
    r_root    = simple_get(career_root)
    root_html = r_root.text if r_root else ""
    root_score = _job_signal_count(root_html) if root_html else 0
    if root_html and not ats_name:
        detected_ats = is_ats_from_html(root_html)
        if detected_ats:
            return career_root, strategy + "+html_fp", detected_ats
    suffix_url, suffix_score = await _probe_suffixes(page, career_root, root_score)
    if suffix_url:
        return suffix_url, strategy + "+suffix", None
    if root_score >= 2:
        return career_root, strategy, ats_name
    if root_html:
        found = _crawl_career_links(root_html, career_root)
        if found: return found, strategy + "+crawl", None
    try:
        await page.goto(career_root, timeout=25000, wait_until="domcontentloaded")
        await page.wait_for_timeout(2500)
        html = await page.content()
        detected_ats = is_ats_from_html(html)
        if detected_ats:
            m = re.search(r'(https?://[^\s"\'<>]*/go/[^\s"\'<>]+)', html, re.I)
            if m: return m.group(1), strategy + "+js_fp_go", detected_ats
            return career_root, strategy + "+js_fp", detected_ats
        for pat, aname in ATS_DOMAINS.items():
            if pat in html.lower():
                m = re.search(r'https?://[^\s"\'<>]*' + re.escape(pat) + r'[^\s"\'<>]*', html, re.I)
                if m: return m.group(0), strategy + "+js_ats", aname
        iframe_src = await _resolve_iframe_ats(page, html, career_root)
        if iframe_src:
            aname = is_ats(iframe_src)
            return iframe_src, strategy + "+js_iframe", aname
        suffix_url2, suffix_score2 = await _probe_suffixes(page, career_root, max(root_score, _job_signal_count(html)))
        if suffix_url2: return suffix_url2, strategy + "+js_suffix", None
        if _job_signal_count(html) >= 2: return career_root, strategy + "+js", None
        found = _crawl_career_links(html, career_root)
        if found: return found, strategy + "+js_crawl", None
    except Exception as e:
        print(f"       JS resolve error: {e}")
    return career_root, strategy + "+unresolved", None

# ─── main career page finder ─────────────────────────────────────────────────
async def find_career_page(page, website):
    base_domain, scheme = get_base(website)
    base = f"{scheme}://{base_domain}"
    for sub in CAREER_SUBDOMAINS:
        url = f"{scheme}://{sub}.{base_domain}"
        r   = simple_get(url)
        if r:
            if _is_blocked_career_url(r.url): continue
            resolved = await _resolve_to_jobs_url(page, r.url, f"subdomain:{sub}", is_ats(r.url))
            if resolved[0] and not _is_blocked_career_url(resolved[0]):
                return resolved
    for path in CAREER_PATHS:
        url = base + path
        r   = simple_get(url)
        if r and r.url not in (base, base+"/", base+"/#"):
            if _is_blocked_career_url(r.url): continue
            resolved = await _resolve_to_jobs_url(page, r.url, f"path:{path}", is_ats(r.url))
            if resolved[0] and not _is_blocked_career_url(resolved[0]):
                return resolved
    r = simple_get(website)
    if r:
        result = _scan_links(BeautifulSoup(r.text,"html.parser"), website, base_domain)
        if result:
            career_root, strat, aname = result
            if not _is_blocked_career_url(career_root):
                return await _resolve_to_jobs_url(page, career_root, strat, aname)
    if _PLAYWRIGHT_OK:
        try:
            await page.goto(website, timeout=25000, wait_until="domcontentloaded")
            await page.wait_for_timeout(2000)
            content = await page.content()
            detected_ats = is_ats_from_html(content)
            if detected_ats:
                m_go = re.search(r'(https?://[^\s"\'<>]*/go/[A-Za-z0-9%-]+/\d+[^\s"\'<>]*)', content, re.I)
                if m_go: return m_go.group(1), "embedded:SF_go", detected_ats
                return website, f"embedded:{detected_ats}", detected_ats
            for pat, ats_name in ATS_DOMAINS.items():
                if pat in content.lower():
                    m = re.search(r'https?://[^\s"\'<>]*' + re.escape(pat) + r'[^\s"\'<>]*', content, re.I)
                    if m:
                        candidate = m.group(0)
                        if not _is_blocked_career_url(candidate):
                            return candidate, f"embedded:{ats_name}", ats_name
            iframe_src = await _resolve_iframe_ats(page, content, website)
            if iframe_src and not _is_blocked_career_url(iframe_src):
                ats_name = is_ats(iframe_src)
                return iframe_src, f"iframe:{ats_name or 'unknown'}", ats_name
            best_score, best_url, best_strat = 0, None, None
            for link in await page.query_selector_all("a[href]"):
                try:
                    href = (await link.get_attribute("href")) or ""
                    text = clean(await link.inner_text(), 80).lower()
                    if not href or href.startswith(("#","mailto:","tel:")): continue
                    full = href if href.startswith("http") else urljoin(website, href)
                    if _is_blocked_career_url(full): continue
                    ats_name = is_ats(full)
                    if ats_name:
                        return await _resolve_to_jobs_url(page, full, f"ats_link:{ats_name}", ats_name)
                    ok, reason = is_career_link(href, text)
                    if ok:
                        score = (10 if reason == "definite" else 5) + (3 if base_domain in full else 0)
                        if score > best_score:
                            best_score, best_url, best_strat = score, full, f"playwright:{reason}"
                except: continue
            if best_url and best_score >= 5 and not _is_blocked_career_url(best_url):
                return await _resolve_to_jobs_url(page, best_url, best_strat, None)
        except: pass
    for sm in [base + "/sitemap.xml", base + "/sitemap_index.xml"]:
        r = simple_get(sm)
        if r:
            hits = re.findall(r'<loc>(https?://[^<]*(?:career|job|vacanc|hiring)[^<]*)</loc>', r.text, re.I)
            valid = [h for h in hits if not _is_blocked_career_url(h)]
            if valid:
                return await _resolve_to_jobs_url(page, valid[0], "sitemap", is_ats(valid[0]))
    try:
        r = simple_get(f"https://html.duckduckgo.com/html/?q={quote_plus(f'site:{base_domain} careers jobs')}")
        if r:
            for el in BeautifulSoup(r.text,"html.parser").select(".result__url"):
                u = el.get_text(strip=True)
                if not u.startswith("http"): u = "https://" + u
                if base_domain in u and any(kw in u.lower() for kw in ["career","job","vacanc"]) and not _is_blocked_career_url(u):
                    return await _resolve_to_jobs_url(page, u, "duckduckgo", is_ats(u))
    except: pass
    return None, "not_found", None

# ═════════════════════════════════════════════════════════════════════════════
#  SF BOARD SCRAPER
# ═════════════════════════════════════════════════════════════════════════════

async def _scrape_sf_board(page, careers_url, company, website, industry):
    stubs = []; root = careers_url.rstrip("/")
    listing_url = None; listing_html = None; best_score = -1
    candidates = [careers_url] + [root + s for s in SF_LISTING_SUFFIXES] + [root + q for q in SF_SEARCH_PARAMS]
    for candidate in candidates:
        r = simple_get(candidate, timeout=15)
        if not r: continue
        html = r.text; soup_c = BeautifulSoup(html,"html.parser")
        table_stubs = _scrape_sf_listing_table(soup_c, r.url)
        if table_stubs:
            print(f"        SF table found at: {r.url}  ({len(table_stubs)} rows)")
            listing_url = r.url; listing_html = html; stubs = table_stubs; break
        score = _job_signal_count(html)
        if score > best_score:
            best_score = score; listing_url = r.url; listing_html = html
    if not stubs and _PLAYWRIGHT_OK:
        try:
            await page.goto(listing_url or careers_url, timeout=30000, wait_until="networkidle")
            await page.wait_for_timeout(2500)
            for _ in range(3):
                await page.evaluate("window.scrollBy(0, 800)")
                await page.wait_for_timeout(600)
            listing_html = await page.content()
            final_url    = page.url
            soup_js      = BeautifulSoup(listing_html, "html.parser")
            table_stubs  = _scrape_sf_listing_table(soup_js, final_url)
            if table_stubs:
                print(f"        SF table found via Playwright ({len(table_stubs)} rows)")
                stubs = table_stubs; listing_url = final_url
        except Exception as e:
            print(f"        SF Playwright render error: {e}")
    if not stubs and listing_html:
        soup_fb = BeautifulSoup(listing_html, "html.parser")
        seen_fb = set()
        for a in soup_fb.find_all("a", href=True):
            href = a.get("href","")
            if not href or href.startswith(("#","mailto:","tel:","javascript")): continue
            full = href if href.startswith("http") else urljoin(listing_url or careers_url, href)
            if full in seen_fb: continue
            seen_fb.add(full)
            if not _is_detail_page_url(full): continue
            text = a.get_text(strip=True)
            if not text or len(text) < 3:
                parts = urlparse(full).path.rstrip("/").split("/")
                slug = (parts[-2] if (len(parts) >= 2 and parts[-1].isdigit()) else parts[-1])
                text = re.sub(r"[-_]", " ", slug).strip()
                text = re.sub(r"\s+\d{5,}$", "", text).strip()
            if text and len(text) >= 3:
                stubs.append({"url": full, "title": text, "location": "", "date_posted": ""})
        print(f"        SF fallback link collect: {len(stubs)} candidates")
    return stubs

# ═════════════════════════════════════════════════════════════════════════════
#  DEEP JOB DETAIL SCRAPER  (website version, from v5 pipeline)
# ═════════════════════════════════════════════════════════════════════════════

async def scrape_job_detail_from_website(page, job_url, company, website, industry,
                                          careers_url, source,
                                          prefill_title="", prefill_location="",
                                          prefill_date=""):
    html = None; final_url = job_url
    vprint(f"  Fetching detail: {job_url[:80]}", indent=3)
    r = simple_get(job_url, timeout=12)
    if r:
        html = r.text; final_url = r.url
        vprint(f"  HTTP OK ({len(html):,} chars)", indent=4)
    elif _PLAYWRIGHT_OK:
        vprint(f"  Plain HTTP failed — using Playwright…", indent=4)
        try:
            await page.goto(job_url, timeout=25000, wait_until="domcontentloaded")
            await page.wait_for_timeout(1800)
            html = await page.content(); final_url = page.url
            vprint(f"  Playwright OK ({len(html):,} chars)", indent=4)
        except Exception as e:
            log.warning(f"Detail fetch failed {job_url}: {e}")
            _stats["detail_failures"] += 1
            return None

    _stats["detail_fetches"] += 1
    if not html: return None

    soup = BeautifulSoup(html, "html.parser")
    ld   = _extract_json_ld(soup)
    sections = _parse_headed_sections(soup)

    for tag in soup.select("nav,footer,header,script,style,iframe,.cookie"):
        tag.decompose()
    full_text = soup.get_text(separator="\n")

    # title
    title = (ld.get("title") or prefill_title or _extract_meta(soup, "og:title")
             or clean(soup.title.get_text() if soup.title else "")
             or clean(soup.select_one("h1").get_text() if soup.select_one("h1") else ""))
    if title and "|" in title: title = title.split("|")[0].strip()
    if title and "-" in title and len(title) > 60: title = title.split("-")[0].strip()

    # department / field
    department = _sanitize_field(
        ld.get("field")
        or _find_text_near_label(soup, "Department","Team","Function","Division")
        or ""
    )
    if not department:
        h2_tags = soup.find_all("h2")
        if h2_tags:
            first_h2 = h2_tags[0].get_text(strip=True)
            if (first_h2 and len(first_h2) < 80
                    and not re.search(r"(apply|job search|search|opportunities|alert|description)", first_h2, re.I)
                    and _sanitize_field(first_h2)):
                department = first_h2

    # job type
    job_type = _sanitize_field(
        ld.get("job_type")
        or _find_text_near_label(soup, "Employment Type","Job Type","Contract Type","Type")
        or _pick_section(sections, "employment type","job type")
        or ""
    )
    if not job_type:
        for jt in ["Full-time","Part-time","Contract","Internship","Freelance","Temporary","Permanent"]:
            if re.search(r"\b" + jt.split("-")[0] + r"\b", full_text, re.I):
                job_type = jt; break

    # location
    location = (ld.get("location") or prefill_location
                or _find_text_near_label(soup, "Location","Job Location","City")
                or "")
    if not location:
        m = LOCATION_PATTERN.search(full_text)
        if m: location = m.group(1)
    location = location or "Saudi Arabia"

    # description
    description = ld.get("description","")
    if not description:
        description = _pick_section(sections,
            "key functional","responsibilities","job purpose","about the role",
            "what you'll do","role overview","job summary","overview","purpose")
    if not description:
        for sel in ["[class*='description']","[class*='job-desc']","[class*='content']",
                    "#job-description","article","main","[class*='detail']"]:
            el = soup.select_one(sel)
            if el and len(el.get_text()) > 100:
                description = clean(el.get_text(), 2000); break
    if not description and len(full_text) > 200:
        description = clean(full_text, 2000)

    # qualifications
    qualifications = (
        ld.get("qualifications")
        or _pick_section(sections, "qualif","requirement","education","what you need")
        or _find_text_near_label(soup, "Qualifications","Requirements","Education","Degree")
        or ""
    )
    if not qualifications:
        qualifications = _extract_bold_field(soup, "Education","Qualifications","Requirements")
    if _GARBAGE_FIELD_RE.match(qualifications.strip()[:80]):
        qualifications = ""

    # experience
    experience = _sanitize_field(
        ld.get("experience")
        or _find_text_near_label(soup, "Experience","Years of Experience")
        or _pick_section(sections, "experience")
        or ""
    )
    if not experience:
        experience = _sanitize_field(_extract_bold_field(soup, "Experience","Years of Experience"))
    if not experience:
        m = EXPERIENCE_RE.search(full_text)
        if m: experience = m.group(0)

    # field
    field = _sanitize_field(
        ld.get("field") or department
        or _find_text_near_label(soup, "Field","Category","Department","Function")
        or industry
    )

    # dates
    date_posted = ld.get("date_posted","") or prefill_date
    if not date_posted:
        m = DATE_POSTED_RE.search(full_text)
        if m: date_posted = next((g for g in m.groups() if g), "")
    if not date_posted:
        time_el = soup.find("time")
        if time_el: date_posted = time_el.get("datetime") or time_el.get_text(strip=True)
    if date_posted and _GARBAGE_FIELD_RE.match(date_posted.strip()):
        date_posted = ""

    deadline = ld.get("deadline","")
    if not deadline:
        m = DEADLINE_RE.search(full_text)
        if m: deadline = next((g for g in m.groups() if g), "")
    if deadline and _GARBAGE_FIELD_RE.match(deadline.strip()):
        deadline = ""

    estimated_deadline = deadline or _estimate_deadline(date_posted)
    if estimated_deadline and not re.match(r"\d{4}-\d{2}-\d{2}", estimated_deadline.strip()):
        estimated_deadline = ""

    # salary
    salary_range = (
        ld.get("salary_range")
        or _find_text_near_label(soup, "Salary","Compensation")
        or ""
    )
    if not salary_range:
        m = SALARY_RE.search(full_text)
        if m: salary_range = m.group(0)

    # company metadata
    company_logo   = ld.get("company_logo","") or _extract_logo(soup, final_url)
    company_type   = _sanitize_field(ld.get("company_type","") or _find_text_near_label(soup,"Company Type","Organization Type") or "")
    if company_type and company_type.lower() in ("organization","legalservice","thing"): company_type = ""
    company_address = _find_text_near_label(soup, "Address","Headquarters","Head Office")
    company_founded = _sanitize_field(_find_text_near_label(soup, "Founded","Established","Year Founded"))
    company_details = _pick_section(sections, "about the company","about us","company overview","who we are")

    # apply URL
    apply_url = final_url
    apply_patterns = ["talentcommunity/apply","/apply/","?action=apply","/application/"]
    for a in soup.find_all("a", href=True):
        href = a.get("href",""); text_a = a.get_text(strip=True).lower()
        if any(p in href.lower() for p in apply_patterns):
            apply_url = href if href.startswith("http") else urljoin(final_url, href); break
        if "apply" in text_a and href and href not in ("#","javascript:void(0)"):
            candidate = href if href.startswith("http") else urljoin(final_url, href)
            if candidate != final_url: apply_url = candidate

    if salary_range: _stats["jobs_with_salary"] += 1
    if description:  _stats["jobs_with_description"] += 1

    return make_job(
        company=company, website=website, industry=industry,
        careers_url=careers_url, source=source + "+website_detail",
        title=title, location=location, job_type=job_type,
        department=department, apply_url=apply_url,
        qualifications=qualifications, experience=experience, field=field,
        date_posted=date_posted, deadline=deadline, description=description,
        company_logo=company_logo, company_founded=company_founded,
        company_type=company_type, company_address=company_address,
        company_details=company_details,
        estimated_deadline=estimated_deadline, salary_range=salary_range,
    )

# ═════════════════════════════════════════════════════════════════════════════
#  ATS JOB EXTRACTOR  (uses native APIs where possible)
# ═════════════════════════════════════════════════════════════════════════════

async def extract_jobs_from_website(page, careers_url, company, website, industry, ats_name):
    stub_jobs = []

    if ats_name == "Greenhouse" or "greenhouse.io" in careers_url:
        m = re.search(r"greenhouse\.io/([^/?#]+)", careers_url)
        if m:
            r = simple_get(f"https://boards.greenhouse.io/{m.group(1)}/jobs.json")
            if r:
                try:
                    for j in r.json().get("jobs",[]):
                        stub_jobs.append({"url": j.get("absolute_url",""), "title": j.get("title","")})
                except: pass

    if ats_name == "Lever" or "lever.co" in careers_url:
        m = re.search(r"lever\.co/([^/?#]+)", careers_url)
        if m:
            r = simple_get(f"https://api.lever.co/v0/postings/{m.group(1)}?mode=json")
            if r:
                try:
                    for j in r.json():
                        stub_jobs.append({"url": j.get("hostedUrl",""), "title": j.get("text","")})
                except: pass

    if ats_name == "Ashby" or "ashbyhq.com" in careers_url:
        m = re.search(r"ashbyhq\.com/([^/?#]+)", careers_url)
        if m:
            r = simple_get(
                f"https://api.ashbyhq.com/posting-public/job-board/all"
                f"?organizationHostedJobsPageName={m.group(1)}"
            )
            if r:
                try:
                    for j in r.json().get("jobPostings",[]):
                        stub_jobs.append({"url": j.get("jobUrl",""), "title": j.get("title","")})
                except: pass

    if ats_name == "SmartRecruiters" or "smartrecruiters.com" in careers_url:
        m = re.search(r"smartrecruiters\.com/([^/?#]+)", careers_url)
        if m:
            r = simple_get(f"https://api.smartrecruiters.com/v1/companies/{m.group(1)}/postings")
            if r:
                try:
                    for j in r.json().get("content",[]):
                        stub_jobs.append({
                            "url": f"https://jobs.smartrecruiters.com/{m.group(1)}/{j.get('id','')}",
                            "title": j.get("name",""),
                        })
                except: pass

    if ats_name == "Workday" or "myworkdayjobs.com" in careers_url:
        m = re.search(r"(https?://[^/]+myworkdayjobs\.com/[^/?#]+)", careers_url)
        if m:
            r = simple_get(m.group(1) + "/jobs?format=json")
            if r:
                try:
                    for j in r.json().get("jobPostings",[]):
                        stub_jobs.append({
                            "url": m.group(1) + "/job/" + j.get("externalPath",""),
                            "title": j.get("title",""),
                        })
                except: pass

    if ats_name == "Wynt" or "wynt.ai" in careers_url:
        m = re.search(r"wynt\.ai/([^/?#]+)", careers_url)
        if m:
            org_slug = m.group(1)
            r = simple_get(f"https://apply.wynt.ai/api/v1/jobs?organization={org_slug}&limit=200")
            if r:
                try:
                    data = r.json()
                    job_list = data if isinstance(data, list) else data.get("results", data.get("jobs",[]))
                    for j in job_list:
                        jid = j.get("id","") or j.get("slug","")
                        title = j.get("title","") or j.get("name","")
                        jurl = (j.get("url","") or j.get("apply_url","")
                                or f"https://apply.wynt.ai/{org_slug}/jobs/{jid}")
                        stub_jobs.append({"url": jurl, "title": title, "_wynt_data": j})
                except: pass

    if not stub_jobs and ("successfactors" in careers_url.lower()
                           or "/go/job-search/" in careers_url.lower()
                           or "talentcommunity" in careers_url.lower()):
        print(f"        Using SuccessFactors board scraper…")
        sf_stubs = await _scrape_sf_board(page, careers_url, company, website, industry)
        stub_jobs.extend(sf_stubs)
        print(f"        SF scraper: {len(stub_jobs)} job stubs")

    if not stub_jobs and _PLAYWRIGHT_OK:
        try:
            await page.goto(careers_url, timeout=30000, wait_until="networkidle")
            await page.wait_for_timeout(2500)
            for _ in range(3):
                await page.evaluate("window.scrollBy(0, 800)")
                await page.wait_for_timeout(700)
            html = await page.content()
            soup = BeautifulSoup(html, "html.parser")
            for tag in soup.select("nav,footer,header,script,style"):
                tag.decompose()
            ats_detected = is_ats_from_html(html)
            if ats_detected == "SuccessFactors":
                table_stubs = _scrape_sf_listing_table(soup, careers_url)
                if table_stubs:
                    stub_jobs.extend(table_stubs)
            if not stub_jobs:
                career_domain = get_domain(careers_url)
                job_links = _collect_job_links(soup, careers_url, career_domain)
                for jurl, jtext in job_links[:200]:
                    stub_jobs.append({"url": jurl, "title": jtext, "location": "", "date_posted": ""})
            if not stub_jobs:
                for next_url in _pagination(soup, careers_url)[:4]:
                    await page.goto(next_url, timeout=20000, wait_until="networkidle")
                    await page.wait_for_timeout(1500)
                    nsoup = BeautifulSoup(await page.content(), "html.parser")
                    extra = _collect_job_links(nsoup, next_url, career_domain)
                    for jurl, jtext in extra:
                        stub_jobs.append({"url": jurl, "title": jtext, "location": "", "date_posted": ""})
        except Exception as e:
            print(f"        DOM error: {e}")

    if not stub_jobs:
        # plain requests fallback
        r = simple_get(careers_url)
        if r:
            soup = BeautifulSoup(r.text, "html.parser")
            career_domain = get_domain(careers_url)
            job_links = _collect_job_links(soup, careers_url, career_domain)
            for jurl, jtext in job_links[:200]:
                stub_jobs.append({"url": jurl, "title": jtext, "location": "", "date_posted": ""})

    if not stub_jobs:
        vprint(f"  0 job URLs found on listing page", indent=2)
        return []

    stub_jobs = [s for s in stub_jobs if _is_detail_page_url(s.get("url",""))]
    if not stub_jobs:
        vprint(f"  0 valid detail URLs after validation", indent=2)
        return []

    vprint(f"  Scraping {len(stub_jobs)} detail pages…", indent=2)
    jobs = []; seen_urls = set()

    for stub in stub_jobs[:200]:
        jurl = stub.get("url","")
        if not jurl or jurl in seen_urls: continue
        seen_urls.add(jurl)

        wynt_raw = stub.get("_wynt_data")
        if wynt_raw:
            loc_raw = wynt_raw.get("location", {})
            location = (loc_raw.get("city","") + " " + loc_raw.get("country","")).strip() \
                       if isinstance(loc_raw, dict) else str(loc_raw)
            sal = wynt_raw.get("salary", {})
            salary_range = ""
            if isinstance(sal, dict) and (sal.get("min") or sal.get("max")):
                salary_range = f"{sal.get('currency','')} {sal.get('min','')}–{sal.get('max','')}".strip()
            job = make_job(
                company=company, website=website, industry=industry,
                careers_url=careers_url, source="Wynt API",
                title=stub.get("title",""), location=location or "Saudi Arabia",
                job_type=wynt_raw.get("employment_type",""),
                department=wynt_raw.get("department",""),
                apply_url=jurl,
                qualifications=wynt_raw.get("qualifications",""),
                experience=wynt_raw.get("experience",""),
                field=wynt_raw.get("category",""),
                date_posted=wynt_raw.get("created_at",""),
                deadline=wynt_raw.get("deadline",""),
                description=wynt_raw.get("description",""),
                salary_range=salary_range,
            )
            _JOB_COUNTER[0] += 1; _stats["jobs_total"] += 1
            jobs.append(job)
            continue

        detail = await scrape_job_detail_from_website(
            page, jurl, company, website, industry, careers_url, source="website",
            prefill_title=stub.get("title",""),
            prefill_location=stub.get("location",""),
            prefill_date=stub.get("date_posted",""),
        )
        if detail:
            _JOB_COUNTER[0] += 1; _stats["jobs_total"] += 1
            if detail.get("Job Title"): _stats["jobs_with_title"] += 1
            jobs.append(detail)
        else:
            fallback = make_job(
                company=company, website=website, industry=industry,
                careers_url=careers_url, source="listing_only",
                title=stub.get("title",""), location=stub.get("location",""),
                date_posted=stub.get("date_posted",""), apply_url=jurl,
            )
            _JOB_COUNTER[0] += 1; _stats["jobs_total"] += 1
            jobs.append(fallback)

        await asyncio.sleep(random.uniform(0.4, 1.0))

    vprint(f"  {len(jobs)} jobs scraped from website", indent=2)
    return jobs

# ═════════════════════════════════════════════════════════════════════════════
#  LINKEDIN JOB DETAIL SCRAPER  (v4 logic + website enrichment)
# ═════════════════════════════════════════════════════════════════════════════

async def scrape_job_details_async(page, job_url: str) -> dict | None:
    """Full job scrape: LinkedIn page → company page → website career crawl."""
    log.info(f"Scraping LinkedIn job: {job_url}")
    try:
        resp = requests.get(job_url, headers=_next_headers(), timeout=20)
        if resp.status_code != 200: return None
        html = resp.text
        soup = BeautifulSoup(html, "html.parser")
    except Exception as e:
        log.warning(f"Job fetch failed: {e}"); return None

    def sel_text(*selectors):
        for s in selectors:
            el = soup.select_one(s)
            if el:
                t = el.get_text(strip=True)
                if t: return t
        return ""

    title = sel_text(".top-card-layout__title","h1.topcard__title",
                     ".job-details-jobs-unified-top-card__job-title","h1")
    company_name_fallback = sel_text(".topcard__org-name-link",
                                     ".job-details-jobs-unified-top-card__company-name",
                                     ".topcard__flavor")
    company_url_el = (soup.select_one(".topcard__org-name-link") or
                      soup.select_one(".job-details-jobs-unified-top-card__company-name a"))
    company_url = company_url_el.get("href","") if company_url_el else ""
    location = sel_text(".topcard__flavor--bullet",
                        ".job-details-jobs-unified-top-card__bullet")
    workplace_type = get_workplace_type(soup)

    time_el = soup.find("time")
    raw_posted = (time_el.get("datetime","") if time_el else "") or \
                 sel_text(".posted-time-ago__text",".job-details-jobs-unified-top-card__posted-date")
    posted_date = resolve_posted_date(raw_posted)

    raw_desc = sel_text(".show-more-less-html__markup",".description__text")
    description = clean_description(raw_desc)

    salary = ""
    for s in [".compensation__salary",".salary","[class*='salary']","[class*='compensation']"]:
        el = soup.select_one(s)
        if el: salary = el.get_text(strip=True); break
    if not salary:
        for chip in soup.select(".job-details-jobs-unified-top-card__job-insight"):
            t = chip.get_text(strip=True)
            if re.search(r"\$|MUR|Rs\.?|SAR|salary|/yr|/hour|per month", t, re.I):
                salary = t; break

    raw_job_type = get_job_criteria(soup, "Employment type") or workplace_type
    job_type = raw_job_type or "Full-time"
    linkedin_function = get_job_criteria(soup, "Job function")
    linkedin_industry = get_job_criteria(soup, "Industries")

    real_deadline      = parse_deadline(soup)
    estimated_deadline = estimate_deadline_from_posted(posted_date) if not real_deadline else ""
    effective_deadline = real_deadline or estimated_deadline

    # ── Layer 1: extract company data from job page ───────────────────────────
    job_page_co = extract_company_from_job_page(html, soup)
    ld          = _parse_jsonld(html)

    # ── Layer 2: LinkedIn company page ───────────────────────────────────────
    time.sleep(0.5)
    company = scrape_company_details(company_url)

    def _first(*vals):
        for v in vals:
            if v and str(v).strip(): return str(v).strip()
        return ""

    merged_name     = _first(company.get("name"),     job_page_co.get("company_name"),  company_name_fallback)
    merged_industry = _first(company.get("industry"), job_page_co.get("company_industry"), linkedin_industry)
    merged_logo     = _first(company.get("logo"),     job_page_co.get("company_logo"))
    merged_website  = _first(company.get("website"),  job_page_co.get("company_website"))
    merged_hq       = _first(company.get("headquarters"), job_page_co.get("company_address"))
    merged_founded  = _first(company.get("founded"),  job_page_co.get("company_founded"))
    merged_type     = _first(company.get("type"))
    merged_about    = _first(company.get("about"),    job_page_co.get("company_about"))

    # ── Layer 3 (NEW): find & crawl the company's career page ────────────────
    website_jobs = []
    if merged_website and not should_skip_crawl(merged_website):
        log.info(f"Finding career page for: {merged_website}")
        careers_url, strategy, ats_name = await find_career_page(page, merged_website)
        if careers_url and not _is_blocked_career_url(careers_url):
            log.info(f"Career page found ({strategy}, ATS={ats_name}): {careers_url}")
            _stats["companies_with_careers"] += 1
            if ats_name:
                _stats["ats_hits"][ats_name] = _stats["ats_hits"].get(ats_name, 0) + 1
            website_jobs = await extract_jobs_from_website(
                page, careers_url, merged_name, merged_website,
                merged_industry, ats_name,
            )
        else:
            _stats["companies_no_careers"] += 1
    else:
        careers_url = ""; ats_name = None

    # ── Layer 4: determine apply link for THIS specific LinkedIn job ──────────
    # First try website jobs to find a matching role
    apply_link = ""
    if website_jobs:
        # Try to match by title
        for wj in website_jobs:
            if _title_similarity(title, wj.get("Job Title","")) >= 0.4:
                apply_link = wj.get("Application","") or wj.get("Job URL","")
                break
        if not apply_link:
            apply_link = website_jobs[0].get("Application","") or website_jobs[0].get("Job URL","")

    # Fallback to v4 extraction methods
    if not apply_link:
        desc_text = ""
        for sel in [".show-more-less-html__markup",".description__text"]:
            el = soup.select_one(sel)
            if el: desc_text = el.get_text(); break

        if ld.get("apply_url") and not is_bad_url(ld["apply_url"]):
            apply_link = ld["apply_url"]
        if not apply_link:
            btn = follow_linkedin_apply_button(soup, job_url)
            if btn: apply_link = btn
        if not apply_link:
            for script in soup.find_all("script"):
                txt = script.string or ""
                for pat in [r'"applyStartUrl"\s*:\s*"([^"]+)"',
                            r'"applicationUrl"\s*:\s*"([^"]+)"']:
                    m = re.search(pat, txt)
                    if m:
                        cand = decode_html_entities(m.group(1)).replace("\\","")
                        if cand.startswith("http") and not is_bad_url(cand):
                            apply_link = cand; break
        if not apply_link:
            desc_el = soup.select_one(".show-more-less-html__markup") or soup.select_one(".description__text")
            if desc_el:
                for a in desc_el.find_all("a", href=True):
                    h = a.get("href","")
                    if not is_bad_url(h): apply_link = h; break
        if not apply_link:
            em = extract_email_from_text(desc_text)
            if em: apply_link = em
        if not apply_link and careers_url and not _is_blocked_career_url(careers_url):
            apply_link = careers_url
        if not apply_link and merged_website:
            apply_link = merged_website

    apply_link = clean_application_link(apply_link)

    job_field = linkedin_function or standardise_field("", title, description, merged_industry)
    qualifications = standardise_qualification("", description)
    experience     = standardise_experience(description)

    linkedin_job = {
        "Job Title":          title,
        "Job Type":           job_type,
        "Job Qualifications": qualifications,
        "Job Experience":     experience,
        "Job Location":       location,
        "Job Field":          job_field,
        "Date Posted":        posted_date,
        "Deadline":           effective_deadline,
        "Job Description":    description,
        "Application":        apply_link,
        "Company URL":        company_url,
        "Company Name":       merged_name,
        "Company Logo":       clean_logo_url(merged_logo),
        "Company Industry":   merged_industry,
        "Company Founded":    merged_founded,
        "Company Type":       merged_type,
        "Company Website":    merged_website,
        "Company Address":    merged_hq,
        "Company Details":    merged_about,
        "Job URL":            job_url,
        "Estimated Deadline": estimated_deadline,
        "Salary Range":       salary,
        "source":             "linkedin",
    }

    return linkedin_job, website_jobs

# ═════════════════════════════════════════════════════════════════════════════
#  VERBOSE PRINTER
# ═════════════════════════════════════════════════════════════════════════════

def print_job_verbose(job: dict, index: int, total: int):
    desc = job.get("Job Description","")
    desc_preview = (desc[:400] + " [...]") if len(desc) > 400 else desc
    desc_indented = "\n".join("   " + line for line in desc_preview.splitlines() if line.strip())
    apply = job.get("Application","")
    logo  = job.get("Company Logo","")
    print()
    print(C_DIVIDER())
    print(C_HEADER(f"  JOB {index}/{total}"))
    print(C_DIVIDER())
    print(f"  {C_LABEL('Title')}          : {C_VALUE(job.get('Job Title',''))}")
    print(f"  {C_LABEL('Job Type')}       : {job.get('Job Type','')}")
    print(f"  {C_LABEL('Field')}          : {job.get('Job Field','') or C_DIM('—')}")
    print(f"  {C_LABEL('Location')}       : {job.get('Job Location','') or C_DIM('—')}")
    print(f"  {C_LABEL('Experience')}     : {job.get('Job Experience','') or C_DIM('—')}")
    print(f"  {C_LABEL('Qualifications')} : {job.get('Job Qualifications','') or C_DIM('—')}")
    print(f"  {C_LABEL('Salary')}         : {job.get('Salary Range','') or C_DIM('—')}")
    print(f"  {C_LABEL('Date Posted')}    : {job.get('Date Posted','') or C_DIM('—')}")
    print(f"  {C_LABEL('Deadline')}       : {job.get('Deadline','') or C_DIM('—')}")
    print(f"  {C_LABEL('Apply Link')}     : {C_GREEN(apply) if apply else C_DIM('— not found —')}")
    print()
    print(f"  {C_BLUE('── COMPANY ──────────────────────────────────────────────')}")
    print(f"  {C_LABEL('Name')}           : {C_VALUE(job.get('Company Name','') or C_DIM('—'))}")
    print(f"  {C_LABEL('Industry')}       : {job.get('Company Industry','') or C_DIM('—')}")
    print(f"  {C_LABEL('Type')}           : {job.get('Company Type','') or C_DIM('—')}")
    print(f"  {C_LABEL('Founded')}        : {job.get('Company Founded','') or C_DIM('—')}")
    print(f"  {C_LABEL('Headquarters')}   : {job.get('Company Address','') or C_DIM('—')}")
    print(f"  {C_LABEL('Website')}        : {job.get('Company Website','') or C_DIM('—')}")
    print(f"  {C_LABEL('LinkedIn URL')}   : {job.get('Company URL','') or C_DIM('—')}")
    print(f"  {C_LABEL('Logo')}           : {logo if logo else C_DIM('— none —')}")
    about = job.get("Company Details","")
    if about:
        preview = (about[:200] + " [...]") if len(about) > 200 else about
        print(f"  {C_LABEL('About')}          : {preview}")
    print()
    print(f"  {C_BLUE('── DESCRIPTION PREVIEW ──────────────────────────────────')}")
    print(desc_indented if desc_indented else C_DIM("   — no description —"))
    print()
    print(f"  {C_LABEL('Job URL')}        : {C_DIM(job.get('Job URL',''))}")
    print(C_DIVIDER())

# ═════════════════════════════════════════════════════════════════════════════
#  CHECKPOINT / FLUSH
# ═════════════════════════════════════════════════════════════════════════════

def load_checkpoint():
    if Path(CHECKPOINT_FILE).exists():
        return json.loads(Path(CHECKPOINT_FILE).read_text())
    return {"processed_domains": [], "jobs_count": 0, "processed_li_urls": []}

def save_checkpoint(cp):
    Path(CHECKPOINT_FILE).write_text(json.dumps(cp, indent=2))

def flush_all(jobs: list):
    if not jobs: return
    df = pd.DataFrame(jobs)
    # Deduplicate
    dedup_cols = ["Job Title","Company Name","Job Location"]
    existing   = [c for c in dedup_cols if c in df.columns]
    if existing:
        df = df.drop_duplicates(subset=existing)
    # Save CSV
    df.to_csv(OUTPUT_CSV, index=False)
    # Save XLSX
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    headers = [
        "Job Title","Job Type","Job Qualifications","Job Experience",
        "Job Location","Job Field","Date Posted","Deadline",
        "Job Description","Application","Company URL","Company Name",
        "Company Logo","Company Industry","Company Founded","Company Type",
        "Company Website","Company Address","Company Details","Job URL",
        "Estimated Deadline","Salary Range",
    ]
    ws.append(headers)
    for _, row in df.iterrows():
        ws.append([row.get(h,"") for h in headers])
    wb.save(OUTPUT_XLSX)
    log.info(f"Flushed {len(df)} rows → {OUTPUT_XLSX} + {OUTPUT_CSV}")

# ═════════════════════════════════════════════════════════════════════════════
#  LINKEDIN URL COLLECTION  (unchanged from v4)
# ═════════════════════════════════════════════════════════════════════════════

def _build_guest_api_url(keyword: str, start: int) -> str:
    kw = quote_plus(keyword)
    return (
        "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search"
        f"?location=Saudi+Arabia&f_TPR=r604800&keywords={kw}&start={start}"
    )

def _collect_job_urls_from_cards(html: str, seen: set) -> list:
    found = []
    for raw_href in re.findall(r'href="(https?://[^"]*?/jobs/view/\d+[^"]*?)"', html):
        c = canonicalise_job_url(raw_href)
        if c and c not in seen: seen.add(c); found.append(c)
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup.find_all("a", href=True):
        href = tag["href"]
        if "/jobs/view/" not in href: continue
        if not href.startswith("http"): href = "https://www.linkedin.com" + href
        c = canonicalise_job_url(href)
        if c and c not in seen: seen.add(c); found.append(c)
    for sel in ["a.base-card__full-link","a.base-main-card__full-link",
                "a[data-tracking-control-name='public_jobs_jserp-name_click']",
                "a.job-card-list__title","a.job-card-container__link"]:
        for tag in soup.select(sel):
            href = tag.get("href","")
            if "/jobs/view/" not in href: continue
            if not href.startswith("http"): href = "https://www.linkedin.com" + href
            c = canonicalise_job_url(href)
            if c and c not in seen: seen.add(c); found.append(c)
    return found

def _fetch_guest_api_page(keyword: str, start: int, retries: int = 3) -> str | None:
    url = _build_guest_api_url(keyword, start)
    for attempt in range(retries):
        try:
            time.sleep(DELAY_S + attempt * 3)
            r = requests.get(url, headers=_next_headers(), allow_redirects=True, timeout=25)
            if r.status_code == 429:
                wait = 60 + attempt * 60
                print(C_RED(f"  Rate limited (429) — waiting {wait}s ..."))
                time.sleep(wait); continue
            if r.status_code in (400,403,999):
                log.warning(f"Blocked ({r.status_code}): {url}"); return None
            if r.status_code != 200:
                log.warning(f"HTTP {r.status_code}: {url}"); return None
            text = r.text.strip()
            if not text:
                log.info(f"Empty body (start={start}, kw='{keyword}') — end of results.")
                return None
            return text
        except Exception as e:
            log.warning(f"Guest API error (attempt {attempt+1}, kw='{keyword}'): {e}")
            time.sleep(3 + attempt * 3)
    return None

def _paginate_keyword(keyword: str, seen: set) -> list:
    urls = []; page = 0; empty_streak = 0
    label = keyword if keyword else "(all)"
    while True:
        if MAX_PAGES and page >= MAX_PAGES: break
        start = page * 25
        print(f"  {C_DIM(f'[{label}] page {page+1} (start={start}) ...')}", flush=True)
        html = _fetch_guest_api_page(keyword, start)
        if html is None: break
        new_urls = _collect_job_urls_from_cards(html, seen)
        log.info(f"[{label}] page {page+1}: {len(new_urls)} new (total seen={len(seen)})")
        if new_urls: urls.extend(new_urls); empty_streak = 0
        else:
            empty_streak += 1
            if empty_streak >= MAX_EMPTY_PAGES: break
        if start >= 975: break
        page += 1
        if page % 10 == 0:
            print(C_DIM("  Pausing 20s (every 10 pages) ..."))
            time.sleep(20)
    return urls

# ═════════════════════════════════════════════════════════════════════════════
#  MAIN ASYNC CRAWL
# ═════════════════════════════════════════════════════════════════════════════

async def crawl_async():
    global all_jobs, company_results

    start_time = time.time()
    print()
    print(C_HEADER("=" * 72))
    print(C_HEADER("  LINKEDIN JOB SCRAPER v5 — Saudi Arabia (with website crawl)"))
    print(C_HEADER("=" * 72))
    print(f"  Keywords    : {len(SEARCH_KEYWORDS)}")
    print(f"  Max pages   : {'unlimited' if not MAX_PAGES else MAX_PAGES} per keyword")
    print(f"  Job cap     : {'none' if not JOB_LIMIT else JOB_LIMIT}")
    print(f"  Playwright  : {'enabled' if _PLAYWRIGHT_OK else 'DISABLED (install playwright)'}")
    print(f"  Started     : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(C_HEADER("=" * 72))

    cp = load_checkpoint()
    if Path(OUTPUT_CSV).exists():
        all_jobs = pd.read_csv(OUTPUT_CSV).to_dict("records")
        print(f"  Resuming — {len(all_jobs)} jobs already saved")
    if Path(COMPANIES_FILE).exists():
        company_results = pd.read_csv(COMPANIES_FILE).to_dict("records")

    seen_urls   : set = set(cp.get("processed_li_urls", []))
    seen_content: set = set()
    all_job_urls: list = []

    # ── Phase 1: collect LinkedIn URLs ───────────────────────────────────────
    for qi, keyword in enumerate(SEARCH_KEYWORDS):
        label = keyword if keyword else "(all)"
        print()
        print(C_BLUE(f"┌─ Keyword {qi+1}/{len(SEARCH_KEYWORDS)}: '{label}' ─────────────────"))
        new_urls = _paginate_keyword(keyword, seen_urls)
        all_job_urls.extend(new_urls)
        print(C_BLUE(f"└─ Found {len(new_urls)} new jobs (running total: {len(all_job_urls)})"))
        if JOB_LIMIT and len(all_job_urls) >= JOB_LIMIT: break
        time.sleep(DELAY_S * 2)

    if JOB_LIMIT and len(all_job_urls) > JOB_LIMIT:
        all_job_urls = all_job_urls[:JOB_LIMIT]

    print()
    print(C_HEADER(f"  Total unique LinkedIn URLs collected: {len(all_job_urls)}"))
    print()

    # ── Phase 2: scrape jobs + crawl company websites ────────────────────────
    errors = 0

    if _PLAYWRIGHT_OK:
        async with async_playwright() as pw:
            browser = await pw.chromium.launch(
                headless=True,
                args=["--no-sandbox","--disable-dev-shm-usage","--disable-gpu"],
            )
            context = await browser.new_context(
                user_agent=HEADERS["User-Agent"],
                locale="en-US",
                viewport={"width": 1280, "height": 800},
            )
            page = await context.new_page()
            await page.route(
                "**/*.{png,jpg,jpeg,gif,svg,woff,woff2,ttf,eot,ico}",
                lambda r: r.abort(),
            )

            for j, url in enumerate(all_job_urls):
                print(f"\n{C_HEADER(f'>>> Scraping job {j+1}/{len(all_job_urls)} ...')}")
                log.info(f"URL: {url}")
                try:
                    result = await scrape_job_details_async(page, url)
                    if result:
                        linkedin_job, website_jobs = result
                        if linkedin_job and linkedin_job.get("Job Title"):
                            fp = (
                                (linkedin_job.get("Job Title","")).lower().strip(),
                                (linkedin_job.get("Company Name","")).lower().strip(),
                                (linkedin_job.get("Job Location","")).lower().strip(),
                            )
                            if fp not in seen_content:
                                seen_content.add(fp)
                                all_jobs.append(linkedin_job)
                                print_job_verbose(linkedin_job, j+1, len(all_job_urls))
                            else:
                                print(C_DIM(f"  Duplicate LinkedIn job — skipped"))

                        # Also add any additional jobs found on the company website
                        for wj in website_jobs:
                            wfp = (
                                (wj.get("Job Title","")).lower().strip(),
                                (wj.get("Company Name","")).lower().strip(),
                                (wj.get("Job Location","")).lower().strip(),
                            )
                            if wfp not in seen_content:
                                seen_content.add(wfp)
                                all_jobs.append(wj)
                                print(C_GREEN(f"  + Website job: {wj.get('Job Title','')} @ {wj.get('Company Name','')}"))
                    else:
                        print(C_RED("  ✗  No title found — skipped"))
                except Exception as e:
                    errors += 1
                    print(C_RED(f"  ✗  ERROR: {e}"))
                    log.warning(f"Job error: {e}")

                cp.setdefault("processed_li_urls",[]).append(url)
                time.sleep(DELAY_S)
                if len(all_jobs) % 50 == 0 and len(all_jobs) > 0:
                    flush_all(all_jobs)
                    save_checkpoint(cp)

            await browser.close()
    else:
        # Playwright not available: synchronous fallback (website crawl disabled)
        print(C_RED("  WARNING: Playwright not available. Website crawl disabled."))
        print(C_RED("  Install: pip install playwright && playwright install chromium"))
        for j, url in enumerate(all_job_urls):
            print(f"\n{C_HEADER(f'>>> Scraping job {j+1}/{len(all_job_urls)} ...')}")
            try:
                # synchronous LinkedIn-only scrape
                resp = requests.get(url, headers=_next_headers(), timeout=20)
                if resp.status_code != 200: continue
                html = resp.text
                soup = BeautifulSoup(html, "html.parser")
                def sel(*ss):
                    for s in ss:
                        el = soup.select_one(s)
                        if el:
                            t = el.get_text(strip=True)
                            if t: return t
                    return ""
                title = sel(".top-card-layout__title","h1.topcard__title","h1")
                if not title: continue
                company_name = sel(".topcard__org-name-link",".job-details-jobs-unified-top-card__company-name")
                location = sel(".topcard__flavor--bullet",".job-details-jobs-unified-top-card__bullet")
                raw_desc = sel(".show-more-less-html__markup",".description__text")
                description = clean_description(raw_desc)
                time_el = soup.find("time")
                raw_posted = (time_el.get("datetime","") if time_el else "")
                posted_date = resolve_posted_date(raw_posted)
                ld = _parse_jsonld(html)
                job_page_co = extract_company_from_job_page(html, soup)
                apply_link = (follow_linkedin_apply_button(soup, url)
                              or ld.get("apply_url","")
                              or job_page_co.get("company_website","")
                              or "")
                job = make_job(
                    company=company_name, website=job_page_co.get("company_website",""),
                    industry=job_page_co.get("company_industry",""),
                    careers_url="", source="linkedin_only",
                    title=title, location=location,
                    apply_url=clean_application_link(apply_link),
                    description=description, date_posted=posted_date,
                    company_logo=clean_logo_url(job_page_co.get("company_logo","")),
                )
                fp = (title.lower().strip(), company_name.lower().strip(), location.lower().strip())
                if fp not in seen_content:
                    seen_content.add(fp)
                    all_jobs.append(job)
                    print_job_verbose(job, j+1, len(all_job_urls))
            except Exception as e:
                errors += 1
                print(C_RED(f"  ✗  ERROR: {e}"))
            time.sleep(DELAY_S)

    flush_all(all_jobs)
    save_checkpoint(cp)

    # ── Final summary ─────────────────────────────────────────────────────────
    mins = round((time.time() - start_time) / 60, 1)
    print()
    print(C_HEADER("=" * 72))
    print(C_HEADER("  SCRAPE COMPLETE"))
    print(C_HEADER("=" * 72))
    print(f"  {C_LABEL('Total jobs')}      : {C_GREEN(str(len(all_jobs)))}")
    print(f"  {C_LABEL('Errors')}          : {C_RED(str(errors)) if errors else '0'}")
    print(f"  {C_LABEL('Duration')}        : ~{mins} min")
    print(f"  {C_LABEL('XLSX output')}     : {OUTPUT_XLSX}")
    print(f"  {C_LABEL('CSV output')}      : {OUTPUT_CSV}")

    if all_jobs:
        from collections import Counter
        fields = Counter(j.get("Job Field") or "Unknown" for j in all_jobs)
        print(f"\n  {C_LABEL('Jobs by field:')}")
        for field, count in fields.most_common(10):
            print(f"    {field:<35} {'█'*min(count,40)} {count}")

        with_apply = sum(1 for j in all_jobs if j.get("Application"))
        with_email = sum(1 for j in all_jobs if "@" in (j.get("Application") or ""))
        with_url   = with_apply - with_email
        no_apply   = len(all_jobs) - with_apply
        print(f"\n  {C_LABEL('Application links:')}")
        print(f"    URL found    : {with_url}")
        print(f"    Email found  : {with_email}")
        print(f"    Not found    : {no_apply}")

        sources = Counter(j.get("source","?") for j in all_jobs)
        print(f"\n  {C_LABEL('Jobs by source:')}")
        for src, cnt in sources.most_common():
            print(f"    {src:<35} {cnt}")

        if _stats["ats_hits"]:
            print(f"\n  {C_LABEL('ATS platforms detected:')}")
            for ats, cnt in sorted(_stats["ats_hits"].items(), key=lambda x: -x[1]):
                print(f"    {ats:<22} {cnt}")

    print(C_HEADER("=" * 72))

# ═════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════

def main():
    asyncio.run(crawl_async())

if __name__ == "__main__":
    main()
