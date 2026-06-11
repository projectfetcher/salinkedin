"""
Saudi Jobs Pipeline — Unified v1
═══════════════════════════════════════════════════════════════
Two complementary scrapers running in the same process, writing
to ONE deduplicated output file:

  SOURCE A — LinkedIn guest API
    Paginates 245 keyword × Saudi Arabia search queries.
    Scrapes each job detail page for full metadata.

  SOURCE B — Wikipedia company discovery → career page scraper
    Discovers Saudi companies from Wikipedia category/list pages.
    Finds each company's career page and scrapes jobs directly.

Deduplication (two layers):
  1. URL-level  : canonical /jobs/view/{id}/ for LinkedIn URLs;
                  exact apply_url for Wikipedia-sourced jobs.
  2. Content-level: (title, company, location) fingerprint shared
                  across BOTH sources.  When a job is found in
                  both pipelines the richer record wins and blank
                  fields are back-filled from the weaker one.

Output: saudi_jobs.xlsx  (22 columns, same schema as before)

Requirements:
    pip install requests beautifulsoup4 openpyxl pandas
                playwright nest_asyncio tqdm
    playwright install chromium
"""

# ── stdlib ────────────────────────────────────────────────────
import asyncio
import base64
import csv
import importlib
import json
import logging
import random
import re
import site
import subprocess
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import (quote_plus, unquote, urljoin, urlparse)

# ── install deps (Colab-safe) ─────────────────────────────────
print("📦  Installing / verifying dependencies…")
subprocess.run(["apt-get", "install", "-y",
    "libatk1.0-0", "libatk-bridge2.0-0", "libcups2", "libdrm2",
    "libxkbcommon0", "libxcomposite1", "libxdamage1", "libxfixes3",
    "libxrandr2", "libgbm1", "libasound2"], capture_output=True)
subprocess.run([sys.executable, "-m", "pip", "install",
    "playwright", "pandas", "beautifulsoup4",
    "requests", "nest_asyncio", "tqdm", "openpyxl", "-q"], check=True)
subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"],
    capture_output=True)
importlib.invalidate_caches()
for _sp in site.getsitepackages():
    if _sp not in sys.path:
        sys.path.insert(0, _sp)

import nest_asyncio
import openpyxl
import pandas as pd
import requests
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright
from tqdm import tqdm

nest_asyncio.apply()
print("✅  All imports OK\n")

# ═══════════════════════════════════════════════════════════════
# ▌ LOGGING
# ═══════════════════════════════════════════════════════════════
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

import sys as _sys
_USE_COLOUR = _sys.stdout.isatty()
def _c(code, t): return f"\033[{code}m{t}\033[0m" if _USE_COLOUR else t
C_H  = lambda t: _c("1;36", t)
C_L  = lambda t: _c("1;33", t)
C_G  = lambda t: _c("1;32", t)
C_R  = lambda t: _c("1;31", t)
C_B  = lambda t: _c("1;34", t)
C_D  = lambda t: _c("2",    t)
def _ts(): return datetime.now().strftime("%H:%M:%S")
def vlog(msg, indent=0): print(f"[{_ts()}] {'   '*indent}{msg}", flush=True)

# ═══════════════════════════════════════════════════════════════
# ▌ OUTPUT FILE + COLUMNS
# ═══════════════════════════════════════════════════════════════
OUTPUT_FILE      = "saudi_jobs.xlsx"
CHECKPOINT_FILE  = "pipeline_checkpoint.json"
COMPANIES_FILE   = "saudi_companies_found.csv"

# Canonical 22-column schema used throughout
COLUMNS = [
    "Job Title", "Job Type", "Job Qualifications", "Job Experience",
    "Job Location", "Job Field", "Date Posted", "Deadline",
    "Job Description", "Application", "Company URL", "Company Name",
    "Company Logo", "Company Industry", "Company Founded", "Company Type",
    "Company Website", "Company Address", "Company Details", "Job URL",
    "Estimated Deadline", "Salary Range",
]

# ═══════════════════════════════════════════════════════════════
# ▌ SHARED STATE  (populated by both pipelines)
# ═══════════════════════════════════════════════════════════════
# Each entry is a dict keyed by COLUMNS above.
all_jobs: list[dict] = []

# Dedup sets — shared across BOTH pipelines
# url_seen  : canonical job URL strings
# fp_seen   : (title.lower, company.lower, location.lower) tuples
url_seen:  set[str]   = set()
fp_seen:   set[tuple] = set()

# Company domains already processed by the Wikipedia pipeline
discovered_domains: set[str] = set()
company_results:    list     = []

# ═══════════════════════════════════════════════════════════════
# ▌ UNIFIED JOB RECORD FACTORY
# ═══════════════════════════════════════════════════════════════

def make_record(**kwargs) -> dict:
    """
    Build a canonical job record from arbitrary keyword arguments.
    Any key not in COLUMNS is silently dropped.
    All values are string-coerced and stripped.
    """
    rec = {col: "" for col in COLUMNS}
    for col in COLUMNS:
        # Accept both canonical names and legacy camelCase aliases
        for alias in (col, _ALIASES.get(col, "")):
            if alias and alias in kwargs and kwargs[alias]:
                rec[col] = str(kwargs[alias]).strip()[:2000]
                break
    return rec

# camelCase → Title Case aliases for LinkedIn scraper output
_ALIASES = {
    "Job Title":          "jobTitle",
    "Job Type":           "jobType",
    "Job Qualifications": "jobQualifications",
    "Job Experience":     "jobExperience",
    "Job Location":       "jobLocation",
    "Job Field":          "jobField",
    "Date Posted":        "datePosted",
    "Deadline":           "deadline",
    "Job Description":    "jobDescription",
    "Application":        "application",
    "Company URL":        "companyUrl",
    "Company Name":       "companyName",
    "Company Logo":       "companyLogo",
    "Company Industry":   "companyIndustry",
    "Company Founded":    "companyFounded",
    "Company Type":       "companyType",
    "Company Website":    "companyWebsite",
    "Company Address":    "companyAddress",
    "Company Details":    "companyDetails",
    "Job URL":            "jobUrl",
    "Estimated Deadline": "estimatedDeadline",
    "Salary Range":       "salaryRange",
}


def _merge_records(existing: dict, incoming: dict) -> dict:
    """
    Merge two records for the same logical job.
    For each field: keep existing if non-empty, else take incoming.
    Longer description / details always wins.
    """
    merged = dict(existing)
    for col in COLUMNS:
        ev = (existing.get(col) or "").strip()
        iv = (incoming.get(col) or "").strip()
        if not ev and iv:
            merged[col] = iv
        elif col in ("Job Description", "Company Details") and len(iv) > len(ev):
            merged[col] = iv
    return merged


def _fingerprint(rec: dict) -> tuple:
    return (
        (rec.get("Job Title")    or "").lower().strip(),
        (rec.get("Company Name") or "").lower().strip(),
        (rec.get("Job Location") or "").lower().strip(),
    )


def _canonical_url(url: str) -> str:
    """Normalise a LinkedIn /jobs/view/NNN URL; pass others through clean."""
    if not url:
        return ""
    m = re.search(r"/jobs/view/(\d+)", url)
    if m:
        return f"https://www.linkedin.com/jobs/view/{m.group(1)}/"
    return re.sub(r"[?#].*$", "", url).rstrip("/")


def register_job(rec: dict, source: str = "") -> bool:
    """
    Add rec to all_jobs after dedup check.
    Returns True if added/merged, False if fully duplicate.
    """
    if not rec.get("Job Title"):
        return False

    # ── URL dedup ────────────────────────────────────────────
    job_url = _canonical_url(rec.get("Job URL") or rec.get("Application") or "")
    if job_url and job_url in url_seen:
        # URL match — merge fields into existing record
        for i, existing in enumerate(all_jobs):
            ex_url = _canonical_url(existing.get("Job URL") or existing.get("Application") or "")
            if ex_url == job_url:
                all_jobs[i] = _merge_records(existing, rec)
                log.info(f"[merge/url] {rec.get('Job Title','')[:50]}")
                return True
        return False

    # ── Content fingerprint dedup ────────────────────────────
    fp = _fingerprint(rec)
    if fp in fp_seen:
        # Same logical job — merge
        for i, existing in enumerate(all_jobs):
            if _fingerprint(existing) == fp:
                all_jobs[i] = _merge_records(existing, rec)
                log.info(f"[merge/fp] {rec.get('Job Title','')[:50]}")
                return True
        return False

    # ── New job ──────────────────────────────────────────────
    if job_url:
        url_seen.add(job_url)
    fp_seen.add(fp)
    all_jobs.append(rec)
    return True


def save_output():
    if not all_jobs:
        return
    df = pd.DataFrame(all_jobs, columns=COLUMNS)
    df.to_excel(OUTPUT_FILE, index=False)
    log.info(f"Saved {len(all_jobs)} rows → {OUTPUT_FILE}")


# ═══════════════════════════════════════════════════════════════
# ▌ SHARED CONSTANTS & HELPERS
# ═══════════════════════════════════════════════════════════════
DELAY_S     = 2.0
FETCH_LIMIT = 120_000
MONTH_MAP   = {
    "jan":0,"feb":1,"mar":2,"apr":3,"may":4,"jun":5,
    "jul":6,"aug":7,"sep":8,"oct":9,"nov":10,"dec":11,
}
SAUDI_CITIES = [
    "Riyadh","Jeddah","Dammam","Mecca","Medina","Khobar",
    "Tabuk","Abha","Jubail","Yanbu","Taif","Buraidah",
    "Khamis Mushait","Hail","Najran","Jizan","Dhahran",
]
INDUSTRIES = [
    "technology","software","fintech","banking","finance",
    "healthcare","hospital","construction","real estate",
    "retail","manufacturing","oil gas","energy","telecom",
    "logistics","education","hospitality","consulting",
    "engineering","automotive","ecommerce","cybersecurity",
]
SKIP_DOMAINS = {
    "google","facebook","twitter","linkedin","wikipedia","youtube",
    "instagram","tiktok","snapchat","amazon","duckduck","bing",
    "yahoo","reddit","quora","trustpilot","glassdoor","indeed",
    "zawya","arabnews","saudigazette","bloomberg","reuters",
}
BAD_DOMAINS = [
    "linkedin.com","google.com","youtube.com","facebook.com",
    "twitter.com","x.com","instagram.com","t.co","example.com",
    "w3.org","sentry.io","schema.org",
]
NOISE_EMAIL_DOMAINS = [
    "example.com","sentry.io","google.com","w3.org",
    "schema.org","wixpress.com","squarespace.com",
]
FAKE_LOCAL_RE  = re.compile(
    r"^(name|user|email|mail|yourname|your[-_.]?email|sample|test|info|hello"
    r"|noreply|no[-_.]?reply|admin|webmaster|support|contact|example)$", re.I)
FAKE_DOMAIN_RE = re.compile(
    r"^(domain|example|yoursite|yourdomain|yourbrand|company|mycompany"
    r"|website|yourcompany|mysite|placeholder|site)\.[a-z]{2,}$", re.I)

_LI_UA_IDX = 0
_USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_4) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
]
def _next_headers() -> dict:
    global _LI_UA_IDX
    ua = _USER_AGENTS[_LI_UA_IDX % len(_USER_AGENTS)]
    _LI_UA_IDX += 1
    return {
        "User-Agent": ua,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Cache-Control": "no-cache",
        "X-Li-Lang": "en_US",
        "X-Requested-With": "XMLHttpRequest",
    }

HEADERS = _next_headers()

def get_domain(url):
    try: return urlparse(str(url)).netloc.lower().replace("www.","")
    except: return ""

def get_base(website):
    p = urlparse(str(website))
    return p.netloc.lower().replace("www.",""), p.scheme or "https"

def simple_get(url, timeout=10):
    try:
        r = requests.get(str(url), headers=HEADERS, timeout=timeout, allow_redirects=True)
        if r.status_code == 200:
            return r
    except: pass
    return None

def parse_domain(url):
    try:
        p = urlparse(url if url.startswith("http") else "https://"+url)
        netloc = p.netloc.lower().replace("www.","")
        if not netloc or any(s in netloc for s in SKIP_DOMAINS):
            return None, None
        return netloc, f"https://www.{netloc}"
    except:
        return None, None

def is_bad_url(url):
    if not url or not url.startswith("http"): return True
    lower = url.lower()
    return any(d in lower for d in BAD_DOMAINS)

def decode_html_entities(s):
    if not s: return ""
    for old,new in [("&amp;","&"),("&lt;","<"),("&gt;",">"),("&quot;",'"'),
                    ("&#39;","'"),("\\u0026","&"),("\\u003D","="),
                    ("\\u003A",":"),("\\u002F","/")]:
        s = s.replace(old,new)
    return s

def clean(text, max_len=300):
    return re.sub(r"\s+", " ", str(text or "")).strip()[:max_len]

# ── checkpoint ───────────────────────────────────────────────
def load_checkpoint():
    if Path(CHECKPOINT_FILE).exists():
        return json.loads(Path(CHECKPOINT_FILE).read_text())
    return {"processed_domains":[], "jobs_count":0, "done_sources":[]}

def save_checkpoint(cp):
    Path(CHECKPOINT_FILE).write_text(json.dumps(cp, indent=2))

# ═══════════════════════════════════════════════════════════════
# ▌ STANDARDISATION (shared by both pipelines)
# ═══════════════════════════════════════════════════════════════
FIELD_KEYWORD_MAP = [
    ("Information Technology",
     ["software engineer","developer","devops","frontend","backend",
      "full stack","fullstack","sysadmin","cloud","cybersecurity",
      "data engineer","machine learning","artificial intelligence",
      "ai/ml","it support","network engineer","database","kubernetes",
      "docker","aws","azure","react","node.js","python developer",
      "java developer","it manager","systems analyst","erp","sap",
      "technology","tech lead","infrastructure engineer","qa engineer",
      "automation engineer","business intelligence","data analyst"],
     ["programming","coding","api","agile","scrum","git","linux",
      "server","infrastructure","software","digital","tech"]),

    ("Finance & Accounting",
     ["accountant","auditor","finance manager","financial analyst",
      "cfo","treasurer","tax","bookkeeper","payroll","budget analyst",
      "credit analyst","investment","portfolio manager","risk analyst",
      "forex","actuary","acca","cfa","cpa","finance officer",
      "financial controller","internal audit","external audit",
      "accounts payable","accounts receivable","treasury"],
     ["financial","accounting","balance sheet","p&l","reconciliation",
      "ifrs","gaap","ledger","invoicing","fiscal","budget","revenue"]),

    ("Sales & Business Development",
     ["sales executive","sales manager","business development",
      "account manager","sales representative","bd manager",
      "regional sales","key account","sales director",
      "commercial manager","sales officer","revenue manager",
      "partnerships manager","channel manager","pre-sales"],
     ["revenue","pipeline","crm","leads","prospects","quota","target",
      "upsell","cross-sell","b2b","b2c","salesforce","negotiation"]),

    ("Marketing & Communications",
     ["marketing manager","digital marketing","seo","sem",
      "content marketer","social media manager","brand manager",
      "marketing executive","communications manager","pr manager",
      "copywriter","growth hacker","email marketing","campaign manager",
      "marketing director","brand strategist","media buyer",
      "public relations","communications officer"],
     ["marketing","branding","advertising","social media","content",
      "campaign","analytics","google ads","facebook ads","influencer",
      "awareness","positioning","messaging"]),

    ("Human Resources",
     ["hr manager","human resources","recruiter","talent acquisition",
      "hr business partner","hrbp","hr officer","compensation",
      "benefits manager","organisational development",
      "learning and development","l&d","hr generalist","payroll manager",
      "people operations","talent management","workforce planning",
      "employee engagement","hr director"],
     ["recruitment","onboarding","performance management",
      "employee relations","hr","workforce","headhunting","staffing",
      "saudization","nitaqat","labor law"]),

    ("Engineering",
     ["mechanical engineer","civil engineer","electrical engineer",
      "structural engineer","process engineer","project engineer",
      "maintenance engineer","production engineer","quality engineer",
      "safety engineer","site engineer","design engineer",
      "petroleum engineer","chemical engineer","industrial engineer",
      "instrumentation engineer","piping engineer","hvac engineer"],
     ["engineering","cad","autocad","solidworks","manufacturing",
      "plant","machinery","commissioning","maintenance","iso","asme"]),

    ("Healthcare & Medicine",
     ["doctor","physician","nurse","pharmacist","medical officer",
      "surgeon","anaesthetist","physiotherapist","radiographer",
      "lab technician","clinical","healthcare manager",
      "occupational therapist","dentist","midwife","radiologist",
      "oncologist","cardiologist","icu","emergency medicine",
      "infection control","medical director"],
     ["hospital","clinic","patient","medical","health",
      "pharmaceutical","diagnosis","treatment","ward","jci","cbahi"]),

    ("Education & Training",
     ["teacher","lecturer","professor","trainer","educator","tutor",
      "school principal","academic","curriculum","e-learning",
      "instructional designer","teaching assistant","academic advisor",
      "dean","faculty","research fellow"],
     ["school","university","college","classroom","students",
      "pedagogy","curriculum","education","training","accreditation"]),

    ("Hospitality & Tourism",
     ["hotel manager","front desk","housekeeping","chef","sous chef",
      "food and beverage","f&b manager","restaurant manager",
      "bartender","waiter","concierge","tour guide","travel agent",
      "events coordinator","catering","revenue manager hotel",
      "guest relations","front office manager"],
     ["hospitality","hotel","resort","tourism","guest",
      "accommodation","restaurant","kitchen","culinary","five star"]),

    ("Logistics & Supply Chain",
     ["supply chain manager","logistics coordinator","warehouse manager",
      "fleet manager","procurement manager","purchasing manager",
      "import export","freight","shipping coordinator",
      "inventory manager","demand planner","customs clearance",
      "logistics manager","distribution manager","last mile"],
     ["logistics","supply chain","warehouse","inventory","freight",
      "procurement","sourcing","distribution","customs","3pl","sap mm"]),

    ("Legal",
     ["lawyer","attorney","legal counsel","paralegal",
      "compliance officer","legal advisor","solicitor","barrister",
      "corporate counsel","legal manager","contract manager",
      "legal officer","in-house counsel","data protection officer"],
     ["legal","law","contracts","litigation","regulatory",
      "compliance","gdpr","intellectual property","arbitration"]),

    ("Administration & Operations",
     ["office manager","executive assistant","administrative officer",
      "operations manager","personal assistant","receptionist",
      "data entry","office administrator","company secretary",
      "business analyst","operations officer","facility manager",
      "administrative coordinator","executive secretary"],
     ["administration","operations","office","coordination",
      "scheduling","reporting","clerical","filing","facilities"]),

    ("Customer Service",
     ["customer service","call centre","customer success",
      "customer support","help desk","service advisor",
      "client relations","customer experience","contact centre",
      "customer care","cx specialist","complaints officer"],
     ["customer","support","helpdesk","tickets","escalation",
      "satisfaction","service level","inbound","outbound","nps"]),

    ("Construction & Real Estate",
     ["quantity surveyor","site supervisor","project manager construction",
      "architect","draughtsman","property manager","estate agent",
      "real estate","building inspector","land surveyor",
      "construction manager","project director","bim engineer",
      "fit out","interior designer","facilities engineer"],
     ["construction","building","property","real estate","site",
      "contractor","tender","drawings","neom","vision 2030 project"]),

    ("Manufacturing & Production",
     ["production manager","quality control","quality assurance",
      "qa","qc","factory manager","plant manager",
      "production supervisor","assembly","cnc operator","technician",
      "operations technician","manufacturing engineer","process technician"],
     ["production","manufacturing","factory","assembly","quality",
      "lean","six sigma","safety","line","output","throughput"]),

    ("Design & Creative",
     ["graphic designer","ui/ux","product designer","art director",
      "creative director","animator","illustrator","photographer",
      "videographer","motion designer","web designer","ux researcher",
      "visual designer","brand designer"],
     ["design","creative","adobe","figma","photoshop","illustrator",
      "indesign","sketch","branding","visual","wireframe","prototype"]),

    ("Research & Science",
     ["research scientist","data scientist","lab researcher",
      "research analyst","clinical researcher","environmental scientist",
      "chemist","biologist","statistician","epidemiologist",
      "geologist","geophysicist","reservoir engineer","r&d"],
     ["research","analysis","data","laboratory","science","experiment",
      "findings","methodology","survey","phd","publication"]),

    ("Security",
     ["security officer","security guard","security manager","cctv",
      "loss prevention","risk manager","health and safety",
      "hse officer","osh","fire safety","security analyst",
      "information security","cyber analyst","soc analyst"],
     ["security","safety","risk","surveillance","patrol",
      "access control","emergency","incident","iso 27001"]),

    ("Media & Journalism",
     ["journalist","editor","reporter","broadcast","news anchor",
      "content creator","media manager","radio","television",
      "producer","scriptwriter","editorial producer","multimedia journalist",
      "content producer","broadcast operations"],
     ["media","journalism","broadcast","news","editorial",
      "publishing","press","interview","newsroom","on-air"]),

    ("Non-Profit & Social Work",
     ["social worker","ngo","charity","programme coordinator",
      "community development","welfare officer","case manager",
      "development officer","fundraiser","volunteer coordinator"],
     ["social","ngo","community","welfare","beneficiary",
      "donor","impact","charity","development","fellowship"]),

    ("Oil & Gas",
     ["petroleum engineer","drilling engineer","reservoir engineer",
      "production engineer oil","subsurface","geoscientist",
      "upstream","downstream","refinery","petrochemical",
      "gas plant","field operator","well engineer","hse oil"],
     ["oil","gas","petroleum","refinery","upstream","downstream",
      "aramco","sabic","petrochemical","drilling","reservoir",
      "pipeline","lng","lpg","ngl","fracking","well"]),
]


def standardise_field(raw_field="", title="", description="", industry="") -> str:
    combined = " ".join([raw_field, title, description[:800], industry]).lower()
    best_label, best_score = "", 0
    for label, high_kws, low_kws in FIELD_KEYWORD_MAP:
        score = sum(3 for kw in high_kws if kw in combined)
        score += sum(1 for kw in low_kws if kw in combined)
        if score > best_score:
            best_score, best_label = score, label
    return best_label if best_score >= 1 else ""


_NO_EXP_KW = [
    "no experience","no prior experience","fresh graduate","freshers",
    "entry level","entry-level","0 years","zero experience",
    "training provided","will train","no experience required",
]
_LT1_KW = [
    "less than 1 year","under 1 year","6 months","less than a year",
    "some experience","minimal experience",
]

def _years_to_band(n):
    if n <= 0:  return "No Experience Required"
    if n <= 2:  return "1 - 2 Years"
    if n <= 5:  return "3 - 5 Years"
    if n <= 10: return "6 - 10 Years"
    return "10+ Years"

def standardise_experience(raw="") -> str:
    if not raw: return ""
    text = raw.lower().strip()
    for kw in _NO_EXP_KW:
        if kw in text: return "No Experience Required"
    for kw in _LT1_KW:
        if kw in text: return "Less than 1 Year"
    m = re.search(r"(\d+)\s*(?:\+|[-–]to]*\s*\d+)?\s*years?", text, re.I)
    if m:
        return _years_to_band(int(m.group(1)))
    if re.search(r"\b(senior|sr\.?|lead|principal|head of|director|vp)\b", text, re.I):
        return "6 - 10 Years"
    if re.search(r"\b(mid.?level|intermediate|associate)\b", text, re.I):
        return "3 - 5 Years"
    if re.search(r"\b(junior|jr\.?|graduate|intern|trainee|fresh)\b", text, re.I):
        return "Less than 1 Year"
    return ""

QUALIFICATION_TIERS = [
    ("PhD / Doctorate",           ["phd","ph.d","doctorate","doctoral","doctor of philosophy"]),
    ("Master's Degree",           ["master","msc","m.sc","mba","m.b.a","meng","m.eng","mphil",
                                    "postgraduate","post-graduate","post graduate"]),
    ("Bachelor's Degree",         ["bachelor","bsc","b.sc","beng","b.eng","bcom","b.com","bba",
                                    "llb","degree in","undergraduate degree","honours degree","hons",
                                    "b.tech","btech"]),
    ("Higher National Diploma",   ["hnd","hnc","higher national diploma","higher national certificate",
                                    "higher diploma","advanced diploma"]),
    ("Diploma",                   ["diploma","associate degree","foundation degree"]),
    ("Professional Certification",["acca","cpa","cfa","cima","pmp","prince2","cissp","aws certified",
                                    "comptia","cisco","ccna","ccnp","shrm","cipd","chartered",
                                    "certified public","certified financial","certified project",
                                    "professional certification","professional certificate"]),
    ("A-Levels / HSC",            ["a-level","a level","hsc","higher school certificate",
                                    "ib diploma","international baccalaureate","gce advanced"]),
    ("O-Levels / School Certificate", ["o-level","o level","igcse","gcse","school certificate"]),
    ("No Formal Qualification Required",
                                  ["no qualification","no degree","no formal","school leaver",
                                    "no experience required","training provided","will train"]),
]

def standardise_qualification(raw="", full_text="") -> str:
    corpus = ((raw or "") + " " + (full_text or "")[:2000]).lower()
    for label, keywords in QUALIFICATION_TIERS:
        if any(kw in corpus for kw in keywords):
            return label
    return ""

# ═══════════════════════════════════════════════════════════════
# ▌ DATE HELPERS  (shared)
# ═══════════════════════════════════════════════════════════════
def normalise_date_text(text):
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

def resolve_posted_date(raw):
    if not raw: return ""
    text = normalise_date_text(raw)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", text.strip()): return text.strip()
    try:
        return datetime.fromisoformat(text).strftime("%Y-%m-%d")
    except: pass
    base = datetime.now()
    m = re.search(r"(\d+)\s*(hour|day|week|month|year)", text, re.I)
    if m:
        n, unit = int(m.group(1)), m.group(2).lower()
        if "hour"  in unit: base -= timedelta(hours=n)
        elif "day"  in unit: base -= timedelta(days=n)
        elif "week" in unit: base -= timedelta(weeks=n)
        elif "month" in unit:
            mo = base.month - n; yr = base.year + mo//12; mo = mo%12 or 12
            base = base.replace(year=yr, month=mo)
        elif "year" in unit: base = base.replace(year=base.year-n)
        return base.strftime("%Y-%m-%d")
    if re.search(r"just\s*now|today", text, re.I):
        return datetime.now().strftime("%Y-%m-%d")
    return datetime.now().strftime("%Y-%m-%d")

def estimate_deadline(date_posted_str):
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

def estimate_deadline_from_posted(posted_text):
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
            mo = base.month-n; yr = base.year+mo//12; mo = mo%12 or 12
            base = base.replace(year=yr, month=mo)
    mo = base.month+3; yr = base.year+(mo-1)//12; mo = (mo-1)%12+1
    return base.replace(year=yr, month=mo).strftime("%Y-%m-%d")

# ═══════════════════════════════════════════════════════════════
# ▌ EMAIL / URL CLEANERS  (shared)
# ═══════════════════════════════════════════════════════════════
def clean_email(raw):
    if not raw: return ""
    em = raw
    em = re.sub(r"^mailto:", "", em, flags=re.I)
    em = re.sub(r"\?.*$", "", em)
    for pat, repl in [(r"\\u003[Ee]",""), (r"\\u003[Cc]",""), (r"\\u0040","@"),
                       (r"\\u002[Ee]","."), (r"\\u0026",""), (r"u003[Ee]",""),
                       (r"u003[Cc]",""), (r"u0040","@"), (r"&amp;",""),
                       (r"&lt;",""), (r"&gt;",""), (r"&#64;","@"), (r"&#46;","."),
                       (r"&nbsp;",""), (r"%40","@"), (r"%2[Ee]","."), (r"%20",""),
                       (r"[>]+$",""), (r"[<]+$","")]:
        em = re.sub(pat, repl, em, flags=re.I)
    em = em.strip().lower()
    if not em or "@" not in em or "." not in em: return ""
    at = em.rfind("@"); local, domain = em[:at], em[at+1:]
    domain = re.sub(r"(\.[a-z]{2,6})[a-z0-9\-_/?#+]*$", r"\1", domain, flags=re.I)
    em = local + "@" + domain
    if not re.match(r"^[a-zA-Z0-9]", em): return ""
    return em

def extract_email_from_text(text):
    if not text: return ""
    for raw_em in re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", text):
        em = clean_email(raw_em)
        if not em or "@" not in em: continue
        parts = em.split("@")
        if len(parts) != 2: continue
        if any(em.find(d) != -1 for d in NOISE_EMAIL_DOMAINS): continue
        if FAKE_LOCAL_RE.match(parts[0]) or FAKE_DOMAIN_RE.match(parts[1]): continue
        return em
    return ""

def clean_application_link(raw):
    if not raw: return ""
    raw = raw.strip()
    if "@" in raw and not raw.startswith("http"): return clean_email(raw)
    if raw.startswith("http"):
        url = re.sub(r"#.*$","",raw)
        url = re.sub(r"[.,;:!?)]+$","",url)
        return url.strip()
    return raw

def decode_linkedin_apply_url(raw):
    if not raw: return ""
    raw = decode_html_entities(raw)
    if raw.startswith("http") and "linkedin.com" not in raw: return raw
    m = re.search(r"[?&]url=([^&]+)", raw)
    if m:
        decoded = unquote(m.group(1))
        if "%" in decoded: decoded = unquote(decoded)
        if decoded.startswith("http") and "linkedin.com" not in decoded: return decoded
    return ""

# ═══════════════════════════════════════════════════════════════
# ▌ DESCRIPTION CLEANER  (shared)
# ═══════════════════════════════════════════════════════════════
def clean_description(raw):
    if not raw: return ""
    text = raw.replace("\u00a0"," ").replace("\u200b","")
    text = text.replace("\r\n","\n").replace("\r","\n")
    text = re.sub(r"([a-z])([A-Z])",r"\1 \2",text)
    text = re.sub(r"(\d)([A-Za-z])",r"\1 \2",text)
    text = re.sub(r"([A-Za-z])(\d)",r"\1 \2",text)
    text = re.sub(r"([.,:;!?])([A-Za-z0-9])",r"\1 \2",text)
    text = re.sub(r"\s*[•·▪◦]\s*","\n• ",text)
    text = re.sub(r"\n\s*[-–—]\s+","\n• ",text)
    paragraphs = re.split(r"\n{2,}", text)
    cleaned = []
    for para in paragraphs:
        lines = []
        for line in para.split("\n"):
            line = line.strip()
            if not line: continue
            if (not re.search(r"[.!?:;,]$",line) and
                    not re.match(r"^[A-Z\s]{3,30}$",line) and
                    len(line) > 8 and
                    not re.match(r"^[•\-–]",line) and
                    not re.match(r"^\w+:$",line)):
                line = line + "."
            lines.append(line)
        cleaned.append("\n".join(lines))
    return re.sub(r" {2,}"," ","\n\n".join(p for p in cleaned if p.strip())).strip()

# ═══════════════════════════════════════════════════════════════
# ▌ LINKEDIN PIPELINE  (Source A)
# ═══════════════════════════════════════════════════════════════

# ── Config ────────────────────────────────────────────────────
LI_MAX_PAGES       = 0   # 0 = unlimited per keyword
LI_MAX_EMPTY_PAGES = 5
LI_JOB_LIMIT       = 0   # 0 = no cap

SEARCH_KEYWORDS = [
    # ── Broad sweep ──────────────────────────────────────────
    "",
    # ── Engineering & Technical ──────────────────────────────
    "engineer","civil engineer","mechanical engineer","electrical engineer",
    "structural engineer","process engineer","chemical engineer","petroleum engineer",
    "instrumentation engineer","piping engineer","safety engineer","HSE","QA QC",
    "commissioning","maintenance engineer","automation engineer","controls engineer",
    "HVAC engineer","fire protection engineer","geotechnical engineer",
    # ── Information Technology ────────────────────────────────
    "developer","software engineer","full stack developer","frontend developer",
    "backend developer","DevOps","cloud engineer","cybersecurity","network engineer",
    "database administrator","IT","IT support","systems administrator","data engineer",
    "data scientist","machine learning","AI engineer","ERP consultant","SAP","Oracle DBA",
    # ── Business & Management ─────────────────────────────────
    "manager","project manager","operations manager","general manager","country manager",
    "business development","commercial manager","strategy","consultant","director",
    "VP","CEO","COO","CFO",
    # ── Finance & Accounting ──────────────────────────────────
    "finance","accountant","financial analyst","auditor","tax","treasury",
    "budget analyst","cost controller","credit analyst","investment analyst",
    "ACCA","CPA","CFA","VAT","IFRS",
    # ── Sales & Marketing ─────────────────────────────────────
    "sales","marketing","brand manager","digital marketing","SEO","social media",
    "sales executive","account manager","key account","trade marketing",
    "retail manager","e-commerce",
    # ── Human Resources ───────────────────────────────────────
    "HR","human resources","recruiter","talent acquisition","HR business partner",
    "payroll","organizational development","learning and development","compensation benefits",
    # ── Healthcare & Medicine ─────────────────────────────────
    "doctor","physician","nurse","pharmacist","dentist","surgeon","radiologist",
    "physiotherapist","lab technician","medical officer","healthcare","clinical",
    "paramedic","dietitian","occupational therapist",
    # ── Construction & Real Estate ────────────────────────────
    "construction","site engineer","quantity surveyor","architect","site supervisor",
    "project controls","planning engineer","draughtsman","BIM","real estate",
    "property manager","facilities manager","MEP engineer",
    # ── Logistics & Supply Chain ──────────────────────────────
    "logistics","supply chain","procurement","warehouse","inventory","freight",
    "customs clearance","shipping","import export","fleet manager","demand planner","sourcing",
    # ── Administration & Operations ───────────────────────────
    "operations","receptionist","executive assistant","personal assistant",
    "office manager","administrative","data entry","document controller","company secretary",
    # ── Customer Service ──────────────────────────────────────
    "customer service","call center","customer success","help desk","customer experience",
    # ── Education & Training ──────────────────────────────────
    "teacher","lecturer","trainer","academic","curriculum","principal","tutor",
    "instructional designer","e-learning",
    # ── Hospitality & Tourism ─────────────────────────────────
    "chef","hotel manager","food and beverage","sous chef","front desk","housekeeping",
    "catering","hospitality","tour operator","events coordinator",
    # ── Legal & Compliance ────────────────────────────────────
    "lawyer","legal counsel","compliance","paralegal","contract manager","risk manager","GDPR",
    # ── Design & Creative ─────────────────────────────────────
    "graphic designer","UX designer","UI designer","product designer","interior designer",
    "videographer","content creator","motion designer",
    # ── Manufacturing & Production ────────────────────────────
    "production manager","quality control","plant manager","factory manager",
    "manufacturing","lean","six sigma","CNC","welding","foreman",
    # ── Oil & Gas (Saudi-specific) ────────────────────────────
    "petroleum","reservoir engineer","drilling engineer","production engineer oil",
    "refinery","upstream","downstream","LNG","gas plant","subsea","pipeline engineer",
    "well completion","aramco","SABIC","NEOM",
    # ── Transport & Driving ───────────────────────────────────
    "driver","delivery driver","truck driver","heavy vehicle","chauffeur",
    "transport coordinator",
    # ── Security ─────────────────────────────────────────────
    "security","security officer","CCTV","loss prevention","fire safety",
    # ── Research & Science ────────────────────────────────────
    "researcher","scientist","chemist","biologist","geologist","environmental",
    "laboratory","statistician",
    # ── Media & Communications ────────────────────────────────
    "journalist","editor","copywriter","PR","communications","media","photographer","broadcast",
    # ── Banking & Insurance ───────────────────────────────────
    "banker","relationship manager","credit officer","insurance","actuarial",
    "investment banker","branch manager","teller",
    # ── Retail & Consumer ─────────────────────────────────────
    "retail","store manager","merchandiser","visual merchandiser","cashier","sales associate",
    # ── Vision 2030 / Emerging Sectors ───────────────────────
    "renewable energy","solar","wind energy","sustainability","ESG","smart city",
    "digital transformation","blockchain","fintech","startup",
    "tourism development","entertainment","sports management",
]


def _build_li_url(keyword, start):
    return (
        "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search"
        f"?location=Saudi+Arabia&f_TPR=r604800&keywords={quote_plus(keyword)}&start={start}"
    )


def _collect_li_urls(html, seen):
    found = []
    for raw in re.findall(r'href="(https?://[^"]*?/jobs/view/\d+[^"]*?)"', html):
        c = _canonical_url(raw)
        if c and c not in seen:
            seen.add(c); found.append(c)
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup.find_all("a", href=True):
        href = tag["href"]
        if "/jobs/view/" not in href: continue
        if not href.startswith("http"): href = "https://www.linkedin.com" + href
        c = _canonical_url(href)
        if c and c not in seen:
            seen.add(c); found.append(c)
    return found


def _fetch_li_page(keyword, start, retries=3):
    url = _build_li_url(keyword, start)
    for attempt in range(retries):
        try:
            time.sleep(DELAY_S + attempt*3)
            r = requests.get(url, headers=_next_headers(), allow_redirects=True, timeout=25)
            if r.status_code == 429:
                wait = 60 + attempt*60
                print(C_R(f"  ⏳ Rate limited — waiting {wait}s …"))
                time.sleep(wait); continue
            if r.status_code in (400,403,999): return None
            if r.status_code != 200: return None
            text = r.text.strip()
            return text if text else None
        except Exception as e:
            log.warning(f"LI page error attempt {attempt+1}: {e}")
            time.sleep(3 + attempt*3)
    return None


def _paginate_li_keyword(keyword, url_seen_set):
    urls, page, empty_streak = [], 0, 0
    label = keyword or "(all)"
    while True:
        if LI_MAX_PAGES and page >= LI_MAX_PAGES: break
        start = page * 25
        html  = _fetch_li_page(keyword, start)
        if html is None: break
        new = _collect_li_urls(html, url_seen_set)
        if new:
            urls.extend(new); empty_streak = 0
        else:
            empty_streak += 1
            if empty_streak >= LI_MAX_EMPTY_PAGES: break
        if start >= 975: break
        page += 1
        if page % 10 == 0: time.sleep(20)
    return urls


# ── LinkedIn job detail scraper ───────────────────────────────

SKIP_CRAWL_DOMAINS = [
    "dhl.com","fedex.com","ups.com","amazon.com","amazon.jobs",
    "google.com","microsoft.com","apple.com","meta.com","ibm.com",
    "oracle.com","sap.com","accenture.com","deloitte.com","pwc.com",
    "kpmg.com","ey.com","mckinsey.com","bcg.com","bain.com",
    "citibank.com","hsbc.com","barclays.com","bnpparibas.com",
    "airbus.com","boeing.com","siemens.com","ge.com",
    "unilever.com","nestle.com","pg.com","shell.com","bp.com",
]

def _should_skip_crawl(url):
    if not url: return True
    return any(d in url.lower() for d in SKIP_CRAWL_DOMAINS)

def _is_career_url(url):
    return any(k in url.lower() for k in
               ["career","jobs","apply","vacanci","recruit","opening","hiring","work-with"])

def _is_contact_url(url):
    return any(k in url.lower() for k in
               ["contact","about","reach","get-in","enquir","support"])

def _make_absolute(href, root):
    if not href: return ""
    href = href.strip()
    if href.startswith("http"): return href
    if href.startswith("//"): return "https:"+href
    if href.startswith("/"): return root.rstrip("/")+href
    return ""

def _scan_page_for_email(soup, raw_html=""):
    for tag in soup.find_all("a", href=re.compile(r"^mailto:",re.I)):
        em = clean_email(tag.get("href",""))
        if em and not any(d in em for d in NOISE_EMAIL_DOMAINS):
            parts = em.split("@")
            if len(parts)==2 and not FAKE_LOCAL_RE.match(parts[0]) and not FAKE_DOMAIN_RE.match(parts[1]):
                return em
    found = extract_email_from_text(
        "".join(" "+tag.get_text() for tag in soup.select("footer,#footer,.footer,#contact,.contact"))
    )
    if found: return found
    return extract_email_from_text(soup.get_text())

def _crawl_company_website(website_url, job_title):
    if _should_skip_crawl(website_url):
        return {"url": website_url, "email": "", "method": "fallback_website"}
    deadline = time.time() + 12
    root = website_url.rstrip("/")
    def _get(url):
        if time.time() > deadline: return None
        try:
            time.sleep(0.5)
            r = requests.get(url, headers=HEADERS, timeout=10, allow_redirects=True)
            return r.text[:FETCH_LIMIT] if r.status_code == 200 else None
        except: return None
    home_html = _get(root)
    if not home_html: return {"url":"","email":"","method":""}
    soup_h = BeautifulSoup(home_html, "html.parser")
    home_email = _scan_page_for_email(soup_h, home_html)
    if home_email: return {"url":"","email":home_email,"method":"s7_homepage_email"}
    careers_url = contact_url = ""
    for tag in soup_h.find_all("a", href=True):
        href = _make_absolute(tag.get("href",""), root)
        if not href or is_bad_url(href) or href==root or root not in href: continue
        if not careers_url and _is_career_url(href): careers_url = href
        if not contact_url and (_is_contact_url(href) or "contact" in tag.get_text().lower()):
            contact_url = href
        if careers_url and contact_url: break
    for url, method in [(careers_url,"s7_careers_email"),(contact_url,"s7_contact_email")]:
        if url and time.time() < deadline:
            h = _get(url)
            if h:
                email = _scan_page_for_email(BeautifulSoup(h,"html.parser"), h)
                if email: return {"url":"","email":email,"method":method}
    if careers_url: return {"url":careers_url,"email":"","method":"s7_careers_page"}
    return {"url":root,"email":"","method":"fallback_website"}


def _get_job_criteria(soup, label):
    lower = label.lower()
    for li in soup.select(".description__job-criteria-list > li"):
        h3 = li.find("h3")
        if h3 and lower in h3.get_text().strip().lower():
            spans = li.select(".description__job-criteria-text, span")
            if spans: return spans[-1].get_text(strip=True)
    for chip in soup.select(".job-details-jobs-unified-top-card__job-insight,.jobs-unified-top-card__job-insight"):
        text = chip.get_text(strip=True).lower()
        if "employment" in lower or "type" in lower:
            if re.search(r"full[\-\s]?time|part[\-\s]?time|contract|temporary|internship|freelance",text,re.I):
                return chip.get_text(strip=True)
        elif "seniority" in lower:
            if re.search(r"entry|associate|mid[\-\s]?senior|senior|director|executive|intern",text,re.I):
                return chip.get_text(strip=True)
    meta_map = {
        "employment type": soup.find("meta",{"name":"employmentType"}),
        "seniority level":  soup.find("meta",{"name":"seniorityLevel"}),
        "industries":       soup.find("meta",{"name":"industry"}),
    }
    tag = meta_map.get(lower)
    if tag: return tag.get("content","")
    return ""


def _get_workplace_type(soup):
    for sel in [".topcard__workplace-type",
                ".job-details-jobs-unified-top-card__workplace-type",
                ".jobs-unified-top-card__workplace-type"]:
        el = soup.select_one(sel)
        if el: return el.get_text(strip=True)
    for chip in soup.select(".job-details-jobs-unified-top-card__job-insight,.jobs-unified-top-card__job-insight"):
        t = chip.get_text(strip=True)
        if re.match(r"^(remote|on[\-\s]?site|hybrid)$",t,re.I): return t
    return ""


def _scrape_company(company_url):
    if not company_url: return {}
    try:
        resp = requests.get(company_url, headers=HEADERS, timeout=15)
        if resp.status_code != 200: return {}
        html = resp.text; soup = BeautifulSoup(html,"html.parser")
        logo_m = re.search(r'<meta[^>]+property=["\']og:image["\'][^>]+content=["\']([^"\']+)["\']',html,re.I)
        raw_logo = logo_m.group(1) if logo_m else ""
        ws_tag = soup.select_one("a[data-tracking-control-name='about_website']")
        website = ws_tag.get("href","") if ws_tag else ""
        name_el = soup.find("h1")
        def _detail(lbl):
            for div in soup.select("section.core-section-container dl > div"):
                dt = div.find("dt")
                if dt and lbl.lower() in dt.get_text().strip().lower():
                    dd = div.find("dd")
                    if dd: return dd.get_text(strip=True)
            return ""
        return {
            "name": name_el.get_text(strip=True) if name_el else "",
            "industry": _detail("Industry"),
            "size": _detail("Company size"),
            "headquarters": _detail("Headquarters"),
            "type": _detail("Type"),
            "founded": _detail("Founded"),
            "website": website,
            "logo": raw_logo,
            "about": (soup.select_one("section.about-us p") or
                      soup.select_one(".core-section-container__content p") or
                      type("_",(),{"get_text":lambda s,**k:""})()
                     ).get_text(strip=True),
        }
    except: return {}


def _scrape_li_job(job_url):
    try:
        resp = requests.get(job_url, headers=HEADERS, timeout=15)
        if resp.status_code != 200: return None
        html = resp.text; soup = BeautifulSoup(html,"html.parser")
    except: return None

    def _sel(*selectors):
        for s in selectors:
            el = soup.select_one(s)
            if el:
                t = el.get_text(strip=True)
                if t: return t
        return ""

    title        = _sel(".top-card-layout__title","h1.topcard__title",
                         ".job-details-jobs-unified-top-card__job-title","h1")
    company_name = _sel(".topcard__org-name-link",
                         ".job-details-jobs-unified-top-card__company-name",".topcard__flavor")
    co_url_el    = (soup.select_one(".topcard__org-name-link") or
                    soup.select_one(".job-details-jobs-unified-top-card__company-name a"))
    co_url       = co_url_el.get("href","") if co_url_el else ""
    location     = _sel(".topcard__flavor--bullet",
                         ".job-details-jobs-unified-top-card__bullet")
    workplace    = _get_workplace_type(soup)
    time_el      = soup.find("time")
    raw_posted   = (time_el.get("datetime","") if time_el else "") or \
                    _sel(".posted-time-ago__text",
                          ".job-details-jobs-unified-top-card__posted-date")
    posted_date  = resolve_posted_date(raw_posted)
    raw_desc     = _sel(".show-more-less-html__markup",".description__text")
    description  = clean_description(raw_desc)
    salary = ""
    for sel in [".compensation__salary",".salary","[class*='salary']","[class*='compensation']"]:
        el = soup.select_one(sel)
        if el: salary = el.get_text(strip=True); break
    if not salary:
        for chip in soup.select(".job-details-jobs-unified-top-card__job-insight"):
            t = chip.get_text(strip=True)
            if re.search(r"\$|SAR|SR|salary|/yr|/hour|per month",t,re.I):
                salary = t; break
    raw_job_type = _get_job_criteria(soup,"Employment type") or workplace
    job_type     = raw_job_type or "Full-time"
    li_function  = _get_job_criteria(soup,"Job function")
    li_industry  = _get_job_criteria(soup,"Industries")
    job_field    = standardise_field(li_function or li_industry or "",
                                      title, description, li_industry)
    real_dl      = ""
    for pat in [r"closes?\s+on\s+(\S.{0,30})",r"apply\s+by\s+(\S.{0,30})",
                r"deadline[:\s]+(\S.{0,30})"]:
        m = re.search(pat, soup.get_text(), re.I)
        if m:
            try:
                d = datetime.strptime(m.group(1).strip()[:10],"%Y-%m-%d")
                if d > datetime.now(): real_dl = d.strftime("%Y-%m-%d"); break
            except: pass
    est_dl  = estimate_deadline_from_posted(posted_date) if not real_dl else ""
    eff_dl  = real_dl or est_dl
    company = _scrape_company(co_url)
    # apply link
    apply_link = ""
    for tag in soup.find_all("a", href=True):
        href = tag.get("href","")
        ctrl = tag.get("data-tracking-control-name","")
        if "offsite" in ctrl.lower() or "apply" in ctrl.lower():
            resolved = decode_linkedin_apply_url(href)
            if resolved and not is_bad_url(resolved):
                apply_link = resolved; break
    if not apply_link:
        for script in soup.find_all("script"):
            txt = script.string or ""
            for pat in [r'"applyStartUrl"\s*:\s*"([^"]+)"',
                        r'"applicationUrl"\s*:\s*"([^"]+)"']:
                m = re.search(pat,txt)
                if m:
                    candidate = decode_html_entities(m.group(1)).replace("\\","")
                    if candidate.startswith("http") and not is_bad_url(candidate):
                        apply_link = candidate; break
    if not apply_link:
        co_website = company.get("website","")
        if co_website and not is_bad_url(co_website):
            if _should_skip_crawl(co_website):
                apply_link = co_website
            else:
                result = _crawl_company_website(co_website, title)
                apply_link = result.get("email","") or result.get("url","") or co_website
    apply_link = clean_application_link(apply_link)
    qualifications = standardise_qualification("", description)
    experience     = standardise_experience(description)
    return make_record(
        **{
            "Job Title":       title,
            "Job Type":        job_type,
            "Job Qualifications": qualifications,
            "Job Experience":  experience,
            "Job Location":    location,
            "Job Field":       job_field,
            "Date Posted":     posted_date,
            "Deadline":        eff_dl,
            "Job Description": description,
            "Application":     apply_link,
            "Company URL":     co_url,
            "Company Name":    company.get("name") or company_name,
            "Company Logo":    company.get("logo",""),
            "Company Industry":company.get("industry") or li_industry,
            "Company Founded": company.get("founded",""),
            "Company Type":    company.get("type",""),
            "Company Website": company.get("website",""),
            "Company Address": company.get("headquarters",""),
            "Company Details": company.get("about",""),
            "Job URL":         job_url,
            "Estimated Deadline": est_dl,
            "Salary Range":    salary,
        }
    )


def run_linkedin_pipeline():
    """Synchronous driver for the LinkedIn guest-API pipeline."""
    print(C_H("\n" + "="*72))
    print(C_H("  SOURCE A — LinkedIn guest API (Saudi Arabia)"))
    print(C_H("="*72))
    print(f"  Keywords : {len(SEARCH_KEYWORDS)}")
    print(f"  Job cap  : {'none' if not LI_JOB_LIMIT else LI_JOB_LIMIT}")

    li_url_seen: set[str] = set()  # local set fed into global url_seen after canonicalise
    all_li_urls: list[str] = []

    # Phase 1 — collect URLs
    for qi, kw in enumerate(SEARCH_KEYWORDS):
        label = kw or "(all)"
        print(C_B(f"\n┌─ [{qi+1}/{len(SEARCH_KEYWORDS)}] keyword='{label}'"))
        new = _paginate_li_keyword(kw, li_url_seen)
        all_li_urls.extend(new)
        print(C_B(f"└─ {len(new)} new  (total {len(all_li_urls)})"))
        if LI_JOB_LIMIT and len(all_li_urls) >= LI_JOB_LIMIT:
            all_li_urls = all_li_urls[:LI_JOB_LIMIT]; break
        time.sleep(DELAY_S * 2)

    print(C_H(f"\n  Total LinkedIn URLs collected: {len(all_li_urls)}"))

    # Phase 2 — scrape details
    added = skipped = errors = 0
    for j, url in enumerate(all_li_urls, 1):
        print(f"\n{C_H(f'>>> LI job {j}/{len(all_li_urls)} ...')}")
        try:
            rec = _scrape_li_job(url)
            if rec and rec.get("Job Title"):
                if register_job(rec, source="linkedin"):
                    added += 1
                    print(C_G(f"  ✔  {rec['Job Title'][:60]} @ {rec.get('Company Name','')[:40]}"))
                else:
                    skipped += 1
                    print(C_D(f"  ⟳  dup skipped: {rec.get('Job Title','')[:60]}"))
            else:
                print(C_R("  ✗  no title"))
        except Exception as e:
            errors += 1
            print(C_R(f"  ✗  ERROR: {e}"))
        time.sleep(DELAY_S)
        if (added + skipped) % 50 == 0 and added > 0:
            save_output()

    save_output()
    print(C_H(f"\n  LinkedIn done — added:{added}  merged:{skipped-errors}  errors:{errors}"))


# ═══════════════════════════════════════════════════════════════
# ▌ WIKIPEDIA PIPELINE  (Source B)
# ═══════════════════════════════════════════════════════════════
# (Everything below is the v5 career-page scraper, adapted to
#  call register_job() instead of appending to a local list.)

NON_SAUDI_DOMAINS = {
    "fortune.com","forbes.com","bloomberg.com","reuters.com",
    "arabnews.com","aljazeera.com","bbc.com","bbc.co.uk",
    "cnn.com","nytimes.com","wsj.com","ft.com","c-span.org",
    "federalreserve.gov","sec.gov","state.gov","treasury.gov",
    "londonstockexchange.com","lseg.com","nyse.com","nasdaq.com",
    "imf.org","worldbank.org","opec.org","un.org","unesco.org",
    "pepsico.com","linkedin.com","facebook.com","twitter.com",
    "x.com","instagram.com","tiktok.com","snapchat.com",
    "amazon.com","google.com","microsoft.com","apple.com",
    "hsbc.com","barclays.com","houseofsaud.com",
}
SAUDI_DOMAIN_INDICATORS = [
    r"\.sa$",r"\.gov\.sa$",r"\.edu\.sa$",r"\.com\.sa$",r"\.org\.sa$",
    r"aramco",r"sabic",r"stc\.com",r"alrajhi",r"samba",r"riyad",
    r"ncb",r"jarir",r"maaden",r"tasnee",r"mobily",r"zain\.sa",
    r"saudia",r"flynas",r"flyadeal",r"neom",r"vision2030",r"pif\.gov",
]
NON_COMPANY_WIKI = re.compile(
    r"^(women'?s?\s+rights|transport\s+in|arabic$|ardah$|"
    r"custodian\s+of|council\s+of\s+ministers|consultative\s+assembly|"
    r"royal\s+saudi\s+(navy|air|army|guard|force|defense)|"
    r"committee\s+for\s+the\s+promotion|general\s+intelligence|"
    r"supreme\s+economic|capital\s+market\s+authority|"
    r"us\s+dollar|saudi\s+riyal|opec|arabic\s+language|"
    r"saudi\s+arabia\s+(economy|culture|history|geography|politics|religion)|"
    r"islam\s+in|hajj|umrah|mecca$|medina$)",
    re.I,
)

def _is_saudi_domain(domain):
    d = domain.lower()
    if d in NON_SAUDI_DOMAINS: return False
    return any(re.search(p,d) for p in SAUDI_DOMAIN_INDICATORS)

def _looks_like_saudi_company(name, website):
    domain = get_domain(website)
    if domain in NON_SAUDI_DOMAINS: return False
    if NON_COMPANY_WIKI.match(name.strip()): return False
    return True

BLOCKED_CAREER_DOMAINS = {"linkedin.com","www.linkedin.com"}

def _is_blocked_career(url):
    return get_domain(url) in BLOCKED_CAREER_DOMAINS

CAREER_SUBDOMAINS = [
    "careers","jobs","career","job","work","hiring",
    "apply","talent","recruitment","hr","people","vacancies","opportunities","join",
]
CAREER_PATHS = [
    "/careers","/jobs","/career","/job","/careers/","/jobs/",
    "/en/careers","/en/jobs","/ar/careers","/about/careers","/about/jobs",
    "/company/careers","/join-us","/join","/work-with-us",
    "/openings","/vacancies","/opportunities","/employment",
    "/hiring","/recruitment","/apply","/careers/search","/jobs/search",
]
ATS_DOMAINS = {
    "greenhouse.io":"Greenhouse","lever.co":"Lever",
    "myworkdayjobs.com":"Workday","ashbyhq.com":"Ashby",
    "smartrecruiters.com":"SmartRecruiters","bamboohr.com":"BambooHR",
    "icims.com":"iCIMS","taleo.net":"Taleo",
    "successfactors.com":"SuccessFactors","sapsf.com":"SuccessFactors",
    "oraclecloud.com":"Oracle","recruitee.com":"Recruitee",
    "workable.com":"Workable","jobvite.com":"Jobvite",
    "breezy.hr":"Breezy","zohorecruit.com":"Zoho",
    "bayt.com":"Bayt","gulftalent.com":"GulfTalent",
    "teamtailor.com":"Teamtailor","comeet.com":"Comeet",
    "pageuppeople.com":"PageUp","jobsoid.com":"Jobsoid",
    "wynt.ai":"Wynt","easy.jobs":"EasyJobs",
    "ats.sa":"ATS.sa","peoplehr.net":"PeopleHR",
    "smarterp.sa":"SmartERP",
}
ATS_HTML_FP = {
    "rmkcdn.successfactors.com":"SuccessFactors",
    "talentcommunity/apply":"SuccessFactors",
    "/go/Job-Search/":"SuccessFactors",
    "greenhouse-io":"Greenhouse","myworkdayjobs":"Workday",
    "lever.co/v0/postings":"Lever","ashbyhq.com":"Ashby",
    "icims.com":"iCIMS","bamboohr.com":"BambooHR",
    "smartrecruiters.com":"SmartRecruiters",
    "teamtailor.com":"Teamtailor","recruitee.com":"Recruitee",
    "pageuppeople.com":"PageUp","jobvite.com":"Jobvite",
    "taleo.net":"Taleo","wynt.ai":"Wynt","easy.jobs":"EasyJobs",
}
CAREER_KEYWORDS = [
    "career","careers","job","jobs","vacancy","vacancies","hiring",
    "work with us","join us","join our team","employment","opportunities",
    "opening","openings","وظائف","التوظيف","انضم إلينا","فرص عمل",
]
JOB_PAGE_SIGNALS = [
    "apply now","apply for","job description","requirements",
    "qualifications","responsibilities","we are looking","we are hiring",
    "open position","full-time","part-time","remote","salary","benefits",
    "وظيفة","تقديم","المتطلبات","job opportunities","job search",
    "talentcommunity","create alert","sort by title","sort by location",
]
JOB_LISTING_SUFFIXES = [
    "/go/Job-Search/","/go/All-Jobs/",
    "/job-search-results","/job-search-results/",
    "/en/job-search-results","/search","/search/",
    "/job-search","/job-search/","/openings","/openings/",
    "/current-openings","/positions","/positions/",
    "/open-positions","/listings","/listings/",
    "/all-jobs","/all-jobs/","/list","/list/",
    "/vacancies","/vacancies/","/opportunities","/opportunities/",
    "/jobs","/jobs/","/apply","/apply/","/join","/join/",
    "/en/jobs","/en/careers","/en/search","/en/openings",
    "/ar/jobs","/ar/careers","/ar/search","/ar/vacancies",
]
LOCATION_PATTERN = re.compile(
    r"(Riyadh|Jeddah|Dammam|Khobar|Mecca|Medina|Saudi Arabia|KSA|Remote|"
    r"Dhahran|Jubail|Yanbu|Taif|Abha|Buraidah|Hail|Tabuk|"
    r"الرياض|جدة|الدمام|مكة|المدينة|Central Province|Eastern Province|"
    r"Western Province|Makkah|Madinah)",re.I
)
SALARY_RE = re.compile(
    r"(?:SAR|SR|USD|\$|€|£)\s?[\d,]+(?:\s?[-–]\s?[\d,]+)?(?:\s?[Kk])?|"
    r"[\d,]+(?:\s?[-–]\s?[\d,]+)?\s?(?:SAR|SR|USD|per\s+month|/month|monthly)",re.I
)
DATE_POSTED_RE = re.compile(
    r"(?:posted|date posted|published)[:\s]*([^\n<]{1,40})|"
    r"(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})|"
    r"(\d+\s+(?:day|hour|week|month)s?\s+ago)|"
    r"(\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2},?\s+\d{4})",re.I
)
DEADLINE_RE = re.compile(
    r"(?:deadline|closing date|apply by|applications? close)[:\s]*([^\n<]{1,40})|"
    r"(?:expires?|expiry)[:\s]*([^\n<]{1,40})",re.I
)
EXPERIENCE_RE = re.compile(
    r"(\d+\+?\s*(?:–|-|to)?\s*\d*\s*years?\s*(?:of\s+)?experience|"
    r"experience[:\s]*(\d+[\+\s\-]*\d*\s*years?)|"
    r"(?:minimum|min\.?)\s+\d+\s+years?)",re.I
)
HARD_BLOCKED = re.compile(
    r"(/legal/|/user.agreement|/privacy|/terms|/cookie|/uas/request.password"
    r"|/forgot.password|/sign.in|/login|/register|/signup|/sign.up"
    r"|/new.cv|/myworkspace|/dashboard|\.svg$|\.png$|\.jpg$|\.jpeg$"
    r"|\.gif$|\.ico$|/sitemap|/feed\.xml|javascript:|mailto:|tel:"
    r"|/cdn.cgi/access)",re.I
)
JOB_URL_ATS = re.compile(
    r"(/go/[A-Za-z0-9%-]+/\d+|/\d{6,}/?$|/job-detail/\d+"
    r"|/req/\d+|/posting/[A-Za-z0-9-]+|/application/\d+|/apply/\d+)",re.I
)
JOB_PATH_CORE = re.compile(r"/jobs?/",re.I)
DEFINITE_PATS = [
    r"\bcareers?\b",r"\bjobs?\b",r"join\s+us",r"work\s+with\s+us",
    r"we'?re?\s+hiring",r"open\s+positions?",r"current\s+openings?",
    r"وظائف",r"التوظيف",
]
_GARBAGE_FP = re.compile(
    r"^(job\s*description|key\s*accountabilities?|requirements?|qualifications?"
    r"|responsibilities|overview|role\s*purpose|job\s*purpose|about\s*the\s*role"
    r"|what\s*you|we\s*are\s*looking|others?|not\s*applicable|n/?a|—|-)$",re.I
)


def _is_likely_job_url(url, career_base=""):
    if not url or len(url)<10: return False
    lower = url.lower(); path = urlparse(url).path.lower()
    if HARD_BLOCKED.search(lower): return False
    if re.search(r"\.(svg|png|jpg|jpeg|gif|ico|pdf|zip|xml|json)$",path,re.I): return False
    if JOB_URL_ATS.search(path):
        if re.search(r"/go/(job.?search|all.?jobs)/\d+/?$",path,re.I): return False
        return True
    return bool(JOB_PATH_CORE.search(path))

def _is_detail_url(url):
    if not url or len(url)<10: return False
    path = urlparse(url).path.rstrip("/").lower()
    if HARD_BLOCKED.search(url.lower()): return False
    if re.search(r"\.(svg|png|jpg|jpeg|gif|ico|pdf|zip|xml)$",path,re.I): return False
    if re.search(r"/go/(job.?search|all.?jobs)/\d+/?$",path,re.I): return False
    if re.match(r"^/(jobs|job.search|openings|vacancies|all.jobs|opportunities|positions|search)$",path,re.I):
        return False
    return _is_likely_job_url(url)

def _sanitize(v, mx=300):
    if not v: return ""
    v = re.sub(r"\s+"," ",str(v)).strip()
    if _GARBAGE_FP.match(v): return ""
    return v[:mx] if len(v)>=2 else ""

def _job_signal_count(html):
    lower = html.lower()
    return sum(1 for s in JOB_PAGE_SIGNALS if s in lower)

def _is_ats(url):
    url = str(url).lower()
    if "linkedin.com" in url: return None
    for pat,name in ATS_DOMAINS.items():
        if pat in url: return name
    return None

def _is_ats_html(html):
    lower = html.lower()
    for fp,name in ATS_HTML_FP.items():
        if fp.lower() in lower: return name
    return None

def _is_career_link(href, text):
    combined = (str(href)+" "+str(text)).lower()
    if "linkedin.com" in combined: return False, None
    for pat in DEFINITE_PATS:
        if re.search(pat,text.lower()): return True,"definite"
    for kw in CAREER_KEYWORDS:
        if kw in combined: return True,"keyword"
    return False, None

def _extract_json_ld(soup):
    data = {}
    for script in soup.find_all("script",type="application/ld+json"):
        try:
            obj = json.loads(script.string or "{}")
            items = obj.get("@graph",[obj])
            for item in items:
                if "JobPosting" not in str(item.get("@type","")): continue
                data["title"]       = item.get("title","")
                data["description"] = item.get("description","")
                data["date_posted"] = item.get("datePosted","")
                data["deadline"]    = item.get("validThrough","")
                data["job_type"]    = item.get("employmentType","")
                bs = item.get("baseSalary",{})
                if isinstance(bs,dict):
                    val = bs.get("value",{})
                    if isinstance(val,dict):
                        mn,mx2,unit = val.get("minValue",""),val.get("maxValue",""),val.get("unitText","")
                        cur = bs.get("currency","")
                        data["salary_range"] = f"{cur} {mn}–{mx2} ({unit})".strip() if (mn or mx2) else ""
                loc = item.get("jobLocation",{})
                if isinstance(loc,dict):
                    addr = loc.get("address",{})
                    if isinstance(addr,dict):
                        parts = [addr.get("addressLocality",""),addr.get("addressRegion",""),addr.get("addressCountry","")]
                        data["location"] = ", ".join(p for p in parts if p)
                org = item.get("hiringOrganization",{})
                if isinstance(org,dict):
                    data["company_name"]  = org.get("name","")
                    data["company_logo"]  = org.get("logo","")
                    data["company_website"] = org.get("sameAs","")
                data["qualifications"] = item.get("qualifications","")
                data["experience"]     = item.get("experienceRequirements","")
                data["field"]          = item.get("occupationalCategory","") or item.get("industry","")
        except: pass
    return data

def _find_near_label(soup, *labels):
    for label in labels:
        pat = re.compile(re.escape(label),re.I)
        for dt in soup.find_all(["dt","th","strong","label","span","td","b"]):
            if pat.search(dt.get_text()):
                for nxt in [dt.find_next_sibling(), dt.find_next("td")]:
                    if nxt:
                        v = _sanitize(nxt.get_text())
                        if v: return clean(v)
    return ""

def _extract_logo(soup, base_url):
    og = soup.find("meta",property="og:image")
    if og and og.get("content"): return og["content"]
    for sel in ["img.logo","img[class*='logo']","img[alt*='logo']",
                ".logo img","header img",".navbar-brand img"]:
        el = soup.select_one(sel)
        if el and el.get("src"):
            src = el["src"]
            return src if src.startswith("http") else urljoin(base_url,src)
    return ""

def _pick_section(sections, *keys):
    for k in keys:
        kl = k.lower()
        for key,val in sections.items():
            if kl in key and val:
                v = _sanitize(val[:300])
                if v: return val
    return ""

def _parse_sections(soup):
    sections = {}
    for h in soup.find_all(["h1","h2","h3","h4"]):
        key = h.get_text(strip=True).lower()
        if not key or len(key)>120: continue
        parts = []
        for sib in h.find_next_siblings():
            if sib.name in ("h1","h2","h3","h4"): break
            t = sib.get_text(separator=" ",strip=True)
            if t: parts.append(t)
        sections[key] = " ".join(parts)
    return sections

def _collect_job_links(soup, base_url, career_domain):
    seen, results = set(), []
    for a in soup.find_all("a", href=True):
        href = a.get("href","")
        if not href or href.startswith(("#","mailto:","tel:","javascript")): continue
        full = href if href.startswith("http") else urljoin(base_url,href)
        if full in seen: continue
        seen.add(full)
        if not _is_likely_job_url(full, career_domain): continue
        text = a.get_text(strip=True)
        if not text or len(text)<3:
            parts = urlparse(full).path.rstrip("/").split("/")
            slug  = parts[-2] if (len(parts)>=2 and parts[-1].isdigit()) else parts[-1]
            text  = re.sub(r"[-_]"," ",slug).strip()
            text  = re.sub(r"\s+\d{5,}$","",text).strip()
        if text and len(text)>=3:
            results.append((full,text))
    return results

def _pagination(soup, base_url):
    urls = []
    for a in soup.find_all("a",href=True):
        text = a.get_text(strip=True).lower()
        href = a.get("href","")
        if not href: continue
        full = href if href.startswith("http") else urljoin(base_url,href)
        if any(kw in text for kw in ["next","›","»","load more","show more"]): urls.append(full)
        elif re.match(r"^\d+$",text) and int(text)<=50: urls.append(full)
    return list(dict.fromkeys(urls))

def _crawl_career_links(html, career_root):
    career_domain = get_domain(career_root)
    soup = BeautifulSoup(html,"html.parser")
    candidates = {}
    for a in soup.find_all("a",href=True):
        href = a.get("href",""); text = a.get_text(strip=True).lower()
        if not href or href.startswith(("#","mailto","tel")): continue
        full = href if href.startswith("http") else urljoin(career_root,href)
        if "linkedin.com" in full.lower(): continue
        if career_domain not in full.lower() and not _is_ats(full): continue
        score = 0
        if re.search(r"/(job.search.results?|openings?|positions?|listings?|vacancies|all.jobs|current.opening)",full.lower()): score+=8
        if re.search(r"/(job|career|role|vacanc|posting)s?[/_-]",full.lower()): score+=5
        if re.search(r"/(en|ar)/(job|career|vacanc|opening|position|search)",full.lower()): score+=7
        if re.search(r"\b(view|see|all|browse|search|explore)\b.*\b(job|position|opening|vacanc|career)",text): score+=6
        if re.search(r"\b(job|position|opening|vacanc|career)s?\b",text): score+=3
        if re.search(r"/go/[A-Za-z0-9%-]+/\d+",full.lower()): score+=10
        if score>=3: candidates[full] = max(candidates.get(full,0),score)
    for url,score in sorted(candidates.items(),key=lambda x:-x[1])[:1]:
        return url
    return None


async def _probe_suffixes(page, career_root, root_score=0):
    root = career_root.rstrip("/")
    best_url, best_score = None, root_score
    for suffix in JOB_LISTING_SUFFIXES:
        candidate = root+suffix
        if candidate.rstrip("/")==root: continue
        html = None; final_url = candidate
        r = simple_get(candidate)
        if r and r.url.rstrip("/")!=root:
            html = r.text; final_url = r.url
        else:
            try:
                resp = await page.goto(candidate,timeout=15000,wait_until="domcontentloaded")
                if resp and resp.status<400:
                    final_url = page.url
                    if final_url.rstrip("/")!=root:
                        await page.wait_for_timeout(1200)
                        html = await page.content()
            except: pass
        if not html: continue
        score = _job_signal_count(html)
        if score>best_score:
            best_score=score; best_url=final_url
    return best_url, best_score


async def _resolve_iframe_ats(page, html, base_url):
    soup = BeautifulSoup(html,"html.parser")
    for iframe in soup.find_all("iframe"):
        src = iframe.get("src","") or iframe.get("data-src","")
        if not src: continue
        full = src if src.startswith("http") else urljoin(base_url,src)
        if _is_ats(full): return full
    try:
        for frame in page.frames:
            furl = frame.url
            if furl and furl!=base_url and furl!="about:blank" and _is_ats(furl): return furl
    except: pass
    m = re.search(r'<iframe[^>]+src=["\']?(https?://[^"\'>\s]+)["\']?',html,re.I)
    if m and _is_ats(m.group(1)): return m.group(1)
    return None


async def _resolve_to_jobs_url(page, career_root, strategy, ats_name):
    if _is_blocked_career(career_root): return None, strategy+"+blocked", None
    if ats_name: return career_root, strategy, ats_name
    r_root = simple_get(career_root)
    root_html = r_root.text if r_root else ""
    root_score = _job_signal_count(root_html) if root_html else 0
    if root_html:
        det = _is_ats_html(root_html)
        if det: return career_root, strategy+"+html_fp", det
    suffix_url, suffix_score = await _probe_suffixes(page, career_root, root_score)
    if suffix_url: return suffix_url, strategy+"+suffix", None
    if root_score>=2: return career_root, strategy, ats_name
    if root_html:
        found = _crawl_career_links(root_html, career_root)
        if found: return found, strategy+"+crawl", None
    try:
        await page.goto(career_root, timeout=25000, wait_until="domcontentloaded")
        await page.wait_for_timeout(2500)
        html = await page.content()
        det = _is_ats_html(html)
        if det:
            m = re.search(r'(https?://[^\s"\'<>]*/go/[^\s"\'<>]+)',html,re.I)
            if m: return m.group(1), strategy+"+js_fp_go", det
            return career_root, strategy+"+js_fp", det
        for pat,aname in ATS_DOMAINS.items():
            if pat in html.lower():
                m = re.search(r'https?://[^\s"\'<>]*'+re.escape(pat)+r'[^\s"\'<>]*',html,re.I)
                if m: return m.group(0), strategy+"+js_ats", aname
        iframe_src = await _resolve_iframe_ats(page, html, career_root)
        if iframe_src: return iframe_src, strategy+"+js_iframe", _is_ats(iframe_src)
        su2,ss2 = await _probe_suffixes(page, career_root, max(root_score,_job_signal_count(html)))
        if su2: return su2, strategy+"+js_suffix", None
        if _job_signal_count(html)>=2: return career_root, strategy+"+js", None
        found = _crawl_career_links(html, career_root)
        if found: return found, strategy+"+js_crawl", None
    except: pass
    return career_root, strategy+"+unresolved", None


async def _find_career_page(page, website):
    base_domain, scheme = get_base(website)
    base = f"{scheme}://{base_domain}"
    for sub in CAREER_SUBDOMAINS:
        url = f"{scheme}://{sub}.{base_domain}"
        r = simple_get(url)
        if r and not _is_blocked_career(r.url):
            res = await _resolve_to_jobs_url(page, r.url, f"subdomain:{sub}", _is_ats(r.url))
            if res[0] and not _is_blocked_career(res[0]): return res
    for path in CAREER_PATHS:
        url = base+path
        r = simple_get(url)
        if r and r.url not in (base,base+"/",base+"/#") and not _is_blocked_career(r.url):
            res = await _resolve_to_jobs_url(page, r.url, f"path:{path}", _is_ats(r.url))
            if res[0] and not _is_blocked_career(res[0]): return res
    r = simple_get(website)
    if r:
        soup = BeautifulSoup(r.text,"html.parser")
        best_s, best_u, best_st, best_a = 0, None, None, None
        for a in soup.find_all("a",href=True):
            href = a.get("href",""); text = clean(a.get_text(),80).lower()
            if not href or href.startswith(("#","mailto","tel")): continue
            full = href if href.startswith("http") else urljoin(website,href)
            if _is_blocked_career(full): continue
            aname = _is_ats(full)
            if aname:
                return await _resolve_to_jobs_url(page, full, f"ats_link:{aname}", aname)
            ok,reason = _is_career_link(href,text)
            if ok:
                score = (10 if reason=="definite" else 5)+(3 if base_domain in full else 0)
                if score>best_s: best_s,best_u,best_st = score,full,f"link:{reason}"
        if best_u and best_s>=5 and not _is_blocked_career(best_u):
            return await _resolve_to_jobs_url(page, best_u, best_st, None)
    try:
        await page.goto(website, timeout=25000, wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)
        content = await page.content()
        det = _is_ats_html(content)
        if det:
            m_go = re.search(r'(https?://[^\s"\'<>]*/go/[A-Za-z0-9%-]+/\d+[^\s"\'<>]*)',content,re.I)
            if m_go: return m_go.group(1), f"embedded:SF_go", det
            return website, f"embedded:{det}", det
        for pat,aname in ATS_DOMAINS.items():
            if pat in content.lower():
                m = re.search(r'https?://[^\s"\'<>]*'+re.escape(pat)+r'[^\s"\'<>]*',content,re.I)
                if m and not _is_blocked_career(m.group(0)):
                    return m.group(0), f"embedded:{aname}", aname
        iframe_src = await _resolve_iframe_ats(page, content, website)
        if iframe_src and not _is_blocked_career(iframe_src):
            return iframe_src, f"iframe:{_is_ats(iframe_src) or 'unknown'}", _is_ats(iframe_src)
    except: pass
    return None, "not_found", None


async def _scrape_wiki_job_detail(page, job_url, company, website, industry,
                                   careers_url, prefill_title="",
                                   prefill_location="", prefill_date=""):
    html = None; final_url = job_url
    r = simple_get(job_url, timeout=12)
    if r:
        html = r.text; final_url = r.url
    else:
        try:
            await page.goto(job_url, timeout=25000, wait_until="domcontentloaded")
            await page.wait_for_timeout(1800)
            html = await page.content(); final_url = page.url
        except: return None
    if not html: return None

    soup = BeautifulSoup(html,"html.parser")
    ld   = _extract_json_ld(soup)
    sections = _parse_sections(soup)
    for tag in soup.select("nav,footer,header,script,style,iframe,.cookie"):
        tag.decompose()
    full_text = soup.get_text(separator="\n")

    title = (ld.get("title") or prefill_title
             or (soup.select_one("h1").get_text(strip=True) if soup.select_one("h1") else ""))
    if title and "|" in title:  title = title.split("|")[0].strip()
    if title and "-" in title and len(title)>60: title = title.split("-")[0].strip()

    job_type = _sanitize(
        ld.get("job_type") or
        _find_near_label(soup,"Employment Type","Job Type","Contract Type","نوع الوظيفة") or
        _pick_section(sections,"employment type","job type") or ""
    )
    if not job_type:
        for jt in ["Full-time","Part-time","Contract","Internship","Freelance","Temporary","Permanent"]:
            if re.search(r"\b"+jt.split("-")[0]+r"\b",full_text,re.I):
                job_type=jt; break

    location = (ld.get("location") or prefill_location or
                _find_near_label(soup,"Location","Job Location","City","الموقع","المدينة") or "")
    if not location:
        m = LOCATION_PATTERN.search(full_text)
        if m: location = m.group(1)
    location = location or "Saudi Arabia"

    description = (ld.get("description","") or
                   _pick_section(sections,"responsibilities","job purpose","about the role",
                                  "what you'll do","role overview","job summary","overview",
                                  "المهام","المسؤوليات") or "")
    if not description:
        for sel in ["[class*='description']","[class*='job-desc']","[class*='content']",
                    "#job-description","article","main","[class*='detail']"]:
            el = soup.select_one(sel)
            if el and len(el.get_text())>100:
                description = clean(el.get_text(), 2000); break
    if not description and len(full_text)>200:
        description = clean(full_text, 2000)

    qualifications = (ld.get("qualifications","") or
                      _pick_section(sections,"qualif","requirement","education",
                                     "you should have","minimum requirement","المؤهلات") or
                      _find_near_label(soup,"Qualifications","Requirements","Education","المؤهلات") or "")
    experience = _sanitize(
        ld.get("experience","") or
        _find_near_label(soup,"Experience","Years of Experience","الخبرة","سنوات الخبرة") or
        _pick_section(sections,"experience") or ""
    )
    if not experience:
        m_exp = EXPERIENCE_RE.search(full_text)
        if m_exp: experience = m_exp.group(0)

    department = _sanitize(ld.get("field","") or
                            _find_near_label(soup,"Department","Team","Function","القسم") or "")
    field = _sanitize(department or _find_near_label(soup,"Field","Category","التخصص") or industry)

    date_posted = ld.get("date_posted","") or prefill_date
    if not date_posted:
        m_dp = DATE_POSTED_RE.search(full_text)
        if m_dp: date_posted = next((g for g in m_dp.groups() if g),"")
        if not date_posted:
            tel = soup.find("time")
            if tel: date_posted = tel.get("datetime") or tel.get_text(strip=True)

    deadline = ld.get("deadline","")
    if not deadline:
        m_dl = DEADLINE_RE.search(full_text)
        if m_dl: deadline = next((g for g in m_dl.groups() if g),"")
    est_dl = deadline or estimate_deadline(date_posted)
    if est_dl and not re.match(r"\d{4}-\d{2}-\d{2}",est_dl.strip()): est_dl=""

    salary_range = (ld.get("salary_range","") or
                    _find_near_label(soup,"Salary","Compensation","الراتب") or "")
    if not salary_range:
        m_sal = SALARY_RE.search(full_text)
        if m_sal: salary_range = m_sal.group(0)

    company_logo    = ld.get("company_logo","") or _extract_logo(soup, final_url)
    company_type    = _sanitize(ld.get("company_type","") or
                                 _find_near_label(soup,"Company Type","Organization Type","نوع الشركة") or "")
    if company_type and company_type.lower() in ("organization","legalservice","thing"):
        company_type=""
    company_address = _find_near_label(soup,"Address","Headquarters","Head Office","العنوان")
    company_founded = _sanitize(_find_near_label(soup,"Founded","Established","Year Founded","تأسست"))
    company_details = _pick_section(sections,"about the company","about us","company overview","من نحن")

    apply_url = final_url
    for a in soup.find_all("a",href=True):
        href = a.get("href",""); text_a = a.get_text(strip=True).lower()
        if any(p in href.lower() for p in ["talentcommunity/apply","/apply/","?action=apply"]):
            apply_url = href if href.startswith("http") else urljoin(final_url,href); break
        if "apply" in text_a and href and href not in ("#","javascript:void(0)"):
            candidate = href if href.startswith("http") else urljoin(final_url,href)
            if candidate!=final_url: apply_url=candidate

    return make_record(
        **{
            "Job Title":       title,
            "Job Type":        job_type,
            "Job Qualifications": standardise_qualification(qualifications, description),
            "Job Experience":  standardise_experience(experience),
            "Job Location":    location,
            "Job Field":       standardise_field(field, title, description, industry),
            "Date Posted":     date_posted,
            "Deadline":        deadline or est_dl,
            "Job Description": description,
            "Application":     apply_url,
            "Company URL":     careers_url,
            "Company Name":    ld.get("company_name","") or company,
            "Company Logo":    company_logo,
            "Company Industry":industry,
            "Company Founded": company_founded,
            "Company Type":    company_type,
            "Company Website": website,
            "Company Address": company_address,
            "Company Details": company_details,
            "Job URL":         apply_url,
            "Estimated Deadline": est_dl,
            "Salary Range":    salary_range,
        }
    )


async def _extract_wiki_jobs(page, careers_url, company, website, industry, ats_name):
    stub_jobs = []

    if ats_name=="Greenhouse" or "greenhouse.io" in careers_url:
        m = re.search(r"greenhouse\.io/([^/?#]+)",careers_url)
        if m:
            r = simple_get(f"https://boards.greenhouse.io/{m.group(1)}/jobs.json")
            if r:
                try:
                    for j in r.json().get("jobs",[]):
                        stub_jobs.append({"url":j.get("absolute_url",""),"title":j.get("title","")})
                except: pass

    if ats_name=="Lever" or "lever.co" in careers_url:
        m = re.search(r"lever\.co/([^/?#]+)",careers_url)
        if m:
            r = simple_get(f"https://api.lever.co/v0/postings/{m.group(1)}?mode=json")
            if r:
                try:
                    for j in r.json():
                        stub_jobs.append({"url":j.get("hostedUrl",""),"title":j.get("text","")})
                except: pass

    if ats_name=="Ashby" or "ashbyhq.com" in careers_url:
        m = re.search(r"ashbyhq\.com/([^/?#]+)",careers_url)
        if m:
            r = simple_get(f"https://api.ashbyhq.com/posting-public/job-board/all?organizationHostedJobsPageName={m.group(1)}")
            if r:
                try:
                    for j in r.json().get("jobPostings",[]):
                        stub_jobs.append({"url":j.get("jobUrl",""),"title":j.get("title","")})
                except: pass

    if ats_name=="SmartRecruiters" or "smartrecruiters.com" in careers_url:
        m = re.search(r"smartrecruiters\.com/([^/?#]+)",careers_url)
        if m:
            r = simple_get(f"https://api.smartrecruiters.com/v1/companies/{m.group(1)}/postings")
            if r:
                try:
                    for j in r.json().get("content",[]):
                        stub_jobs.append({"url":f"https://jobs.smartrecruiters.com/{m.group(1)}/{j.get('id','')}","title":j.get("name","")})
                except: pass

    if not stub_jobs:
        try:
            await page.goto(careers_url, timeout=30000, wait_until="networkidle")
            await page.wait_for_timeout(2500)
            for _ in range(3):
                await page.evaluate("window.scrollBy(0,800)")
                await page.wait_for_timeout(700)
            html = await page.content()
            soup = BeautifulSoup(html,"html.parser")
            for tag in soup.select("nav,footer,header,script,style"): tag.decompose()
            career_domain = get_domain(careers_url)
            job_links = _collect_job_links(soup, careers_url, career_domain)
            vlog(f"   {len(job_links)} job link candidates", indent=2)
            for jurl,jtext in job_links[:200]:
                stub_jobs.append({"url":jurl,"title":jtext,"location":"","date_posted":""})
            if not stub_jobs:
                for nxt in _pagination(soup, careers_url)[:4]:
                    await page.goto(nxt, timeout=20000, wait_until="networkidle")
                    await page.wait_for_timeout(1500)
                    nsoup = BeautifulSoup(await page.content(),"html.parser")
                    for jurl,jtext in _collect_job_links(nsoup,nxt,career_domain):
                        stub_jobs.append({"url":jurl,"title":jtext,"location":"","date_posted":""})
        except Exception as e:
            vlog(f"DOM error: {e}", indent=2)

    stub_jobs = [s for s in stub_jobs if _is_detail_url(s.get("url",""))]
    if not stub_jobs: return []

    jobs = []; seen_detail = set()
    for stub in stub_jobs[:200]:
        jurl = stub.get("url","")
        if not jurl or jurl in seen_detail: continue
        seen_detail.add(jurl)
        detail = await _scrape_wiki_job_detail(
            page, jurl, company, website, industry, careers_url,
            prefill_title=stub.get("title",""),
            prefill_location=stub.get("location",""),
            prefill_date=stub.get("date_posted",""),
        )
        if detail:
            jobs.append(detail)
        else:
            jobs.append(make_record(**{
                "Job Title":    stub.get("title",""),
                "Job Location": stub.get("location","") or "Saudi Arabia",
                "Date Posted":  stub.get("date_posted",""),
                "Application":  jurl,
                "Job URL":      jurl,
                "Company Name": company,
                "Company Website": website,
                "Company Industry": industry,
                "Company URL":  careers_url,
            }))
        await asyncio.sleep(random.uniform(0.4,1.0))
    return jobs


async def _process_wiki_company(page, name, website, industry, cp):
    if not _looks_like_saudi_company(name, website):
        vlog(f"⏭  Non-Saudi skip: {name}", indent=0); return

    domain = get_domain(website)
    if not domain or domain in discovered_domains: return
    discovered_domains.add(domain)
    if domain in cp.get("processed_domains",[]):
        vlog(f"⏭  Already done: {name} ({domain})"); return

    vlog(f"\n{'═'*60}")
    vlog(f"  🏢  {name}  |  {website}  |  {industry}")

    careers_url, strategy, ats_name = await _find_career_page(page, website)

    if careers_url and _is_blocked_career(careers_url):
        vlog(f"  ⏭  Career page blocked ({get_domain(careers_url)})", indent=1)
        careers_url = None

    if not careers_url:
        vlog(f"  ❌  No career page", indent=1)
        cp.setdefault("processed_domains",[]).append(domain)
        save_checkpoint(cp)
        company_results.append({"name":name,"website":website,"industry":industry,
                                  "careers_url":None,"ats":None,"jobs_found":0})
        save_output()
        return

    vlog(f"  ✅  {careers_url}", indent=1)
    vlog(f"  🔧  {strategy}  |  ATS: {ats_name or 'Custom'}", indent=1)

    jobs = await _extract_wiki_jobs(page, careers_url, name, website, industry, ats_name)

    added = 0
    for rec in jobs:
        if register_job(rec, source="wikipedia"):
            added += 1

    cp.setdefault("processed_domains",[]).append(domain)
    cp["jobs_count"] = cp.get("jobs_count",0) + added
    save_checkpoint(cp)
    company_results.append({"name":name,"website":website,"industry":industry,
                              "careers_url":careers_url,"ats":ats_name or "Custom",
                              "jobs_found":added})
    save_output()
    vlog(f"  📋  {added} new jobs registered (from {len(jobs)} found)", indent=1)
    vlog(f"{'═'*60}")


WIKIPEDIA_PAGES = [
    "https://en.wikipedia.org/wiki/Category:Companies_of_Saudi_Arabia",
    "https://en.wikipedia.org/wiki/Category:Banks_of_Saudi_Arabia",
    "https://en.wikipedia.org/wiki/Category:Technology_companies_of_Saudi_Arabia",
    "https://en.wikipedia.org/wiki/Category:Oil_companies_of_Saudi_Arabia",
    "https://en.wikipedia.org/wiki/Category:Retail_companies_of_Saudi_Arabia",
    "https://en.wikipedia.org/wiki/Category:Construction_companies_of_Saudi_Arabia",
    "https://en.wikipedia.org/wiki/List_of_companies_of_Saudi_Arabia",
    "https://en.wikipedia.org/wiki/List_of_largest_companies_in_Saudi_Arabia",
]


async def run_wikipedia_pipeline(page, cp):
    if "wikipedia" in cp.get("done_sources",[]):
        print("⏭  Wikipedia: already done"); return
    print(C_H("\n"+"="*72))
    print(C_H("  SOURCE B — Wikipedia company discovery → career page scraper"))
    print(C_H("="*72))

    for wiki_url in tqdm(WIKIPEDIA_PAGES, desc="  Wiki pages"):
        r = simple_get(wiki_url)
        if not r: continue
        soup = BeautifulSoup(r.text,"html.parser")
        for link in soup.select(".mw-category a, .wikitable a, #mw-content-text a"):
            title = link.get("title",""); href = link.get("href","")
            if not title or "Category:" in title or "List" in title: continue
            if NON_COMPANY_WIKI.match(title.strip()): continue
            if not href.startswith("/wiki/"): continue
            await asyncio.sleep(random.uniform(0.5,1.5))
            wr = simple_get("https://en.wikipedia.org"+href)
            if not wr: continue
            wsoup = BeautifulSoup(wr.text,"html.parser")
            web_tag = wsoup.select_one(".infobox a.external")
            if not web_tag: continue
            raw_url = web_tag.get("href","")
            domain_check = get_domain(raw_url)
            if domain_check in NON_SAUDI_DOMAINS: continue
            domain, website = parse_domain(raw_url)
            if not domain: continue
            cats = [c.get_text() for c in wsoup.select("#mw-normal-catlinks a")]
            industry = "General"
            for cat in cats:
                for ind in INDUSTRIES:
                    if ind.lower() in cat.lower(): industry=ind; break
            await _process_wiki_company(page, title, website, industry, cp)
        await asyncio.sleep(random.uniform(1,2))

    cp.setdefault("done_sources",[]).append("wikipedia")
    save_checkpoint(cp)


# ═══════════════════════════════════════════════════════════════
# ▌ MAIN
# ═══════════════════════════════════════════════════════════════
async def main():
    global all_jobs, company_results

    print(C_H("="*72))
    print(C_H("  Saudi Jobs Pipeline — Unified  (LinkedIn + Wikipedia)"))
    print(C_H("="*72))

    cp = load_checkpoint()

    # Resume from previous run
    if Path(OUTPUT_FILE).exists():
        df_prev = pd.read_excel(OUTPUT_FILE)
        for _, row in df_prev.iterrows():
            rec = {col: str(row.get(col,"") or "") for col in COLUMNS}
            if rec.get("Job Title"):
                ju = _canonical_url(rec.get("Job URL","") or rec.get("Application",""))
                if ju: url_seen.add(ju)
                fp_seen.add(_fingerprint(rec))
                all_jobs.append(rec)
        print(f"   ▶  Resumed — {len(all_jobs)} jobs already in output")

    discovered_domains.update(cp.get("processed_domains",[]))

    # ── Source A : LinkedIn ──────────────────────────────────
    run_linkedin_pipeline()

    # ── Source B : Wikipedia → career pages ─────────────────
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=["--no-sandbox","--disable-dev-shm-usage","--disable-gpu"]
        )
        context = await browser.new_context(
            user_agent=HEADERS["User-Agent"],
            locale="en-US",
            viewport={"width":1280,"height":800},
        )
        page = await context.new_page()
        await page.route(
            "**/*.{png,jpg,jpeg,gif,svg,woff,woff2,ttf,eot,ico}",
            lambda r: r.abort()
        )
        await run_wikipedia_pipeline(page, cp)
        await browser.close()

    # ── Final save ───────────────────────────────────────────
    save_output()
    if company_results:
        pd.DataFrame(company_results).to_csv(COMPANIES_FILE, index=False)

    # ── Summary ──────────────────────────────────────────────
    total = len(all_jobs)
    W = 72
    print(f"\n{'═'*W}")
    print(C_H("  PIPELINE COMPLETE"))
    print(f"{'─'*W}")
    print(f"  {'Total unique jobs:':<32} {total:>6,}")
    if all_jobs:
        df = pd.DataFrame(all_jobs, columns=COLUMNS)
        with_apply   = df["Application"].astype(bool).sum()
        with_salary  = df["Salary Range"].astype(bool).sum()
        with_desc    = df["Job Description"].astype(bool).sum()
        print(f"  {'  with apply link:':<32} {with_apply:>6,}  ({100*with_apply//max(total,1)}%)")
        print(f"  {'  with salary data:':<32} {with_salary:>6,}  ({100*with_salary//max(total,1)}%)")
        print(f"  {'  with description:':<32} {with_desc:>6,}  ({100*with_desc//max(total,1)}%)")
        print(f"{'─'*W}")
        sources = df.get("source", pd.Series(dtype=str)) if "source" in df else None
        print(f"\n  Top companies by jobs:")
        top = df.groupby("Company Name")["Job Title"].count().sort_values(ascending=False).head(15)
        for co,cnt in top.items():
            print(f"    {co[:35]:<35}  {cnt:>4}")
        if "Job Field" in df.columns:
            print(f"\n  Top job fields:")
            for field,cnt in df["Job Field"].dropna().replace("",pd.NA).dropna().value_counts().head(10).items():
                print(f"    {field[:35]:<35}  {cnt:>4}")
    print(f"\n  Output  : {OUTPUT_FILE}")
    print(f"  Companies: {COMPANIES_FILE}")
    print(f"{'═'*W}")

    try:
        from google.colab import files
        files.download(OUTPUT_FILE)
        files.download(COMPANIES_FILE)
    except: pass


asyncio.run(main())

