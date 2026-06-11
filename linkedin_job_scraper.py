# ═══════════════════════════════════════════════════════════════
# SAUDI JOBS PIPELINE — v6  (MERGED)
# Discovery: Wikipedia only
# Sources  : Company career pages (v5) + LinkedIn Guest API (v2)
# Dedup    : Cross-source deduplication by title+company+location
#
# v6 CHANGES (merged from v5 + LinkedIn v2):
#  1. WIKIPEDIA ONLY — DDG, Google Maps, Kompass removed.
#  2. DUAL SOURCE per company — after finding a career page,
#     ALSO searches LinkedIn guest API for the same company.
#  3. CROSS-SOURCE DEDUP — jobs from both sources are merged;
#     duplicates detected by (title, company, location) fingerprint.
#     When a duplicate is found, fields are merged (career page
#     detail enriches the LinkedIn stub, or vice-versa).
#  4. CLEAN OUTPUT — single CSV with unified 22-column schema.
# ═══════════════════════════════════════════════════════════════

import subprocess, sys, os

print("📦 Installing dependencies...")
subprocess.run(["apt-get", "install", "-y",
    "libatk1.0-0", "libatk-bridge2.0-0", "libcups2", "libdrm2",
    "libxkbcommon0", "libxcomposite1", "libxdamage1", "libxfixes3",
    "libxrandr2", "libgbm1", "libasound2"],
    capture_output=True)

subprocess.run([sys.executable, "-m", "pip", "install",
    "playwright", "pandas", "beautifulsoup4",
    "requests", "nest_asyncio", "tqdm", "openpyxl", "-q"], check=True)

subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"],
    capture_output=True)

import importlib, site
importlib.invalidate_caches()
for sp in site.getsitepackages():
    if sp not in sys.path:
        sys.path.insert(0, sp)

import asyncio, re, json, random, time, csv, base64, logging
from pathlib import Path
from urllib.parse import urlparse, urljoin, quote_plus, unquote
from datetime import datetime, timedelta

import requests
import pandas as pd
import nest_asyncio
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright
from tqdm import tqdm

nest_asyncio.apply()
print("✅ All imports successful\n")

logging.basicConfig(level=logging.WARNING, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════
# ▌ SAUDI ARABIA GEOGRAPHIC FILTER
# ══════════════════════════════════════════════════════════════

NON_SAUDI_DOMAINS = {
    "fortune.com", "forbes.com", "bloomberg.com", "reuters.com",
    "arabnews.com", "aljazeera.com", "bbc.com", "bbc.co.uk",
    "cnn.com", "nytimes.com", "wsj.com", "ft.com",
    "c-span.org", "cspan.org", "pbs.org", "npr.org",
    "federalreserve.gov", "sec.gov", "state.gov", "treasury.gov",
    "loc.gov", "whitehouse.gov", "congress.gov",
    "londonstockexchange.com", "lseg.com", "nyse.com", "nasdaq.com",
    "imf.org", "worldbank.org",
    "opec.org", "un.org", "unesco.org", "ich.unesco.org",
    "pepsico.com", "middleeast.pepsico.com",
    "linkedin.com", "facebook.com", "twitter.com", "x.com",
    "instagram.com", "tiktok.com", "snapchat.com",
    "amazon.com", "google.com", "microsoft.com", "apple.com",
    "hsbc.com", "barclays.com",
    "houseofsaud.com",
}

SAUDI_DOMAIN_INDICATORS = [
    r"\.sa$", r"\.gov\.sa$", r"\.edu\.sa$", r"\.com\.sa$", r"\.org\.sa$",
    r"aramco", r"sabic", r"stc\.com", r"alrajhi", r"samba", r"riyad",
    r"ncb", r"jarir", r"maaden", r"tasnee", r"mobily", r"zain\.sa",
    r"saudia", r"flynas", r"flyadeal", r"neom", r"vision2030", r"pif\.gov",
]

NON_COMPANY_WIKI_PATTERNS = re.compile(
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


def _is_saudi_domain(domain: str) -> bool:
    d = domain.lower()
    if d in NON_SAUDI_DOMAINS:
        return False
    for pat in SAUDI_DOMAIN_INDICATORS:
        if re.search(pat, d):
            return True
    return False


def _looks_like_saudi_company(name: str, website: str) -> bool:
    domain = get_domain(website)
    if domain in NON_SAUDI_DOMAINS:
        return False
    if NON_COMPANY_WIKI_PATTERNS.match(name.strip()):
        return False
    return True


# ══════════════════════════════════════════════════════════════
# ▌ STANDARDISATION TABLES  (shared by both sources)
# ══════════════════════════════════════════════════════════════

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
      "operations technician","broadcast technician","studio technician",
      "manufacturing engineer","process technician"],
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
      "development officer","fundraiser","volunteer coordinator",
      "network associate"],
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


def standardise_field(raw_field: str, title: str = "", description: str = "",
                       industry: str = "") -> str:
    combined = " ".join([
        (raw_field or ""), (title or ""), (description or "")[:800],
        (industry or ""),
    ]).lower()
    best_label, best_score = "", 0
    for label, high_kws, low_kws in FIELD_KEYWORD_MAP:
        score = sum(3 for kw in high_kws if kw in combined)
        score += sum(1 for kw in low_kws if kw in combined)
        if score > best_score:
            best_score, best_label = score, label
    return best_label if best_score >= 1 else ""


# ── EXPERIENCE ────────────────────────────────────────────────
_NO_EXP_KW = [
    "no experience","no prior experience","fresh graduate","freshers",
    "entry level","entry-level","0 years","zero experience",
    "training provided","will train","no experience required",
    "no experience needed",
]
_LT1_KW = [
    "less than 1 year","under 1 year","6 months","less than a year",
    "some experience","minimal experience","up to 1 year",
]
_EXP_RE = re.compile(
    r"(?:minimum|min\.?|at\s+least|over|more\s+than)?\s*"
    r"(\d+)\s*(?:\+|plus)?\s*"
    r"(?:[-–to]+\s*(\d+))?\s*"
    r"years?(?:\s+of)?(?:\s+(?:relevant\s+)?experience)?",
    re.I,
)


def _years_to_band(n: int) -> str:
    if n <= 0:  return "No Experience Required"
    if n <= 2:  return "1 - 2 Years"
    if n <= 5:  return "3 - 5 Years"
    if n <= 10: return "6 - 10 Years"
    return "10+ Years"


def standardise_experience(raw: str) -> str:
    if not raw:
        return ""
    text = raw.lower().strip()
    for kw in _NO_EXP_KW:
        if kw in text: return "No Experience Required"
    for kw in _LT1_KW:
        if kw in text: return "Less than 1 Year"
    matches = _EXP_RE.findall(text)
    if matches:
        nums = [int(g) for m in matches for g in m if g]
        if nums: return _years_to_band(min(nums))
    m = re.search(r"(\d+)\s*\+?\s*years?", text, re.I)
    if m: return _years_to_band(int(m.group(1)))
    if re.search(r"\b(senior|sr\.?|lead|principal|head of|director|vp|vice president)\b", text, re.I):
        return "6 - 10 Years"
    if re.search(r"\b(mid.?level|intermediate|associate)\b", text, re.I):
        return "3 - 5 Years"
    if re.search(r"\b(junior|jr\.?|graduate|intern|trainee|fresh)\b", text, re.I):
        return "Less than 1 Year"
    return ""


# ── QUALIFICATION ─────────────────────────────────────────────
QUALIFICATION_TIERS = [
    ("PhD / Doctorate",
     ["phd","ph.d","doctorate","doctoral","doctor of philosophy"]),
    ("Master's Degree",
     ["master","msc","m.sc","mba","m.b.a","meng","m.eng","mphil",
      "postgraduate","post-graduate","post graduate"]),
    ("Bachelor's Degree",
     ["bachelor","bsc","b.sc","beng","b.eng","bcom","b.com","bba",
      "llb","degree in","undergraduate degree","honours degree","hons",
      "b.tech","btech"]),
    ("Higher National Diploma",
     ["hnd","hnc","higher national diploma","higher national certificate",
      "higher diploma","advanced diploma"]),
    ("Diploma",
     ["diploma","associate degree","foundation degree"]),
    ("Professional Certification",
     ["acca","cpa","cfa","cima","pmp","prince2","cissp","aws certified",
      "comptia","cisco","ccna","ccnp","shrm","cipd","chartered",
      "certified public","certified financial","certified project",
      "professional certification","professional certificate"]),
    ("A-Levels / HSC",
     ["a-level","a level","hsc","higher school certificate",
      "ib diploma","international baccalaureate","gce advanced"]),
    ("O-Levels / School Certificate",
     ["o-level","o level","igcse","gcse","school certificate"]),
    ("No Formal Qualification Required",
     ["no qualification","no degree","no formal","school leaver",
      "no experience required","training provided","will train"]),
]


def standardise_qualification(raw: str, full_text: str = "") -> str:
    corpus = ((raw or "") + " " + (full_text or "")[:2000]).lower()
    for label, keywords in QUALIFICATION_TIERS:
        for kw in keywords:
            if kw in corpus:
                return label
    return ""


# ══════════════════════════════════════════════════════════════
# ▌ VERBOSE LOGGING
# ══════════════════════════════════════════════════════════════
VERBOSE = True

_stats = {
    "companies_seen": 0,
    "companies_with_careers": 0,
    "companies_no_careers": 0,
    "companies_skipped_non_saudi": 0,
    "jobs_total": 0,
    "jobs_from_career_pages": 0,
    "jobs_from_linkedin": 0,
    "jobs_duplicates_merged": 0,
    "jobs_with_salary": 0,
    "jobs_with_description": 0,
    "detail_fetches": 0,
    "detail_failures": 0,
    "ats_hits": {},
    "strategy_hits": {},
}

_JOB_COUNTER = [0]


def _ts():
    return datetime.now().strftime("%H:%M:%S")


def vlog(msg, indent=0):
    print(f"[{_ts()}] {'   ' * indent}{msg}", flush=True)


def vprint(msg, indent=0):
    if VERBOSE:
        vlog(msg, indent)


def _bar(filled, total, width=20):
    if total == 0:
        return "[" + "─" * width + "] 0/0"
    n = int(width * filled / total)
    return f"[{'█' * n}{'─' * (width - n)}] {filled}/{total}"


def print_job_card(job: dict, serial: int, source_tag: str = ""):
    if not VERBOSE:
        return
    W = 72
    SEP = "─" * W

    def row(label, value, icon=""):
        label_str = f"{icon} {label}:" if icon else f"   {label}:"
        val_str   = str(value).strip() if value else "—"
        max_val   = W - len(label_str) - 3
        if len(val_str) > max_val:
            val_str = val_str[:max_val - 1] + "…"
        return f"  {label_str:<26} {val_str}"

    src = f"[{source_tag}]" if source_tag else ""
    print(f"\n  ┌{SEP}┐")
    print(f"  │  JOB #{serial:04d}  {src}" + " " * max(0, W - 14 - len(src)) + "│")
    print(f"  ├{SEP}┤")
    print(f"  │{row('Title',         job.get('Job Title'),          '📌'):<{W+2}}│")
    print(f"  │{row('Type',          job.get('Job Type'),           '📋'):<{W+2}}│")
    print(f"  │{row('Location',      job.get('Job Location'),       '📍'):<{W+2}}│")
    print(f"  │{row('Field',         job.get('Job Field'),          '🏷️'):<{W+2}}│")
    print(f"  │{row('Experience',    job.get('Job Experience'),     '🎯'):<{W+2}}│")
    print(f"  │{row('Qualification', job.get('Job Qualifications'), '🎓'):<{W+2}}│")
    print(f"  │{row('Salary Range',  job.get('Salary Range'),       '💰'):<{W+2}}│")
    print(f"  │{row('Date Posted',   job.get('Date Posted'),        '📅'):<{W+2}}│")
    print(f"  │{row('Deadline',      job.get('Deadline'),           '⏰'):<{W+2}}│")
    print(f"  │{row('Est. Deadline', job.get('Estimated Deadline'), '🗓️'):<{W+2}}│")
    print(f"  ├{SEP}┤")
    desc = (job.get("Job Description") or "").strip()
    lines = [desc[i:i+W-4] for i in range(0, min(len(desc), (W-4)*3), W-4)] if desc else []
    print(f"  │  📝 Description:  {' ' * (W - 19)}│")
    for ln in lines:
        print(f"  │    {ln:<{W-4}}│")
    if not lines:
        print(f"  │    {'—':<{W-4}}│")
    print(f"  ├{SEP}┤")
    print(f"  │{row('Company',       job.get('Company Name'),       '🏢'):<{W+2}}│")
    print(f"  │{row('Industry',      job.get('Company Industry'),   '🏭'):<{W+2}}│")
    print(f"  │{row('Careers URL',   job.get('Company URL'),        '🔗'):<{W+2}}│")
    print(f"  ├{SEP}┤")
    apply = job.get("Application") or job.get("Job URL") or ""
    print(f"  │  🚀 Apply: {apply[:W-11]:<{W-11}}│")
    print(f"  │  🔧 Source: {job.get('source',''):<{W-12}}│")
    print(f"  └{SEP}┘")


def print_company_header(name, website, industry, n):
    W = 72
    print(f"\n{'═'*W}")
    print(f"  🏢  #{n}  {name}")
    print(f"  🌐  {website}   |   🏭  {industry}")
    print(f"{'─'*W}")


def print_company_summary(name, careers_url, strategy, ats_name,
                           career_jobs, li_jobs, merged):
    W = 72
    total = career_jobs + li_jobs - merged
    print(f"  {'─'*W}")
    if careers_url:
        print(f"  ✅  Career page : {careers_url[:W-18]}")
    else:
        print(f"  ❌  No career page found")
    print(f"  🔧  Strategy    : {strategy}")
    print(f"  🤖  ATS         : {ats_name or 'Custom'}")
    print(f"  📋  Career jobs : {career_jobs}   LinkedIn jobs: {li_jobs}   "
          f"Merged duplicates: {merged}   Net new: {total}")
    print(f"{'═'*W}")


def print_live_stats():
    s = _stats
    W = 72
    print(f"\n  ┌── LIVE STATS {'─'*(W-15)}┐")
    print(f"  │  Companies : seen={s['companies_seen']}  "
          f"careers={s['companies_with_careers']}  "
          f"no_careers={s['companies_no_careers']}  "
          f"skipped={s['companies_skipped_non_saudi']}")
    print(f"  │  Jobs      : total={s['jobs_total']}  "
          f"career_page={s['jobs_from_career_pages']}  "
          f"linkedin={s['jobs_from_linkedin']}  "
          f"merged_dupes={s['jobs_duplicates_merged']}")
    print(f"  │  Quality   : salary={s['jobs_with_salary']}  "
          f"desc={s['jobs_with_description']}")
    if s["ats_hits"]:
        ats_str = "  ".join(
            f"{k}:{v}" for k, v in sorted(s["ats_hits"].items(), key=lambda x: -x[1])
        )
        print(f"  │  ATS hits  : {ats_str}")
    print(f"  └{'─'*W}┘")


# ══════════════════════════════════════════════════════════════
# ▌ OUTPUT FILES & SCHEMA
# ══════════════════════════════════════════════════════════════
JOBS_FILE       = "saudi_jobs_v6.csv"
COMPANIES_FILE  = "saudi_companies_v6.csv"
CHECKPOINT_FILE = "pipeline_v6_checkpoint.json"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
}

# Rotating user-agents for LinkedIn requests
_LI_USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_4) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
]
_ua_idx = [0]


def _next_li_headers() -> dict:
    ua = _LI_USER_AGENTS[_ua_idx[0] % len(_LI_USER_AGENTS)]
    _ua_idx[0] += 1
    return {
        "User-Agent": ua,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Cache-Control": "no-cache",
        "X-Li-Lang": "en_US",
        "X-Requested-With": "XMLHttpRequest",
    }


SAUDI_CITIES = [
    "Riyadh", "Jeddah", "Dammam", "Mecca", "Medina", "Khobar",
    "Tabuk", "Abha", "Jubail", "Yanbu", "Taif", "Buraidah",
    "Khamis Mushait", "Hail", "Najran", "Jizan", "Dhahran",
]

INDUSTRIES = [
    "technology", "software", "fintech", "banking", "finance",
    "healthcare", "hospital", "construction", "real estate",
    "retail", "manufacturing", "oil gas", "energy", "telecom",
    "logistics", "education", "hospitality", "consulting",
    "engineering", "automotive", "ecommerce", "cybersecurity",
]

SKIP_DOMAINS = {
    "google", "facebook", "twitter", "linkedin", "wikipedia", "youtube",
    "instagram", "tiktok", "snapchat", "amazon", "duckduck", "bing",
    "yahoo", "reddit", "quora", "trustpilot", "glassdoor", "indeed",
    "zawya", "arabnews", "saudigazette", "bloomberg", "reuters",
}

BLOCKED_CAREER_DOMAINS = {"linkedin.com", "www.linkedin.com"}

CAREER_SUBDOMAINS = [
    "careers", "jobs", "career", "job", "work", "hiring",
    "apply", "talent", "recruitment", "hr", "people",
    "vacancies", "opportunities", "join",
]

CAREER_PATHS = [
    "/careers", "/jobs", "/career", "/job",
    "/careers/", "/jobs/",
    "/en/careers", "/en/jobs", "/ar/careers",
    "/about/careers", "/about/jobs", "/company/careers",
    "/join-us", "/join", "/work-with-us",
    "/openings", "/vacancies", "/opportunities", "/employment",
    "/hiring", "/recruitment", "/apply",
    "/careers/search", "/jobs/search",
]

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
    "wizara.sa": "Wizara",
    "erecruit.com": "eRecruit",
    "trakstar.com": "Trakstar",
    "pinpointhq.com": "Pinpoint",
    "teamtailor.com": "Teamtailor",
    "comeet.com": "Comeet",
    "apply.workable.com": "Workable",
    "jobs.lever.co": "Lever",
    "boards.eu.greenhouse.io": "Greenhouse",
    "rmkcdn.successfactors.com": "SuccessFactors",
    "talentcommunity/apply": "SuccessFactors",
    "successfactors": "SuccessFactors",
    "ats.sa": "ATS.sa",
    "hiringsolved.com": "HiringSolved",
    "peoplehr.net": "PeopleHR",
    "oracle.com/taleo": "Taleo",
    "careers.pageuppeople.com": "PageUp",
    "pageuppeople.com": "PageUp",
    "jobsoid.com": "Jobsoid",
    "freshteam.com": "Freshteam",
    "cornerstone": "Cornerstone",
    "teamtailor.com": "Teamtailor",
    "hibob.com": "HiBob",
    "rippling.com": "Rippling",
    "personio.com": "Personio",
    "apply.wynt.ai": "Wynt",
    "wynt.ai": "Wynt",
    "recruitcrm.io": "RecruitCRM",
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
}

CAREER_KEYWORDS = [
    "career", "careers", "job", "jobs", "vacancy", "vacancies",
    "hiring", "work with us", "join us", "join our team",
    "employment", "opportunities", "opening", "openings",
    "وظائف", "التوظيف", "انضم إلينا", "فرص عمل",
]

DEFINITE_PATTERNS = [
    r"\bcareers?\b", r"\bjobs?\b", r"join\s+us",
    r"work\s+with\s+us", r"we'?re?\s+hiring",
    r"open\s+positions?", r"current\s+openings?",
    r"وظائف", r"التوظيف",
]

JOB_PAGE_SIGNALS = [
    "apply now", "apply for", "job description", "requirements",
    "qualifications", "responsibilities", "we are looking",
    "we are hiring", "open position", "full-time", "part-time",
    "remote", "salary", "benefits", "وظيفة", "تقديم", "المتطلبات",
    "job opportunities", "job search", "talentcommunity", "create alert",
    "sort by title", "sort by location",
]

LOCATION_PATTERN = re.compile(
    r"(Riyadh|Jeddah|Dammam|Khobar|Mecca|Medina|Saudi Arabia|KSA|Remote|"
    r"Dhahran|Jubail|Yanbu|Taif|Abha|Buraidah|Hail|Tabuk|"
    r"الرياض|جدة|الدمام|مكة|المدينة|Central Province|Eastern Province|"
    r"Western Province|Makkah|Madinah)", re.I
)

JOB_LISTING_SUFFIXES = [
    "/go/Job-Search/", "/go/All-Jobs/",
    "/job-search-results", "/job-search-results/",
    "/en/job-search-results", "/en/job-search-results/",
    "/ar/job-search-results",
    "/search", "/search/",
    "/job-search", "/job-search/",
    "/jobs-search", "/jobs-search/",
    "/career-search", "/career-search/",
    "/openings", "/openings/",
    "/current-openings", "/current-openings/",
    "/positions", "/positions/",
    "/open-positions", "/open-positions/",
    "/listings", "/listings/",
    "/job-listings", "/job-listings/",
    "/all-jobs", "/all-jobs/",
    "/list", "/list/",
    "/vacancies", "/vacancies/",
    "/opportunities", "/opportunities/",
    "/jobs", "/jobs/",
    "/apply", "/apply/",
    "/join", "/join/",
    "/en/jobs", "/en/jobs/",
    "/en/careers", "/en/careers/",
    "/en/search", "/en/search/",
    "/en/job-search", "/en/job-search/",
    "/en/openings", "/en/openings/",
    "/en/positions", "/en/positions/",
    "/en/vacancies", "/en/vacancies/",
    "/ar/jobs", "/ar/jobs/",
    "/ar/careers", "/ar/careers/",
    "/ar/search", "/ar/search/",
    "/ar/vacancies", "/ar/vacancies/",
]

SF_LISTING_SUFFIXES = [
    "/go/Job-Search/", "/go/All-Jobs/",
    "/job-search-results", "/job-search-results/",
    "/en/job-search-results",
    "/search", "/search/",
    "/jobs", "/jobs/",
    "/openings", "/openings/",
    "/all-jobs", "/all-jobs/",
]

SF_SEARCH_PARAMS = [
    "/?createNewAlert=false&q=&locationsearch=",
    "/search/?createNewAlert=false&q=",
    "/search/?q=&locationsearch=",
    "/?q=",
    "/search?q=",
]

JOB_PATH_CORE = re.compile(r"/jobs?/", re.I)

JOB_URL_ATS_SPECIFIC = re.compile(
    r"("
    r"/go/[A-Za-z0-9%-]+/\d+"
    r"|/\d{6,}/?$"
    r"|/job-detail/\d+"
    r"|/req/\d+"
    r"|/posting/[A-Za-z0-9-]+"
    r"|/application/\d+"
    r"|/apply/\d+"
    r")",
    re.I,
)

HARD_BLOCKED_PATTERNS = re.compile(
    r"("
    r"/legal/|/user.agreement|/privacy|/terms|/cookie"
    r"|/uas/request.password|/forgot.password|/sign.in|/login"
    r"|/register|/signup|/sign.up"
    r"|/new.cv|/myworkspace|/my.workspace|/dashboard"
    r"|/comments/feed|\.svg$|\.png$|\.jpg$|\.jpeg$|\.gif$|\.ico$"
    r"|/sitemap|/feed\.xml|/rss"
    r"|javascript:|mailto:|tel:"
    r"|/cdn.cgi/access"
    r"|objectstorage.*\.png"
    r")",
    re.I,
)

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
    r"(?:expires?|expiry)[:\s]*([^\n<]{1,40})",
    re.I,
)

EXPERIENCE_RE = re.compile(
    r"(\d+\+?\s*(?:–|-|to)?\s*\d*\s*years?\s*(?:of\s+)?experience|"
    r"experience[:\s]*(\d+[\+\s\-]*\d*\s*years?)|"
    r"(?:minimum|min\.?)\s+\d+\s+years?)",
    re.I,
)

_GARBAGE_FIELD_PATTERNS = re.compile(
    r"^(job\s*description|key\s*accountabilities?|key\s*functional|"
    r"requirements?|qualifications?|responsibilities|overview|"
    r"role\s*purpose|job\s*purpose|about\s*the\s*role|"
    r"what\s*you|we\s*are\s*looking|our\s*team|"
    r"education\s*&|experience\s*&|skills\s*&|"
    r"others?|not\s*applicable|n/?a|—|-)$",
    re.I,
)

MONTH_MAP = {
    "jan": 0, "feb": 1, "mar": 2, "apr": 3, "may": 4, "jun": 5,
    "jul": 6, "aug": 7, "sep": 8, "oct": 9, "nov": 10, "dec": 11,
}

BAD_URL_DOMAINS = [
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

SKIP_CRAWL_DOMAINS = [
    "dhl.com","fedex.com","ups.com","amazon.com","amazon.jobs",
    "google.com","microsoft.com","apple.com","meta.com","ibm.com",
    "oracle.com","sap.com","accenture.com","deloitte.com","pwc.com",
    "kpmg.com","ey.com","mckinsey.com","bcg.com","bain.com",
    "citibank.com","hsbc.com","barclays.com","bnpparibas.com",
    "airbus.com","boeing.com","siemens.com","ge.com",
    "unilever.com","nestle.com","pg.com","shell.com","bp.com",
]


# ══════════════════════════════════════════════════════════════
# ▌ SHARED STATE
# ══════════════════════════════════════════════════════════════
discovered_domains = set()
all_jobs           = []
company_results    = []


# ══════════════════════════════════════════════════════════════
# ▌ DEDUPLICATION ENGINE
# ══════════════════════════════════════════════════════════════
def _job_fingerprint(job: dict) -> tuple:
    """
    Canonical fingerprint for deduplication.
    Uses normalised (title, company, location) — insensitive to minor
    differences in whitespace or casing.
    """
    title   = re.sub(r"\s+", " ", (job.get("Job Title") or "").lower().strip())
    company = re.sub(r"\s+", " ", (job.get("Company Name") or "").lower().strip())
    loc     = re.sub(r"\s+", " ", (job.get("Job Location") or "").lower().strip())
    # Strip common location suffixes like ", saudi arabia" so "Riyadh, Saudi Arabia"
    # matches "Riyadh"
    loc = re.sub(r",?\s*(saudi arabia|ksa)$", "", loc, flags=re.I).strip()
    return (title, company, loc)


def _merge_jobs(primary: dict, secondary: dict) -> dict:
    """
    Merge two job records that represent the same position.
    `primary` wins on non-empty fields; `secondary` fills any blanks.
    Richer description / longer text fields are preferred.
    """
    merged = dict(primary)
    for key, val in secondary.items():
        if key == "source":
            # Concatenate sources so we know both contributed
            existing = merged.get("source", "")
            if val and val not in existing:
                merged["source"] = f"{existing}+{val}" if existing else val
            continue
        existing_val = merged.get(key, "")
        if not existing_val and val:
            merged[key] = val
        elif val and key in ("Job Description", "Company Details"):
            # Prefer the longer / richer text
            if len(str(val)) > len(str(existing_val)):
                merged[key] = val
        elif val and key in ("Company Logo", "Application", "Job URL"):
            # Prefer non-empty; already set above — also prefer https
            if not existing_val:
                merged[key] = val
            elif val.startswith("https") and not existing_val.startswith("https"):
                merged[key] = val
    return merged


def dedup_and_merge(jobs: list) -> list:
    """
    Deduplicate a flat list of job dicts.
    Returns a new list with duplicates merged (not dropped).
    """
    seen: dict[tuple, dict] = {}
    order = []
    for job in jobs:
        fp = _job_fingerprint(job)
        if fp in seen:
            seen[fp] = _merge_jobs(seen[fp], job)
            _stats["jobs_duplicates_merged"] += 1
        else:
            seen[fp] = job
            order.append(fp)
    return [seen[fp] for fp in order]


# ══════════════════════════════════════════════════════════════
# ▌ GENERIC HELPERS
# ══════════════════════════════════════════════════════════════
def clean(text, max_len=300):
    return re.sub(r"\s+", " ", str(text or "")).strip()[:max_len]


def get_domain(url):
    try:
        return urlparse(str(url)).netloc.lower().replace("www.", "")
    except:
        return ""


def get_base(website):
    p = urlparse(str(website))
    return p.netloc.lower().replace("www.", ""), p.scheme or "https"


def is_ats(url):
    url = str(url).lower()
    if "linkedin.com" in url:
        return None
    for pat, name in ATS_DOMAINS.items():
        if pat in url:
            return name
    return None


def is_ats_from_html(html):
    lower = html.lower()
    for fingerprint, name in ATS_HTML_FINGERPRINTS.items():
        if fingerprint.lower() in lower:
            return name
    return None


def is_career_link(href, text):
    combined = (str(href) + " " + str(text)).lower()
    if "linkedin.com" in combined:
        return False, None
    for pat in DEFINITE_PATTERNS:
        if re.search(pat, text.lower()):
            return True, "definite"
    for kw in CAREER_KEYWORDS:
        if kw in combined:
            return True, "keyword"
    return False, None


def page_has_jobs(html):
    lower = html.lower()
    return sum(1 for s in JOB_PAGE_SIGNALS if s in lower) >= 2


def simple_get(url, timeout=10):
    try:
        r = requests.get(str(url), headers=HEADERS, timeout=timeout, allow_redirects=True)
        if r.status_code == 200:
            return r
    except:
        pass
    return None


def parse_domain(url):
    try:
        p = urlparse(url if url.startswith("http") else "https://" + url)
        netloc = p.netloc.lower().replace("www.", "")
        if not netloc or any(s in netloc for s in SKIP_DOMAINS):
            return None, None
        return netloc, f"https://www.{netloc}"
    except:
        return None, None


def load_checkpoint():
    if Path(CHECKPOINT_FILE).exists():
        return json.loads(Path(CHECKPOINT_FILE).read_text())
    return {"processed_domains": [], "jobs_count": 0}


def save_checkpoint(cp):
    Path(CHECKPOINT_FILE).write_text(json.dumps(cp, indent=2))


def _sanitize_field(value, max_len=300):
    if not value:
        return ""
    v = re.sub(r"\s+", " ", str(value)).strip()
    if _GARBAGE_FIELD_PATTERNS.match(v):
        return ""
    if len(v) < 2:
        return ""
    return v[:max_len]


def flush_jobs():
    if all_jobs:
        deduped = dedup_and_merge(all_jobs)
        pd.DataFrame(deduped).drop_duplicates(subset="Job URL", keep="first").to_csv(
            JOBS_FILE, index=False
        )
    if company_results:
        pd.DataFrame(company_results).to_csv(COMPANIES_FILE, index=False)


def make_job(company, website, industry, careers_url, source,
             title="", location="", job_type="", department="", apply_url="",
             qualifications="", experience="", field="",
             date_posted="", deadline="", description="",
             company_logo="", company_founded="", company_type="",
             company_address="", company_details="",
             estimated_deadline="", salary_range=""):

    std_field = standardise_field(field or department, title, description, industry)
    std_exp   = standardise_experience(experience)
    std_qual  = standardise_qualification(qualifications, description)

    return {
        "Job Title":          clean(title),
        "Job Type":           _sanitize_field(job_type),
        "Job Qualifications": std_qual,
        "Job Experience":     std_exp,
        "Job Location":       clean(location) or "Saudi Arabia",
        "Job Field":          std_field,
        "Date Posted":        _sanitize_field(date_posted),
        "Deadline":           _sanitize_field(deadline),
        "Job Description":    clean(description, 2000),
        "Application":        apply_url,
        "Company URL":        careers_url,
        "Company Name":       clean(company),
        "Company Logo":       company_logo,
        "Company Industry":   clean(industry),
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


# ══════════════════════════════════════════════════════════════
# ▌ DATE / DEADLINE HELPERS
# ══════════════════════════════════════════════════════════════
def _estimate_deadline(date_posted_str):
    if not date_posted_str:
        return ""
    try:
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%B %d, %Y",
                    "%b %d, %Y", "%b %d, %y", "%B %d %Y"):
            try:
                dt = datetime.strptime(date_posted_str.strip(), fmt)
                return (dt + timedelta(days=30)).strftime("%Y-%m-%d")
            except:
                continue
        m = re.search(r"(\d+)\s+days?\s+ago", date_posted_str, re.I)
        if m:
            posted = datetime.now() - timedelta(days=int(m.group(1)))
            return (posted + timedelta(days=30)).strftime("%Y-%m-%d")
    except:
        pass
    return ""


def _resolve_posted_date(raw: str) -> str:
    if not raw:
        return ""
    text = raw.strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}$", text):
        return text
    try:
        return datetime.fromisoformat(text).strftime("%Y-%m-%d")
    except:
        pass
    base = datetime.now()
    m = re.search(r"(\d+)\s*(hour|day|week|month|year)", text, re.I)
    if m:
        n, unit = int(m.group(1)), m.group(2).lower()
        if "hour"  in unit: base -= timedelta(hours=n)
        elif "day"  in unit: base -= timedelta(days=n)
        elif "week" in unit: base -= timedelta(weeks=n)
        elif "month" in unit:
            mo = base.month - n
            base = base.replace(year=base.year + mo // 12, month=mo % 12 or 12)
        elif "year" in unit:
            base = base.replace(year=base.year - n)
        return base.strftime("%Y-%m-%d")
    if re.search(r"just\s*now|today", text, re.I):
        return datetime.now().strftime("%Y-%m-%d")
    return datetime.now().strftime("%Y-%m-%d")


def _try_parse_date(s: str):
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%B %d, %Y", "%d %B %Y"):
        try:
            return datetime.strptime(s.strip(), fmt)
        except:
            pass
    m = re.match(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", s)
    if m:
        mo = MONTH_MAP.get(m.group(2)[:3].lower())
        if mo is not None:
            return datetime(int(m.group(3)), mo + 1, int(m.group(1)))
    return None


def _parse_deadline(soup) -> str:
    full_text = soup.get_text()
    patterns = [
        r"closes?\s+on\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"apply\s+by\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"applications?\s+close[sd]?\s*(?:on)?\s+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"deadline[:\s]+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
        r"closing\s+date[:\s]+([A-Za-z]+\s+\d{1,2},?\s+\d{4})",
    ]
    now = datetime.now()
    for pattern in patterns:
        m = re.search(pattern, full_text, re.I)
        if m:
            d = _try_parse_date(m.group(1))
            if d and d > now:
                return d.strftime("%Y-%m-%d")
    return ""


# ══════════════════════════════════════════════════════════════
# ▌ LOGO / META HELPERS
# ══════════════════════════════════════════════════════════════
def _extract_logo(soup, base_url):
    og = soup.find("meta", property="og:image")
    if og and og.get("content"):
        return og["content"]
    for sel in ["img.logo", "img[class*='logo']", "img[alt*='logo']",
                ".logo img", "header img", ".navbar-brand img"]:
        el = soup.select_one(sel)
        if el and el.get("src"):
            src = el["src"]
            return src if src.startswith("http") else urljoin(base_url, src)
    return ""


def _extract_json_ld(soup):
    data = {}
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            obj = json.loads(script.string or "{}")
            if "@graph" in obj:
                for item in obj["@graph"]:
                    data.update(_parse_json_ld_job(item))
            else:
                data.update(_parse_json_ld_job(obj))
        except:
            pass
    return data


def _parse_json_ld_job(obj):
    out = {}
    t = obj.get("@type", "")
    if "JobPosting" not in str(t):
        return out
    out["title"]        = obj.get("title", "")
    out["description"]  = obj.get("description", "")
    out["date_posted"]  = obj.get("datePosted", "")
    out["deadline"]     = obj.get("validThrough", "")
    out["job_type"]     = obj.get("employmentType", "")
    out["salary_range"] = ""
    bs = obj.get("baseSalary", {})
    if isinstance(bs, dict):
        val = bs.get("value", {})
        if isinstance(val, dict):
            mn  = val.get("minValue", "")
            mx  = val.get("maxValue", "")
            unit = val.get("unitText", "")
            cur = bs.get("currency", "")
            if mn or mx:
                out["salary_range"] = f"{cur} {mn}–{mx} ({unit})".strip()
        elif val:
            out["salary_range"] = str(val)
    loc = obj.get("jobLocation", {})
    if isinstance(loc, dict):
        addr = loc.get("address", {})
        if isinstance(addr, dict):
            parts = [addr.get("addressLocality",""), addr.get("addressRegion",""),
                     addr.get("addressCountry","")]
            out["location"] = ", ".join(p for p in parts if p)
        else:
            out["location"] = str(addr)
    out["qualifications"] = obj.get("qualifications", "")
    out["experience"]     = obj.get("experienceRequirements", "")
    org = obj.get("hiringOrganization", {})
    if isinstance(org, dict):
        out["company_name"]    = org.get("name", "")
        out["company_logo"]    = org.get("logo", "")
        out["company_website"] = org.get("sameAs", "")
        out["company_type"]    = org.get("@type", "")
    out["field"] = obj.get("occupationalCategory", "") or obj.get("industry", "")
    return out


def _extract_meta(soup, key):
    for attr in ["name", "property"]:
        tag = soup.find("meta", {attr: key})
        if tag and tag.get("content"):
            return tag["content"]
    return ""


def _find_text_near_label(soup, *labels):
    for label in labels:
        pat = re.compile(re.escape(label), re.I)
        for dt in soup.find_all(["dt", "th", "strong", "label", "span", "td", "b"]):
            if pat.search(dt.get_text()):
                sib = dt.find_next_sibling()
                if sib:
                    val = _sanitize_field(sib.get_text())
                    if val:
                        return clean(val)
                nxt = dt.find_next("td")
                if nxt:
                    val = _sanitize_field(nxt.get_text())
                    if val:
                        return clean(val)
    return ""


def _parse_headed_sections(soup):
    sections = {}
    for h in soup.find_all(["h1", "h2", "h3", "h4"]):
        key = h.get_text(strip=True).lower()
        if not key or len(key) > 120:
            continue
        body_parts = []
        for sib in h.find_next_siblings():
            if sib.name in ("h1", "h2", "h3", "h4"):
                break
            text = sib.get_text(separator=" ", strip=True)
            if text:
                body_parts.append(text)
        sections[key] = " ".join(body_parts).strip()
    return sections


def _pick_section(sections, *keywords):
    for kw in keywords:
        kw_lower = kw.lower()
        for key, val in sections.items():
            if kw_lower in key and val:
                cleaned = _sanitize_field(val[:300])
                if cleaned:
                    return val
    return ""


def _extract_bold_field(soup, *labels):
    for label in labels:
        pat = re.compile(re.escape(label) + r"\s*:?", re.I)
        for el in soup.find_all(["strong", "b", "p", "li"]):
            text = el.get_text()
            if pat.search(text):
                cleaned = pat.sub("", text).strip()
                if cleaned and len(cleaned) > 2:
                    val = _sanitize_field(cleaned, 300)
                    if val:
                        return val
                sib = el.find_next_sibling()
                if sib:
                    val = _sanitize_field(sib.get_text(), 300)
                    if val:
                        return val
    return ""


# ══════════════════════════════════════════════════════════════
# ▌ URL HELPERS
# ══════════════════════════════════════════════════════════════
def is_likely_job_url(url, career_base_domain=""):
    if not url or len(url) < 10:
        return False
    url_lower = url.lower()
    path = urlparse(url).path.lower()
    if HARD_BLOCKED_PATTERNS.search(url_lower):
        return False
    if re.search(r"\.(svg|png|jpg|jpeg|gif|ico|pdf|zip|xml|json)$", path, re.I):
        return False
    segments = [s for s in path.strip("/").split("/") if s]
    if len(segments) < 1:
        return False
    if JOB_URL_ATS_SPECIFIC.search(path):
        if re.search(r"/go/job.?search/\d+/?$", path, re.I):
            return False
        if re.search(r"/go/all.?jobs/\d+/?$", path, re.I):
            return False
        return True
    if JOB_PATH_CORE.search(path):
        return True
    return False


def _is_detail_page_url(url):
    if not url or len(url) < 10:
        return False
    url_lower = url.lower()
    path = urlparse(url).path.rstrip("/").lower()
    if HARD_BLOCKED_PATTERNS.search(url_lower):
        return False
    if re.search(r"\.(svg|png|jpg|jpeg|gif|ico|pdf|zip|xml)$", path, re.I):
        return False
    if "/job-application" in path:
        if not re.search(r"jb_id=\d+", url_lower):
            return False
    if re.search(r"/go/job.?search/\d+/?$", path, re.I):
        return False
    if re.search(r"/go/all.?jobs/\d+/?$", path, re.I):
        return False
    if re.search(
        r"^/(jobs|job.search|job.search.results|openings"
        r"|vacancies|all.jobs|opportunities|positions|search)$",
        path, re.I,
    ):
        return False
    return is_likely_job_url(url)


def _is_blocked_career_url(url: str) -> bool:
    return get_domain(url) in BLOCKED_CAREER_DOMAINS


def is_bad_url(url: str) -> bool:
    if not url or not url.startswith("http"):
        return True
    lower = url.lower()
    return any(d in lower for d in BAD_URL_DOMAINS)


def decode_html_entities(s: str) -> str:
    if not s:
        return ""
    for old, new in [("&amp;","&"),("&lt;","<"),("&gt;",">"),("&quot;",'"'),
                     ("&#39;","'"),("\\u0026","&"),("\\u003D","="),
                     ("\\u003A",":"),("\\u002F","/")]:
        s = s.replace(old, new)
    return s


def canonicalise_linkedin_url(url: str) -> str:
    if not url:
        return ""
    m = re.search(r"/jobs/view/(\d+)", url)
    if m:
        return f"https://www.linkedin.com/jobs/view/{m.group(1)}/"
    return re.sub(r"[?#].*$", "", url)


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
        except:
            pass
    return ""


# ══════════════════════════════════════════════════════════════
# ▌ EMAIL HELPERS  (from LinkedIn v2)
# ══════════════════════════════════════════════════════════════
def clean_email(raw: str) -> str:
    if not raw:
        return ""
    em = re.sub(r"^mailto:", "", raw, flags=re.I)
    em = re.sub(r"\?.*$", "", em)
    em = em.strip().lower()
    if not em or "@" not in em or "." not in em:
        return ""
    at_idx = em.rfind("@")
    local  = em[:at_idx]
    domain = em[at_idx + 1:]
    em = local + "@" + domain
    if not re.match(r"^[a-zA-Z0-9]", em):
        return ""
    return em


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
        if FAKE_LOCAL_RE.match(parts[0]) or FAKE_DOMAIN_RE.match(parts[1]):
            continue
        return em
    return ""


# ══════════════════════════════════════════════════════════════
# ▌ DESCRIPTION CLEANER  (from LinkedIn v2)
# ══════════════════════════════════════════════════════════════
def clean_description(raw: str, max_len: int = 2000) -> str:
    if not raw:
        return ""
    text = raw.replace("\u00a0", " ").replace("\u200b", "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"([a-z])([A-Z])", r"\1 \2", text)
    text = re.sub(r"\s*[•·▪◦]\s*", "\n• ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r" {2,}", " ", text)
    return text.strip()[:max_len]


# ══════════════════════════════════════════════════════════════
# ▌ SF TABLE SCRAPER
# ══════════════════════════════════════════════════════════════
def _scrape_sf_listing_table(soup, listing_url):
    results = []
    table = soup.find("table")
    if not table:
        return results
    for row in table.find_all("tr"):
        cells = row.find_all("td")
        if len(cells) < 2:
            continue
        title_cell = cells[0]
        links = title_cell.find_all("a", href=True)
        if not links:
            continue
        seen_urls = set()
        for link in links:
            href = link.get("href", "")
            if not href or href in seen_urls:
                continue
            seen_urls.add(href)
            full_url = href if href.startswith("http") else urljoin(listing_url, href)
            if not _is_detail_page_url(full_url):
                continue
            title       = link.get_text(strip=True)
            location    = clean(cells[1].get_text()) if len(cells) >= 2 else ""
            date_posted = clean(cells[2].get_text()) if len(cells) >= 3 else ""
            if title and len(title) > 2:
                results.append({"url": full_url, "title": title,
                                 "location": location, "date_posted": date_posted})
            break
    return results


# ══════════════════════════════════════════════════════════════
# ▌ DEEP JOB DETAIL SCRAPER  (career pages)
# ══════════════════════════════════════════════════════════════
async def scrape_job_detail(page, job_url, company, website, industry,
                             careers_url, source,
                             prefill_title="", prefill_location="",
                             prefill_date=""):
    html      = None
    final_url = job_url

    vprint(f"🌍  Fetching detail: {job_url[:80]}", indent=3)
    r = simple_get(job_url, timeout=12)
    if r:
        html      = r.text
        final_url = r.url
        vprint(f"✔  HTTP OK  ({len(html):,} chars)", indent=4)
    else:
        vprint(f"⚠  Plain HTTP failed — using Playwright…", indent=4)
        try:
            await page.goto(job_url, timeout=25000, wait_until="domcontentloaded")
            await page.wait_for_timeout(1800)
            html      = await page.content()
            final_url = page.url
            vprint(f"✔  Playwright OK  ({len(html):,} chars)", indent=4)
        except Exception as e:
            vlog(f"⚠️  Detail fetch failed {job_url}: {e}", indent=3)
            _stats["detail_failures"] += 1
            return None

    _stats["detail_fetches"] += 1
    if not html:
        return None

    soup = BeautifulSoup(html, "html.parser")
    ld       = _extract_json_ld(soup)
    sections = _parse_headed_sections(soup)

    for tag in soup.select("nav, footer, header, script, style, iframe, .cookie"):
        tag.decompose()
    full_text = soup.get_text(separator="\n")

    title = (
        ld.get("title") or prefill_title
        or _extract_meta(soup, "og:title")
        or clean(soup.title.get_text() if soup.title else "")
        or clean(soup.select_one("h1").get_text() if soup.select_one("h1") else "")
    )
    if title and "|" in title:
        title = title.split("|")[0].strip()
    if title and "-" in title and len(title) > 60:
        title = title.split("-")[0].strip()

    department = _sanitize_field(ld.get("field") or _find_text_near_label(
        soup, "Department", "Team", "Function", "Division", "القسم", "الإدارة") or "")

    job_type = _sanitize_field(
        ld.get("job_type")
        or _find_text_near_label(soup, "Employment Type", "Job Type", "Contract Type", "Type", "نوع الوظيفة")
        or _pick_section(sections, "employment type", "job type") or "")
    if not job_type:
        for jt in ["Full-time", "Part-time", "Contract", "Internship", "Freelance", "Temporary"]:
            if re.search(r"\b" + jt.split("-")[0] + r"\b", full_text, re.I):
                job_type = jt
                break

    location = (
        ld.get("location") or prefill_location
        or _find_text_near_label(soup, "Location", "Job Location", "City", "الموقع", "المدينة") or "")
    if not location:
        m = LOCATION_PATTERN.search(full_text)
        if m:
            location = m.group(1)
    location = location or "Saudi Arabia"

    description = ld.get("description", "") or _pick_section(
        sections, "key functional", "responsibilities", "job purpose", "about the role",
        "what you'll do", "role overview", "job summary", "overview", "المهام", "المسؤوليات")
    if not description:
        for sel in ["[class*='description']", "[class*='job-desc']", "[class*='content']",
                    "#job-description", "article", "main"]:
            el = soup.select_one(sel)
            if el and len(el.get_text()) > 100:
                description = clean_description(el.get_text(), 2000)
                break
    if not description and len(full_text) > 200:
        description = clean_description(full_text, 2000)

    qualifications = (
        ld.get("qualifications")
        or _pick_section(sections, "qualif", "requirement", "education",
                         "what you need", "you should have", "minimum requirement",
                         "المؤهلات", "المتطلبات")
        or _find_text_near_label(soup, "Qualifications", "Requirements", "Education", "Degree")
        or _extract_bold_field(soup, "Education", "Qualifications", "Requirements") or "")
    if _GARBAGE_FIELD_PATTERNS.match(qualifications.strip()[:80]):
        qualifications = ""

    experience = _sanitize_field(
        ld.get("experience")
        or _find_text_near_label(soup, "Experience", "Years of Experience", "الخبرة")
        or _pick_section(sections, "experience")
        or _extract_bold_field(soup, "Experience", "Years of Experience") or "")
    if not experience:
        m = EXPERIENCE_RE.search(full_text)
        if m:
            experience = m.group(0)

    field = _sanitize_field(
        ld.get("field") or department
        or _find_text_near_label(soup, "Field", "Category", "Department", "Function", "التخصص")
        or industry)

    date_posted = ld.get("date_posted", "") or prefill_date
    if not date_posted:
        m = DATE_POSTED_RE.search(full_text)
        if m:
            date_posted = next((g for g in m.groups() if g), "")
        if not date_posted:
            time_el = soup.find("time")
            if time_el:
                date_posted = time_el.get("datetime") or time_el.get_text(strip=True)
    if date_posted and _GARBAGE_FIELD_PATTERNS.match(date_posted.strip()):
        date_posted = ""

    deadline = ld.get("deadline", "")
    if not deadline:
        m = DEADLINE_RE.search(full_text)
        if m:
            deadline = next((g for g in m.groups() if g), "")
    if deadline and _GARBAGE_FIELD_PATTERNS.match(deadline.strip()):
        deadline = ""

    estimated_deadline = deadline or _estimate_deadline(date_posted)
    if estimated_deadline and not re.match(r"\d{4}-\d{2}-\d{2}", estimated_deadline.strip()):
        estimated_deadline = ""

    salary_range = (
        ld.get("salary_range")
        or _find_text_near_label(soup, "Salary", "Compensation", "الراتب", "المرتب") or "")
    if not salary_range:
        m = SALARY_RE.search(full_text)
        if m:
            salary_range = m.group(0)

    company_logo    = ld.get("company_logo", "") or _extract_logo(soup, final_url)
    company_type    = _sanitize_field(ld.get("company_type", "") or "")
    if company_type and company_type.lower() in ("organization", "legalservice", "thing"):
        company_type = ""
    company_address = _find_text_near_label(soup, "Address", "Headquarters", "العنوان")
    company_founded = _sanitize_field(_find_text_near_label(soup, "Founded", "Established"))
    company_details = _pick_section(sections, "about the company", "about us",
                                    "company overview", "who we are", "من نحن")

    apply_url = final_url
    for a in soup.find_all("a", href=True):
        href       = a.get("href", "")
        text_a     = a.get_text(strip=True).lower()
        href_lower = href.lower()
        if any(p in href_lower for p in ["talentcommunity/apply", "/apply/",
                                          "?action=apply", "/application/"]):
            apply_url = href if href.startswith("http") else urljoin(final_url, href)
            break
        if "apply" in text_a and href and href not in ("#", "javascript:void(0)"):
            candidate = href if href.startswith("http") else urljoin(final_url, href)
            if candidate != final_url:
                apply_url = candidate

    vprint(f"  ┄ title        : {title[:60] or '—'}", indent=4)
    vprint(f"  ┄ location     : {location or '—'}", indent=4)
    vprint(f"  ┄ field        : {standardise_field(field or department, title, description, industry) or '—'}", indent=4)
    vprint(f"  ┄ salary       : {salary_range or '—'}", indent=4)

    if salary_range:     _stats["jobs_with_salary"] += 1
    if description:      _stats["jobs_with_description"] += 1

    return make_job(
        company=company, website=website, industry=industry,
        careers_url=careers_url, source=source + "+detail",
        title=title, location=location, job_type=job_type,
        department=department, apply_url=apply_url,
        qualifications=qualifications, experience=experience, field=field,
        date_posted=date_posted, deadline=deadline, description=description,
        company_logo=company_logo, company_founded=company_founded,
        company_type=company_type, company_address=company_address,
        company_details=company_details,
        estimated_deadline=estimated_deadline, salary_range=salary_range,
    )


# ══════════════════════════════════════════════════════════════
# ▌ LINKEDIN GUEST API  (from LinkedIn v2, stripped of non-Saudi)
# ══════════════════════════════════════════════════════════════
_LI_DELAY   = 2.0
_LI_MAX_PAGES = 0       # 0 = unlimited (up to LinkedIn's 1000-result cap)
_LI_MAX_EMPTY = 3       # stop after N consecutive empty pages per keyword

# LinkedIn seniority / employment-type mappings
_LI_JOB_TYPE_MAP = {
    "full-time": "Full-time", "part-time": "Part-time",
    "contract": "Contract", "temporary": "Temporary",
    "internship": "Internship", "freelance": "Freelance",
    "permanent": "Permanent",
}


def _li_build_url(keyword: str, start: int) -> str:
    kw_enc = quote_plus(keyword)
    return (
        "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search"
        f"?location=Saudi+Arabia&f_TPR=r604800&keywords={kw_enc}&start={start}"
    )


def _li_collect_urls(html: str, seen: set) -> list:
    found = []
    for raw_href in re.findall(r'href="(https?://[^"]*?/jobs/view/\d+[^"]*?)"', html):
        c = canonicalise_linkedin_url(raw_href)
        if c and c not in seen:
            seen.add(c)
            found.append(c)
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup.find_all("a", href=True):
        href = tag["href"]
        if "/jobs/view/" not in href:
            continue
        if not href.startswith("http"):
            href = "https://www.linkedin.com" + href
        c = canonicalise_linkedin_url(href)
        if c and c not in seen:
            seen.add(c)
            found.append(c)
    return found


def _li_fetch_page(keyword: str, start: int, retries: int = 3):
    url = _li_build_url(keyword, start)
    for attempt in range(retries):
        try:
            time.sleep(_LI_DELAY + attempt * 3)
            r = requests.get(url, headers=_next_li_headers(),
                             allow_redirects=True, timeout=25)
            if r.status_code == 429:
                wait = 60 + attempt * 60
                vlog(f"  ⏳ LinkedIn 429 — waiting {wait}s ...")
                time.sleep(wait)
                continue
            if r.status_code in (400, 403, 999):
                return None
            if r.status_code != 200:
                return None
            text = r.text.strip()
            if not text:
                return None
            return text
        except Exception as e:
            time.sleep(3 + attempt * 3)
    return None


def _li_paginate_keyword(keyword: str, seen: set) -> list:
    urls, page, empty_streak = [], 0, 0
    while True:
        if _LI_MAX_PAGES and page >= _LI_MAX_PAGES:
            break
        start = page * 25
        html  = _li_fetch_page(keyword, start)
        if html is None:
            break
        new_urls = _li_collect_urls(html, seen)
        if new_urls:
            urls.extend(new_urls)
            empty_streak = 0
        else:
            empty_streak += 1
            if empty_streak >= _LI_MAX_EMPTY:
                break
        if start >= 975:
            break
        page += 1
        if page % 10 == 0:
            time.sleep(20)
    return urls


def _li_scrape_job_detail(job_url: str, company_name_hint: str = "") -> dict | None:
    """
    Scrape a single LinkedIn job detail page using the public guest view.
    Returns a unified job dict (same schema as make_job).
    """
    try:
        r = requests.get(job_url, headers=_next_li_headers(), timeout=15)
        if r.status_code != 200:
            return None
        html = r.text
        soup = BeautifulSoup(html, "html.parser")
    except Exception as e:
        return None

    def sel_text(*selectors):
        for s in selectors:
            el = soup.select_one(s)
            if el:
                t = el.get_text(strip=True)
                if t:
                    return t
        return ""

    title         = sel_text(".top-card-layout__title", "h1.topcard__title",
                              ".job-details-jobs-unified-top-card__job-title", "h1")
    company_name  = sel_text(".topcard__org-name-link",
                              ".job-details-jobs-unified-top-card__company-name",
                              ".topcard__flavor") or company_name_hint
    company_url_el = (soup.select_one(".topcard__org-name-link") or
                      soup.select_one(".job-details-jobs-unified-top-card__company-name a"))
    company_li_url = company_url_el.get("href", "") if company_url_el else ""

    location  = sel_text(".topcard__flavor--bullet",
                          ".job-details-jobs-unified-top-card__bullet")
    if not location:
        m = LOCATION_PATTERN.search(soup.get_text())
        if m:
            location = m.group(1)
    location = location or "Saudi Arabia"

    time_el  = soup.find("time")
    raw_posted = (time_el.get("datetime", "") if time_el else "") or sel_text(
        ".posted-time-ago__text",
        ".job-details-jobs-unified-top-card__posted-date")
    date_posted = _resolve_posted_date(raw_posted)

    raw_desc = sel_text(".show-more-less-html__markup", ".description__text")
    description = clean_description(raw_desc, 2000)

    salary = ""
    for sel in [".compensation__salary", ".salary", "[class*='salary']"]:
        el = soup.select_one(sel)
        if el:
            salary = el.get_text(strip=True)
            break
    if not salary:
        for chip in soup.select(".job-details-jobs-unified-top-card__job-insight"):
            t = chip.get_text(strip=True)
            if re.search(r"\$|SAR|SR|salary|/yr|/hour|per month", t, re.I):
                salary = t
                break

    # Job type
    def _get_criteria(label):
        lower = label.lower()
        for li in soup.select(".description__job-criteria-list > li"):
            h3 = li.find("h3")
            if h3 and lower in h3.get_text().strip().lower():
                spans = li.select(".description__job-criteria-text, span")
                if spans:
                    return spans[-1].get_text(strip=True)
        return ""

    job_type   = _get_criteria("Employment type") or "Full-time"
    seniority  = _get_criteria("Seniority level")
    li_function = _get_criteria("Job function")
    li_industry = _get_criteria("Industries")

    # Deadline
    real_deadline = _parse_deadline(soup)
    estimated_deadline = _estimate_deadline(date_posted) if not real_deadline else ""
    effective_deadline = real_deadline or estimated_deadline

    # Logo from og:image
    company_logo = ""
    og_img = soup.find("meta", property="og:image")
    if og_img and og_img.get("content"):
        company_logo = og_img["content"]

    # Apply link — try the offsite apply button
    apply_url = job_url
    for tag in soup.find_all("a", href=True):
        href = tag.get("href", "")
        control = tag.get("data-tracking-control-name", "")
        if "offsite" in control.lower() or "apply" in control.lower():
            resolved = decode_linkedin_apply_url(href)
            if resolved and not is_bad_url(resolved):
                apply_url = resolved
                break
    # Script tag fallback
    if apply_url == job_url:
        for script in soup.find_all("script"):
            txt = script.string or ""
            for pat in [r'"applyStartUrl"\s*:\s*"([^"]+)"',
                         r'"applicationUrl"\s*:\s*"([^"]+)"']:
                mm = re.search(pat, txt)
                if mm:
                    candidate = decode_html_entities(mm.group(1)).replace("\\", "")
                    if candidate.startswith("http") and not is_bad_url(candidate):
                        apply_url = candidate
                        break

    qualifications = standardise_qualification("", description)
    experience     = standardise_experience(
        _get_criteria("Seniority level") or seniority or "")
    if not experience:
        experience = standardise_experience(description[:500])

    field = standardise_field(li_function or li_industry, title, description, li_industry)
    industry = li_industry or ""

    if salary:   _stats["jobs_with_salary"] += 1
    if description: _stats["jobs_with_description"] += 1

    return make_job(
        company=company_name, website="", industry=industry,
        careers_url=company_li_url, source="LinkedIn",
        title=title, location=location, job_type=job_type,
        department="", apply_url=apply_url,
        qualifications="", experience=experience, field=field,
        date_posted=date_posted, deadline=effective_deadline,
        description=description,
        company_logo=company_logo, company_founded="", company_type="",
        company_address="", company_details="",
        estimated_deadline=estimated_deadline, salary_range=salary,
    )


async def fetch_linkedin_jobs_for_company(company_name: str) -> list:
    """
    Use LinkedIn guest API to find jobs posted by this company in Saudi Arabia.
    Returns a list of unified job dicts.
    """
    vprint(f"🔗  LinkedIn search for: {company_name}", indent=1)
    seen_urls: set = set()
    # Search by company name as keyword — most targeted approach
    urls = _li_paginate_keyword(company_name, seen_urls)
    vprint(f"  ↳ {len(urls)} LinkedIn job URLs found", indent=1)

    jobs = []
    for url in urls[:100]:   # cap per company to avoid abuse
        job = _li_scrape_job_detail(url, company_name_hint=company_name)
        if job and job.get("Job Title"):
            # Only keep if company name roughly matches
            scraped_co = (job.get("Company Name") or "").lower()
            query_co   = company_name.lower()
            # Accept if at least one word overlaps (handles "Saudi Aramco" vs "Aramco")
            co_words   = set(re.findall(r"\w{3,}", query_co))
            sc_words   = set(re.findall(r"\w{3,}", scraped_co))
            if co_words & sc_words or not scraped_co:
                jobs.append(job)
        time.sleep(_LI_DELAY)

    _stats["jobs_from_linkedin"] += len(jobs)
    vprint(f"  ↳ {len(jobs)} LinkedIn jobs matched for {company_name}", indent=1)
    return jobs


# ══════════════════════════════════════════════════════════════
# ▌ CAREER PAGE FINDER  (v5, unchanged)
# ══════════════════════════════════════════════════════════════
def _job_signal_count(html):
    lower = html.lower()
    return sum(1 for s in JOB_PAGE_SIGNALS if s in lower)


async def _probe_suffixes(page, career_root, root_score=0):
    root = career_root.rstrip("/")
    best_url, best_score = None, root_score
    for suffix in JOB_LISTING_SUFFIXES:
        candidate = root + suffix
        if candidate.rstrip("/") == root:
            continue
        html      = None
        final_url = candidate
        r = simple_get(candidate)
        if r and r.url.rstrip("/") != root:
            html      = r.text
            final_url = r.url
        else:
            try:
                resp = await page.goto(candidate, timeout=15000, wait_until="domcontentloaded")
                if resp and resp.status < 400:
                    final_url = page.url
                    if final_url.rstrip("/") != root:
                        await page.wait_for_timeout(1200)
                        html = await page.content()
            except:
                pass
        if not html:
            continue
        score = _job_signal_count(html)
        if score > best_score:
            best_score = score
            best_url   = final_url
    return best_url, best_score


def _crawl_career_links(html, career_root):
    career_domain = get_domain(career_root)
    soup          = BeautifulSoup(html, "html.parser")
    candidates    = {}
    for a in soup.find_all("a", href=True):
        href = a.get("href", "")
        text = a.get_text(strip=True).lower()
        if not href or href.startswith(("#", "mailto", "tel")):
            continue
        full       = href if href.startswith("http") else urljoin(career_root, href)
        full_lower = full.lower()
        if "linkedin.com" in full_lower:
            continue
        if career_domain not in full_lower and not is_ats(full):
            continue
        score = 0
        if re.search(
            r"/(job.search.results?|job.search|jobs.search|career.search|"
            r"search|openings?|positions?|listings?|vacancies|all.jobs|current.opening)",
            full_lower,
        ):
            score += 8
        if re.search(r"/(job|career|role|vacanc|posting)s?[/_-]", full_lower):
            score += 5
        if re.search(r"/(en|ar)/(job|career|vacanc|opening|position|search)", full_lower):
            score += 7
        if re.search(
            r"\b(view|see|all|browse|search|explore)\b.*\b(job|position|opening|vacanc|career)",
            text,
        ):
            score += 6
        if re.search(r"\b(job|position|opening|vacanc|career)s?\b", text):
            score += 3
        if re.search(r"/go/[A-Za-z0-9%-]+/\d+", full_lower):
            score += 10
        if score >= 3:
            candidates[full] = max(candidates.get(full, 0), score)
    for url, score in sorted(candidates.items(), key=lambda x: -x[1])[:5]:
        return url
    return None


async def _resolve_iframe_ats(page, html, base_url):
    soup = BeautifulSoup(html, "html.parser")
    for iframe in soup.find_all("iframe"):
        src = iframe.get("src", "") or iframe.get("data-src", "")
        if not src:
            continue
        full = src if src.startswith("http") else urljoin(base_url, src)
        if is_ats(full):
            return full
    try:
        for frame in page.frames:
            furl = frame.url
            if furl and furl != base_url and furl != "about:blank":
                if is_ats(furl):
                    return furl
    except:
        pass
    m = re.search(r'<iframe[^>]+src=["\']?(https?://[^"\'>\s]+)["\']?', html, re.I)
    if m:
        src = m.group(1)
        if is_ats(src):
            return src
    return None


async def _resolve_to_jobs_url(page, career_root, strategy, ats_name):
    if _is_blocked_career_url(career_root):
        return None, strategy + "+blocked", None
    if ats_name:
        return career_root, strategy, ats_name
    r_root     = simple_get(career_root)
    root_html  = r_root.text if r_root else ""
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
        if found:
            return found, strategy + "+crawl", None
    try:
        await page.goto(career_root, timeout=25000, wait_until="domcontentloaded")
        await page.wait_for_timeout(2500)
        html = await page.content()
        detected_ats = is_ats_from_html(html)
        if detected_ats:
            m = re.search(r'(https?://[^\s"\'<>]*/go/[^\s"\'<>]+)', html, re.I)
            if m:
                return m.group(1), strategy + "+js_fp_go", detected_ats
            return career_root, strategy + "+js_fp", detected_ats
        for pat, aname in ATS_DOMAINS.items():
            if pat in html.lower():
                m = re.search(
                    r'https?://[^\s"\'<>]*' + re.escape(pat) + r'[^\s"\'<>]*', html, re.I
                )
                if m:
                    return m.group(0), strategy + "+js_ats", aname
        iframe_src = await _resolve_iframe_ats(page, html, career_root)
        if iframe_src:
            aname = is_ats(iframe_src)
            return iframe_src, strategy + "+js_iframe", aname
        js_score = _job_signal_count(html)
        suffix_url2, _ = await _probe_suffixes(page, career_root, max(root_score, js_score))
        if suffix_url2:
            return suffix_url2, strategy + "+js_suffix", None
        if js_score >= 2:
            return career_root, strategy + "+js", None
        found = _crawl_career_links(html, career_root)
        if found:
            return found, strategy + "+js_crawl", None
    except Exception as e:
        pass
    return career_root, strategy + "+unresolved", None


async def find_career_page(page, website):
    base_domain, scheme = get_base(website)
    base = f"{scheme}://{base_domain}"
    for sub in CAREER_SUBDOMAINS:
        url = f"{scheme}://{sub}.{base_domain}"
        r   = simple_get(url)
        if r:
            if _is_blocked_career_url(r.url):
                continue
            resolved = await _resolve_to_jobs_url(page, r.url, f"subdomain:{sub}", is_ats(r.url))
            if resolved[0] and not _is_blocked_career_url(resolved[0]):
                return resolved
    for path in CAREER_PATHS:
        url = base + path
        r   = simple_get(url)
        if r and r.url not in (base, base + "/", base + "/#"):
            if _is_blocked_career_url(r.url):
                continue
            resolved = await _resolve_to_jobs_url(page, r.url, f"path:{path}", is_ats(r.url))
            if resolved[0] and not _is_blocked_career_url(resolved[0]):
                return resolved
    r = simple_get(website)
    if r:
        result = _scan_links(BeautifulSoup(r.text, "html.parser"), website, base_domain)
        if result:
            career_root, strat, aname = result
            if not _is_blocked_career_url(career_root):
                return await _resolve_to_jobs_url(page, career_root, strat, aname)
    try:
        await page.goto(website, timeout=25000, wait_until="domcontentloaded")
        await page.wait_for_timeout(2000)
        content = await page.content()
        detected_ats = is_ats_from_html(content)
        if detected_ats:
            m_go = re.search(
                r'(https?://[^\s"\'<>]*/go/[A-Za-z0-9%-]+/\d+[^\s"\'<>]*)', content, re.I
            )
            if m_go:
                return m_go.group(1), "embedded:SF_go", detected_ats
            return website, f"embedded:{detected_ats}", detected_ats
        for pat, ats_name in ATS_DOMAINS.items():
            if pat in content.lower():
                m = re.search(
                    r'https?://[^\s"\'<>]*' + re.escape(pat) + r'[^\s"\'<>]*', content, re.I
                )
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
                if not href or href.startswith(("#", "mailto", "tel")):
                    continue
                full = href if href.startswith("http") else urljoin(website, href)
                if _is_blocked_career_url(full):
                    continue
                ats_name = is_ats(full)
                if ats_name:
                    return await _resolve_to_jobs_url(page, full, f"ats_link:{ats_name}", ats_name)
                ok, reason = is_career_link(href, text)
                if ok:
                    score = (10 if reason == "definite" else 5) + (3 if base_domain in full else 0)
                    if score > best_score:
                        best_score, best_url, best_strat = score, full, f"playwright:{reason}"
            except:
                continue
        if best_url and best_score >= 5 and not _is_blocked_career_url(best_url):
            return await _resolve_to_jobs_url(page, best_url, best_strat, None)
    except:
        pass
    for sm in [base + "/sitemap.xml", base + "/sitemap_index.xml"]:
        r = simple_get(sm)
        if r:
            hits = re.findall(
                r'<loc>(https?://[^<]*(?:career|job|vacanc|hiring)[^<]*)</loc>', r.text, re.I
            )
            valid = [h for h in hits if not _is_blocked_career_url(h)]
            if valid:
                return await _resolve_to_jobs_url(page, valid[0], "sitemap", is_ats(valid[0]))
    return None, "not_found", None


def _scan_links(soup, website, base_domain):
    best_score, best_url, best_strat, best_ats = 0, None, None, None
    for a in soup.find_all("a", href=True):
        href     = a.get("href", "")
        text     = clean(a.get_text(), 80).lower()
        if not href or href.startswith(("#", "mailto", "tel")):
            continue
        full     = href if href.startswith("http") else urljoin(website, href)
        if _is_blocked_career_url(full):
            continue
        ats_name = is_ats(full)
        if ats_name:
            return full, f"ats_link:{ats_name}", ats_name
        ok, reason = is_career_link(href, text)
        if ok:
            score = (10 if reason == "definite" else 5) + (3 if base_domain in full else 0)
            if score > best_score:
                best_score, best_url, best_strat = score, full, f"link:{reason}"
    return (best_url, best_strat, best_ats) if best_url and best_score >= 5 else None


# ══════════════════════════════════════════════════════════════
# ▌ SF BOARD SCRAPER
# ══════════════════════════════════════════════════════════════
async def _scrape_sf_board(page, careers_url, company, website, industry):
    stubs        = []
    root         = careers_url.rstrip("/")
    listing_url  = None
    listing_html = None
    best_score   = -1

    candidates = [careers_url]
    for s in SF_LISTING_SUFFIXES:
        candidates.append(root + s)
    for qp in SF_SEARCH_PARAMS:
        candidates.append(root + qp)

    for candidate in candidates:
        r = simple_get(candidate, timeout=15)
        if not r:
            continue
        html   = r.text
        soup_c = BeautifulSoup(html, "html.parser")
        table_stubs = _scrape_sf_listing_table(soup_c, r.url)
        if table_stubs:
            listing_url  = r.url
            listing_html = html
            stubs        = table_stubs
            break
        score = _job_signal_count(html)
        if score > best_score:
            best_score   = score
            listing_url  = r.url
            listing_html = html

    if not stubs:
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
                stubs       = table_stubs
                listing_url = final_url
        except Exception as e:
            pass

    if not stubs and listing_html:
        soup_fb = BeautifulSoup(listing_html, "html.parser")
        seen_fb = set()
        for a in soup_fb.find_all("a", href=True):
            href = a.get("href", "")
            if not href or href.startswith(("#", "mailto:", "tel:", "javascript")):
                continue
            full = href if href.startswith("http") else urljoin(listing_url or careers_url, href)
            if full in seen_fb:
                continue
            seen_fb.add(full)
            if not _is_detail_page_url(full):
                continue
            text = a.get_text(strip=True)
            if not text or len(text) < 3:
                parts = urlparse(full).path.rstrip("/").split("/")
                slug  = (parts[-2] if (len(parts) >= 2 and parts[-1].isdigit())
                         else parts[-1])
                text  = re.sub(r"[-_]", " ", slug).strip()
                text  = re.sub(r"\s+\d{5,}$", "", text).strip()
            if text and len(text) >= 3:
                stubs.append({"url": full, "title": text, "location": "", "date_posted": ""})

    return stubs


# ══════════════════════════════════════════════════════════════
# ▌ DOM HELPERS
# ══════════════════════════════════════════════════════════════
def _collect_job_links(soup, base_url, career_domain):
    seen, results = set(), []
    for a in soup.find_all("a", href=True):
        href = a.get("href", "")
        if not href or href.startswith(("#", "mailto:", "tel:", "javascript")):
            continue
        full = href if href.startswith("http") else urljoin(base_url, href)
        if full in seen:
            continue
        seen.add(full)
        if not is_likely_job_url(full, career_domain):
            continue
        text = a.get_text(strip=True)
        if not text or len(text) < 3:
            parent = a.parent
            for _ in range(3):
                if parent is None:
                    break
                heading = parent.find(["h1", "h2", "h3", "h4", "strong", "b"])
                if heading:
                    text = heading.get_text(strip=True)
                    break
                parent = parent.parent
        if not text or len(text) < 3:
            parts = urlparse(full).path.rstrip("/").split("/")
            slug  = parts[-2] if (len(parts) >= 2 and parts[-1].isdigit()) else parts[-1]
            text  = re.sub(r"[-_]", " ", slug).strip()
            text  = re.sub(r"\s+\d{5,}$", "", text).strip()
        if text and len(text) >= 3:
            results.append((full, text))
    return results


def _pagination(soup, base_url):
    urls = []
    for a in soup.find_all("a", href=True):
        text = a.get_text(strip=True).lower()
        href = a.get("href", "")
        if not href:
            continue
        full = href if href.startswith("http") else urljoin(base_url, href)
        if any(kw in text for kw in ["next", "›", "»", "load more", "show more"]):
            urls.append(full)
        elif re.match(r"^\d+$", text) and int(text) <= 50:
            urls.append(full)
    return list(dict.fromkeys(urls))


# ══════════════════════════════════════════════════════════════
# ▌ JOB EXTRACTOR (career pages)
# ══════════════════════════════════════════════════════════════
async def extract_jobs_from_career_page(page, careers_url, company, website, industry, ats_name):
    stub_jobs = []

    if ats_name == "Greenhouse" or "greenhouse.io" in careers_url:
        m = re.search(r"greenhouse\.io/([^/?#]+)", careers_url)
        if m:
            r = simple_get(f"https://boards.greenhouse.io/{m.group(1)}/jobs.json")
            if r:
                try:
                    for j in r.json().get("jobs", []):
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
                    for j in r.json().get("jobPostings", []):
                        stub_jobs.append({"url": j.get("jobUrl",""), "title": j.get("title","")})
                except: pass

    if ats_name == "SmartRecruiters" or "smartrecruiters.com" in careers_url:
        m = re.search(r"smartrecruiters\.com/([^/?#]+)", careers_url)
        if m:
            r = simple_get(f"https://api.smartrecruiters.com/v1/companies/{m.group(1)}/postings")
            if r:
                try:
                    for j in r.json().get("content", []):
                        stub_jobs.append({
                            "url": f"https://jobs.smartrecruiters.com/{m.group(1)}/{j.get('id','')}",
                            "title": j.get("name",""),
                        })
                except: pass

    if ats_name == "Wynt" or "wynt.ai" in careers_url:
        m = re.search(r"wynt\.ai/([^/?#]+)", careers_url)
        if m:
            org_slug = m.group(1)
            api_url  = f"https://apply.wynt.ai/api/v1/jobs?organization={org_slug}&limit=200"
            r        = simple_get(api_url)
            wynt_ok  = False
            if r:
                try:
                    data     = r.json()
                    job_list = data if isinstance(data, list) else data.get("results", data.get("jobs", []))
                    for j in job_list:
                        jid   = j.get("id", "") or j.get("slug", "")
                        title = j.get("title", "") or j.get("name", "")
                        jurl  = (j.get("url", "") or j.get("apply_url", "")
                                 or f"https://apply.wynt.ai/{org_slug}/jobs/{jid}")
                        stub_jobs.append({"url": jurl, "title": title, "_wynt_data": j})
                    if stub_jobs:
                        wynt_ok = True
                except: pass
            if not wynt_ok:
                try:
                    wynt_board = f"https://apply.wynt.ai/{org_slug}"
                    await page.goto(wynt_board, timeout=30000, wait_until="networkidle")
                    await page.wait_for_timeout(3000)
                    html      = await page.content()
                    soup      = BeautifulSoup(html, "html.parser")
                    job_links = _collect_job_links(soup, wynt_board, "wynt.ai")
                    for jurl, jtext in job_links[:200]:
                        stub_jobs.append({"url": jurl, "title": jtext})
                except:
                    pass

    if not stub_jobs and (
        ats_name == "SuccessFactors"
        or "successfactors" in careers_url.lower()
        or "/go/job-search/" in careers_url.lower()
        or "talentcommunity" in careers_url.lower()
    ):
        sf_stubs = await _scrape_sf_board(page, careers_url, company, website, industry)
        stub_jobs.extend(sf_stubs)

    if not stub_jobs:
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
                job_links     = _collect_job_links(soup, careers_url, career_domain)
                for jurl, jtext in job_links[:200]:
                    stub_jobs.append({"url": jurl, "title": jtext, "location": "", "date_posted": ""})
            if not stub_jobs:
                for next_url in _pagination(soup, careers_url)[:4]:
                    await page.goto(next_url, timeout=20000, wait_until="networkidle")
                    await page.wait_for_timeout(1500)
                    nsoup = BeautifulSoup(await page.content(), "html.parser")
                    extra = _collect_job_links(nsoup, next_url, get_domain(careers_url))
                    for jurl, jtext in extra:
                        stub_jobs.append({"url": jurl, "title": jtext,
                                           "location": "", "date_posted": ""})
        except Exception as e:
            pass

    if not stub_jobs:
        return []

    stub_jobs = [s for s in stub_jobs if _is_detail_page_url(s.get("url", ""))]
    if not stub_jobs:
        return []

    jobs      = []
    seen_urls = set()

    for stub in stub_jobs[:200]:
        jurl = stub.get("url", "")
        if not jurl or jurl in seen_urls:
            continue
        seen_urls.add(jurl)

        wynt_raw = stub.get("_wynt_data")
        if wynt_raw:
            loc_raw = wynt_raw.get("location", {})
            location = (loc_raw.get("city", "") + " " + loc_raw.get("country", "")).strip() \
                       if isinstance(loc_raw, dict) else str(loc_raw)
            sal = wynt_raw.get("salary", {})
            salary_range = ""
            if isinstance(sal, dict) and (sal.get("min") or sal.get("max")):
                salary_range = f"{sal.get('currency','')} {sal.get('min','')}–{sal.get('max','')}".strip()
            job = make_job(
                company=company, website=website, industry=industry,
                careers_url=careers_url, source="Wynt API",
                title=stub.get("title", ""),
                location=location or "Saudi Arabia",
                job_type=wynt_raw.get("employment_type", ""),
                department=wynt_raw.get("department", ""),
                apply_url=jurl,
                qualifications=wynt_raw.get("qualifications", ""),
                experience=wynt_raw.get("experience", ""),
                field=wynt_raw.get("category", ""),
                date_posted=wynt_raw.get("created_at", ""),
                deadline=wynt_raw.get("deadline", ""),
                description=wynt_raw.get("description", ""),
                salary_range=salary_range,
            )
            _JOB_COUNTER[0] += 1
            _stats["jobs_total"] += 1
            _stats["jobs_from_career_pages"] += 1
            if salary_range: _stats["jobs_with_salary"] += 1
            if job.get("Job Description"): _stats["jobs_with_description"] += 1
            print_job_card(job, _JOB_COUNTER[0], "career")
            jobs.append(job)
            continue

        detail = await scrape_job_detail(
            page, jurl, company, website, industry, careers_url, source="career_page",
            prefill_title=stub.get("title", ""),
            prefill_location=stub.get("location", ""),
            prefill_date=stub.get("date_posted", ""),
        )

        if detail:
            _JOB_COUNTER[0] += 1
            _stats["jobs_total"] += 1
            _stats["jobs_from_career_pages"] += 1
            print_job_card(detail, _JOB_COUNTER[0], "career")
            jobs.append(detail)
        else:
            fallback = make_job(
                company=company, website=website, industry=industry,
                careers_url=careers_url, source="career_listing_only",
                title=stub.get("title", ""),
                location=stub.get("location", ""),
                date_posted=stub.get("date_posted", ""),
                apply_url=jurl,
            )
            _JOB_COUNTER[0] += 1
            _stats["jobs_total"] += 1
            _stats["jobs_from_career_pages"] += 1
            print_job_card(fallback, _JOB_COUNTER[0], "career")
            jobs.append(fallback)

        await asyncio.sleep(random.uniform(0.4, 1.0))

    return jobs


# ══════════════════════════════════════════════════════════════
# ▌ PROCESS ONE COMPANY  (dual source: career page + LinkedIn)
# ══════════════════════════════════════════════════════════════
async def process_company(page, name, website, industry, cp):
    if not _looks_like_saudi_company(name, website):
        _stats["companies_skipped_non_saudi"] += 1
        vprint(f"⏭  Skipped (non-Saudi): {name} ({get_domain(website)})")
        return

    domain = get_domain(website)
    if not domain or domain in discovered_domains:
        return
    discovered_domains.add(domain)
    if domain in cp.get("processed_domains", []):
        vprint(f"⏭  Already processed: {name} ({domain})")
        return

    _stats["companies_seen"] += 1
    n = _stats["companies_seen"]
    print_company_header(name, website, industry, n)

    # ── Source 1: Company career page ─────────────────────────
    careers_url, strategy, ats_name = await find_career_page(page, website)

    if careers_url and _is_blocked_career_url(careers_url):
        careers_url = None

    career_jobs = []
    if not careers_url:
        _stats["companies_no_careers"] += 1
        vlog(f"❌  No career page found for {name}", indent=1)
    else:
        _stats["companies_with_careers"] += 1
        strat_key = strategy.split("+")[0]
        _stats["strategy_hits"][strat_key] = _stats["strategy_hits"].get(strat_key, 0) + 1
        if ats_name:
            _stats["ats_hits"][ats_name] = _stats["ats_hits"].get(ats_name, 0) + 1
        vlog(f"✅  Career page   : {careers_url}", indent=1)
        vlog(f"🔧  Strategy      : {strategy}", indent=1)
        vlog(f"🤖  ATS           : {ats_name or 'Custom'}", indent=1)
        career_jobs = await extract_jobs_from_career_page(
            page, careers_url, name, website, industry, ats_name
        )

    # ── Source 2: LinkedIn guest API ──────────────────────────
    li_jobs = await fetch_linkedin_jobs_for_company(name)
    # Enrich LinkedIn jobs with company website/industry from our known data
    for j in li_jobs:
        if not j.get("Company Website"):
            j["Company Website"] = website
        if not j.get("Company Industry"):
            j["Company Industry"] = industry
        _JOB_COUNTER[0] += 1
        _stats["jobs_total"] += 1
        print_job_card(j, _JOB_COUNTER[0], "LinkedIn")

    # ── Merge and dedup this company's jobs ───────────────────
    combined = career_jobs + li_jobs
    merged   = dedup_and_merge(combined)
    dupes    = len(combined) - len(merged)
    _stats["jobs_duplicates_merged"] += dupes
    # Adjust total counter for duplicates removed
    _stats["jobs_total"] -= dupes

    all_jobs.extend(merged)

    cp.setdefault("processed_domains", []).append(domain)
    cp["jobs_count"] = cp.get("jobs_count", 0) + len(merged)
    save_checkpoint(cp)

    company_results.append({
        "name": name, "website": website, "industry": industry,
        "careers_url": careers_url, "ats": ats_name or "Custom",
        "career_jobs": len(career_jobs), "linkedin_jobs": len(li_jobs),
        "merged_dupes": dupes, "net_jobs": len(merged),
    })
    flush_jobs()
    print_company_summary(name, careers_url, strategy, ats_name,
                           len(career_jobs), len(li_jobs), dupes)
    print_live_stats()


# ══════════════════════════════════════════════════════════════
# ▌ WIKIPEDIA DISCOVERY  (only source in v6)
# ══════════════════════════════════════════════════════════════
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


async def wikipedia_companies(page, cp, process_fn):
    if "wikipedia" in cp.get("done_sources", []):
        print("⏭  Wikipedia: already done")
        return
    print("\n📖  Discovery: Wikipedia (only source in v6)\n")
    for wiki_url in tqdm(WIKIPEDIA_PAGES, desc="  Wiki pages"):
        r = simple_get(wiki_url)
        if not r:
            continue
        soup = BeautifulSoup(r.text, "html.parser")
        for link in soup.select(".mw-category a, .wikitable a, #mw-content-text a"):
            title = link.get("title", "")
            href  = link.get("href", "")
            if not title or "Category:" in title or "List" in title:
                continue
            if NON_COMPANY_WIKI_PATTERNS.match(title.strip()):
                vprint(f"  ⏭  Wiki skip (non-company): {title}", indent=1)
                continue
            if not href.startswith("/wiki/"):
                continue
            await asyncio.sleep(random.uniform(0.5, 1.5))
            wr = simple_get("https://en.wikipedia.org" + href)
            if not wr:
                continue
            wsoup   = BeautifulSoup(wr.text, "html.parser")
            web_tag = wsoup.select_one(".infobox a.external")
            if not web_tag:
                continue
            raw_url = web_tag.get("href", "")
            domain_check = get_domain(raw_url)
            if domain_check in NON_SAUDI_DOMAINS:
                vprint(f"  ⏭  Wiki skip (blocked domain): {title} → {domain_check}", indent=1)
                continue
            domain, website = parse_domain(raw_url)
            if not domain:
                continue
            cats     = [c.get_text() for c in wsoup.select("#mw-normal-catlinks a")]
            industry = "General"
            for cat in cats:
                for ind in INDUSTRIES:
                    if ind.lower() in cat.lower():
                        industry = ind
                        break
            await process_fn(page, title, website, industry, cp)
        await asyncio.sleep(random.uniform(1, 2))
    cp.setdefault("done_sources", []).append("wikipedia")
    save_checkpoint(cp)


# ══════════════════════════════════════════════════════════════
# ▌ MAIN
# ══════════════════════════════════════════════════════════════
async def main():
    global all_jobs, company_results

    print("🚀  Saudi Jobs Pipeline v6 — Wikipedia · Career Pages + LinkedIn · Deduped\n")

    cp = load_checkpoint()

    if Path(JOBS_FILE).exists():
        all_jobs = pd.read_csv(JOBS_FILE).to_dict("records")
        print(f"   ▶  Resuming — {len(all_jobs)} jobs already saved")
    if Path(COMPANIES_FILE).exists():
        company_results = pd.read_csv(COMPANIES_FILE).to_dict("records")
    discovered_domains.update(cp.get("processed_domains", []))

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"]
        )
        context = await browser.new_context(
            user_agent=HEADERS["User-Agent"],
            locale="en-US",
            viewport={"width": 1280, "height": 800},
        )
        page = await context.new_page()

        await page.route(
            "**/*.{png,jpg,jpeg,gif,svg,woff,woff2,ttf,eot,ico}",
            lambda r: r.abort()
        )

        # Wikipedia is the ONLY discovery source
        await wikipedia_companies(page, cp, process_company)

        await browser.close()

    # Final global dedup pass across entire dataset
    vlog("🔄  Running final global dedup pass…")
    before_dedup = len(all_jobs)
    all_jobs = dedup_and_merge(all_jobs)
    after_dedup = len(all_jobs)
    if before_dedup != after_dedup:
        vlog(f"   Removed {before_dedup - after_dedup} additional cross-company duplicates")

    flush_jobs()

    # ── Summary ───────────────────────────────────────────────
    total_jobs      = len(all_jobs)
    total_companies = len(company_results)
    with_careers    = sum(1 for c in company_results if c.get("careers_url"))
    W = 72

    print(f"\n{'═'*W}")
    print(f"  🏁  PIPELINE v6 COMPLETE")
    print(f"{'─'*W}")
    print(f"  {'Companies discovered:':<30} {total_companies:>6,}")
    print(f"  {'  with career pages:':<30} {with_careers:>6,}  ({100*with_careers//max(total_companies,1)}%)")
    print(f"  {'  no career page:':<30} {total_companies-with_careers:>6,}")
    print(f"  {'  skipped (non-Saudi):':<30} {_stats['companies_skipped_non_saudi']:>6,}")
    print(f"{'─'*W}")
    print(f"  {'Total jobs (deduped):':<30} {total_jobs:>6,}")
    print(f"  {'  from career pages:':<30} {_stats['jobs_from_career_pages']:>6,}")
    print(f"  {'  from LinkedIn:':<30} {_stats['jobs_from_linkedin']:>6,}")
    print(f"  {'  duplicates merged:':<30} {_stats['jobs_duplicates_merged']:>6,}")
    print(f"  {'  with salary data:':<30} {_stats['jobs_with_salary']:>6,}  ({100*_stats['jobs_with_salary']//max(total_jobs,1)}%)")
    print(f"  {'  with description:':<30} {_stats['jobs_with_description']:>6,}  ({100*_stats['jobs_with_description']//max(total_jobs,1)}%)")
    print(f"  {'Detail fetches:':<30} {_stats['detail_fetches']:>6,}")
    print(f"  {'Detail failures:':<30} {_stats['detail_failures']:>6,}  ({100*_stats['detail_failures']//max(_stats['detail_fetches'],1)}%)")
    print(f"{'─'*W}")

    if _stats["ats_hits"]:
        print(f"  ATS breakdown:")
        for ats, count in sorted(_stats["ats_hits"].items(), key=lambda x: -x[1]):
            print(f"    {ats:<22} {count:>4}  {_bar(count, total_companies, width=30)}")
        print(f"{'─'*W}")

    if all_jobs:
        jobs_df = pd.DataFrame(all_jobs)
        jobs_df.to_csv(JOBS_FILE, index=False)

        print(f"\n  Top 15 companies by jobs scraped:")
        top = (jobs_df.groupby("Company Name")["Job Title"]
               .count().sort_values(ascending=False).head(15))
        for co, cnt in top.items():
            print(f"    {co[:30]:<30}  {cnt:>4}  {_bar(cnt, top.iloc[0], width=25)}")

        if "Job Field" in jobs_df.columns:
            print(f"\n  Top job fields:")
            fields = (jobs_df["Job Field"].dropna()
                      .replace("", pd.NA).dropna()
                      .value_counts().head(10))
            for field, cnt in fields.items():
                print(f"    {field[:30]:<30}  {cnt:>4}")

        print(f"\n  Source breakdown:")
        srcs = jobs_df["source"].str.split("+").str[0].value_counts()
        for src, cnt in srcs.items():
            print(f"    {src:<30}  {cnt:>4}")

    print(f"\n{'═'*W}")
    print(f"  💾  {JOBS_FILE}  |  {COMPANIES_FILE}")
    print(f"{'═'*W}")

    try:
        from google.colab import files
        files.download(JOBS_FILE)
        files.download(COMPANIES_FILE)
    except:
        print(f"\n   Files saved locally.")


asyncio.run(main())
