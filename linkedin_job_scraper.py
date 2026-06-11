# ═══════════════════════════════════════════════════════════════
# SAUDI JOBS PIPELINE — v6
# Discover company → find career page → scrape jobs (all at once)
#
# v6 NEW DISCOVERY SOURCES (on top of v5):
#  1. eArabicMarket  — Saudi company directory with websites
#  2. Wikidata SPARQL — structured Saudi company data (cleaner than
#                       Wikipedia HTML scraping)
#  3. Saudi Chambers of Commerce (chamber.org.sa) — official registry
#  4. Zawya           — MENA premium business directory
#  5. GulfTalent      — Gulf job board, employer profile pages
#  6. Bayt.com        — MENA's largest job board, employer pages
#  7. Indeed Saudi    — employer profile discovery
#  8. Naukrigulf      — Gulf-specific job board with company pages
#  9. LinkedIn (discovery-only) — search results page for Saudi
#                       company names + websites (NOT scraped for jobs)
# 10. Glassdoor       — Saudi company discovery only
# 11. SaudiCommerce / Exporters lists (modon.gov.sa, saudiexporters.sa)
# 12. Tadawul (Argaam) — Saudi stock exchange listed companies
# 13. CrunchBase-style — wamda.com / magnitt.com Saudi startups
#
# v6 DEDUPLICATION — global domain registry prevents any company
#  being processed twice regardless of which source found it first.
#
# All v5 fixes retained:
#  • Saudi-only geographic filter
#  • LinkedIn career-page block
#  • Clean standardised field logging
#  • Field inference score ≥ 1
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
    "requests", "nest_asyncio", "tqdm", "lxml", "-q"], check=True)

subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"],
    capture_output=True)

import importlib, site
importlib.invalidate_caches()
for sp in site.getsitepackages():
    if sp not in sys.path:
        sys.path.insert(0, sp)

import asyncio, re, json, random, time, csv
from pathlib import Path
from urllib.parse import urlparse, urljoin, quote_plus
from datetime import datetime, timedelta

import requests
import pandas as pd
import nest_asyncio
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright
from tqdm import tqdm

nest_asyncio.apply()
print("✅ All imports successful\n")


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
    # Job boards — valid for discovery but not as company websites
    "indeed.com", "glassdoor.com", "bayt.com", "gulftalent.com",
    "naukrigulf.com", "monster.com", "ziprecruiter.com",
    "zawya.com", "argaam.com", "mubasher.info",
    "wamda.com", "magnitt.com",
    "earabicmarket.com", "kompass.com",
    "chamber.org.sa",        # directory itself, not a company
    "modon.gov.sa",
    "saudiexporters.sa",
}

SAUDI_DOMAIN_INDICATORS = [
    r"\.sa$",
    r"\.gov\.sa$",
    r"\.edu\.sa$",
    r"\.com\.sa$",
    r"\.org\.sa$",
    r"aramco",
    r"sabic",
    r"stc\.com",
    r"alrajhi",
    r"samba",
    r"riyad",
    r"ncb",
    r"jarir",
    r"maaden",
    r"tasnee",
    r"mobily",
    r"zain\.sa",
    r"saudia",
    r"flynas",
    r"flyadeal",
    r"neom",
    r"vision2030",
    r"pif\.gov",
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
# ▌ STANDARDISATION TABLES (unchanged from v5)
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

    best_label = ""
    best_score = 0

    for label, high_kws, low_kws in FIELD_KEYWORD_MAP:
        score = 0
        for kw in high_kws:
            if kw in combined:
                score += 3
        for kw in low_kws:
            if kw in combined:
                score += 1
        if score > best_score:
            best_score = score
            best_label = label

    if best_score >= 1:
        return best_label
    return ""


# ── EXPERIENCE ────────────────────────────────────────────────
_NO_EXP_KW = [
    "no experience", "no prior experience", "fresh graduate", "freshers",
    "entry level", "entry-level", "0 years", "zero experience",
    "training provided", "will train", "no experience required",
    "no experience needed",
]
_LT1_KW = [
    "less than 1 year", "under 1 year", "6 months", "less than a year",
    "some experience", "minimal experience", "up to 1 year",
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
        if kw in text:
            return "No Experience Required"
    for kw in _LT1_KW:
        if kw in text:
            return "Less than 1 Year"
    matches = _EXP_RE.findall(text)
    if matches:
        nums = []
        for m in matches:
            for g in m:
                if g:
                    try: nums.append(int(g))
                    except: pass
        if nums:
            return _years_to_band(min(nums))
    m = re.search(r"(\d+)\s*\+?\s*years?", text, re.I)
    if m:
        return _years_to_band(int(m.group(1)))
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
# VERBOSE LOGGING SYSTEM
# ══════════════════════════════════════════════════════════════
VERBOSE = True

_stats = {
    "companies_seen": 0,
    "companies_with_careers": 0,
    "companies_no_careers": 0,
    "companies_skipped_non_saudi": 0,
    "jobs_total": 0,
    "jobs_with_title": 0,
    "jobs_with_salary": 0,
    "jobs_with_description": 0,
    "detail_fetches": 0,
    "detail_failures": 0,
    "ai_assists": 0,
    "ats_hits": {},
    "strategy_hits": {},
    "source_companies": {},   # v6: track per-source company counts
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


def print_job_card(job: dict, serial: int, company_serial: int, total_this_company: int):
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

    print(f"\n  ┌{SEP}┐")
    print(f"  │  JOB #{serial:04d}  [{company_serial}/{total_this_company} this company]" + " " * (W - 43) + "│")
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
    if desc:
        lines = [desc[i:i+W-4] for i in range(0, min(len(desc), (W-4)*3), W-4)]
        print(f"  │  📝 Description:  {' ' * (W - 19)}│")
        for ln in lines:
            print(f"  │    {ln:<{W-4}}│")
    else:
        print(f"  │  📝 Description:  {'—':<{W-20}}│")
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


def print_company_summary(name, careers_url, strategy, ats_name, jobs_found):
    W = 72
    ats_str  = ats_name or "Custom"
    jobs_bar = _bar(jobs_found, max(jobs_found, 1))
    print(f"  {'─'*W}")
    if careers_url:
        print(f"  ✅  Career page : {careers_url[:W-18]}")
    else:
        print(f"  ❌  No career page found")
    print(f"  🔧  Strategy    : {strategy}")
    print(f"  🤖  ATS         : {ats_str}")
    print(f"  📋  Jobs found  : {jobs_found}  {jobs_bar}")
    print(f"{'═'*W}")


def print_live_stats():
    s = _stats
    W = 72
    print(f"\n  ┌── LIVE STATS {'─'*(W-15)}┐")
    print(f"  │  Companies : seen={s['companies_seen']}  "
          f"with_careers={s['companies_with_careers']}  "
          f"no_careers={s['companies_no_careers']}  "
          f"skipped={s['companies_skipped_non_saudi']}")
    print(f"  │  Jobs      : total={s['jobs_total']}  "
          f"with_salary={s['jobs_with_salary']}  "
          f"with_desc={s['jobs_with_description']}")
    print(f"  │  Detail    : fetches={s['detail_fetches']}  "
          f"failures={s['detail_failures']}  "
          f"ai_assists={s['ai_assists']}")
    if s["ats_hits"]:
        ats_str = "  ".join(
            f"{k}:{v}" for k, v in sorted(s["ats_hits"].items(), key=lambda x: -x[1])
        )
        print(f"  │  ATS hits  : {ats_str}")
    if s["source_companies"]:
        src_str = "  ".join(
            f"{k}:{v}" for k, v in sorted(s["source_companies"].items(), key=lambda x: -x[1])[:6]
        )
        print(f"  │  Sources   : {src_str[:W-14]}")
    print(f"  └{'─'*W}┘")


# ── Output files ──────────────────────────────────────────────
JOBS_FILE        = "saudi_jobs.csv"
COMPANIES_FILE   = "saudi_companies_found.csv"
CHECKPOINT_FILE  = "pipeline_checkpoint.json"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
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
    "bayt", "naukrigulf", "gulftalent", "monster", "ziprecruiter",
    "wamda", "magnitt", "argaam", "mubasher", "earabicmarket",
}

BLOCKED_CAREER_DOMAINS = {
    "linkedin.com",
    "www.linkedin.com",
}

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
    "kenexa.com": "Kenexa",
    "silkroad.com": "SilkRoad",
    "ceipal.com": "Ceipal",
    "cornerstone": "Cornerstone",
    "lumesse.com": "Lumesse",
    "talentsoft.com": "TalentSoft",
    "hrcloud.com": "HRCloud",
    "hibob.com": "HiBob",
    "rippling.com": "Rippling",
    "personio.com": "Personio",
    "easy.jobs": "EasyJobs",
    "apply.wynt.ai": "Wynt",
    "wynt.ai": "Wynt",
    "recruitcrm.io": "RecruitCRM",
    "hirepos.com": "HirePos",
    "hirize.hr": "Hirize",
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


_GARBAGE_FIELD_PATTERNS = re.compile(
    r"^(job\s*description|key\s*accountabilities?|key\s*functional|"
    r"requirements?|qualifications?|responsibilities|overview|"
    r"role\s*purpose|job\s*purpose|about\s*the\s*role|"
    r"what\s*you|we\s*are\s*looking|our\s*team|"
    r"education\s*&|experience\s*&|skills\s*&|"
    r"others?|not\s*applicable|n/?a|—|-)$",
    re.I,
)


def _sanitize_field(value, max_len=300):
    if not value:
        return ""
    v = re.sub(r"\s+", " ", str(value)).strip()
    if _GARBAGE_FIELD_PATTERNS.match(v):
        return ""
    if len(v) < 2:
        return ""
    return v[:max_len]


# ══════════════════════════════════════════════════════════════
# SHARED STATE
# ══════════════════════════════════════════════════════════════
discovered_domains = set()
all_jobs           = []
company_results    = []


# ══════════════════════════════════════════════════════════════
# HELPERS
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


def flush_jobs():
    if all_jobs:
        pd.DataFrame(all_jobs).drop_duplicates(subset="Job URL").to_csv(JOBS_FILE, index=False)
    if company_results:
        pd.DataFrame(company_results).to_csv(COMPANIES_FILE, index=False)


# ══════════════════════════════════════════════════════════════
# REGEX EXTRACTORS  (unchanged)
# ══════════════════════════════════════════════════════════════
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
            mn   = val.get("minValue", "")
            mx   = val.get("maxValue", "")
            unit = val.get("unitText", "")
            cur  = bs.get("currency", "")
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


# ══════════════════════════════════════════════════════════════
# SF TABLE SCRAPER / SECTION PARSER / DETAIL SCRAPER
# (unchanged from v5 — full code retained)
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

    department = _sanitize_field(
        ld.get("field")
        or _find_text_near_label(soup, "Department", "Team", "Function",
                                  "Division", "القسم", "الإدارة")
        or ""
    )
    if not department:
        h2_tags = soup.find_all("h2")
        if h2_tags:
            first_h2 = h2_tags[0].get_text(strip=True)
            if (first_h2 and len(first_h2) < 80
                    and not re.search(
                        r"(apply|job search|search|opportunities|alert|description)",
                        first_h2, re.I)
                    and _sanitize_field(first_h2)):
                department = first_h2

    job_type = _sanitize_field(
        ld.get("job_type")
        or _find_text_near_label(soup, "Employment Type", "Job Type",
                                  "Contract Type", "Type", "نوع الوظيفة")
        or _pick_section(sections, "employment type", "job type")
        or ""
    )
    if not job_type:
        for jt in ["Full-time", "Part-time", "Contract", "Internship",
                   "Freelance", "Temporary", "Permanent"]:
            if re.search(r"\b" + jt.split("-")[0] + r"\b", full_text, re.I):
                job_type = jt
                break

    location = (
        ld.get("location") or prefill_location
        or _find_text_near_label(soup, "Location", "Job Location", "City",
                                  "الموقع", "المدينة")
        or ""
    )
    if not location:
        m = LOCATION_PATTERN.search(full_text)
        if m:
            location = m.group(1)
    location = location or "Saudi Arabia"

    description = ld.get("description", "")
    if not description:
        description = _pick_section(
            sections,
            "key functional", "responsibilities", "job purpose", "about the role",
            "what you'll do", "role overview", "job summary", "overview", "purpose",
            "المهام", "المسؤوليات",
        )
    if not description:
        for sel in ["[class*='description']", "[class*='job-desc']",
                    "[class*='content']", "#job-description", "article", "main",
                    "[class*='detail']", "[class*='body']"]:
            el = soup.select_one(sel)
            if el and len(el.get_text()) > 100:
                description = clean(el.get_text(), 2000)
                break
    if not description and len(full_text) > 200:
        description = clean(full_text, 2000)

    qualifications = (
        ld.get("qualifications")
        or _pick_section(sections, "qualif", "requirement", "education",
                         "what you need", "you should have", "minimum requirement",
                         "المؤهلات", "المتطلبات")
        or _find_text_near_label(soup, "Qualifications", "Requirements",
                                  "Education", "Degree", "المؤهلات")
        or ""
    )
    if not qualifications:
        qualifications = _extract_bold_field(
            soup, "Education", "Qualifications", "Requirements", "Minimum Qualifications"
        )
    if _GARBAGE_FIELD_PATTERNS.match(qualifications.strip()[:80]):
        qualifications = ""

    experience = _sanitize_field(
        ld.get("experience")
        or _find_text_near_label(soup, "Experience", "Years of Experience",
                                  "الخبرة", "سنوات الخبرة")
        or _pick_section(sections, "experience")
        or ""
    )
    if not experience:
        experience = _sanitize_field(
            _extract_bold_field(soup, "Experience", "Years of Experience",
                                "Minimum Experience")
        )
    if not experience:
        m = EXPERIENCE_RE.search(full_text)
        if m:
            experience = m.group(0)

    field = _sanitize_field(
        ld.get("field") or department
        or _find_text_near_label(soup, "Field", "Category", "Department",
                                  "Function", "التخصص", "القسم")
        or industry
    )

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
        or _find_text_near_label(soup, "Salary", "Compensation", "الراتب", "المرتب")
        or ""
    )
    if not salary_range:
        m = SALARY_RE.search(full_text)
        if m:
            salary_range = m.group(0)
    if salary_range:
        if re.search(r"(linkedin|user.agreement|terms\s+of\s+service)",
                     full_text[:500].lower()):
            salary_range = ""

    company_logo   = ld.get("company_logo", "") or _extract_logo(soup, final_url)
    company_type   = _sanitize_field(
        ld.get("company_type", "")
        or _find_text_near_label(soup, "Company Type", "Organization Type", "نوع الشركة")
        or ""
    )
    if company_type and company_type.lower() in ("organization", "legalservice", "thing"):
        company_type = ""
    company_address = _find_text_near_label(
        soup, "Address", "Headquarters", "Head Office", "العنوان", "المقر الرئيسي"
    )
    company_founded = _sanitize_field(
        _find_text_near_label(soup, "Founded", "Established", "Year Founded", "تأسست")
    )
    company_details = _pick_section(
        sections, "about the company", "about ceer", "about us",
        "company overview", "who we are", "من نحن"
    )

    apply_url = final_url
    apply_patterns = ["talentcommunity/apply", "/apply/", "?action=apply", "/application/"]
    for a in soup.find_all("a", href=True):
        href       = a.get("href", "")
        text_a     = a.get_text(strip=True).lower()
        href_lower = href.lower()
        if any(p in href_lower for p in apply_patterns):
            apply_url = href if href.startswith("http") else urljoin(final_url, href)
            break
        if "apply" in text_a and href and href not in ("#", "javascript:void(0)"):
            candidate = href if href.startswith("http") else urljoin(final_url, href)
            if candidate != final_url:
                apply_url = candidate

    vprint(f"  ┄ title        : {title[:60] or '—'}", indent=4)
    vprint(f"  ┄ type         : {job_type or '—'}", indent=4)
    vprint(f"  ┄ location     : {location or '—'}", indent=4)
    vprint(f"  ┄ field        : {standardise_field(field or department, title, description, industry) or '—'}", indent=4)
    vprint(f"  ┄ experience   : {standardise_experience(experience) or '—'}", indent=4)
    vprint(f"  ┄ qualification: {standardise_qualification(qualifications, description) or '—'}", indent=4)
    vprint(f"  ┄ salary       : {salary_range or '—'}", indent=4)
    vprint(f"  ┄ date posted  : {date_posted or '—'}", indent=4)
    vprint(f"  ┄ deadline     : {deadline or '—'}", indent=4)
    vprint(f"  ┄ est deadline : {estimated_deadline or '—'}", indent=4)
    vprint(f"  ┄ description  : {description[:80] or '—'}{'…' if len(description)>80 else ''}", indent=4)
    vprint(f"  ┄ logo         : {company_logo[:60] or '—'}", indent=4)
    vprint(f"  ┄ apply url    : {apply_url[:80]}", indent=4)

    if salary_range:
        _stats["jobs_with_salary"] += 1
    if description:
        _stats["jobs_with_description"] += 1

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
# IFRAME / CAREER PAGE FINDER / JOB EXTRACTOR
# (unchanged from v5 — full logic retained)
# ══════════════════════════════════════════════════════════════
async def _resolve_iframe_ats(page, html, base_url):
    soup = BeautifulSoup(html, "html.parser")
    for iframe in soup.find_all("iframe"):
        src = iframe.get("src", "") or iframe.get("data-src", "")
        if not src:
            continue
        full = src if src.startswith("http") else urljoin(base_url, src)
        if is_ats(full):
            print(f"       🖼️  Iframe ATS detected: {full}")
            return full
    try:
        for frame in page.frames:
            furl = frame.url
            if furl and furl != base_url and furl != "about:blank":
                if is_ats(furl):
                    print(f"       🖼️  Live iframe ATS frame: {furl}")
                    return furl
    except:
        pass
    m = re.search(r'<iframe[^>]+src=["\']?(https?://[^"\'>\s]+)["\']?', html, re.I)
    if m:
        src = m.group(1)
        if is_ats(src):
            print(f"       🖼️  Regex iframe ATS: {src}")
            return src
    return None


def _job_signal_count(html):
    lower = html.lower()
    return sum(1 for s in JOB_PAGE_SIGNALS if s in lower)


def _is_blocked_career_url(url: str) -> bool:
    domain = get_domain(url)
    return domain in BLOCKED_CAREER_DOMAINS


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
            print(f"       🔍  Better jobs URL via suffix '{suffix}': score={score}  →  {final_url}")
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
        print(f"       🔗  Jobs listing candidate (score={score}): {url}")
        return url
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
        suffix_url2, suffix_score2 = await _probe_suffixes(
            page, career_root, max(root_score, js_score)
        )
        if suffix_url2:
            return suffix_url2, strategy + "+js_suffix", None
        if js_score >= 2:
            return career_root, strategy + "+js", None
        found = _crawl_career_links(html, career_root)
        if found:
            return found, strategy + "+js_crawl", None
    except Exception as e:
        print(f"       ⚠️  JS resolve error: {e}")
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
    try:
        r = simple_get(
            f"https://html.duckduckgo.com/html/?q={quote_plus(f'site:{base_domain} careers jobs')}"
        )
        if r:
            for el in BeautifulSoup(r.text, "html.parser").select(".result__url"):
                u = el.get_text(strip=True)
                if not u.startswith("http"):
                    u = "https://" + u
                if (base_domain in u
                        and any(kw in u.lower() for kw in ["career", "job", "vacanc"])
                        and not _is_blocked_career_url(u)):
                    return await _resolve_to_jobs_url(page, u, "duckduckgo", is_ats(u))
    except:
        pass
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
            print(f"        📊  SF table found at: {r.url}  ({len(table_stubs)} rows)")
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
                print(f"        📊  SF table found via Playwright ({len(table_stubs)} rows)")
                stubs       = table_stubs
                listing_url = final_url
        except Exception as e:
            print(f"        ⚠️  SF Playwright render error: {e}")

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
        print(f"        🔎  SF fallback link collect (filtered): {len(stubs)} candidates")

    if stubs and listing_html:
        soup_pg   = BeautifulSoup(listing_html, "html.parser")
        per_page_m = re.search(r"Page\s+1\s+of\s+(\d+)", soup_pg.get_text(), re.I)
        if per_page_m:
            total_pages = int(per_page_m.group(1))
            if total_pages > 1:
                print(f"        📄  SF pagination: {total_pages} pages")
                for pg in range(2, min(total_pages + 1, 20)):
                    pg_url = re.sub(r"[?&]page=\d+", "", listing_url or careers_url)
                    sep    = "&" if "?" in pg_url else "?"
                    pg_url = f"{pg_url}{sep}page={pg}"
                    r2     = simple_get(pg_url)
                    if not r2:
                        break
                    soup2      = BeautifulSoup(r2.text, "html.parser")
                    page_stubs = _scrape_sf_listing_table(soup2, listing_url or careers_url)
                    if not page_stubs:
                        break
                    stubs.extend(page_stubs)
                    await asyncio.sleep(0.5)

    return stubs


async def extract_jobs(page, careers_url, company, website, industry, ats_name):
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

    if ats_name == "Workday" or "myworkdayjobs.com" in careers_url:
        m = re.search(r"(https?://[^/]+myworkdayjobs\.com/[^/?#]+)", careers_url)
        if m:
            r = simple_get(m.group(1) + "/jobs?format=json")
            if r:
                try:
                    for j in r.json().get("jobPostings", []):
                        stub_jobs.append({
                            "url": m.group(1) + "/job/" + j.get("externalPath",""),
                            "title": j.get("title",""),
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
                        print(f"        ✅  Wynt API: {len(stub_jobs)} jobs for {org_slug}")
                except: pass
            if not wynt_ok:
                try:
                    wynt_board = f"https://apply.wynt.ai/{org_slug}"
                    await page.goto(wynt_board, timeout=30000, wait_until="networkidle")
                    await page.wait_for_timeout(3000)
                    for _ in range(4):
                        await page.evaluate("window.scrollBy(0, 800)")
                        await page.wait_for_timeout(600)
                    html         = await page.content()
                    soup         = BeautifulSoup(html, "html.parser")
                    job_links    = _collect_job_links(soup, wynt_board, "wynt.ai")
                    for jurl, jtext in job_links[:200]:
                        stub_jobs.append({"url": jurl, "title": jtext})
                except Exception as e:
                    print(f"        ⚠️  Wynt SPA render error: {e}")

    if not stub_jobs and (
        ats_name == "SuccessFactors"
        or "successfactors" in careers_url.lower()
        or "/go/job-search/" in careers_url.lower()
        or "/go/all-jobs/" in careers_url.lower()
        or "talentcommunity" in careers_url.lower()
    ):
        print(f"        🔧  Using SuccessFactors board scraper…")
        sf_stubs = await _scrape_sf_board(page, careers_url, company, website, industry)
        stub_jobs.extend(sf_stubs)
        print(f"        ✅  SF scraper: {len(stub_jobs)} job stubs")

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
                    print(f"        📊  Late SF table detection: {len(table_stubs)} jobs")
            if not stub_jobs:
                career_domain = get_domain(careers_url)
                job_links     = _collect_job_links(soup, careers_url, career_domain)
                print(f"        🔎  {len(job_links)} job URL candidates on listing page")
                for jurl, jtext in job_links[:200]:
                    stub_jobs.append({"url": jurl, "title": jtext, "location": "", "date_posted": ""})
            if not stub_jobs:
                for next_url in _pagination(soup, careers_url)[:4]:
                    await page.goto(next_url, timeout=20000, wait_until="networkidle")
                    await page.wait_for_timeout(1500)
                    nsoup = BeautifulSoup(await page.content(), "html.parser")
                    extra = _collect_job_links(nsoup, next_url, career_domain)
                    for jurl, jtext in extra:
                        stub_jobs.append({"url": jurl, "title": jtext,
                                           "location": "", "date_posted": ""})
        except Exception as e:
            print(f"        ⚠️  DOM error: {e}")

    if not stub_jobs:
        print(f"        ⚠️  0 job URLs found on listing page")
        return []

    before    = len(stub_jobs)
    stub_jobs = [s for s in stub_jobs if _is_detail_page_url(s.get("url", ""))]
    after     = len(stub_jobs)
    if before != after:
        vlog(f"🔍  Stub validation: {before} → {after} (removed {before-after} non-detail URLs)", indent=2)

    if not stub_jobs:
        print(f"        ⚠️  0 valid detail URLs after validation")
        return []

    vlog(f"📄  Scraping detail pages for {len(stub_jobs)} jobs…", indent=2)
    jobs      = []
    seen_urls = set()

    for stub in stub_jobs[:200]:
        jurl = stub.get("url", "")
        if not jurl or jurl in seen_urls:
            continue
        seen_urls.add(jurl)

        wynt_raw = stub.get("_wynt_data")
        if wynt_raw:
            loc_raw  = wynt_raw.get("location", {})
            location = (loc_raw.get("city", "") + " " + loc_raw.get("country", "")).strip() \
                       if isinstance(loc_raw, dict) else str(loc_raw)
            sal      = wynt_raw.get("salary", {})
            salary_range = ""
            if isinstance(sal, dict) and (sal.get("min") or sal.get("max")):
                salary_range = f"{sal.get('currency','')} {sal.get('min','')}–{sal.get('max','')}".strip()
            job = make_job(
                company=company, website=website, industry=industry,
                careers_url=careers_url, source="Wynt API",
                title=stub.get("title", ""),
                location=location or "Saudi Arabia",
                job_type=wynt_raw.get("employment_type", "") or wynt_raw.get("job_type", ""),
                department=wynt_raw.get("department", "") or wynt_raw.get("team", ""),
                apply_url=jurl,
                qualifications=wynt_raw.get("qualifications", "") or wynt_raw.get("requirements", ""),
                experience=wynt_raw.get("experience", "") or wynt_raw.get("experience_level", ""),
                field=wynt_raw.get("category", "") or wynt_raw.get("field", ""),
                date_posted=wynt_raw.get("created_at", "") or wynt_raw.get("posted_at", ""),
                deadline=wynt_raw.get("deadline", "") or wynt_raw.get("expires_at", ""),
                description=wynt_raw.get("description", ""),
                salary_range=salary_range,
            )
            _JOB_COUNTER[0] += 1
            _stats["jobs_total"] += 1
            if salary_range: _stats["jobs_with_salary"] += 1
            if job.get("Job Description"): _stats["jobs_with_description"] += 1
            print_job_card(job, _JOB_COUNTER[0], len(jobs) + 1, len(stub_jobs))
            jobs.append(job)
            continue

        detail = await scrape_job_detail(
            page, jurl, company, website, industry, careers_url, source="detail",
            prefill_title=stub.get("title", ""),
            prefill_location=stub.get("location", ""),
            prefill_date=stub.get("date_posted", ""),
        )

        if detail:
            _JOB_COUNTER[0] += 1
            _stats["jobs_total"] += 1
            if detail.get("Job Title"): _stats["jobs_with_title"] += 1
            print_job_card(detail, _JOB_COUNTER[0], len(jobs) + 1, len(stub_jobs))
            jobs.append(detail)
        else:
            fallback = make_job(
                company=company, website=website, industry=industry,
                careers_url=careers_url, source="listing_only",
                title=stub.get("title", ""),
                location=stub.get("location", ""),
                date_posted=stub.get("date_posted", ""),
                apply_url=jurl,
            )
            _JOB_COUNTER[0] += 1
            _stats["jobs_total"] += 1
            vlog(f"  ⚠️  Detail failed — stub record: {stub.get('title','')[:60]}", indent=3)
            print_job_card(fallback, _JOB_COUNTER[0], len(jobs) + 1, len(stub_jobs))
            jobs.append(fallback)

        await asyncio.sleep(random.uniform(0.4, 1.0))

    vlog(f"✅  {len(jobs)} jobs scraped for this company", indent=2)
    return jobs


# ══════════════════════════════════════════════════════════════
# DOM HELPERS
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
# PROCESS ONE COMPANY  (v6: track discovery source)
# ══════════════════════════════════════════════════════════════
async def process_company(page, name, website, industry, cp, source_tag=""):
    if not _looks_like_saudi_company(name, website):
        _stats["companies_skipped_non_saudi"] += 1
        vprint(f"⏭  Skipped (non-Saudi): {name} ({get_domain(website)})", indent=0)
        return

    domain = get_domain(website)
    if not domain or domain in discovered_domains:
        return
    discovered_domains.add(domain)
    if domain in cp.get("processed_domains", []):
        vprint(f"⏭  Already processed: {name} ({domain})")
        return

    # v6: track per-source company counts
    if source_tag:
        _stats["source_companies"][source_tag] = _stats["source_companies"].get(source_tag, 0) + 1

    _stats["companies_seen"] += 1
    n = _stats["companies_seen"]
    print_company_header(name, website, industry, n)

    careers_url, strategy, ats_name = await find_career_page(page, website)

    if careers_url and _is_blocked_career_url(careers_url):
        vlog(f"⏭  Career page leads to blocked domain ({get_domain(careers_url)}) — skipping", indent=1)
        careers_url = None

    if not careers_url:
        _stats["companies_no_careers"] += 1
        vlog(f"❌  No career page found for {name}", indent=1)
        cp.setdefault("processed_domains", []).append(domain)
        save_checkpoint(cp)
        company_results.append({
            "name": name, "website": website, "industry": industry,
            "careers_url": None, "ats": None, "jobs_found": 0, "source": source_tag,
        })
        flush_jobs()
        print_company_summary(name, None, strategy, ats_name, 0)
        print_live_stats()
        return

    _stats["companies_with_careers"] += 1
    strat_key = strategy.split("+")[0]
    _stats["strategy_hits"][strat_key] = _stats["strategy_hits"].get(strat_key, 0) + 1
    if ats_name:
        _stats["ats_hits"][ats_name] = _stats["ats_hits"].get(ats_name, 0) + 1

    vlog(f"✅  Career page   : {careers_url}", indent=1)
    vlog(f"🔧  Strategy      : {strategy}", indent=1)
    vlog(f"🤖  ATS           : {ats_name or 'Custom'}", indent=1)

    jobs = await extract_jobs(page, careers_url, name, website, industry, ats_name)

    all_jobs.extend(jobs)
    cp.setdefault("processed_domains", []).append(domain)
    cp["jobs_count"] = cp.get("jobs_count", 0) + len(jobs)
    save_checkpoint(cp)
    company_results.append({
        "name": name, "website": website, "industry": industry,
        "careers_url": careers_url, "ats": ats_name or "Custom",
        "jobs_found": len(jobs), "source": source_tag,
    })
    flush_jobs()
    print_company_summary(name, careers_url, strategy, ats_name, len(jobs))
    print_live_stats()


# ══════════════════════════════════════════════════════════════
# ▌▌▌  DISCOVERY SOURCES  ▌▌▌
# ══════════════════════════════════════════════════════════════

# ── 1. WIKIPEDIA (same as v5) ─────────────────────────────────
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
    print("\n📖  Discovery: Wikipedia")
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
            await process_fn(page, title, website, industry, cp, source_tag="wikipedia")
        await asyncio.sleep(random.uniform(1, 2))
    cp.setdefault("done_sources", []).append("wikipedia")
    save_checkpoint(cp)


# ── 2. WIKIDATA SPARQL ────────────────────────────────────────
# Returns structured company data: name, website, industry label
WIKIDATA_SPARQL_URL = "https://query.wikidata.org/sparql"

WIKIDATA_QUERY = """
SELECT DISTINCT ?company ?companyLabel ?website ?industryLabel WHERE {
  ?company wdt:P31 wd:Q4830453 .         # instance of business enterprise
  ?company wdt:P17 wd:Q851 .             # country Saudi Arabia
  OPTIONAL { ?company wdt:P856 ?website . }
  OPTIONAL { ?company wdt:P452 ?industry .
             ?industry rdfs:label ?industryLabel .
             FILTER(LANG(?industryLabel) = "en") }
  SERVICE wikibase:label { bd:serviceParam wikibase:language "en". }
}
LIMIT 2000
"""


async def wikidata_companies(page, cp, process_fn):
    if "wikidata" in cp.get("done_sources", []):
        print("⏭  Wikidata: already done")
        return
    print("\n🔬  Discovery: Wikidata SPARQL")
    try:
        r = requests.get(
            WIKIDATA_SPARQL_URL,
            params={"query": WIKIDATA_QUERY, "format": "json"},
            headers={"User-Agent": "SaudiJobsScraper/1.0 (educational research)"},
            timeout=60,
        )
        if not r or r.status_code != 200:
            print("  ⚠️  Wikidata SPARQL failed")
            return
        results = r.json().get("results", {}).get("bindings", [])
        print(f"  ✅  Wikidata returned {len(results)} entries")
        for row in tqdm(results, desc="  Wikidata entries"):
            name    = row.get("companyLabel", {}).get("value", "")
            website = row.get("website", {}).get("value", "")
            industry= row.get("industryLabel", {}).get("value", "General")
            if not name or not website:
                continue
            if re.match(r"^Q\d+$", name):   # skip unlabelled QIDs
                continue
            domain_check = get_domain(website)
            if domain_check in NON_SAUDI_DOMAINS:
                continue
            domain, norm_website = parse_domain(website)
            if not domain:
                continue
            await process_fn(page, name, norm_website, industry or "General", cp,
                             source_tag="wikidata")
            await asyncio.sleep(random.uniform(0.3, 0.8))
    except Exception as e:
        print(f"  ⚠️  Wikidata error: {e}")
    cp.setdefault("done_sources", []).append("wikidata")
    save_checkpoint(cp)


# ── 3. DUCKDUCKGO (same as v5) ────────────────────────────────
async def duckduckgo_companies(page, cp, process_fn):
    if "duckduckgo" in cp.get("done_sources", []):
        print("⏭  DuckDuckGo: already done")
        return
    print("\n🦆  Discovery: DuckDuckGo")
    queries = []
    for city in SAUDI_CITIES:
        for ind in INDUSTRIES:
            queries.append(
                (f"top {ind} companies in {city} Saudi Arabia site:*.com OR site:*.sa", ind)
            )
    random.shuffle(queries)
    for query, industry in tqdm(queries, desc="  DDG queries"):
        try:
            r = simple_get(f"https://html.duckduckgo.com/html/?q={quote_plus(query)}")
            if not r:
                continue
            soup = BeautifulSoup(r.text, "html.parser")
            for card in soup.select(".result"):
                name_el = card.select_one(".result__title")
                url_el  = card.select_one(".result__url")
                if not name_el or not url_el:
                    continue
                name    = name_el.get_text(strip=True)
                raw_url = url_el.get_text(strip=True)
                domain_check = get_domain(raw_url if raw_url.startswith("http") else "https://" + raw_url)
                if domain_check in NON_SAUDI_DOMAINS:
                    continue
                domain, website = parse_domain(raw_url)
                if not domain:
                    continue
                await process_fn(page, name, website, industry, cp, source_tag="duckduckgo")
            await asyncio.sleep(random.uniform(3, 7))
        except Exception as e:
            print(f"  ⚠️  DDG error: {e}")
            await asyncio.sleep(10)
    cp.setdefault("done_sources", []).append("duckduckgo")
    save_checkpoint(cp)


# ── 4. KOMPASS ────────────────────────────────────────────────
async def kompass_companies(page, cp, process_fn):
    if "kompass" in cp.get("done_sources", []):
        print("⏭  Kompass: already done")
        return
    print("\n📋  Discovery: Kompass")
    for pg in tqdm(range(1, 201), desc="  Kompass pages"):
        url  = f"https://sa.kompass.com/en/saudi-arabia/companies/?page={pg}"
        r    = simple_get(url)
        if not r:
            break
        soup  = BeautifulSoup(r.text, "html.parser")
        cards = soup.select(".company-list .company")
        if not cards:
            break
        for card in cards:
            name_el = card.select_one(".company-name")
            url_el  = card.select_one(".company-url a")
            if not name_el:
                continue
            name    = name_el.get_text(strip=True)
            raw_url = url_el.get("href", "") if url_el else ""
            domain, website = parse_domain(raw_url)
            if not domain:
                continue
            card_text = card.get_text()
            industry  = "General"
            for ind in INDUSTRIES:
                if ind.lower() in card_text.lower():
                    industry = ind
                    break
            await process_fn(page, name, website, industry, cp, source_tag="kompass")
        await asyncio.sleep(random.uniform(2, 4))
    cp.setdefault("done_sources", []).append("kompass")
    save_checkpoint(cp)


# ── 5. eARABICMARKET ─────────────────────────────────────────
# https://saudiarabia.earabicmarket.com/  — categorised Saudi company directory
EARABIC_BASE = "https://saudiarabia.earabicmarket.com"
EARABIC_CATEGORIES = [
    "/companies/", "/companies/?page=2",
    "/technology/", "/finance-banking/", "/healthcare/",
    "/construction-real-estate/", "/oil-gas/", "/manufacturing/",
    "/retail/", "/education/", "/hospitality-tourism/",
    "/logistics-transport/", "/telecommunications/",
    "/consulting/", "/engineering/",
]


async def earabicmarket_companies(page, cp, process_fn):
    if "earabicmarket" in cp.get("done_sources", []):
        print("⏭  eArabicMarket: already done")
        return
    print("\n🌐  Discovery: eArabicMarket")
    visited_cats = set()
    try:
        # First fetch the homepage to discover all category links
        r = simple_get(EARABIC_BASE)
        if r:
            soup = BeautifulSoup(r.text, "html.parser")
            for a in soup.find_all("a", href=True):
                href = a.get("href", "")
                if re.search(r"earabicmarket\.com/([\w-]+)/?$", href):
                    EARABIC_CATEGORIES.append(href if href.startswith("http") else EARABIC_BASE + href)
    except:
        pass

    for cat_path in tqdm(list(dict.fromkeys(EARABIC_CATEGORIES)), desc="  eArabic cats"):
        cat_url = cat_path if cat_path.startswith("http") else EARABIC_BASE + cat_path
        if cat_url in visited_cats:
            continue
        visited_cats.add(cat_url)
        industry_guess = "General"
        for ind in INDUSTRIES:
            if ind.lower().replace(" ", "-") in cat_url.lower():
                industry_guess = ind
                break

        page_num = 1
        while page_num <= 30:
            url = cat_url if page_num == 1 else f"{cat_url.rstrip('/')}/?page={page_num}"
            r   = simple_get(url)
            if not r:
                break
            soup = BeautifulSoup(r.text, "html.parser")
            # eArabicMarket company cards typically have company name + external link
            found_any = False
            for card in soup.select(".company-card, .listing-item, article, .business-item"):
                name_el = (card.select_one("h2") or card.select_one("h3")
                           or card.select_one(".company-name") or card.select_one(".title"))
                link_el = card.select_one("a[href*='http']")
                if not name_el:
                    continue
                name    = name_el.get_text(strip=True)
                raw_url = link_el.get("href", "") if link_el else ""
                # Skip links pointing back to earabicmarket itself
                if "earabicmarket.com" in raw_url:
                    # Try to find the company's own website in the card text
                    for a in card.find_all("a", href=True):
                        h = a.get("href", "")
                        if h.startswith("http") and "earabicmarket.com" not in h:
                            raw_url = h
                            break
                if not raw_url or "earabicmarket.com" in raw_url:
                    continue
                domain, website = parse_domain(raw_url)
                if not domain:
                    continue
                found_any = True
                await process_fn(page, name, website, industry_guess, cp,
                                 source_tag="earabicmarket")
            if not found_any:
                break
            # Check for next page link
            next_link = soup.select_one("a[rel='next'], .next-page a, .pagination .next")
            if not next_link:
                break
            page_num += 1
            await asyncio.sleep(random.uniform(1.5, 3))
        await asyncio.sleep(random.uniform(1, 2))

    cp.setdefault("done_sources", []).append("earabicmarket")
    save_checkpoint(cp)


# ── 6. TADAWUL / ARGAAM (Saudi Stock Exchange listed companies) ──
# Argaam has a public list of all Saudi-listed companies with their websites
ARGAAM_URL = "https://www.argaam.com/en/company/companies-in-saudi-market"


async def tadawul_companies(page, cp, process_fn):
    if "tadawul" in cp.get("done_sources", []):
        print("⏭  Tadawul/Argaam: already done")
        return
    print("\n📈  Discovery: Tadawul (Argaam listed companies)")
    try:
        await page.goto(ARGAAM_URL, timeout=30000, wait_until="domcontentloaded")
        await page.wait_for_timeout(3000)
        for _ in range(10):
            await page.evaluate("window.scrollBy(0, 800)")
            await page.wait_for_timeout(500)
        html = await page.content()
        soup = BeautifulSoup(html, "html.parser")

        companies_found = 0
        # Argaam renders a table of listed companies
        for row in soup.select("table tr, .company-row, .list-item"):
            cells = row.find_all(["td", "div"])
            name_el = row.select_one("a, .company-name, h3, h4")
            if not name_el:
                continue
            name = name_el.get_text(strip=True)
            if not name or len(name) < 2:
                continue
            # Try to find the company detail link on Argaam
            detail_link = name_el.get("href", "") if name_el.name == "a" else ""
            if not detail_link:
                a_tag = row.find("a", href=True)
                detail_link = a_tag.get("href", "") if a_tag else ""
            if not detail_link:
                continue
            full_detail = detail_link if detail_link.startswith("http") else "https://www.argaam.com" + detail_link
            if "argaam.com" not in full_detail:
                continue
            # Visit each company's Argaam profile to find its actual website
            await asyncio.sleep(random.uniform(0.8, 1.5))
            dr = simple_get(full_detail)
            if not dr:
                continue
            dsoup = BeautifulSoup(dr.text, "html.parser")
            web_el = (dsoup.select_one("a[href*='http']:not([href*='argaam'])") or
                      dsoup.select_one(".company-website a") or
                      dsoup.select_one(".website a"))
            if not web_el:
                # Scan all external links in the profile
                for a in dsoup.find_all("a", href=True):
                    h = a.get("href", "")
                    if h.startswith("http") and "argaam.com" not in h and "argaam" not in h:
                        if not any(s in get_domain(h) for s in SKIP_DOMAINS):
                            web_el = a
                            break
            if not web_el:
                continue
            raw_url = web_el.get("href", "")
            domain, website = parse_domain(raw_url)
            if not domain:
                continue
            # Industry from sector label in Argaam profile
            industry = "General"
            sector_el = dsoup.select_one(".sector, .industry, [class*='sector']")
            if sector_el:
                sec_text = sector_el.get_text(strip=True).lower()
                for ind in INDUSTRIES:
                    if ind.lower() in sec_text:
                        industry = ind
                        break
            await process_fn(page, name, website, industry, cp, source_tag="tadawul")
            companies_found += 1

        print(f"  ✅  Tadawul: {companies_found} companies found")
    except Exception as e:
        print(f"  ⚠️  Tadawul/Argaam error: {e}")

    cp.setdefault("done_sources", []).append("tadawul")
    save_checkpoint(cp)


# ── 7. GULFTALENT (employer profiles) ───────────────────────
GULFTALENT_EMPLOYERS = "https://www.gulftalent.com/saudi-arabia/employers"


async def gulftalent_companies(page, cp, process_fn):
    if "gulftalent" in cp.get("done_sources", []):
        print("⏭  GulfTalent: already done")
        return
    print("\n🌴  Discovery: GulfTalent Employers")
    try:
        for pg_num in tqdm(range(1, 51), desc="  GulfTalent pages"):
            url = GULFTALENT_EMPLOYERS if pg_num == 1 else f"{GULFTALENT_EMPLOYERS}?page={pg_num}"
            r   = simple_get(url)
            if not r:
                break
            soup = BeautifulSoup(r.text, "html.parser")
            cards = soup.select(".employer-card, .company-card, .employer-item, article")
            if not cards:
                break
            found_any = False
            for card in cards:
                name_el = (card.select_one("h2") or card.select_one("h3")
                           or card.select_one(".employer-name"))
                if not name_el:
                    continue
                name = name_el.get_text(strip=True)
                # GulfTalent employer pages have a link to the company's own site
                # Try the card's external link first
                raw_url = ""
                for a in card.find_all("a", href=True):
                    h = a.get("href", "")
                    if h.startswith("http") and "gulftalent.com" not in h:
                        raw_url = h
                        break
                if not raw_url:
                    # Fall back to fetching the GulfTalent employer page
                    detail_a = card.find("a", href=re.compile(r"/employers/"))
                    if detail_a:
                        detail_url = "https://www.gulftalent.com" + detail_a.get("href", "")
                        dr = simple_get(detail_url)
                        if dr:
                            ds = BeautifulSoup(dr.text, "html.parser")
                            for a in ds.find_all("a", href=True):
                                h = a.get("href", "")
                                if h.startswith("http") and "gulftalent.com" not in h:
                                    raw_url = h
                                    break
                if not raw_url:
                    continue
                domain, website = parse_domain(raw_url)
                if not domain:
                    continue
                industry = "General"
                card_text = card.get_text().lower()
                for ind in INDUSTRIES:
                    if ind.lower() in card_text:
                        industry = ind
                        break
                found_any = True
                await process_fn(page, name, website, industry, cp, source_tag="gulftalent")
                await asyncio.sleep(random.uniform(0.3, 0.7))
            if not found_any:
                break
            await asyncio.sleep(random.uniform(2, 4))
    except Exception as e:
        print(f"  ⚠️  GulfTalent error: {e}")
    cp.setdefault("done_sources", []).append("gulftalent")
    save_checkpoint(cp)


# ── 8. BAYT.COM (Saudi employer profiles) ────────────────────
BAYT_EMPLOYERS = "https://www.bayt.com/en/saudi-arabia/jobs/companies/"


async def bayt_companies(page, cp, process_fn):
    if "bayt" in cp.get("done_sources", []):
        print("⏭  Bayt: already done")
        return
    print("\n💼  Discovery: Bayt.com Companies")
    try:
        for pg_num in tqdm(range(1, 101), desc="  Bayt pages"):
            url = BAYT_EMPLOYERS if pg_num == 1 else f"{BAYT_EMPLOYERS}?page={pg_num}"
            r   = simple_get(url)
            if not r:
                break
            soup = BeautifulSoup(r.text, "html.parser")
            # Bayt company listing
            cards = soup.select("[class*='company'], [class*='employer'], .list-item, li.t-boxf")
            if not cards:
                # Try finding by header tags in main content
                cards = soup.select("main li, .content li, #results li")
            if not cards:
                break

            found_any = False
            for card in cards:
                name_el = (card.select_one("h2 a") or card.select_one("h3 a")
                           or card.select_one(".company-name a") or card.select_one("a"))
                if not name_el:
                    continue
                name     = name_el.get_text(strip=True)
                prof_url = name_el.get("href", "")
                if not prof_url or not name:
                    continue
                full_prof = prof_url if prof_url.startswith("http") else "https://www.bayt.com" + prof_url
                if "bayt.com" not in full_prof:
                    continue
                # Visit employer profile to extract website
                await asyncio.sleep(random.uniform(0.5, 1.2))
                dr = simple_get(full_prof)
                if not dr:
                    continue
                ds = BeautifulSoup(dr.text, "html.parser")
                raw_url = ""
                for a in ds.find_all("a", href=True):
                    h = a.get("href", "")
                    if h.startswith("http") and "bayt.com" not in h and len(h) > 10:
                        if not any(s in get_domain(h) for s in SKIP_DOMAINS):
                            raw_url = h
                            break
                if not raw_url:
                    continue
                domain, website = parse_domain(raw_url)
                if not domain:
                    continue
                industry = "General"
                prof_text = ds.get_text().lower()
                for ind in INDUSTRIES:
                    if ind.lower() in prof_text:
                        industry = ind
                        break
                found_any = True
                await process_fn(page, name, website, industry, cp, source_tag="bayt")
            if not found_any:
                break
            await asyncio.sleep(random.uniform(2, 5))
    except Exception as e:
        print(f"  ⚠️  Bayt error: {e}")
    cp.setdefault("done_sources", []).append("bayt")
    save_checkpoint(cp)


# ── 9. NAUKRIGULF (Saudi employer profiles) ──────────────────
NAUKRIGULF_BASE = "https://www.naukrigulf.com"
NAUKRIGULF_COMPANIES = [
    "/companies-in-saudi-arabia",
    "/companies-in-riyadh",
    "/companies-in-jeddah",
    "/companies-in-dammam",
    "/companies-in-khobar",
]


async def naukrigulf_companies(page, cp, process_fn):
    if "naukrigulf" in cp.get("done_sources", []):
        print("⏭  Naukrigulf: already done")
        return
    print("\n🌿  Discovery: Naukrigulf Companies")
    try:
        for path in tqdm(NAUKRIGULF_COMPANIES, desc="  Naukrigulf paths"):
            for pg_num in range(1, 51):
                url = (NAUKRIGULF_BASE + path
                       if pg_num == 1
                       else f"{NAUKRIGULF_BASE + path}-{pg_num}")
                r   = simple_get(url)
                if not r:
                    break
                soup = BeautifulSoup(r.text, "html.parser")
                cards = soup.select(".company-list-item, .comp-info, [class*='company'], article")
                if not cards:
                    break
                found_any = False
                for card in cards:
                    name_el = (card.select_one("h2 a") or card.select_one("h3 a")
                               or card.select_one(".comp-name a") or card.select_one("a"))
                    if not name_el:
                        continue
                    name     = name_el.get_text(strip=True)
                    prof_href = name_el.get("href", "")
                    if not name or not prof_href:
                        continue
                    full_prof = prof_href if prof_href.startswith("http") else NAUKRIGULF_BASE + prof_href
                    if "naukrigulf.com" not in full_prof:
                        continue
                    await asyncio.sleep(random.uniform(0.4, 0.9))
                    dr = simple_get(full_prof)
                    if not dr:
                        continue
                    ds = BeautifulSoup(dr.text, "html.parser")
                    raw_url = ""
                    web_label = ds.find(string=re.compile(r"website", re.I))
                    if web_label:
                        parent = web_label.find_parent()
                        if parent:
                            a_tag = parent.find_next("a", href=True)
                            if a_tag:
                                raw_url = a_tag.get("href", "")
                    if not raw_url:
                        for a in ds.find_all("a", href=True):
                            h = a.get("href", "")
                            if h.startswith("http") and "naukrigulf.com" not in h:
                                if not any(s in get_domain(h) for s in SKIP_DOMAINS):
                                    raw_url = h
                                    break
                    if not raw_url:
                        continue
                    domain, website = parse_domain(raw_url)
                    if not domain:
                        continue
                    industry = "General"
                    prof_text = ds.get_text().lower()
                    for ind in INDUSTRIES:
                        if ind.lower() in prof_text:
                            industry = ind
                            break
                    found_any = True
                    await process_fn(page, name, website, industry, cp, source_tag="naukrigulf")
                if not found_any:
                    break
                await asyncio.sleep(random.uniform(2, 4))
    except Exception as e:
        print(f"  ⚠️  Naukrigulf error: {e}")
    cp.setdefault("done_sources", []).append("naukrigulf")
    save_checkpoint(cp)


# ── 10. GLASSDOOR (Saudi company discovery only) ─────────────
# We scrape the Glassdoor company list page for names + websites only.
# We DO NOT use Glassdoor as a job source.
GLASSDOOR_URL = "https://www.glassdoor.com/Explore/browse-companies.htm?overall_rating_low=0&page=1&locId=115&locType=N&locName=Saudi+Arabia&filterType=RATING_OVERALL"


async def glassdoor_companies(page, cp, process_fn):
    if "glassdoor" in cp.get("done_sources", []):
        print("⏭  Glassdoor: already done")
        return
    print("\n🚪  Discovery: Glassdoor (company names only)")
    try:
        for pg_num in tqdm(range(1, 51), desc="  Glassdoor pages"):
            gd_url = GLASSDOOR_URL.replace("page=1", f"page={pg_num}")
            try:
                await page.goto(gd_url, timeout=30000, wait_until="domcontentloaded")
                await page.wait_for_timeout(2500)
                html = await page.content()
            except:
                break
            soup = BeautifulSoup(html, "html.parser")
            cards = soup.select("[data-test='employer-card'], .employer-card, [class*='EmployerCard']")
            if not cards:
                # Try JSON-LD or __NEXT_DATA__
                script = soup.find("script", id="__NEXT_DATA__")
                if script:
                    try:
                        nd = json.loads(script.string)
                        employers = (nd.get("props",{}).get("pageProps",{})
                                       .get("employerResults",[]))
                        for emp in employers:
                            name    = emp.get("name","") or emp.get("shortName","")
                            website = emp.get("website","") or emp.get("websiteURL","")
                            if not name or not website:
                                continue
                            domain, norm = parse_domain(website)
                            if domain:
                                await process_fn(page, name, norm, "General", cp,
                                                 source_tag="glassdoor")
                    except:
                        pass
                break
            found_any = False
            for card in cards:
                name_el = (card.select_one("[data-test='employer-short-name']")
                           or card.select_one("h2") or card.select_one("h3"))
                if not name_el:
                    continue
                name = name_el.get_text(strip=True)
                # Glassdoor doesn't expose company websites directly in list view
                # We use DDG to find the website from the company name
                if not name or len(name) < 2:
                    continue
                q = quote_plus(f"{name} Saudi Arabia official website")
                dr = simple_get(f"https://html.duckduckgo.com/html/?q={q}")
                if not dr:
                    continue
                ds = BeautifulSoup(dr.text, "html.parser")
                for el in ds.select(".result__url"):
                    u = el.get_text(strip=True)
                    if not u.startswith("http"):
                        u = "https://" + u
                    d_check = get_domain(u)
                    if d_check not in NON_SAUDI_DOMAINS and d_check not in SKIP_DOMAINS:
                        domain, website = parse_domain(u)
                        if domain:
                            await process_fn(page, name, website, "General", cp,
                                             source_tag="glassdoor")
                            found_any = True
                            break
                await asyncio.sleep(random.uniform(1, 2))
            if not found_any:
                break
            await asyncio.sleep(random.uniform(2, 4))
    except Exception as e:
        print(f"  ⚠️  Glassdoor error: {e}")
    cp.setdefault("done_sources", []).append("glassdoor")
    save_checkpoint(cp)


# ── 11. INDEED SAUDI ARABIA (employer discovery) ─────────────
INDEED_SA_COMPANIES = "https://sa.indeed.com/companies"


async def indeed_companies(page, cp, process_fn):
    if "indeed" in cp.get("done_sources", []):
        print("⏭  Indeed: already done")
        return
    print("\n🔍  Discovery: Indeed Saudi Arabia Companies")
    try:
        for industry in tqdm(INDUSTRIES, desc="  Indeed industries"):
            url = f"https://sa.indeed.com/companies?q={quote_plus(industry)}&l=Saudi+Arabia"
            try:
                await page.goto(url, timeout=30000, wait_until="domcontentloaded")
                await page.wait_for_timeout(2000)
                html = await page.content()
            except:
                continue
            soup = BeautifulSoup(html, "html.parser")
            for card in soup.select("[data-company-name], .company-card, [class*='Company']"):
                name_el = (card.select_one("[data-company-name]")
                           or card.select_one("h2") or card.select_one("h3"))
                if not name_el:
                    continue
                name = (name_el.get("data-company-name") or
                        name_el.get_text(strip=True))
                if not name:
                    continue
                # Find website via DDG
                q = quote_plus(f"{name} Saudi Arabia official site")
                dr = simple_get(f"https://html.duckduckgo.com/html/?q={q}")
                if not dr:
                    continue
                ds = BeautifulSoup(dr.text, "html.parser")
                for el in ds.select(".result__url"):
                    u = el.get_text(strip=True)
                    if not u.startswith("http"):
                        u = "https://" + u
                    d_check = get_domain(u)
                    if d_check not in NON_SAUDI_DOMAINS and d_check not in SKIP_DOMAINS:
                        domain, website = parse_domain(u)
                        if domain:
                            await process_fn(page, name, website, industry, cp,
                                             source_tag="indeed")
                            break
                await asyncio.sleep(random.uniform(1, 2))
            await asyncio.sleep(random.uniform(3, 6))
    except Exception as e:
        print(f"  ⚠️  Indeed error: {e}")
    cp.setdefault("done_sources", []).append("indeed")
    save_checkpoint(cp)


# ── 12. MODON (Saudi Industrial Development Authority) ───────
# modon.gov.sa lists Saudi industrial companies with websites
MODON_URL = "https://www.modon.gov.sa/en/IndustrialCities/Pages/InvestorGuide.aspx"


async def modon_companies(page, cp, process_fn):
    if "modon" in cp.get("done_sources", []):
        print("⏭  MODON: already done")
        return
    print("\n🏭  Discovery: MODON Industrial Companies")
    try:
        r = simple_get(MODON_URL)
        if not r:
            # Try Playwright
            await page.goto(MODON_URL, timeout=30000, wait_until="domcontentloaded")
            await page.wait_for_timeout(2000)
            html = await page.content()
        else:
            html = r.text
        soup = BeautifulSoup(html, "html.parser")
        for a in soup.find_all("a", href=True):
            href = a.get("href", "")
            if not href.startswith("http") or "modon.gov.sa" in href:
                continue
            name_guess = a.get_text(strip=True) or get_domain(href)
            domain, website = parse_domain(href)
            if domain:
                await process_fn(page, name_guess, website, "manufacturing", cp,
                                 source_tag="modon")
    except Exception as e:
        print(f"  ⚠️  MODON error: {e}")
    cp.setdefault("done_sources", []).append("modon")
    save_checkpoint(cp)


# ── 13. SAUDI EXPORTERS (saudiexporters.sa) ──────────────────
SAUDI_EXPORTERS_URL = "https://www.saudiexporters.sa/en/exporters"


async def saudiexporters_companies(page, cp, process_fn):
    if "saudiexporters" in cp.get("done_sources", []):
        print("⏭  SaudiExporters: already done")
        return
    print("\n📦  Discovery: Saudi Exporters Directory")
    try:
        for pg_num in tqdm(range(1, 101), desc="  Exporters pages"):
            url = SAUDI_EXPORTERS_URL if pg_num == 1 else f"{SAUDI_EXPORTERS_URL}?page={pg_num}"
            r   = simple_get(url)
            if not r:
                # Try Playwright
                try:
                    await page.goto(url, timeout=25000, wait_until="domcontentloaded")
                    await page.wait_for_timeout(1500)
                    html = await page.content()
                except:
                    break
            else:
                html = r.text
            soup = BeautifulSoup(html, "html.parser")
            cards = soup.select(".exporter-card, .company-item, article, .listing-item, tr")
            if not cards:
                break
            found_any = False
            for card in cards:
                name_el = (card.select_one("h2") or card.select_one("h3")
                           or card.select_one(".company-name") or card.select_one("td"))
                link_el = card.find("a", href=re.compile(r"^https?://"))
                if not name_el or not link_el:
                    continue
                name    = name_el.get_text(strip=True)
                raw_url = link_el.get("href", "")
                if "saudiexporters.sa" in raw_url:
                    # Profile page — fetch it for actual website
                    full = raw_url if raw_url.startswith("http") else "https://www.saudiexporters.sa" + raw_url
                    dr   = simple_get(full)
                    if dr:
                        ds = BeautifulSoup(dr.text, "html.parser")
                        for a in ds.find_all("a", href=True):
                            h = a.get("href", "")
                            if h.startswith("http") and "saudiexporters.sa" not in h:
                                raw_url = h
                                break
                domain, website = parse_domain(raw_url)
                if not domain:
                    continue
                industry = "manufacturing"
                card_text = card.get_text().lower()
                for ind in INDUSTRIES:
                    if ind.lower() in card_text:
                        industry = ind
                        break
                found_any = True
                await process_fn(page, name, website, industry, cp, source_tag="saudiexporters")
            if not found_any:
                break
            await asyncio.sleep(random.uniform(1.5, 3))
    except Exception as e:
        print(f"  ⚠️  SaudiExporters error: {e}")
    cp.setdefault("done_sources", []).append("saudiexporters")
    save_checkpoint(cp)


# ── 14. WAMDA / MAGNITT (Saudi startups) ─────────────────────
WAMDA_URL  = "https://wamda.com/companies?country=saudi-arabia"
MAGNITT_URL = "https://magnitt.com/companies?country=saudi-arabia"


async def startup_directories(page, cp, process_fn):
    if "startups" in cp.get("done_sources", []):
        print("⏭  Startup directories: already done")
        return
    print("\n🚀  Discovery: Startup Directories (Wamda + MagniTT)")

    for src_name, src_url in [("wamda", WAMDA_URL), ("magnitt", MAGNITT_URL)]:
        try:
            await page.goto(src_url, timeout=30000, wait_until="domcontentloaded")
            await page.wait_for_timeout(3000)
            for _ in range(8):
                await page.evaluate("window.scrollBy(0, 800)")
                await page.wait_for_timeout(600)
            html = await page.content()
            soup = BeautifulSoup(html, "html.parser")
            count = 0
            for card in soup.select(".company-card, .startup-card, article, [class*='Card']"):
                name_el = (card.select_one("h2") or card.select_one("h3")
                           or card.select_one(".company-name"))
                if not name_el:
                    continue
                name    = name_el.get_text(strip=True)
                raw_url = ""
                for a in card.find_all("a", href=True):
                    h = a.get("href", "")
                    if h.startswith("http") and src_name not in h:
                        raw_url = h
                        break
                if not raw_url:
                    continue
                domain, website = parse_domain(raw_url)
                if not domain:
                    continue
                await process_fn(page, name, website, "technology", cp,
                                 source_tag=f"startup_{src_name}")
                count += 1
            print(f"    ✅  {src_name}: {count} companies")
        except Exception as e:
            print(f"    ⚠️  {src_name} error: {e}")
        await asyncio.sleep(random.uniform(2, 4))

    cp.setdefault("done_sources", []).append("startups")
    save_checkpoint(cp)


# ── 15. GOOGLE MAPS (same as v5) ──────────────────────────────
async def google_maps_companies(page, cp, process_fn):
    if "google_maps" in cp.get("done_sources", []):
        print("⏭  Google Maps: already done")
        return
    print("\n🗺  Discovery: Google Maps")
    for city in tqdm(SAUDI_CITIES, desc="  Cities"):
        for industry in INDUSTRIES:
            query    = f"{industry} companies {city} Saudi Arabia"
            maps_url = f"https://www.google.com/maps/search/{quote_plus(query)}"
            try:
                await page.goto(maps_url, timeout=30000, wait_until="domcontentloaded")
                await page.wait_for_timeout(3000)
                for _ in range(5):
                    await page.keyboard.press("End")
                    await page.wait_for_timeout(1500)
                cards = await page.query_selector_all("a[href*='/maps/place/']")
                for card in cards:
                    try:
                        name = await card.get_attribute("aria-label") or ""
                        if not name:
                            continue
                        await card.click()
                        await page.wait_for_timeout(2000)
                        web_el  = await page.query_selector("a[data-item-id='authority']")
                        raw_url = await web_el.get_attribute("href") if web_el else ""
                        d_check = get_domain(raw_url)
                        if d_check in NON_SAUDI_DOMAINS:
                            await page.go_back()
                            await page.wait_for_timeout(1000)
                            continue
                        domain, website = parse_domain(raw_url)
                        if domain:
                            await process_fn(page, name, website, industry, cp,
                                             source_tag="google_maps")
                        await page.go_back()
                        await page.wait_for_timeout(1000)
                    except:
                        continue
                await page.wait_for_timeout(random.uniform(3000, 6000))
            except Exception as e:
                print(f"    ⚠️  Maps error ({city}/{industry}): {e}")
                await page.wait_for_timeout(5000)
    cp.setdefault("done_sources", []).append("google_maps")
    save_checkpoint(cp)


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════
async def main():
    global all_jobs, company_results

    print("🚀  Saudi Jobs Pipeline v6 — Maximum Discovery · No Duplicates\n")
    print("   Sources: Wikipedia · Wikidata · DDG · Kompass · eArabicMarket")
    print("            Tadawul · GulfTalent · Bayt · Naukrigulf · Glassdoor")
    print("            Indeed · MODON · SaudiExporters · Startups · Google Maps\n")

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

        # ── Discovery order: structured → broad → job-boards ──
        # Run all — global domain set prevents any duplicate processing

        # 1. Structured data (cleanest, most accurate)
        await wikidata_companies(page, cp, process_company)
        await wikipedia_companies(page, cp, process_company)
        await tadawul_companies(page, cp, process_company)

        # 2. Company directories
        await earabicmarket_companies(page, cp, process_company)
        await kompass_companies(page, cp, process_company)
        await modon_companies(page, cp, process_company)
        await saudiexporters_companies(page, cp, process_company)

        # 3. Job boards (employer discovery)
        await gulftalent_companies(page, cp, process_company)
        await bayt_companies(page, cp, process_company)
        await naukrigulf_companies(page, cp, process_company)
        await indeed_companies(page, cp, process_company)
        await glassdoor_companies(page, cp, process_company)

        # 4. Startup directories
        await startup_directories(page, cp, process_company)

        # 5. Search engine discovery
        await duckduckgo_companies(page, cp, process_company)

        # 6. Google Maps (slowest, run last)
        await google_maps_companies(page, cp, process_company)

        await browser.close()

    flush_jobs()

    total_jobs      = len(all_jobs)
    total_companies = len(company_results)
    with_careers    = sum(1 for c in company_results if c.get("careers_url"))

    W = 72
    print(f"\n{'═'*W}")
    print(f"  🏁  PIPELINE COMPLETE — v6")
    print(f"{'─'*W}")
    print(f"  {'Companies discovered:':<30} {total_companies:>6,}")
    print(f"  {'  with career pages:':<30} {with_careers:>6,}  ({100*with_careers//max(total_companies,1)}%)")
    print(f"  {'  no career page:':<30} {total_companies-with_careers:>6,}")
    print(f"  {'  skipped (non-Saudi):':<30} {_stats['companies_skipped_non_saudi']:>6,}")
    print(f"  {'  unique domains only:':<30} {len(discovered_domains):>6,}")
    print(f"{'─'*W}")
    print(f"  {'Total jobs scraped:':<30} {total_jobs:>6,}")
    print(f"  {'  with salary data:':<30} {_stats['jobs_with_salary']:>6,}  ({100*_stats['jobs_with_salary']//max(total_jobs,1)}%)")
    print(f"  {'  with description:':<30} {_stats['jobs_with_description']:>6,}  ({100*_stats['jobs_with_description']//max(total_jobs,1)}%)")
    print(f"  {'Detail fetches:':<30} {_stats['detail_fetches']:>6,}")
    print(f"  {'Detail failures:':<30} {_stats['detail_failures']:>6,}  ({100*_stats['detail_failures']//max(_stats['detail_fetches'],1)}%)")
    print(f"{'─'*W}")

    if _stats["source_companies"]:
        print(f"  Companies by discovery source:")
        for src, cnt in sorted(_stats["source_companies"].items(), key=lambda x: -x[1]):
            bar = _bar(cnt, total_companies, width=25)
            print(f"    {src:<22} {cnt:>4}  {bar}")
        print(f"{'─'*W}")

    if _stats["ats_hits"]:
        print(f"  ATS breakdown:")
        for ats, count in sorted(_stats["ats_hits"].items(), key=lambda x: -x[1]):
            bar = _bar(count, total_companies, width=30)
            print(f"    {ats:<22} {count:>4}  {bar}")
        print(f"{'─'*W}")

    if all_jobs:
        jobs_df = pd.DataFrame(all_jobs).drop_duplicates(subset="Job URL")
        jobs_df.to_csv(JOBS_FILE, index=False)
        print(f"\n  Top 15 companies by jobs scraped:")
        top = (jobs_df.groupby("Company Name")["Job Title"]
               .count().sort_values(ascending=False).head(15))
        for co, cnt in top.items():
            bar = _bar(cnt, top.iloc[0], width=25)
            print(f"    {co[:30]:<30}  {cnt:>4}  {bar}")

        if "Job Field" in jobs_df.columns:
            print(f"\n  Top job fields:")
            fields = (jobs_df["Job Field"].dropna()
                      .replace("", pd.NA).dropna()
                      .value_counts().head(10))
            for field, cnt in fields.items():
                print(f"    {field[:30]:<30}  {cnt:>4}")

        if "Job Location" in jobs_df.columns:
            print(f"\n  Top locations:")
            locs = (jobs_df["Job Location"].dropna()
                    .replace("", pd.NA).dropna()
                    .value_counts().head(10))
            for loc, cnt in locs.items():
                print(f"    {loc[:30]:<30}  {cnt:>4}")

    print(f"\n{'═'*W}")
    print(f"  💾  Files: {JOBS_FILE}  |  {COMPANIES_FILE}")
    print(f"{'═'*W}")

    try:
        from google.colab import files
        print(f"\n   Downloading files…")
        files.download(JOBS_FILE)
        files.download(COMPANIES_FILE)
    except:
        print(f"\n   Files saved locally: {JOBS_FILE}  |  {COMPANIES_FILE}")


asyncio.run(main())
