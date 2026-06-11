"""
LinkedIn Job Scraper — Python v4 + Mistral Paraphrase
======================================================
Fuses v4 scraper (deep company-website crawling) with the Mistral-powered
paraphrase pipeline (title, description, company details, tagline).

What's new vs plain v4:
    ─ Mistral API integration for paraphrasing
        • paraphrase_title()        — rewrites job title
        • paraphrase_description()  — rewrites description paragraph-by-paragraph
        • paraphrase_company()      — rewrites company about/details text
        • paraphrase_tagline()      — rewrites short company tagline
    ─ All v4 scraping layers preserved (JSON-LD, apply-button, deep crawl,
      about/contact/footer, v3 fallback)
    ─ Jobs are paraphrased BEFORE being written to Excel / WordPress
    ─ Duplicate tracker (CSV) records status per job ID
    ─ WordPress posting uses paraphrased title + description

Requirements:
    pip install requests beautifulsoup4 openpyxl sentence-transformers language-tool-python

Usage:
    python linkedin_job_scraper_v4_paraphrase.py

Outputs:
    jobs_output.xlsx          — scraped + paraphrased jobs
    processed_jobs.csv        — duplicate / status tracker
"""

import re
import time
import base64
import json
import hashlib
import logging
import os
from datetime import datetime, timedelta
from urllib.parse import urljoin, urlparse, unquote, quote_plus

import requests
from bs4 import BeautifulSoup
import openpyxl
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
#  OPTIONAL heavy deps — imported lazily so the scraper still runs without them
# ─────────────────────────────────────────────────────────────────────────────
try:
    import language_tool_python
    from sentence_transformers import SentenceTransformer, util
    _NLP_AVAILABLE = True
except ImportError:
    _NLP_AVAILABLE = False

# =============================================================================
#  CONFIG
# =============================================================================

SHEET_NAME  = "Sheet1"
DELAY_S     = 2.0
FETCH_CHAR_LIMIT = 120_000

MAX_PAGES       = 0   # 0 = unlimited
MAX_EMPTY_PAGES = 5
JOB_LIMIT       = 0   # 0 = no cap

OUTPUT_FILE         = "jobs_output.xlsx"
PROCESSED_IDS_FILE  = "processed_jobs.csv"

# ── WordPress ────────────────────────────────────────────────────────────────
WP_URL      = "https://mauritius.mimusjobs.com/wp-json/wp/v2/"
WP_USER     = "calolina"
WP_PASSWORD = "st8a 6mWY wqgV 0syR mB3i y5FQ"
WP_BASE     = WP_URL.rstrip("/")
WP_JOBS_URL    = f"{WP_BASE}/job_listing"
WP_COMPANY_URL = f"{WP_BASE}/company"
WP_MEDIA_URL   = f"{WP_BASE}/media"

# ── Mistral ───────────────────────────────────────────────────────────────────
MISTRAL_API_KEY = "your_mistral_api_key_here"   # ← replace
MISTRAL_MODEL   = "mistral-small-latest"
MISTRAL_URL     = "https://api.mistral.ai/v1/chat/completions"

# Enable / disable paraphrasing globally
ENABLE_PARAPHRASE = True   # set False to skip paraphrasing entirely

# =============================================================================
#  KEYWORDS
# =============================================================================

SEARCH_KEYWORDS = [
    "",
  #  "engineer", "developer", "manager", "finance", "sales", "HR",
  #  "doctor", "construction", "logistics", "operations", "customer service",
  #  "teacher", "chef", "lawyer", "graphic designer", "production manager",
  #  "petroleum", "driver", "security", "researcher", "journalist",
  #  "banker", "retail", "renewable energy",
]

# =============================================================================
#  LOGGING / COLOUR
# =============================================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

import sys
_USE_COLOUR = sys.stdout.isatty()

def _c(code, text):
    return f"\033[{code}m{text}\033[0m" if _USE_COLOUR else text

C_HEADER  = lambda t: _c("1;36",  t)
C_LABEL   = lambda t: _c("1;33",  t)
C_VALUE   = lambda t: _c("97",    t)
C_DIM     = lambda t: _c("2",     t)
C_GREEN   = lambda t: _c("1;32",  t)
C_RED     = lambda t: _c("1;31",  t)
C_BLUE    = lambda t: _c("1;34",  t)
C_DIVIDER = lambda: _c("2", "─" * 72)

# =============================================================================
#  ROTATING USER-AGENTS
# =============================================================================

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

# =============================================================================
#  DOMAIN LISTS
# =============================================================================

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

JOB_TYPE_MAPPING = {
    "full-time": "full-time", "full time": "full-time",
    "part-time": "part-time", "part time": "part-time",
    "contract":  "contract",  "temporary": "temporary",
    "internship":"internship","freelance": "freelance",
    "volunteer": "volunteer",
}

# =============================================================================
#  MOJIBAKE / TEXT HELPERS
# =============================================================================

_MOJIBAKE = [
    ("Â", ""), ("â€™", "'"), ("â€œ", '"'), ("â€\x9d", '"'), ("â€", '"'),
    ("â€¢", "•"), ("â„¢", "™"), ("\u00a0", " "), ("\u200b", ""), ("\ufeff", ""),
]

def _fix_mojibake(text: str) -> str:
    for pattern, replacement in _MOJIBAKE:
        text = text.replace(pattern, replacement)
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)
    return text

def sanitize_text(text, is_url=False, is_email=False) -> str:
    if not isinstance(text, str):
        text = str(text) if (text is not None and str(text) not in ("nan","None","NaN")) else ""
    text = text.strip()
    if text in ("nan", "None", "NaN", "", "N/A", "n/a", "NA", "na"):
        return ""
    text = _fix_mojibake(text)
    if is_url or is_email:
        return re.sub(r"[ \t\r\n\f\v]+", " ", text).strip()
    text = re.sub(r"#+\s*", "", text)
    text = re.sub(r"\*\*", "", text)
    text = re.sub(r"[^\x20-\x7E\n\u00C0-\u017F\u2013\u2014\u2018-\u201D\u2022]", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

# =============================================================================
#  NLP TOOLS (lazy init)
# =============================================================================

_grammar_tool   = None
_sim_model      = None

def _get_grammar_tool():
    global _grammar_tool
    if _grammar_tool is None and _NLP_AVAILABLE:
        try:
            _grammar_tool = language_tool_python.LanguageTool(
                "en-US", remote_server="https://api.languagetool.org")
        except Exception as e:
            log.warning(f"LanguageTool init failed: {e}")
    return _grammar_tool

def _get_sim_model():
    global _sim_model
    if _sim_model is None and _NLP_AVAILABLE:
        try:
            _sim_model = SentenceTransformer("all-MiniLM-L6-v2", device="cpu")
        except Exception as e:
            log.warning(f"SentenceTransformer init failed: {e}")
    return _sim_model

def grammar_correct(text: str) -> str:
    tool = _get_grammar_tool()
    if tool:
        try:
            return language_tool_python.utils.correct(text, tool.check(text))
        except Exception:
            pass
    return text

def similarity_score(a: str, b: str) -> float:
    model = _get_sim_model()
    if model:
        try:
            from sentence_transformers import util as st_util
            emb = model.encode([a, b], convert_to_tensor=True)
            return float(st_util.pytorch_cos_sim(emb[0], emb[1]))
        except Exception:
            pass
    # Fallback: simple token overlap
    def tokens(s):
        return set(re.sub(r"[^a-z0-9 ]", " ", s.lower()).split())
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb: return 0.0
    return len(ta & tb) / max(len(ta), len(tb))

def clean_output(text: str) -> str:
    text = _fix_mojibake(text)
    for pat in [r"\[/?INST\]", r"</?s>",
                r"(?i)(rewritten?|rephrased?|output|paraphrase[d]?)[:\s]+",
                r"\*\*", r"###", r"---"]:
        text = re.sub(pat, "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return grammar_correct(text.strip())

# =============================================================================
#  MISTRAL API
# =============================================================================

def mistral_generate(prompt: str, max_tokens: int = 400, temperature: float = 0.7) -> str:
    if not MISTRAL_API_KEY or MISTRAL_API_KEY == "your_mistral_api_key_here":
        log.warning("Mistral API key not set — skipping paraphrase")
        return ""
    try:
        response = requests.post(
            MISTRAL_URL,
            headers={
                "Authorization": f"Bearer {MISTRAL_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": MISTRAL_MODEL,
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": max_tokens,
                "temperature": temperature,
            },
            timeout=30,
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        log.error(f"Mistral API error: {e}")
        return ""

# =============================================================================
#  PARAPHRASE FUNCTIONS
# =============================================================================

def paraphrase_title(title: str) -> str:
    if not ENABLE_PARAPHRASE:
        return title
    clean = sanitize_text(title)
    if not clean:
        return title

    print(f"\n ┌─ TITLE PARAPHRASE {'─'*45}")
    print(f" │ Original : \"{clean}\"")
    print(f" │ {'─'*60}")

    best_result = None
    best_sim = 0.0

    for attempt in range(4):
        temp = round(0.68 + attempt * 0.06, 2)
        print(f" │ Attempt {attempt+1} (temp={temp}):")

        prompt = (
            f"Rewrite this job title professionally using different words. "
            f"Output ONLY the rewritten title, nothing else. "
            f"Keep it between 4 and 12 words.\n\nJob title: {clean}"
        )

        raw = mistral_generate(prompt, max_tokens=50, temperature=temp)
        result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")

        wc = len(result.split()) if result else 0
        sim = similarity_score(clean, result) if result else 0.0
        is_dup = result.lower().strip() == clean.lower().strip()

        print(f" │    Output  : \"{result}\"")
        print(f" │    Words   : {wc} | Similarity: {sim:.3f} | Duplicate: {'Yes ⚠️' if is_dup else 'No'}")

        valid = bool(result) and 4 <= wc <= 14 and sim >= 0.55 and not is_dup

        if not valid:
            reasons = []
            if not result:           reasons.append("empty output")
            if wc < 4:               reasons.append(f"too short ({wc} words, min=4)")
            if wc > 14:              reasons.append(f"too long ({wc} words, max=14)")
            if sim < 0.55:           reasons.append(f"sim={sim:.3f} < 0.55")
            if is_dup:               reasons.append("identical to original")
            print(f" │    → ❌ REJECTED — {', '.join(reasons)}")
        else:
            if sim > best_sim:
                best_sim = sim
                best_result = result
                print(f" │    → ✅ ACCEPTED — new best candidate (sim={sim:.3f})")
            else:
                print(f" │    → ✅ VALID but not better than current best (best sim={best_sim:.3f})")

        print(f" │ {'─'*60}")
        time.sleep(1)

    if best_result:
        print(f" │ 🏆 FINAL SELECTED : \"{best_result}\"")
        print(f" │    Similarity     : {best_sim:.3f}")
        print(f" └{'─'*65}")
        return best_result
    else:
        print(f" │ ⚠️  No valid paraphrase found → Keeping original: \"{clean}\"")
        print(f" └{'─'*65}")
        return clean


def paraphrase_description(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text

    paragraphs = [p.strip() for p in clean.split("\n") if p.strip()]
    rewritten = []
    success_count = 0

    print(f"\n ┌─ DESCRIPTION PARAPHRASE ({len(paragraphs)} paragraphs) {'─'*25}")

    for i, para in enumerate(paragraphs):
        orig_wc = len(para.split())

        print(f"\n │ ┌─ Paragraph {i+1}/{len(paragraphs)} {'─'*50}")
        print(f" │ │ ORIGINAL ({orig_wc} words):")
        _print_wrapped(para, prefix=" │ │    ")
        print(f" │ │ {'─'*60}")

        prompt = (
            f"Rewrite this job description paragraph professionally. "
            f"Keep ALL facts, requirements, and responsibilities. "
            f"Use different sentence structure and vocabulary. "
            f"Output ONLY the rewritten paragraph — no labels, no explanation.\n\n"
            f"Original:\n{para}"
        )

        best_result = None
        best_sim    = 0.0
        accepted_text = None

        for attempt in range(3):
            temp = round(0.65 + attempt * 0.08, 2)
            print(f" │ │ Attempt {attempt+1}/3 (temp={temp}):")

            raw = mistral_generate(prompt, max_tokens=500, temperature=temp)
            result = clean_output(raw).strip()

            rw  = len(result.split()) if result else 0
            sim = similarity_score(para, result) if result and rw >= 5 else 0.0

            if result:
                print(f" │ │    Paraphrased ({rw} words, sim={sim:.3f}):")
                _print_wrapped(result, prefix=" │ │       ")
            else:
                print(f" │ │    Paraphrased : (no output from model)")

            valid = bool(result) and rw >= 8 and sim >= 0.48

            if not valid:
                reasons = []
                if not result: reasons.append("empty output")
                if rw < 8:     reasons.append(f"too short ({rw} words, min=8)")
                if sim < 0.48: reasons.append(f"sim={sim:.3f} < 0.48")
                print(f" │ │    → ❌ REJECTED — {', '.join(reasons)}")
                if result and sim > best_sim:
                    best_sim = sim
                    best_result = result
                    print(f" │ │       (stored as best fallback, sim={sim:.3f})")
            else:
                print(f" │ │    → ✅ ACCEPTED on attempt {attempt+1}")
                rewritten.append(result)
                success_count += 1
                accepted_text = result
                break

            print(f" │ │ {'─'*60}")
            time.sleep(1)

        if accepted_text is None:
            print(f" │ │ {'─'*60}")
            if best_result and best_sim >= 0.40:
                print(f" │ │ 🔁 FALLBACK — Using best attempt (sim={best_sim:.3f}):")
                _print_wrapped(best_result, prefix=" │ │    ")
                rewritten.append(best_result)
                success_count += 1
            else:
                print(f" │ │ ⚠️  KEPT ORIGINAL — no acceptable paraphrase found (best sim={best_sim:.3f})")
                rewritten.append(para)

        print(f" │ └{'─'*62}")

    print(f"\n │ SUMMARY: {success_count}/{len(paragraphs)} paragraphs successfully paraphrased")
    print(f" └{'─'*80}\n")

    return "\n\n".join(rewritten)


def paraphrase_company(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text

    print(f"\n ┌─ COMPANY PARAPHRASE {'─'*43}")
    orig_wc = len(clean.split())
    print(f" │ Original ({orig_wc} words):")
    _print_wrapped(clean, prefix=" │    ")
    print(f" │ {'─'*60}")

    prompt = (
        f"Rewrite this company description professionally. "
        f"Preserve all facts. Use different wording. "
        f"Output ONLY the rewritten description.\n\nOriginal:\n{clean}"
    )

    raw = mistral_generate(prompt, max_tokens=600, temperature=0.68)
    result = clean_output(raw)
    rw  = len(result.split()) if result else 0
    sim = similarity_score(clean, result) if result and rw >= 10 else 0.0

    if result and rw >= 10:
        print(f" │ Paraphrased ({rw} words, sim={sim:.3f}):")
        _print_wrapped(result, prefix=" │    ")
        print(f" │ → ✅ ACCEPTED")
        print(f" └{'─'*65}")
        time.sleep(1)
        return result
    else:
        reasons = []
        if not result: reasons.append("empty output")
        if rw < 10:    reasons.append(f"too short ({rw} words, min=10)")
        print(f" │ → ❌ REJECTED — {', '.join(reasons)} — keeping original")
        print(f" └{'─'*65}")
        time.sleep(1)
        return clean


def paraphrase_tagline(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text[:300])
    if not clean:
        return text

    print(f"\n ┌─ TAGLINE PARAPHRASE {'─'*43}")
    print(f" │ Original : \"{clean}\"")
    print(f" │ {'─'*60}")

    prompt = (
        f"Rewrite this company tagline as a crisp, professional phrase. "
        f"Output ONLY the rewritten tagline (5–12 words). No explanation.\n\n"
        f"Original: {clean}"
    )

    raw = mistral_generate(prompt, max_tokens=35, temperature=0.75)
    result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")
    wc  = len(result.split()) if result else 0

    print(f" │ Paraphrased : \"{result}\"")
    print(f" │ Words: {wc}")

    if result and 3 <= wc <= 15:
        print(f" │ → ✅ ACCEPTED")
        print(f" └{'─'*65}")
        time.sleep(1)
        return result
    else:
        reasons = []
        if not result: reasons.append("empty output")
        if wc < 3:     reasons.append(f"too short ({wc} words, min=3)")
        if wc > 15:    reasons.append(f"too long ({wc} words, max=15)")
        print(f" │ → ❌ REJECTED — {', '.join(reasons)} — keeping original")
        print(f" └{'─'*65}")
        time.sleep(1)
        return clean


def _print_wrapped(text: str, prefix: str = "   ", width: int = 100):
    """Helper: print long text word-wrapped at `width` chars."""
    words = text.split()
    line  = []
    for w in words:
        line.append(w)
        if len(" ".join(line)) >= width:
            print(f"{prefix}{' '.join(line)}")
            line = []
    if line:
        print(f"{prefix}{' '.join(line)}")

# =============================================================================
#  DUPLICATE TRACKER
# =============================================================================

def _init_tracker():
    if not os.path.exists(PROCESSED_IDS_FILE):
        pd.DataFrame(columns=[
            "Job ID", "Job URL", "Job Title", "Company Name",
            "Status", "Timestamp", "WP ID",
        ]).to_csv(PROCESSED_IDS_FILE, index=False)

def load_processed_ids() -> tuple:
    _init_tracker()
    df = pd.read_csv(PROCESSED_IDS_FILE)
    return (
        set(df["Job ID"].fillna("").astype(str)),
        set(df.get("Job URL", pd.Series()).fillna("").astype(str)),
    )

def _upsert_row(job_id: str, updates: dict):
    _init_tracker()
    df = pd.read_csv(PROCESSED_IDS_FILE)
    mask = df["Job ID"].astype(str) == str(job_id)
    if mask.any():
        for col, val in updates.items():
            if col in df.columns:
                df.loc[mask, col] = val
        df.loc[mask, "Timestamp"] = datetime.now().isoformat()
    else:
        row = {"Job ID": job_id, "Timestamp": datetime.now().isoformat()}
        row.update(updates)
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_csv(PROCESSED_IDS_FILE, index=False)

def make_job_id(job_url: str, title: str = "", company: str = "", idx: int = 0) -> str:
    if job_url:
        return hashlib.md5(job_url.encode()).hexdigest()[:16]
    seed = f"{title}{company}{idx}"
    return hashlib.md5(seed.encode()).hexdigest()[:16]

def mark_scraped(job_id, job_url, title, company):
    _upsert_row(job_id, {"Job URL": job_url, "Job Title": title,
                          "Company Name": company, "Status": "scraped"})

def mark_paraphrased(job_id):
    _upsert_row(job_id, {"Status": "paraphrased"})

def mark_posted(job_id, wp_id, wp_url):
    _upsert_row(job_id, {"Status": f"posted", "WP ID": wp_id})

def mark_failed(job_id, reason):
    _upsert_row(job_id, {"Status": f"failed|{reason}"})

# =============================================================================
#  HELPERS  (unchanged from v4)
# =============================================================================

def should_skip_crawl(url: str) -> bool:
    if not url: return True
    return any(d in url.lower() for d in SKIP_CRAWL_DOMAINS)

def is_bad_url(url: str) -> bool:
    if not url or not url.startswith("http"): return True
    return any(d in url.lower() for d in BAD_DOMAINS)

def is_career_url(url: str) -> bool:
    l = url.lower()
    return any(k in l for k in ["career","jobs","apply","vacanci","recruit","opening",
                                  "hiring","work-with","join-us","join_us","opportunities"])

def is_contact_url(url: str) -> bool:
    l = url.lower()
    return any(k in l for k in ["contact","about","reach","get-in","enquir","support",
                                  "about-us","about_us"])

def is_about_url(url: str) -> bool:
    l = url.lower()
    return any(k in l for k in ["about","who-we-are","our-story","company","our-team",
                                  "overview"])

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

# =============================================================================
#  HTTP  (with retry + back-off)
# =============================================================================

def fetch_page(url: str, follow_redirects: bool = True, retries: int = 3) -> str | None:
    for attempt in range(retries):
        try:
            time.sleep(0.3 + attempt * 1.5)
            r = requests.get(url, headers=_next_headers(),
                             allow_redirects=follow_redirects, timeout=20)
            if r.status_code == 429:
                wait = 30 + attempt * 30
                log.warning(f"Rate-limited (429) — sleeping {wait}s")
                time.sleep(wait); continue
            if r.status_code in (403, 999):
                log.warning(f"Blocked ({r.status_code}): {url}"); return None
            if r.status_code != 200:
                log.warning(f"HTTP {r.status_code}: {url}"); return None
            text = r.text
            if len(text) > FETCH_CHAR_LIMIT: text = text[:FETCH_CHAR_LIMIT]
            return text
        except Exception as e:
            log.warning(f"fetch attempt {attempt+1} failed ({url}): {e}")
            time.sleep(2 + attempt * 2)
    return None

# =============================================================================
#  DATE HELPERS  (unchanged)
# =============================================================================

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
    except Exception: pass
    base = datetime.now()
    m = re.search(r"(\d+)\s*(hour|day|week|month|year)", text, re.I)
    if m:
        n, unit = int(m.group(1)), m.group(2).lower()
        if "hour"  in unit: base -= timedelta(hours=n)
        elif "day" in unit: base -= timedelta(days=n)
        elif "week"in unit: base -= timedelta(weeks=n)
        elif "month"in unit:
            mo = base.month - n; yr = base.year + mo // 12; mo = mo % 12 or 12
            base = base.replace(year=yr, month=mo)
        elif "year"in unit: base = base.replace(year=base.year - n)
        return base.strftime("%Y-%m-%d")
    if re.search(r"just\s*now|today", text, re.I): return datetime.now().strftime("%Y-%m-%d")
    return datetime.now().strftime("%Y-%m-%d")

def try_parse_date(s: str) -> datetime | None:
    if not s: return None
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%B %d, %Y", "%d %B %Y"):
        try: return datetime.strptime(s.strip(), fmt)
        except Exception: pass
    try: return datetime.fromisoformat(s)
    except Exception: pass
    m = re.match(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", s)
    if m:
        mon = MONTH_MAP.get(m.group(2)[:3].lower())
        if mon is not None: return datetime(int(m.group(3)), mon+1, int(m.group(1)))
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
        if "hour" in unit:  base -= timedelta(hours=n)
        elif "day"in unit:  base -= timedelta(days=n)
        elif "week"in unit: base -= timedelta(weeks=n)
        elif "month"in unit:
            mo = base.month - n; yr = base.year + mo // 12; mo = mo % 12 or 12
            base = base.replace(year=yr, month=mo)
    mo = base.month + 3; yr = base.year + (mo - 1) // 12; mo = (mo - 1) % 12 + 1
    return base.replace(year=yr, month=mo).strftime("%Y-%m-%d")

# =============================================================================
#  TEXT CLEANERS  (unchanged)
# =============================================================================

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
                     (r"\\u002[Ee]","."),(r"\\u0026",""),(r"u003[Ee]",""),
                     (r"u003[Cc]",""),(r"u0040","@"),(r"&amp;",""),(r"&lt;",""),
                     (r"&gt;",""),(r"&#64;","@"),(r"&#46;","."),(r"&nbsp;",""),
                     (r"%40","@"),(r"%2[Ee]","."),(r"%20",""),(r"[>]+$",""),
                     (r"[<]+$","")]:
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
    if not em or "@" not in em or "." not in em: return ""
    if not re.match(r"^[a-zA-Z0-9]", em): return ""
    return em

def clean_application_link(raw: str) -> str:
    if not raw: return ""
    raw = raw.strip()
    if "@" in raw and not raw.startswith("http"): return clean_email(raw)
    if raw.startswith("http"):
        url = raw
        if ".mu" in url.lower():
            def mu_replace(m):
                tld, path = m.group(1), m.group(2) or ""
                if path and re.match(r"^/[a-z0-9\-/]+$", path, re.I): return tld + path
                return tld
            url = re.sub(r"(\.mu)(\/[^\s]*)?$", mu_replace, url, flags=re.I)
        url = re.sub(r"#.*$", "", url)
        url = re.sub(r"(subject|applysubject|refno|applyref|applyhere|clickhere|applynow)(\?.*)?$","",url,flags=re.I)
        url = re.sub(r"[.,;:!?)]+$", "", url)
        return url.strip()
    return raw

def clean_logo_url(raw: str) -> str:
    if not raw: return ""
    raw = decode_html_entities(raw).strip()
    if not raw.startswith("http"): return ""
    return re.sub(r"[\"')\s]+$", "", raw)

# =============================================================================
#  EMAIL HELPERS  (unchanged)
# =============================================================================

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
        obs = re.findall(
            r"[a-zA-Z0-9._%+\-]+\s*[\[\(]?\s*at\s*[\]\)]?\s*[a-zA-Z0-9.\-]+"
            r"\s*[\[\(]?\s*dot\s*[\]\)]?\s*[a-zA-Z]{2,}", raw_html, re.I)
        if obs:
            norm = re.sub(r"\s*[\[\(]?\s*at\s*[\]\)]?\s*", "@", obs[0], flags=re.I)
            norm = re.sub(r"\s*[\[\(]?\s*dot\s*[\]\)]?\s*", ".", norm, flags=re.I)
            norm = re.sub(r"\s+", "", norm).lower()
            if "@" in norm and not FAKE_LOCAL_RE.match(norm.split("@")[0]): return norm
        found = extract_email_from_text(raw_html)
        if found: return found
    return ""

# =============================================================================
#  DECODE / FOLLOW LINKEDIN APPLY URL  (unchanged)
# =============================================================================

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
        except Exception: pass
    b64m = re.search(r"[?&]offsiteApplyUrl=([^&]+)", raw)
    if b64m:
        try:
            d2 = base64.b64decode(unquote(b64m.group(1))).decode("utf-8")
            p = json.loads(d2)
            if p and "url" in p: return p["url"]
        except Exception: pass
    return ""

def follow_linkedin_apply_button(soup: BeautifulSoup, job_url: str) -> str:
    for tag in soup.find_all("a", href=True):
        ctrl = tag.get("data-tracking-control-name", "")
        if "offsite" in ctrl.lower() or "apply" in ctrl.lower():
            r = decode_linkedin_apply_url(tag["href"])
            if r and not is_bad_url(r): return r
    for tag in soup.find_all("a", href=True):
        href = tag["href"]; text = tag.get_text().lower()
        if ("apply" in text or "/apply" in href) and "linkedin.com" not in href:
            if href.startswith("http") and not is_bad_url(href): return href
    return ""

# =============================================================================
#  JSON-LD PARSER  (unchanged)
# =============================================================================

def _parse_jsonld(html: str) -> dict:
    result = {}
    for raw in re.findall(
            r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
            html, re.S | re.I):
        try:
            data = json.loads(raw.strip())
        except Exception:
            continue
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
                "company_logo":    clean_logo_url(org.get("logo", "") if isinstance(org.get("logo"), str) else
                                                  org.get("logo", {}).get("url", "") if isinstance(org.get("logo"), dict) else ""),
                "company_url":     org.get("sameAs", "") or org.get("url", ""),
                "company_website": org.get("sameAs", "") or org.get("url", ""),
                "apply_url":       (data.get("url", "") or
                                    (data.get("applicationContact", {}) or {}).get("url", "")),
                "location":        _extract_location_jsonld(data.get("jobLocation", {})),
            })
            addr = data.get("jobLocation", {})
            if isinstance(addr, list): addr = addr[0] if addr else {}
            place = addr.get("address", {}) if isinstance(addr, dict) else {}
            if isinstance(place, dict):
                city = place.get("addressLocality", "")
                country = place.get("addressCountry", "")
                if city or country:
                    result["location"] = ", ".join(filter(None, [city, country]))
        elif schema_type in ("Organization", "Corporation", "LocalBusiness"):
            result.update({
                "company_name":     data.get("name", ""),
                "company_logo":     clean_logo_url(data.get("logo", "") if isinstance(data.get("logo"), str) else
                                                   data.get("logo", {}).get("url", "") if isinstance(data.get("logo"), dict) else ""),
                "company_url":      data.get("sameAs", "") or data.get("url", ""),
                "company_website":  data.get("sameAs", "") or data.get("url", ""),
                "company_industry": data.get("industry", ""),
                "company_founded":  str(data.get("foundingDate", "") or ""),
                "company_address":  _extract_address_jsonld(data.get("address", {})),
                "company_about":    data.get("description", ""),
            })
    return result

def _extract_salary_jsonld(obj) -> str:
    if not obj: return ""
    if isinstance(obj, str): return obj
    if isinstance(obj, dict):
        val = obj.get("value", {}); currency = obj.get("currency", "")
        if isinstance(val, dict):
            lo = val.get("minValue", ""); hi = val.get("maxValue", ""); unit = val.get("unitText", "")
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
            addr.get("addressLocality", ""),
            addr.get("addressRegion", ""),
            addr.get("addressCountry", ""),
        ]))
    return str(addr)

def _extract_address_jsonld(obj) -> str:
    if not obj: return ""
    if isinstance(obj, str): return obj
    if isinstance(obj, dict):
        return ", ".join(filter(None, [
            obj.get("streetAddress", ""),
            obj.get("addressLocality", ""),
            obj.get("addressRegion", ""),
            obj.get("postalCode", ""),
            obj.get("addressCountry", ""),
        ]))
    return ""

# =============================================================================
#  EXTRACT COMPANY DATA FROM JOB PAGE  (unchanged)
# =============================================================================

def extract_company_from_job_page(html: str, soup: BeautifulSoup) -> dict:
    result = {}
    ld = _parse_jsonld(html)
    if ld: result.update({k: v for k, v in ld.items() if v})

    def _meta(name_or_prop: str) -> str:
        tag = (soup.find("meta", attrs={"property": name_or_prop}) or
               soup.find("meta", attrs={"name": name_or_prop}))
        return (tag.get("content", "") if tag else "").strip()

    og_image = _meta("og:image")
    if og_image and not result.get("company_logo"):
        result["company_logo"] = clean_logo_url(og_image)

    def _sel(*selectors) -> str:
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
        if not result.get("company_about"):
            for pat in [r'"tagline"\s*:\s*"([^"]+)"', r'"description"\s*:\s*"([^"]{20,})"']:
                m = re.search(pat, txt)
                if m:
                    about = decode_html_entities(m.group(1)).strip()
                    if about and len(about) > 20: result["company_about"] = about; break

    return result

# =============================================================================
#  DEEP COMPANY WEBSITE CRAWLER  (v4 — unchanged)
# =============================================================================

def crawl_company_website_deep(website_url: str, job_title: str) -> dict:
    result = {"apply_url": "", "email": "", "method": ""}
    if not website_url or should_skip_crawl(website_url): return result

    deadline = time.time() + 20
    root = website_url.rstrip("/")
    parsed_root = urlparse(root)
    root_domain = parsed_root.netloc

    def _get(url: str):
        if time.time() > deadline: return None, None
        try:
            time.sleep(0.6)
            r = requests.get(url, headers=_next_headers(), timeout=12, allow_redirects=True)
            if r.status_code != 200: return None, None
            html = r.text[:FETCH_CHAR_LIMIT]
            return html, BeautifulSoup(html, "html.parser")
        except Exception as e:
            log.debug(f"deep crawl fetch error ({url}): {e}")
            return None, None

    def _links_from(soup: BeautifulSoup, base: str) -> list:
        out = []
        for tag in soup.find_all("a", href=True):
            href = make_absolute(tag.get("href", ""), base)
            if not href or not href.startswith("http"): continue
            if root_domain not in urlparse(href).netloc: continue
            out.append((href, tag.get_text(strip=True).lower()))
        return out

    home_html, home_soup = _get(root)
    if not home_html: return result
    links = _links_from(home_soup, root)

    CAREER_TEXT = re.compile(
        r"career|job|vacanc|opportunit|recruit|hiring|join\s*us|work\s*with\s*us|"
        r"open\s*positions|current\s*opening", re.I)
    careers_url = ""
    for href, txt in links:
        if CAREER_TEXT.search(txt) or is_career_url(href):
            careers_url = href; break

    if not careers_url:
        for path in ["/careers","/jobs","/job-openings","/vacancies",
                     "/work-with-us","/join-us","/opportunities"]:
            candidate = root + path
            if time.time() > deadline: break
            try:
                r = requests.head(candidate, headers=_next_headers(), timeout=8, allow_redirects=True)
                if r.status_code == 200:
                    careers_url = candidate; break
            except Exception:
                pass

    if not careers_url:
        return result

    careers_html, careers_soup = _get(careers_url)
    if not careers_soup: return result

    job_links = _links_from(careers_soup, careers_url)
    best_url, best_score = "", 0.0
    for href, txt in job_links:
        score = _title_similarity(job_title, txt)
        if score > best_score:
            best_score, best_url = score, href

    email_from_careers = scan_page_for_email(careers_soup, careers_html)

    if best_score < 0.3:
        if email_from_careers:
            return {"apply_url": "", "email": email_from_careers, "method": "deep_careers_email"}
        return {"apply_url": careers_url, "email": "", "method": "deep_careers_page"}

    job_html, job_soup = _get(best_url)
    if not job_soup:
        return {"apply_url": best_url, "email": "", "method": "deep_job_page_url"}

    APPLY_TEXT  = re.compile(r"apply\s*now|apply\s*online|apply\s*for|submit\s*(application|cv|resume)|"
                              r"apply\s*here|send\s*(cv|resume|application)", re.I)
    APPLY_CLASS = re.compile(r"apply|btn-apply|cta-apply|job-apply", re.I)

    apply_url = ""
    for tag in job_soup.find_all("a", href=True):
        txt = tag.get_text(strip=True)
        cls = " ".join(tag.get("class", []))
        href = make_absolute(tag.get("href", ""), best_url)
        if not href: continue
        if APPLY_TEXT.search(txt) or APPLY_CLASS.search(cls):
            if href.startswith("mailto:"):
                em = clean_email(href.replace("mailto:", ""))
                if em: return {"apply_url": "", "email": em, "method": "deep_apply_email"}
            if not is_bad_url(href):
                apply_url = href; break

    if not apply_url:
        form = job_soup.find("form")
        if form and form.get("action"):
            action = make_absolute(form["action"], best_url)
            if action and not is_bad_url(action):
                apply_url = action

    if not apply_url:
        em = scan_page_for_email(job_soup, job_html)
        if em: return {"apply_url": "", "email": em, "method": "deep_job_email"}

    if apply_url:
        return {"apply_url": apply_url, "email": "", "method": "deep_apply_button"}

    return {"apply_url": best_url, "email": "", "method": "deep_job_page_url"}

# =============================================================================
#  ABOUT / CONTACT / FOOTER SCRAPER  (v4 — unchanged)
# =============================================================================

def scrape_about_contact_footer(website_url: str) -> dict:
    empty = {"address": "", "phone": "", "email": "", "founded": "",
             "description": "", "social_links": ""}
    if not website_url or should_skip_crawl(website_url): return empty

    root = website_url.rstrip("/")
    parsed_root = urlparse(root)
    root_domain = parsed_root.netloc
    deadline = time.time() + 15

    def _get(url: str):
        if time.time() > deadline: return None, None
        try:
            time.sleep(0.5)
            r = requests.get(url, headers=_next_headers(), timeout=12, allow_redirects=True)
            if r.status_code != 200: return None, None
            html = r.text[:FETCH_CHAR_LIMIT]
            return html, BeautifulSoup(html, "html.parser")
        except Exception as e:
            log.debug(f"about/contact fetch error ({url}): {e}")
            return None, None

    result = dict(empty)

    def _harvest(soup: BeautifulSoup, html: str):
        full_text = soup.get_text(" ", strip=True)
        if not result["email"]:
            em = scan_page_for_email(soup, html)
            if em: result["email"] = em
        if not result["phone"]:
            ph = re.search(r"(?:\+?\d[\d\s\-().]{6,18}\d)(?=\s*(?:$|\n|[^\d]))", full_text)
            if ph:
                candidate = ph.group(0).strip()
                if len(re.sub(r"\D", "", candidate)) >= 7:
                    result["phone"] = candidate
        if not result["address"]:
            for el in soup.select("[itemprop='address'],[itemprop='streetAddress']"):
                t = el.get_text(strip=True)
                if t: result["address"] = t; break
            if not result["address"]:
                addr_m = re.search(
                    r"\d+[\w\s,.-]{5,80}(?:street|st\b|avenue|ave\b|road|rd\b|"
                    r"boulevard|blvd|lane|ln|drive|dr\b|way|close|court|building|"
                    r"floor|suite|tower|plaza|district|zone)", full_text, re.I)
                if addr_m:
                    result["address"] = addr_m.group(0).strip()[:200]
        if not result["founded"]:
            fy = re.search(
                r"(?:founded|established|incorporated|since|est\.?)\s*[:\-]?\s*((?:19|20)\d{2})",
                full_text, re.I)
            if fy: result["founded"] = fy.group(1)
        if not result["description"]:
            og = (soup.find("meta", property="og:description") or
                  soup.find("meta", attrs={"name": "description"}))
            if og:
                desc = og.get("content", "").strip()
                if len(desc) > 40: result["description"] = desc
            if not result["description"]:
                for p in soup.find_all("p"):
                    t = p.get_text(strip=True)
                    if len(t) > 80:
                        result["description"] = t[:500]; break
        if not result["social_links"]:
            socials = []
            for a in soup.find_all("a", href=True):
                href = a["href"].lower()
                for platform in ["twitter.com","x.com","facebook.com","instagram.com",
                                  "linkedin.com","youtube.com","tiktok.com"]:
                    if platform in href and href not in socials:
                        socials.append(a["href"]); break
            if socials: result["social_links"] = ", ".join(socials[:5])

    home_html, home_soup = _get(root)
    if home_soup:
        footer = home_soup.find("footer") or home_soup.select_one("#footer,.footer,[class*='footer']")
        if footer:
            _harvest(BeautifulSoup(str(footer), "html.parser"), str(footer))
        _harvest(home_soup, home_html)
        about_url = contact_url = ""
        for tag in home_soup.find_all("a", href=True):
            href = make_absolute(tag.get("href", ""), root)
            if not href or root_domain not in urlparse(href).netloc: continue
            txt = tag.get_text(strip=True).lower()
            if not about_url and (is_about_url(href) or "about" in txt):
                about_url = href
            if not contact_url and (is_contact_url(href) or "contact" in txt):
                contact_url = href
            if about_url and contact_url: break
        if about_url and time.time() < deadline:
            about_html, about_soup = _get(about_url)
            if about_soup: _harvest(about_soup, about_html)
        if contact_url and time.time() < deadline:
            contact_html, contact_soup = _get(contact_url)
            if contact_soup: _harvest(contact_soup, contact_html)
    return result

# =============================================================================
#  COMPANY PAGE SCRAPER  (unchanged)
# =============================================================================

def scrape_company_details(company_url: str) -> dict:
    empty = {
        "name":"","industry":"","size":"","headquarters":"","type":"",
        "founded":"","specialties":"","website":"","logo":"","about":"",
    }
    if not company_url: return empty
    log.info(f"Scraping company page: {company_url}")
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
                r = requests.get(base_url, headers=_next_headers(),
                                 allow_redirects=True, timeout=20)
                if r.status_code == 429:
                    log.warning("Company page rate-limited — sleeping 60s"); time.sleep(60); continue
                if r.status_code == 200:
                    text = r.text
                    if len(text) > FETCH_CHAR_LIMIT: text = text[:FETCH_CHAR_LIMIT]
                    html = text; break
                log.warning(f"Company page HTTP {r.status_code}: {base_url}"); break
            except Exception as e:
                log.warning(f"Company page fetch error (attempt {attempt+1}): {e}")
                time.sleep(2 + attempt * 2)
    if not html: return empty
    soup = BeautifulSoup(html, "html.parser")
    ld   = _parse_jsonld(html)

    def _sel(*selectors) -> str:
        for s in selectors:
            el = soup.select_one(s)
            if el:
                t = el.get_text(strip=True)
                if t: return t
        return ""

    def _get_detail(label: str) -> str:
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
    raw_logo = (og_img_tag.get("content", "") if og_img_tag else "") or ld.get("company_logo", "")
    if not raw_logo:
        for img in soup.select("img.org-top-card-primary-content__logo, img.artdeco-entity-image"):
            src = img.get("src", "") or img.get("data-delayed-url", "")
            if src and "ghost" not in src.lower():
                raw_logo = src; break
    logo = clean_logo_url(raw_logo)

    ws_tag = soup.select_one("a[data-tracking-control-name='about_website']")
    raw_ws = (ws_tag.get("href", "") if ws_tag else "") or _get_detail("Website") or ld.get("company_website", "")
    website = decode_linkedin_apply_url(raw_ws) or raw_ws

    name = (ld.get("company_name", "") or _sel("h1.org-top-card-summary__title", "h1", "title") or "")
    if " | LinkedIn" in name: name = name.split(" | ")[0].strip()

    about = (ld.get("company_about", "") or
             _sel("section.about-us p", ".core-section-container__content p",
                  ".org-about-us-organization-description__text",
                  ".org-about-module__description") or "")

    hosted_logo = upload_logo_to_wordpress(logo, name) if logo else ""

    return {
        "name":         name,
        "industry":     _get_detail("Industry") or ld.get("company_industry", ""),
        "size":         _get_detail("Company size"),
        "headquarters": _get_detail("Headquarters") or ld.get("company_address", ""),
        "type":         _get_detail("Type"),
        "founded":      _get_detail("Founded") or ld.get("company_founded", ""),
        "specialties":  _get_detail("Specialties"),
        "website":      website,
        "logo":         hosted_logo or logo,
        "about":        about,
    }

# =============================================================================
#  WORDPRESS LOGO UPLOAD  (unchanged)
# =============================================================================

def upload_logo_to_wordpress(logo_url: str, company_name: str) -> str:
    if not logo_url or not logo_url.startswith("http") or not WP_USER: return ""
    try:
        r = requests.get(logo_url, headers={"User-Agent": HEADERS["User-Agent"],
                                             "Referer": "https://www.linkedin.com/"}, timeout=15)
        if r.status_code != 200: return ""
        ct = r.headers.get("Content-Type", "image/jpeg")
        ext = "png" if "png" in ct else "jpg"
        fn  = re.sub(r"[^a-z0-9]", "-", company_name.lower()) + "-logo." + ext
        creds = base64.b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()
        wr = requests.post(WP_URL + "media",
                           headers={"Authorization": "Basic " + creds,
                                    "Content-Disposition": f"attachment; filename={fn}",
                                    "Content-Type": ct},
                           data=r.content, timeout=20)
        if wr.status_code in (200, 201): return wr.json().get("source_url", "")
        log.warning(f"WP upload failed ({wr.status_code})")
    except Exception as e:
        log.warning(f"uploadLogoToWordPress: {e}")
    return ""

# =============================================================================
#  COMPANY WEBSITE CRAWLER  (v3 fallback — unchanged)
# =============================================================================

def crawl_company_website(website_url: str, job_title: str) -> dict:
    log.info(f"Crawling company site (v3 fallback): {website_url}")
    if should_skip_crawl(website_url):
        return {"url": website_url, "email": "", "method": "fallback_website"}
    deadline = time.time() + 12
    root_url = website_url.rstrip("/")

    def get(url):
        if time.time() > deadline: return None
        try:
            time.sleep(0.5)
            r = requests.get(url, headers=HEADERS, timeout=10, allow_redirects=True)
            if r.status_code != 200: return None
            t = r.text
            return t[:FETCH_CHAR_LIMIT] if len(t) > FETCH_CHAR_LIMIT else t
        except Exception: return None

    home_html = get(root_url)
    if not home_html: return {"url": "", "email": "", "method": ""}
    soup_h = BeautifulSoup(home_html, "html.parser")
    home_email = scan_page_for_email(soup_h, home_html)
    if home_email: return {"url": "", "email": home_email, "method": "s7_homepage_email"}

    careers_url = contact_url = ""
    for tag in soup_h.find_all("a", href=True):
        href = make_absolute(tag.get("href", ""), root_url)
        link_text = tag.get_text().lower()
        if not href or is_bad_url(href) or href == root_url: continue
        if root_url not in href: continue
        if not careers_url and is_career_url(href): careers_url = href
        if not contact_url and (is_contact_url(href) or "contact" in link_text): contact_url = href
        if careers_url and contact_url: break

    if careers_url and time.time() < deadline:
        ch = get(careers_url)
        if ch:
            em = scan_page_for_email(BeautifulSoup(ch, "html.parser"), ch)
            if em: return {"url": "", "email": em, "method": "s7_careers_email"}
    if contact_url and time.time() < deadline:
        cth = get(contact_url)
        if cth:
            em = scan_page_for_email(BeautifulSoup(cth, "html.parser"), cth)
            if em: return {"url": "", "email": em, "method": "s7_contact_email"}
    if careers_url: return {"url": careers_url, "email": "", "method": "s7_careers_page"}
    return {"url": root_url, "email": "", "method": "fallback_website"}

# =============================================================================
#  APPLICATION DETAILS EXTRACTOR  (unchanged)
# =============================================================================

def extract_application_details(job_url: str, soup: BeautifulSoup,
                                  company_website: str, ld: dict,
                                  job_title: str = "") -> dict:
    desc_text = ""
    for sel in [".show-more-less-html__markup", ".description__text"]:
        el = soup.select_one(sel)
        if el: desc_text = el.get_text(); break

    if ld.get("apply_url") and not is_bad_url(ld["apply_url"]):
        return {"url": ld["apply_url"], "email": "", "method": "s0_jsonld"}

    apply_btn = follow_linkedin_apply_button(soup, job_url)
    if apply_btn:
        return {"url": apply_btn, "email": "", "method": "s0_apply_button"}

    for script in soup.find_all("script"):
        txt = script.string or ""
        for pat in [r'"applyStartUrl"\s*:\s*"([^"]+)"',
                    r'"applicationUrl"\s*:\s*"([^"]+)"']:
            m = re.search(pat, txt)
            if m:
                cand = decode_html_entities(m.group(1)).replace("\\", "")
                if cand.startswith("http") and not is_bad_url(cand):
                    return {"url": cand, "email": "", "method": "s1b_script_tag"}

    if company_website and not should_skip_crawl(company_website):
        deep = crawl_company_website_deep(company_website, job_title)
        if deep.get("email"):
            return {"url": "", "email": deep["email"], "method": deep["method"]}
        if deep.get("apply_url") and not is_bad_url(deep["apply_url"]):
            return {"url": deep["apply_url"], "email": "", "method": deep["method"]}

    desc_el = (soup.select_one(".show-more-less-html__markup") or
               soup.select_one(".description__text"))
    if desc_el:
        for a in desc_el.find_all("a", href=True):
            h = a.get("href", "")
            if not is_bad_url(h): return {"url": h, "email": "", "method": "s3_desc_link"}

    for u in re.findall(r"https?://[^\s\"'<>)(,\]]+", desc_text):
        u = re.sub(r"[.,;:!?)]+$", "", u)
        if not is_bad_url(u): return {"url": u, "email": "", "method": "s4_desc_url"}

    em = extract_email_from_text(desc_text)
    if em: return {"url": "", "email": em, "method": "s5_desc_email"}

    resolved = decode_linkedin_apply_url(company_website) or company_website
    if resolved and not is_bad_url(resolved):
        if should_skip_crawl(resolved):
            return {"url": resolved, "email": "", "method": "fallback_website"}
        res = crawl_company_website(resolved, job_title)
        if res.get("email") or res.get("url"): return res
        return {"url": resolved, "email": "", "method": "fallback_website"}

    return {"url": "", "email": "", "method": "not_found"}

# =============================================================================
#  JOB CRITERIA HELPERS  (unchanged)
# =============================================================================

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
        elif "seniority" in lower:
            if re.search(r"entry|associate|mid[\-\s]?senior|senior|director|executive|intern", text, re.I):
                return chip.get_text(strip=True)
    meta_map = {"employment type": soup.find("meta", {"name": "employmentType"}),
                "seniority level": soup.find("meta", {"name": "seniorityLevel"}),
                "industries":      soup.find("meta", {"name": "industry"})}
    tag = meta_map.get(lower)
    if tag: return tag.get("content", "")
    return ""

def get_workplace_type(soup: BeautifulSoup) -> str:
    for s in [".topcard__workplace-type",
              ".job-details-jobs-unified-top-card__workplace-type",
              ".jobs-unified-top-card__workplace-type"]:
        el = soup.select_one(s)
        if el: return el.get_text(strip=True)
    for chip in soup.select(
        ".job-details-jobs-unified-top-card__job-insight,"
        ".jobs-unified-top-card__job-insight"):
        t = chip.get_text(strip=True)
        if re.match(r"^(remote|on[\-\s]?site|hybrid)$", t, re.I): return t
    return ""

# =============================================================================
#  JOB FIELD INFERENCE  (unchanged)
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
     ["recruitment","onboarding","performance management","employee relations","hr","workforce"]),
    ("Engineering",
     ["mechanical engineer","civil engineer","electrical engineer","structural engineer",
      "process engineer","project engineer","maintenance engineer","production engineer",
      "quality engineer","safety engineer","site engineer","design engineer"],
     ["engineering","cad","autocad","solidworks","manufacturing","plant","machinery","commissioning"]),
    ("Healthcare & Medicine",
     ["doctor","physician","nurse","pharmacist","medical officer","surgeon","anaesthetist",
      "physiotherapist","radiographer","lab technician","clinical","healthcare manager",
      "occupational therapist","dentist","midwife"],
     ["hospital","clinic","patient","medical","health","pharmaceutical","diagnosis","treatment"]),
    ("Education & Training",
     ["teacher","lecturer","professor","trainer","educator","tutor","school principal",
      "academic","curriculum","e-learning","instructional designer","teaching assistant"],
     ["school","university","college","classroom","students","pedagogy","curriculum","education"]),
    ("Hospitality & Tourism",
     ["hotel manager","front desk","housekeeping","chef","sous chef","food and beverage",
      "f&b manager","restaurant manager","bartender","waiter","concierge","tour guide",
      "travel agent","events coordinator","catering"],
     ["hospitality","hotel","resort","tourism","guest","accommodation","restaurant","kitchen"]),
    ("Logistics & Supply Chain",
     ["supply chain manager","logistics coordinator","warehouse manager","fleet manager",
      "procurement manager","purchasing manager","import export","freight","shipping coordinator",
      "inventory manager","demand planner"],
     ["logistics","supply chain","warehouse","inventory","freight","procurement","sourcing"]),
    ("Legal",
     ["lawyer","attorney","legal counsel","paralegal","compliance officer","legal advisor",
      "solicitor","barrister","corporate counsel","legal manager","contract manager"],
     ["legal","law","contracts","litigation","regulatory","compliance","gdpr"]),
    ("Administration & Operations",
     ["office manager","executive assistant","administrative officer","operations manager",
      "pa","personal assistant","receptionist","data entry","office administrator",
      "company secretary","business analyst"],
     ["administration","operations","office","coordination","scheduling","reporting","clerical"]),
    ("Customer Service",
     ["customer service","call centre","customer success","customer support","help desk",
      "service advisor","client relations","customer experience","contact centre"],
     ["customer","support","helpdesk","tickets","escalation","satisfaction","service level"]),
    ("Construction & Real Estate",
     ["quantity surveyor","site supervisor","project manager construction","architect",
      "draughtsman","property manager","estate agent","real estate","building inspector",
      "land surveyor","construction manager"],
     ["construction","building","property","real estate","site","contractor","tender"]),
    ("Manufacturing & Production",
     ["production manager","quality control","quality assurance","qa","qc","factory manager",
      "plant manager","production supervisor","assembly","cnc operator","technician"],
     ["production","manufacturing","factory","assembly","quality","lean","six sigma"]),
    ("Design & Creative",
     ["graphic designer","ui/ux","product designer","art director","creative director",
      "animator","illustrator","photographer","videographer","motion designer","web designer"],
     ["design","creative","adobe","figma","photoshop","illustrator","indesign","sketch","branding"]),
    ("Research & Science",
     ["research scientist","data scientist","lab researcher","research analyst",
      "clinical researcher","environmental scientist","chemist","biologist","statistician"],
     ["research","analysis","data","laboratory","science","experiment","findings","methodology"]),
    ("Security",
     ["security officer","security guard","security manager","cctv","loss prevention",
      "risk manager","health and safety","hse officer","osh","fire safety"],
     ["security","safety","risk","surveillance","patrol","access control","emergency"]),
    ("Media & Journalism",
     ["journalist","editor","reporter","broadcast","news anchor","content creator",
      "media manager","radio","television","producer","scriptwriter"],
     ["media","journalism","broadcast","news","editorial","publishing","press"]),
    ("Non-Profit & Social Work",
     ["social worker","ngo","charity","programme coordinator","community development",
      "welfare officer","case manager","development officer","fundraiser","volunteer coordinator"],
     ["social","ngo","community","welfare","beneficiary","donor","impact","charity"]),
]

def infer_job_field(title: str, description: str) -> str:
    if not title and not description: return ""
    combined = ((title or "") + " " + (description or "")).lower()
    best_field, best_score = "", 0
    for label, high_keys, supporting in FIELD_KEYWORD_MAP:
        score = sum(3 for k in high_keys if k in combined)
        score += sum(1 for k in supporting if k in combined)
        if score > best_score: best_score, best_field = score, label
    if best_score >= 3: return best_field
    return ""

# =============================================================================
#  QUALIFICATION / EXPERIENCE EXTRACTORS  (unchanged)
# =============================================================================

QUALIFICATION_TIERS = [
    ("PhD / Doctorate",         ["phd","ph.d","doctorate","doctoral","doctor of philosophy"]),
    ("Master's Degree",         ["master","msc","m.sc","ma ","m.a ","mba","m.b.a","meng","m.eng","mphil",
                                  "postgraduate","post-graduate","post graduate"]),
    ("Bachelor's Degree",       ["bachelor","bsc","b.sc","ba ","b.a ","beng","b.eng","bcom","b.com","bba",
                                  "llb","degree in","undergraduate degree","honours degree","hons"]),
    ("Higher National Diploma", ["hnd","hnc","higher national diploma","higher national certificate",
                                  "higher diploma","advanced diploma"]),
    ("Diploma",                 ["diploma","dip ","dip.","associate degree","foundation degree"]),
    ("Professional Certification",["acca","cpa","cfa","cima","pmp","prince2","cissp","aws certified",
                                    "comptia","cisco","ccna","ccnp","shrm","cipd","chartered",
                                    "certified public","certified financial","certified project",
                                    "professional certification","professional certificate"]),
    ("A-Levels / HSC",          ["a-level","a level","hsc","higher school certificate","ib diploma",
                                  "international baccalaureate","gce advanced"]),
    ("O-Levels / School Certificate",["o-level","o level","igcse","gcse","school certificate",
                                       "sc ","cpe","certificate of primary"]),
    ("No Formal Qualification Required",["no qualification","no degree","no formal","school leaver",
                                          "entry level","no experience required","training provided","will train"]),
]

def extract_qualification(text: str) -> str:
    if not text: return ""
    if re.search(r"nursery|primary years|ib pyp|aged between|boys and girls", text, re.I): return ""
    lower = text.lower()
    for label, keywords in QUALIFICATION_TIERS:
        if any(k in lower for k in keywords): return label
    return ""

NO_EXP_KW = ["no experience","no prior experience","fresh graduate","freshers",
              "entry level","entry-level","0 years","zero experience",
              "training provided","will train","no experience required"]
LESS1_KW  = ["less than 1 year","under 1 year","6 months","less than a year",
              "some experience","minimal experience"]

def years_to_band(n: int) -> str:
    if n <= 0:  return "No Experience Required"
    if n <= 2:  return "1 - 2 Years"
    if n <= 5:  return "3 - 5 Years"
    if n <= 10: return "6 - 10 Years"
    return "10+ Years"

def extract_experience(text: str) -> str:
    if not text: return ""
    if re.search(r"aged?\s+between|boys\s+and\s+girls|nursery|primary\s+years|IB\s+PYP", text, re.I): return ""
    lower = text.lower()
    if any(k in lower for k in NO_EXP_KW): return "No Experience Required"
    if any(k in lower for k in LESS1_KW):  return "Less than 1 Year"
    patterns = [
        r"(\d+)\s*[-–to]+\s*(\d+)\s*\+?\s*years?",
        r"(\d+)\s*\+\s*years?\s*(?:of\s+)?(?:experience)?",
        r"(?:minimum|at\s+least|over|more\s+than)\s+(\d+)\s*\+?\s*years?",
        r"(\d+)\s*years?\s*(?:of\s+)?(?:relevant\s+)?(?:work\s+)?experience",
    ]
    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            raw = int(m.group(1))
            if raw > 20: continue
            return years_to_band(raw)
    return ""

# =============================================================================
#  JOB DETAIL SCRAPER  (v4 + paraphrase integrated)
# =============================================================================

def scrape_job_details(job_url: str, processed_ids: set, processed_urls: set) -> dict | None:
    """
    Scrape a single LinkedIn job.
    Returns the job dict (with paraphrased fields) or None.
    processed_ids / processed_urls are used to skip duplicates early.
    """
    job_id = make_job_id(job_url)
    if job_id in processed_ids or job_url in processed_urls:
        print(C_DIM(f"  ⧳ Already processed — skipped ({job_url})"))
        return None

    log.info(f"Scraping job: {job_url}")
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

    title = sel_text(".top-card-layout__title", "h1.topcard__title",
                     ".job-details-jobs-unified-top-card__job-title", "h1")
    if not title:
        return None

    company_name_fallback = sel_text(".topcard__org-name-link",
                                     ".job-details-jobs-unified-top-card__company-name",
                                     ".topcard__flavor")
    company_url_el = (soup.select_one(".topcard__org-name-link") or
                      soup.select_one(".job-details-jobs-unified-top-card__company-name a"))
    company_url  = company_url_el.get("href", "") if company_url_el else ""
    location     = sel_text(".topcard__flavor--bullet",
                             ".job-details-jobs-unified-top-card__bullet")
    workplace_type = get_workplace_type(soup)

    time_el = soup.find("time")
    raw_posted = (time_el.get("datetime", "") if time_el else "") or \
                 sel_text(".posted-time-ago__text",
                          ".job-details-jobs-unified-top-card__posted-date")
    posted_date = resolve_posted_date(raw_posted)

    raw_desc    = sel_text(".show-more-less-html__markup", ".description__text")
    description = clean_description(raw_desc)

    salary = ""
    for s in [".compensation__salary", ".salary", "[class*='salary']", "[class*='compensation']"]:
        el = soup.select_one(s)
        if el: salary = el.get_text(strip=True); break
    if not salary:
        for chip in soup.select(".job-details-jobs-unified-top-card__job-insight"):
            t = chip.get_text(strip=True)
            if re.search(r"\$|MUR|Rs\.?|SAR|salary|/yr|/hour|per month", t, re.I):
                salary = t; break

    raw_job_type = get_job_criteria(soup, "Employment type") or workplace_type
    job_type     = raw_job_type or "Full-time"

    linkedin_function = get_job_criteria(soup, "Job function")
    linkedin_industry = get_job_criteria(soup, "Industries")

    real_deadline      = parse_deadline(soup)
    estimated_deadline = estimate_deadline_from_posted(posted_date) if not real_deadline else ""
    effective_deadline = real_deadline or estimated_deadline

    # ── Layer 1: job-page extraction ─────────────────────────────────────────
    job_page_co = extract_company_from_job_page(html, soup)
    ld          = _parse_jsonld(html)

    # ── Layer 2: LinkedIn company page ───────────────────────────────────────
    time.sleep(0.5)
    company = scrape_company_details(company_url)

    # ── Layer 3: merge LinkedIn sources ──────────────────────────────────────
    def _first(*vals) -> str:
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

    # ── Layer 4: deep company website crawl ──────────────────────────────────
    if merged_website and not should_skip_crawl(merged_website):
        site_info = scrape_about_contact_footer(merged_website)
        if not merged_hq     and site_info.get("address"):     merged_hq      = site_info["address"]
        if not merged_founded and site_info.get("founded"):    merged_founded  = site_info["founded"]
        if not merged_about  and site_info.get("description"): merged_about   = site_info["description"]
        if site_info.get("phone") and not re.search(r"\d{5,}", merged_about or ""):
            phone_note = f"Phone: {site_info['phone']}"
            merged_about = (merged_about + "\n" + phone_note).strip() if merged_about else phone_note
    else:
        site_info = {}

    job_field = linkedin_function or merged_industry or infer_job_field(title, description)

    # ── Layer 5: application details ─────────────────────────────────────────
    time.sleep(0.2)
    apply_data = extract_application_details(
        job_url, soup, merged_website, ld, job_title=title)

    if not apply_data.get("email") and not apply_data.get("url"):
        if site_info.get("email"):
            apply_data = {"url": "", "email": site_info["email"], "method": "site_info_email"}

    raw_apply = ""
    if apply_data.get("email"):
        raw_apply = clean_email(apply_data["email"])
    elif apply_data.get("url") and apply_data.get("method") != "not_found":
        raw_apply = apply_data["url"]
    apply_link = clean_application_link(raw_apply)

    qualifications = extract_qualification(description)
    experience     = extract_experience(description)

    # ─────────────────────────────────────────────────────────────────────────
    #  ▶▶ PARAPHRASE — title, description, company about
    # ─────────────────────────────────────────────────────────────────────────
    mark_scraped(job_id, job_url, title, merged_name)
    processed_ids.add(job_id)
    processed_urls.add(job_url)

    paraphrased_title = title
    paraphrased_desc  = description
    paraphrased_about = merged_about

    if ENABLE_PARAPHRASE and MISTRAL_API_KEY != "your_mistral_api_key_here":
        print(C_BLUE(f"\n  ✍️  Paraphrasing '{title}' ..."))

        paraphrased_title = paraphrase_title(title)
        paraphrased_desc  = paraphrase_description(description)
        if merged_about:
            paraphrased_about = paraphrase_company(merged_about)

        mark_paraphrased(job_id)
    else:
        print(C_DIM("  ⚠️  Paraphrasing skipped (ENABLE_PARAPHRASE=False or no API key)"))

    return {
        # Paraphrased fields (used for WP post + Excel)
        "jobTitle":          paraphrased_title,
        "jobDescription":    paraphrased_desc,
        "companyDetails":    paraphrased_about,
        # Original fields preserved for reference
        "originalTitle":     title,
        "originalDesc":      description,
        # Structured fields
        "jobType":           job_type,
        "jobQualifications": qualifications,
        "jobExperience":     experience,
        "jobLocation":       location,
        "jobField":          job_field,
        "datePosted":        posted_date,
        "deadline":          effective_deadline,
        "application":       apply_link,
        "companyUrl":        company_url,
        "companyName":       merged_name,
        "companyLogo":       clean_logo_url(merged_logo),
        "companyIndustry":   merged_industry,
        "companyFounded":    merged_founded,
        "companyType":       merged_type,
        "companyWebsite":    merged_website,
        "companyAddress":    merged_hq,
        "jobUrl":            job_url,
        "estimatedDeadline": estimated_deadline,
        "salaryRange":       salary,
        "_jobId":            job_id,
    }

# =============================================================================
#  WORDPRESS POSTING
# =============================================================================

def _wp_auth_headers() -> dict:
    token = base64.b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()
    return {"Authorization": f"Basic {token}", "Content-Type": "application/json"}

def get_or_create_term(taxonomy_url: str, name: str) -> int | None:
    if not name or not name.strip(): return None
    slug = re.sub(r"[^a-z0-9-]", "-", name.lower().strip())
    h = _wp_auth_headers()
    try:
        r = requests.get(f"{taxonomy_url}?slug={slug}", headers=h, timeout=10, verify=False)
        terms = r.json()
        if isinstance(terms, list) and terms:
            return terms[0]["id"]
    except Exception:
        pass
    try:
        r = requests.post(taxonomy_url, json={"name": name, "slug": slug},
                          headers=h, auth=(WP_USER, WP_PASSWORD), timeout=10, verify=False)
        return r.json().get("id")
    except Exception as e:
        log.error(f"Term create error '{name}': {e}")
        return None

def post_job_to_wordpress(job: dict) -> tuple:
    h = _wp_auth_headers()
    title       = sanitize_text(job.get("jobTitle", ""))
    description = sanitize_text(job.get("jobDescription", ""))
    if not title or not description:
        return None, None

    slug = re.sub(r"[^a-z0-9-]", "-", title.lower())[:80]
    try:
        r = requests.get(f"{WP_JOBS_URL}?slug={slug}", headers=h, timeout=10, verify=False)
        posts = r.json()
        if isinstance(posts, list) and posts:
            log.info(f"⏭ Job already on WP: {title}")
            return posts[0]["id"], posts[0].get("link")
    except Exception:
        pass

    logo_url    = sanitize_text(job.get("companyLogo", ""), is_url=True)
    location    = sanitize_text(job.get("jobLocation", ""))
    raw_type    = sanitize_text(job.get("jobType", "Full-time"))
    job_type_s  = JOB_TYPE_MAPPING.get(raw_type.lower().strip(), "full-time")
    company     = sanitize_text(job.get("companyName", ""))
    application = sanitize_text(job.get("application", ""), is_url=True)
    deadline    = sanitize_text(job.get("deadline", "")) or sanitize_text(job.get("estimatedDeadline", ""))
    co_website  = sanitize_text(job.get("companyWebsite", ""), is_url=True)
    qualif      = sanitize_text(job.get("jobQualifications", ""))
    experience  = sanitize_text(job.get("jobExperience", ""))
    industry    = sanitize_text(job.get("companyIndustry", ""))
    co_address  = sanitize_text(job.get("companyAddress", ""))
    job_field   = sanitize_text(job.get("jobField", ""))
    job_url_src = sanitize_text(job.get("jobUrl", ""), is_url=True)
    co_founded  = sanitize_text(job.get("companyFounded", ""))
    co_type     = sanitize_text(job.get("companyType", ""))
    salary      = sanitize_text(job.get("salaryRange", ""))
    about       = sanitize_text(job.get("companyDetails", ""))

    is_email = bool(re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", application))
    is_url_v = bool(re.match(r"^https?://[^\s]+$", application))
    if not (is_email or is_url_v):
        application = ""

    # Upload logo
    attachment_id = None
    if logo_url:
        try:
            img_r = requests.get(logo_url, headers={"User-Agent": HEADERS["User-Agent"],
                                                      "Referer": "https://www.linkedin.com/"}, timeout=15)
            if img_r.status_code == 200:
                ct  = img_r.headers.get("Content-Type", "image/jpeg")
                ext = "png" if "png" in ct else "jpg"
                fn  = re.sub(r"[^a-z0-9]", "-", company.lower()) + "-logo." + ext
                up_h = dict(_wp_auth_headers())
                up_h["Content-Disposition"] = f"attachment; filename={fn}"
                up_h["Content-Type"] = ct
                up_r = requests.post(WP_MEDIA_URL, headers=up_h, data=img_r.content,
                                     auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
                if up_r.status_code in (200, 201):
                    attachment_id = up_r.json().get("id")
        except Exception as e:
            log.warning(f"Logo upload failed: {e}")

    region_term_id   = get_or_create_term(f"{WP_BASE}/job_listing_region", location)
    job_type_term_id = get_or_create_term(f"{WP_BASE}/job_listing_type",
                                           job_type_s.replace("-", " ").title())

    # Generate tagline from company about
    tagline = paraphrase_tagline(about[:300]) if about else ""

    payload = {
        "title":          title,
        "content":        description,
        "status":         "publish",
        "featured_media": attachment_id or 0,
        "meta": {
            "_job_title":          title,
            "_job_location":       location,
            "_job_type":           job_type_s,
            "_job_description":    description,
            "_application":        application,
            "_job_expires":        deadline,
            "_company_name":       company,
            "_company_website":    co_website,
            "_company_logo":       str(attachment_id) if attachment_id else "",
            "_company_industry":   industry,
            "_company_address":    co_address,
            "_company_founded":    co_founded,
            "_company_type":       co_type,
            "_company_tagline":    tagline,
            "_company_details":    about,
            "_job_qualifications": qualif,
            "_job_experiences":    experience,
            "_job_field":          job_field,
            "_job_source_url":     job_url_src,
            "_job_salary":         salary,
        },
    }
    if region_term_id:   payload["job_listing_region"] = [region_term_id]
    if job_type_term_id: payload["job_listing_type"]   = [job_type_term_id]

    for attempt in range(3):
        try:
            r = requests.post(WP_JOBS_URL, json=payload, headers=h,
                              auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
            r.raise_for_status()
            post = r.json()
            log.info(f"✅ Job posted: '{title}' → WP ID {post.get('id')}")
            return post.get("id"), post.get("link")
        except Exception as e:
            log.error(f"Job post attempt {attempt+1} failed: {e}")
            if attempt < 2: time.sleep(2 ** attempt)
    return None, None

# =============================================================================
#  VERBOSE PRINTER
# =============================================================================

def print_job_verbose(job: dict, index: int, total: int):
    desc = job.get("jobDescription", "")
    desc_preview = (desc[:400] + " [...]") if len(desc) > 400 else desc
    desc_indented = "\n".join("   " + line for line in desc_preview.splitlines() if line.strip())
    apply = job.get("application", "")
    logo  = job.get("companyLogo", "")
    orig_title = job.get("originalTitle", "")
    print()
    print(C_DIVIDER())
    print(C_HEADER(f"  JOB {index}/{total}"))
    print(C_DIVIDER())
    print(f"  {C_LABEL('Title (original)')}  : {C_VALUE(orig_title)}")
    print(f"  {C_LABEL('Title (paraphrased)')}: {C_GREEN(job.get('jobTitle',''))}")
    print(f"  {C_LABEL('Job Type')}           : {job.get('jobType','')}")
    print(f"  {C_LABEL('Field')}              : {job.get('jobField','') or C_DIM('—')}")
    print(f"  {C_LABEL('Location')}           : {job.get('jobLocation','') or C_DIM('—')}")
    print(f"  {C_LABEL('Seniority')}          : {job.get('jobExperience','') or C_DIM('—')}")
    print(f"  {C_LABEL('Qualifications')}     : {job.get('jobQualifications','') or C_DIM('—')}")
    print(f"  {C_LABEL('Salary')}             : {job.get('salaryRange','') or C_DIM('—')}")
    print(f"  {C_LABEL('Date Posted')}        : {job.get('datePosted','') or C_DIM('—')}")
    print(f"  {C_LABEL('Deadline')}           : {job.get('deadline','') or C_DIM('—')}")
    print(f"  {C_LABEL('Apply Link')}         : {C_GREEN(apply) if apply else C_DIM('— not found —')}")
    print()
    print(f"  {C_BLUE('── COMPANY ──────────────────────────────────────────────')}")
    print(f"  {C_LABEL('Name')}           : {C_VALUE(job.get('companyName','') or C_DIM('—'))}")
    print(f"  {C_LABEL('Industry')}       : {job.get('companyIndustry','') or C_DIM('—')}")
    print(f"  {C_LABEL('Type')}           : {job.get('companyType','') or C_DIM('—')}")
    print(f"  {C_LABEL('Founded')}        : {job.get('companyFounded','') or C_DIM('—')}")
    print(f"  {C_LABEL('Headquarters')}   : {job.get('companyAddress','') or C_DIM('—')}")
    print(f"  {C_LABEL('Website')}        : {job.get('companyWebsite','') or C_DIM('—')}")
    print(f"  {C_LABEL('Logo')}           : {logo if logo else C_DIM('— none —')}")
    about = job.get("companyDetails", "")
    if about:
        preview = (about[:200] + " [...]") if len(about) > 200 else about
        print(f"  {C_LABEL('About')}          : {preview}")
    print()
    print(f"  {C_BLUE('── DESCRIPTION PREVIEW (paraphrased) ───────────────────')}")
    print(desc_indented if desc_indented else C_DIM("   — no description —"))
    print()
    print(f"  {C_LABEL('Job URL')}        : {C_DIM(job.get('jobUrl',''))}")
    print(C_DIVIDER())

# =============================================================================
#  URL COLLECTION — GUEST API  (unchanged)
# =============================================================================

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
            href = tag.get("href", "")
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
                print(C_RED(f"  ⏳ Rate limited (429) — waiting {wait}s ..."))
                time.sleep(wait); continue
            if r.status_code in (400, 403, 999):
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

# =============================================================================
#  EXCEL SAVE  (includes original title/desc columns for audit)
# =============================================================================

def _save_excel(jobs: list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    headers = [
        "Job Title","Original Title","Job Type","Job Qualifications","Job Experience",
        "Job Location","Job Field","Date Posted","Deadline",
        "Job Description","Application","Company URL","Company Name",
        "Company Logo","Company Industry","Company Founded","Company Type",
        "Company Website","Company Address","Company Details","Job URL",
        "Estimated Deadline","Salary Range",
    ]
    ws.append(headers)
    for job in jobs:
        ws.append([
            job["jobTitle"],
            job.get("originalTitle",""),
            job["jobType"],
            job["jobQualifications"],
            job["jobExperience"],
            job["jobLocation"],
            job["jobField"],
            job["datePosted"],
            job["deadline"],
            job["jobDescription"],
            job["application"],
            job["companyUrl"],
            job["companyName"],
            job["companyLogo"],
            job["companyIndustry"],
            job["companyFounded"],
            job["companyType"],
            job["companyWebsite"],
            job["companyAddress"],
            job["companyDetails"],
            job["jobUrl"],
            job["estimatedDeadline"],
            job["salaryRange"],
        ])
    wb.save(OUTPUT_FILE)
    log.info(f"Saved {len(jobs)} rows → {OUTPUT_FILE}")

# =============================================================================
#  MAIN CRAWL
# =============================================================================

def craw():
    start_time = time.time()
    _init_tracker()

    print()
    print(C_HEADER("=" * 72))
    print(C_HEADER("  LINKEDIN JOB SCRAPER v4 + MISTRAL PARAPHRASE"))
    print(C_HEADER("=" * 72))
    print(f"  Keywords      : {len(SEARCH_KEYWORDS)}")
    print(f"  Max pages     : {'unlimited' if not MAX_PAGES else MAX_PAGES} per keyword")
    print(f"  Job cap       : {'none' if not JOB_LIMIT else JOB_LIMIT}")
    print(f"  Paraphrase    : {'✅ enabled' if ENABLE_PARAPHRASE else '❌ disabled'}")
    print(f"  NLP available : {'✅' if _NLP_AVAILABLE else '⚠️  no sentence-transformers / language-tool'}")
    print(f"  Started       : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(C_HEADER("=" * 72))

    processed_ids, processed_urls = load_processed_ids()
    print(f"  Tracker loaded: {len(processed_ids)} previously processed job IDs\n")

    seen_urls = set(); all_job_urls = []; seen_content = set()

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
    print(C_HEADER(f"  Total unique URLs collected: {len(all_job_urls)}"))
    print()

    jobs = []; errors = 0
    for j, url in enumerate(all_job_urls):
        print(f"\n{C_HEADER(f'>>> Scraping job {j+1}/{len(all_job_urls)} ...')}")
        log.info(f"URL: {url}")
        try:
            job = scrape_job_details(url, processed_ids, processed_urls)
            if job and job.get("jobTitle"):
                fp = ((job.get("originalTitle") or "").lower().strip(),
                      (job.get("companyName")   or "").lower().strip(),
                      (job.get("jobLocation")   or "").lower().strip())
                if fp in seen_content:
                    print(C_DIM(f"  ⧳  Duplicate content — skipped"))
                else:
                    seen_content.add(fp)
                    jobs.append(job)
                    print_job_verbose(job, j+1, len(all_job_urls))

                    # Post to WordPress
                    print(C_BLUE(f"\n  📤 Posting to WordPress …"))
                    wp_id, wp_url = post_job_to_wordpress(job)
                    if wp_id:
                        mark_posted(job["_jobId"], wp_id, wp_url or "")
                        print(C_GREEN(f"  ✅ WP ID={wp_id}  🔗 {wp_url}"))
                    else:
                        mark_failed(job["_jobId"], "wp_post_failed")
                        print(C_RED("  ❌ WordPress post failed"))
            else:
                print(C_RED("  ✗  No title found / skipped"))
        except Exception as e:
            errors += 1
            print(C_RED(f"  ✗  ERROR: {e}"))
            log.warning(f"Job error: {e}")

        time.sleep(DELAY_S)
        if len(jobs) % 50 == 0 and len(jobs) > 0:
            _save_excel(jobs)

    _save_excel(jobs)

    mins = round((time.time() - start_time) / 60, 1)
    print()
    print(C_HEADER("=" * 72))
    print(C_HEADER("  SCRAPE COMPLETE"))
    print(C_HEADER("=" * 72))
    print(f"  {C_LABEL('Total scraped')}  : {C_GREEN(str(len(jobs)))} jobs")
    print(f"  {C_LABEL('Errors')}         : {C_RED(str(errors)) if errors else '0'}")
    print(f"  {C_LABEL('Duration')}       : ~{mins} min")
    print(f"  {C_LABEL('Output file')}    : {OUTPUT_FILE}")
    print(f"  {C_LABEL('Tracker file')}   : {PROCESSED_IDS_FILE}")

    if jobs:
        from collections import Counter
        fields = Counter(j.get("jobField") or "Unknown" for j in jobs)
        print(f"\n  {C_LABEL('Jobs by field:')}")
        for field, count in fields.most_common():
            print(f"    {field:<35} {'█'*min(count,40)} {count}")

        with_apply = sum(1 for j in jobs if j.get("application"))
        with_email = sum(1 for j in jobs if "@" in (j.get("application") or ""))
        with_url   = with_apply - with_email
        no_apply   = len(jobs) - with_apply
        print(f"\n  {C_LABEL('Application links:')}")
        print(f"    URL found    : {with_url}")
        print(f"    Email found  : {with_email}")
        print(f"    Not found    : {no_apply}")

        fill_fields = ["companyName","companyIndustry","companyLogo",
                       "companyWebsite","companyAddress","companyFounded","companyDetails"]
        print(f"\n  {C_LABEL('Company field fill-rate:')}")
        for f in fill_fields:
            filled = sum(1 for j in jobs if j.get(f))
            pct = round(filled / len(jobs) * 100) if jobs else 0
            print(f"    {f:<25} {filled}/{len(jobs)} ({pct}%)")

        para_count = sum(1 for j in jobs if j.get("jobTitle") != j.get("originalTitle"))
        print(f"\n  {C_LABEL('Paraphrased titles')} : {para_count}/{len(jobs)}")

    print(C_HEADER("=" * 72))


if __name__ == "__main__":
    craw()
