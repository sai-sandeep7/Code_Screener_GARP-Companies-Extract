import os
import time
import json
import random
import datetime
import pandas as pd
import requests
import re
from bs4 import BeautifulSoup
from collections import defaultdict
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    import openpyxl  # noqa: F401
except ModuleNotFoundError:
    openpyxl = None

BASE = "https://www.screener.in"

CACHE_FILE = "name_cache.json"
INDUSTRY_CACHE_FILE = "industry_cache.json"
METRICS_CACHE_FILE = "metrics_cache.json"
REQUEST_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"}
REQUEST_TIMEOUT = (15, 45)

# ==========================
# RUN CLOCK  (Asia/Kolkata)
# ==========================
# Everything date-related is evaluated in IST, not the machine's local zone,
# so the same scrape buckets identically whether it runs on a laptop in India
# or a UTC cloud box. India has no DST, so a fixed +05:30 offset is exact and
# needs no tzdata package (which Windows does not ship by default).
IST = datetime.timezone(datetime.timedelta(hours=5, minutes=30), "IST")

# Market hours in IST. Prices, P/E and market cap only move between
# MARKET_OPEN_HOUR and MARKET_CLOSE_HOUR; outside that window the source data
# is frozen until the next open. The cache is bucketed to match:
#
#   live   09:00-17:00  values change, so buckets are short (INTRADAY_SLOT_MINUTES)
#   closed 17:00-09:00  values are static, so one bucket spans the evening,
#                       midnight and the following morning
#
# Shorter slots mean fresher numbers and more requests; longer slots mean the
# opposite. 60 minutes keeps a crashed run resumable without ever serving
# prices more than an hour old during trading.
MARKET_OPEN_HOUR = 9
MARKET_CLOSE_HOUR = 17
INTRADAY_SLOT_MINUTES = 60


def ist_now():
    return datetime.datetime.now(IST)


def session_key(dt=None):
    """
    Bucket an IST timestamp according to whether the market is moving.

    During market hours the key advances every INTRADAY_SLOT_MINUTES:

        09:00 -> 2026-07-28#live-0900
        09:59 -> 2026-07-28#live-0900     (same slot, resume works)
        10:00 -> 2026-07-28#live-1000     (new slot, prices re-fetched)

    Outside market hours every timestamp maps to the close it follows, so an
    18:00 run, a 23:00 run and next morning's 07:00 run all share one bucket
    and none of them re-fetch numbers that cannot have changed:

        17:00 on the 28th -> 2026-07-28#closed
        23:59 on the 28th -> 2026-07-28#closed
        07:00 on the 29th -> 2026-07-28#closed   (still the 28th's close)

    Two runs share a metrics cache only when their keys match exactly.
    """
    dt = dt or ist_now()
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=IST)
    else:
        dt = dt.astimezone(IST)

    if MARKET_OPEN_HOUR <= dt.hour < MARKET_CLOSE_HOUR:
        minutes = dt.hour * 60 + dt.minute
        slot = (minutes // INTRADAY_SLOT_MINUTES) * INTRADAY_SLOT_MINUTES
        return f"{dt.date().isoformat()}#live-{slot // 60:02d}{slot % 60:02d}"

    # Before the open the relevant close is the previous calendar day's.
    close_date = dt.date()
    if dt.hour < MARKET_OPEN_HOUR:
        close_date -= datetime.timedelta(days=1)
    return f"{close_date.isoformat()}#closed"


def market_is_open(dt=None):
    dt = (dt or ist_now()).astimezone(IST)
    return MARKET_OPEN_HOUR <= dt.hour < MARKET_CLOSE_HOUR


# Stamped once at import so a run that crosses midnight (or a session
# boundary) stays internally consistent: cache stamps and the output filename
# all refer to the same logical run.
RUN_NOW = ist_now()
RUN_SESSION = session_key(RUN_NOW)              # 2026-07-28#live-0900 | #closed
RUN_MARKET_OPEN = market_is_open(RUN_NOW)
RUN_DATE = RUN_NOW.date()                       # IST calendar date
RUN_DATE_ISO = RUN_DATE.isoformat()             # 2026-07-28
RUN_DATE_FILE = RUN_DATE.strftime("%d_%m_%Y")   # 28_07_2026

# How long each cache stays valid, in days.
#   name      -> a company's official name effectively never changes
#   industry  -> reclassification is rare; refresh monthly
#   metrics   -> price / P-E / market cap change every trading day
NAME_CACHE_TTL_DAYS = None      # None = never expires
INDUSTRY_CACHE_TTL_DAYS = 30    # whole IST days
METRICS_CACHE_TTL_DAYS = 0      # 0 = same IST session only (see session_key)


def ensure_openpyxl_installed():
    if openpyxl is None:
        raise ImportError(
            "openpyxl is required to read/write Excel files. "
            "Install it using: python -m pip install openpyxl"
        )


# ==========================
# SHEET NAME SANITIZER
# ==========================
def sanitize_sheet_name(name, max_length=31):
    """
    Sanitize sheet name to be valid for Excel:
    - Max 31 characters
    - No invalid characters: [ ] : * ? / \\\n    - Replace invalid characters with spaces
    """
    if not name:
        return "Sheet"
    
    # Replace invalid characters with space
    invalid_chars = r'[\[\]:*?/\\]'
    sanitized = re.sub(invalid_chars, ' ', str(name))
    
    # Remove extra spaces
    sanitized = ' '.join(sanitized.split())
    
    # Truncate to max length
    if len(sanitized) > max_length:
        sanitized = sanitized[:max_length-3] + "..."
    
    return sanitized if sanitized else "Sheet"


# ==========================
# TIMER DECORATOR
# ==========================
def timer(func):
    def wrapper(*args, **kwargs):
        start_time = time.time()
        print(f"\n{'='*60}")
        print(f"STARTING: {func.__name__}")
        print(f"{'='*60}\n")
        
        result = func(*args, **kwargs)
        
        end_time = time.time()
        elapsed_time = end_time - start_time
        minutes = int(elapsed_time // 60)
        seconds = elapsed_time % 60
        
        print(f"\n{'='*60}")
        print(f"COMPLETED: {func.__name__}")
        print(f"Total Time: {minutes} minutes {seconds:.2f} seconds")
        print(f"{'='*60}\n")
        
        return result
    return wrapper


# ==========================
# CACHE HELPERS
# ==========================
def _load_stamped_cache(path, ttl_days, label, legacy_ok=False):
    """
    Load a cache file written by _save_stamped_cache.

    The file is
    {"_date": "YYYY-MM-DD", "_session": "YYYY-MM-DD#day", "_version": 3,
     "data": {...}}, with every timestamp in IST.

        ttl_days None -> never expires
        ttl_days 0    -> valid only within the same IST session: a short slot
                         during market hours, one long bucket after the close
                         (see session_key)
        ttl_days N    -> valid for N whole IST days

    Anything older is discarded and an empty cache returned, so the run
    re-fetches rather than serving stale values.

    Undated legacy files cannot have their true age established, so what
    happens to them is a per-cache decision via legacy_ok: adopt them under
    this run's stamp (fine for slow-moving data like names and industries,
    and it avoids re-fetching hundreds of companies once) or discard them
    (the only safe option for prices).
    """
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            blob = json.load(f)
    except (json.JSONDecodeError, OSError) as e:
        print(f"  ⚠ {label} cache unreadable ({e}); starting empty.")
        return {}

    if not isinstance(blob, dict):
        return {}

    # Legacy format: the dict IS the data, with no date stamp.
    if "_date" not in blob or "data" not in blob:
        if ttl_days is None or legacy_ok:
            print(f"  ↻ {label} cache: adopting legacy file ({len(blob)} entries kept, "
                  f"re-stamped for this run).")
            return blob
        print(f"  ↻ {label} cache: legacy file with no date stamp — discarding "
              f"{len(blob)} entries to guarantee freshness.")
        return {}

    data = blob.get("data") or {}
    if ttl_days is None:
        return data

    # Session-scoped caches (metrics): the stamp must match exactly.
    if ttl_days == 0:
        stamped_session = blob.get("_session")
        if not stamped_session:
            print(f"  ↻ {label} cache predates session tracking — discarding "
                  f"{len(data)} entries to guarantee freshness.")
            return {}
        if stamped_session != RUN_SESSION:
            print(f"  ↻ {label} cache is from session {stamped_session}, this run is "
                  f"{RUN_SESSION} — discarding {len(data)} entries.")
            return {}
        if data:
            print(f"  ✓ {label} cache hit: {len(data)} entries from this session "
                  f"({stamped_session}).")
        return data

    # Day-scoped caches (industry): compare IST calendar dates.
    try:
        stamped = datetime.date.fromisoformat(blob["_date"])
    except (ValueError, TypeError):
        return {}

    age = (RUN_DATE - stamped).days
    if age > ttl_days:
        print(f"  ↻ {label} cache is {age} day(s) old (TTL {ttl_days}) — "
              f"discarding {len(data)} entries.")
        return {}
    if data:
        print(f"  ✓ {label} cache hit: {len(data)} entries from {blob['_date']}.")
    return data


def _save_stamped_cache(path, cache, label):
    tmp = path + ".tmp"
    blob = {
        "_date": RUN_DATE_ISO,
        "_session": RUN_SESSION,
        "_stamped_at": RUN_NOW.isoformat(timespec="seconds"),
        "_version": 3,
        "data": cache,
    }
    try:
        # Write via a temp file so an interrupted run cannot leave a
        # half-written cache that fails to parse on the next attempt.
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(blob, f, indent=2, ensure_ascii=False)
        os.replace(tmp, path)
    except OSError as e:
        print(f"  ⚠ could not save {label} cache: {e}")


def load_cache():
    return _load_stamped_cache(CACHE_FILE, NAME_CACHE_TTL_DAYS, "name")


def save_cache(cache):
    _save_stamped_cache(CACHE_FILE, cache, "name")


def load_industry_cache():
    # legacy_ok: sector classification barely moves, so an undated file is
    # worth keeping rather than re-scraping every company once.
    return _load_stamped_cache(INDUSTRY_CACHE_FILE, INDUSTRY_CACHE_TTL_DAYS,
                               "industry", legacy_ok=True)


def save_industry_cache(cache):
    _save_stamped_cache(INDUSTRY_CACHE_FILE, cache, "industry")


def load_metrics_cache():
    """
    Metrics (price / P-E / market cap) are valid only inside one IST session.

    While the market is open (09:00-17:00) that is a short slot, so a run
    crashed at 10:10 and restarted at 10:40 resumes, but one restarted at
    14:00 re-fetches rather than publishing four-hour-old prices. After the
    close the values cannot move, so an 18:00 run and next morning's 07:00
    run share one bucket and skip thousands of pointless requests.
    """
    return _load_stamped_cache(METRICS_CACHE_FILE, METRICS_CACHE_TTL_DAYS, "metrics")


def save_metrics_cache(cache):
    _save_stamped_cache(METRICS_CACHE_FILE, cache, "metrics")


# ==========================
# METRICS CONVERSION HELPER
# ==========================
def convert_metric_to_number(value_str):
    """
    Convert metric strings like '1,23,456', '1.2k', '1.5M', '1.8B' to float numbers.
    Also handles 'Cr' (Crore) and other Indian notations.
    """
    if not value_str or value_str == "-" or value_str == "":
        return None
    
    # Remove commas and extra spaces
    value_str = str(value_str).strip().replace(',', '')
    
    # Handle Crore (Cr) notation
    if 'Cr' in value_str or 'cr' in value_str:
        value_str = value_str.replace('Cr', '').replace('cr', '').strip()
        try:
            return float(value_str) * 100  # Convert Cr to absolute number (1 Cr = 100)
        except ValueError:
            return None
    
    # Handle Lakh (L) notation
    if 'L' in value_str or 'lakh' in value_str.lower():
        value_str = value_str.replace('L', '').replace('lakh', '').replace('Lakh', '').strip()
        try:
            return float(value_str)  # Lakh to number (1 L = 1)
        except ValueError:
            return None
    
    # Handle K (Thousands)
    if value_str.endswith('k') or value_str.endswith('K'):
        value_str = value_str[:-1].strip()
        try:
            return float(value_str) * 1000
        except ValueError:
            return None
    
    # Handle M (Millions)
    if value_str.endswith('m') or value_str.endswith('M'):
        value_str = value_str[:-1].strip()
        try:
            return float(value_str) * 1000000
        except ValueError:
            return None
    
    # Handle B (Billions)
    if value_str.endswith('b') or value_str.endswith('B'):
        value_str = value_str[:-1].strip()
        try:
            return float(value_str) * 1000000000
        except ValueError:
            return None
    
    # Handle plain numbers with possible decimal
    try:
        return float(value_str)
    except ValueError:
        return None


# ==========================
# COMPANY PAGE PARSER
# ==========================
def extract_company_name_from_soup(soup, company_url):
    try:
        h1 = soup.select_one("h1")
        if h1:
            return h1.text.strip()
        title = soup.title.text if soup.title else ""
        return title.split("|")[0].strip()
    except Exception:
        return company_url.rstrip("/").split("/")[-1]


def extract_company_industry_info_from_soup(soup):
    industry_info = {
        "broad_industry": None,
        "industry": None
    }

    try:
        peer_section = soup.find("section", {"id": "peers"})
        if peer_section:
            industry_para = peer_section.find("p", class_="sub")
            if industry_para:
                links = industry_para.find_all("a")
                for link in links:
                    title = link.get("title", "")
                    if title == "Broad Industry":
                        industry_info["broad_industry"] = link.text.strip()
                    elif title == "Industry":
                        industry_info["industry"] = link.text.strip()

                    prev_icon = link.find_previous("i")
                    if prev_icon:
                        icon_class = prev_icon.get("class", [])
                        if "icon-industry" in icon_class and not industry_info["broad_industry"]:
                            industry_info["broad_industry"] = link.text.strip()
                        elif "icon-tools-1" in icon_class and not industry_info["industry"]:
                            industry_info["industry"] = link.text.strip()

        if not industry_info["industry"] or not industry_info["broad_industry"]:
            industry_link = soup.find("a", {"title": "Industry"})
            if industry_link:
                industry_info["industry"] = industry_link.text.strip()

            broad_industry_link = soup.find("a", {"title": "Broad Industry"})
            if broad_industry_link:
                industry_info["broad_industry"] = broad_industry_link.text.strip()

        if not industry_info["industry"] or industry_info["industry"] in ["None", "-", ""]:
            industry_info["industry"] = None
        if not industry_info["broad_industry"] or industry_info["broad_industry"] in ["None", "-", ""]:
            industry_info["broad_industry"] = None
    except Exception as e:
        print(f"  ⚠ Error extracting industry info from page: {str(e)}")

    return industry_info


def extract_company_metrics_from_soup(soup):
    metrics = {
        "Market Cap": None,
        "Stock P/E": None,
        "Current Price": None,
        "Market Cap_Num": None,
        "Stock P/E_Num": None,
        "Current Price_Num": None
    }

    try:
        top_ratios = soup.select_one("#top-ratios")
        if top_ratios:
            ratio_items = top_ratios.select("li")
            for item in ratio_items:
                name_elem = item.select_one(".name")
                value_elem = item.select_one(".value .number")
                if name_elem and value_elem:
                    name = name_elem.text.strip()
                    value = value_elem.text.strip()
                    if "Market Cap" in name:
                        metrics["Market Cap"] = value
                        metrics["Market Cap_Num"] = convert_metric_to_number(value)
                    elif "Stock P/E" in name:
                        metrics["Stock P/E"] = value
                        metrics["Stock P/E_Num"] = convert_metric_to_number(value)
                    elif "Current Price" in name:
                        metrics["Current Price"] = value
                        metrics["Current Price_Num"] = convert_metric_to_number(value)
    except Exception as e:
        print(f"  ⚠ Error extracting metrics from page: {str(e)}")

    return metrics


def fetch_with_retries(session, url, timeout=REQUEST_TIMEOUT, max_retries=6):
    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            response = session.get(url, headers=REQUEST_HEADERS, timeout=timeout)
            response.raise_for_status()
            return response
        except requests.RequestException as exc:
            last_error = exc
            if attempt >= max_retries:
                raise
            delay = min(20, 2 ** (attempt - 1) + random.uniform(0.3, 1.2))
            print(f"  ↻ Retry {attempt}/{max_retries} for {url} after error: {exc}")
            time.sleep(delay)
    if last_error is not None:
        raise last_error
    raise requests.RequestException(f"Unable to fetch {url}")


def get_company_page_details(company_url, page_cache=None):
    """Fetch a company page once and parse all fields into one structured record."""
    if page_cache is None:
        page_cache = {}

    if company_url in page_cache:
        return page_cache[company_url]

    fallback_name = company_url.rstrip("/").split("/")[-1]
    try:
        session = build_session()
        r = fetch_with_retries(session, company_url)
        soup = BeautifulSoup(r.text, "html.parser")
        details = {
            "name": extract_company_name_from_soup(soup, company_url),
            "industry_info": extract_company_industry_info_from_soup(soup),
            "metrics": extract_company_metrics_from_soup(soup),
        }
        page_cache[company_url] = details
        time.sleep(random.uniform(0.3, 0.8))
        return details
    except Exception as e:
        print(f"  ⚠ Error fetching company page {company_url}: {str(e)}")
        details = {
            "name": fallback_name,
            "industry_info": {"broad_industry": None, "industry": None},
            "metrics": {
                "Market Cap": None,
                "Stock P/E": None,
                "Current Price": None,
                "Market Cap_Num": None,
                "Stock P/E_Num": None,
                "Current Price_Num": None
            }
        }
        page_cache[company_url] = details
        return details


# ==========================
# NAME SCRAPER
# ==========================
def get_company_name(company_url, cache, page_cache=None):
    """Retrieve official displayed company name."""
    if company_url in cache:
        return cache[company_url]

    details = get_company_page_details(company_url, page_cache)
    name = details.get("name") or company_url.rstrip("/").split("/")[-1]
    cache[company_url] = name
    return name


# ==========================
# GET BROAD INDUSTRY AND INDUSTRY FROM COMPANY PAGE
# ==========================
def get_company_industry_info(company_url, cache, page_cache=None):
    """
    Extract both Broad Industry and Industry from company page.
    Returns dict with 'broad_industry' and 'industry' keys.
    """
    if company_url in cache:
        cached_value = cache[company_url]
        # Handle case where cache might have stored a string (older version)
        if isinstance(cached_value, str):
            return {"broad_industry": cached_value, "industry": None}
        return cached_value

    details = get_company_page_details(company_url, page_cache)
    industry_info = details.get("industry_info") or {
        "broad_industry": None,
        "industry": None
    }
    cache[company_url] = industry_info
    return industry_info


# ==========================
# EXTRACT MARKET CAP, P/E, AND CURRENT PRICE FROM COMPANY PAGE
# ==========================
def get_company_metrics(company_url, cache, page_cache=None):
    """
    Extract Market Cap, Stock P/E, and Current Price from the company page's top ratios section.
    Returns dict with raw string values and converted numeric values.
    """
    if company_url in cache:
        return cache[company_url]

    details = get_company_page_details(company_url, page_cache)
    metrics = details.get("metrics") or {
        "Market Cap": None,
        "Stock P/E": None,
        "Current Price": None,
        "Market Cap_Num": None,
        "Stock P/E_Num": None,
        "Current Price_Num": None
    }
    cache[company_url] = metrics
    return metrics


# ==========================
# EXTRACT SCREEN LINKS
# ==========================
def build_session():
    session = requests.Session()
    retries = Retry(
        total=6,
        connect=6,
        read=4,
        backoff_factor=2,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(REQUEST_HEADERS)
    return session


def extract_company_urls_from_screen(screen_url, session=None):
    """Extract company URLs from Screener screen with pagination using HTML, not tables."""
    all_urls = set()
    page = 1
    last_batch = set()
    last_error = None
    if session is None:
        session = build_session()

    while True:
        paged = f"{screen_url}?page={page}"
        print(f"  Fetching listing page {page}: {paged}")

        try:
            r = fetch_with_retries(session, paged, timeout=REQUEST_TIMEOUT)
            soup = BeautifulSoup(r.text, "html.parser")

            rows = soup.select("tbody tr a[href^='/company/']")
            urls = {BASE + a.get("href") for a in rows if "/company/" in a.get("href", "")}

            if not urls:
                print(f"  No valid rows on page {page}. Stopping.")
                break
            if urls == last_batch:
                print(f"  Duplicate data detected page {page}. Stopping.")
                break

            all_urls |= urls
            last_batch = urls
            page += 1
            time.sleep(random.uniform(0.6, 1.4))
        except Exception as e:
            last_error = str(e)
            print(f"  ⚠ Error fetching page {page}: {str(e)}")
            break

    return sorted(all_urls), last_error


# ==========================
# PROCESS SINGLE SCREEN
# ==========================
def process_screen(screen_id, url, name_cache=None, page_cache=None):
    print("\n====================================================")
    print(f"START → SCREEN {screen_id.upper()}")
    print(f"URL   → {url}")
    print("====================================================\n")

    # Extract all company page URLs
    company_urls, last_error = extract_company_urls_from_screen(url)
    if last_error:
        print(f"⚠ SCREEN {screen_id} COULD NOT BE FETCHED: {last_error}")
        return {"__error__": last_error}

    if not company_urls:
        print(f"NO COMPANIES FOUND FOR SCREEN {screen_id}")
        return None

    print(f"Total companies found: {len(company_urls)}\n")

    # Retrieve raw text from screen AND official names
    mapping_rows = []
    if name_cache is None:
        name_cache = load_cache()

    # LIVE MAPPING LOGS
    print(f"Mapping company names for {screen_id}...\n")

    for company_url in company_urls:
        raw_name = company_url.rstrip("/").split("/")[-1]
        official = get_company_name(company_url, name_cache, page_cache)

        print(f"  NEW → {raw_name} : {official}")

        mapping_rows.append({
            "Screen_Name": raw_name,
            "Official_Name": official,
            "URL": company_url
        })

    return {row["Official_Name"]: row["URL"] for row in mapping_rows}


# ==========================
# CONSOLIDATE ALL SCREENS
# ==========================
@timer
def consolidate_screens():
    """Consolidate data from all screens and create consolidated Excel with industry information"""
    SCREENS = {
        "p1_main": "https://www.screener.in/screens/3418971/p1-garp-early-fast-growers-pre-institutional/",
        "p2_main": "https://www.screener.in/screens/3419099/p2-garp-aggressive-10-bagger-hunters/",
        "p3_main": "https://www.screener.in/screens/3419001/p3-garp-core-lynch-compounders/",
        'p1_v1': "https://www.screener.in/screens/3529190/p1_updated/",
        'p2_v1': "https://www.screener.in/screens/3529191/p2_updated/",
        'p3_v1': "https://www.screener.in/screens/3529193/p3_updated/",
        'p1_v2': "https://www.screener.in/screens/3529201/p1_updated_v2/",
        'p2_v2': "https://www.screener.in/screens/3529202/p2_updated_v2/",
        'p3_v2': "https://www.screener.in/screens/3529204/p3_updated_v2/",
        'p1_v3': "https://www.screener.in/screens/3529206/p1_updated_v3/",
        'p2_v3': "https://www.screener.in/screens/3529208/p2_updated_v3/",
        'p3_v3': "https://www.screener.in/screens/3529209/p3_updated_v3/",
        'p1_v4': "https://www.screener.in/screens/3529211/p1_updated_v4/",
        'p2_v4': "https://www.screener.in/screens/3529213/p2_updated_v4/",
        'p3_v4': "https://www.screener.in/screens/3529214/p3_updated_v4/",
        'p1_v5': "https://www.screener.in/screens/3529216/p1_updated_v5/",
        'p2_v5': "https://www.screener.in/screens/3529217/p2_updated_v5/",
        'p3_v5': "https://www.screener.in/screens/3529219/p3_updated_v5/",
        'p1_v6': "https://www.screener.in/screens/3529222/p1_updated_v6/",
        'p2_v6': "https://www.screener.in/screens/3529223/p2_updated_v6/",
        'p3_v6': "https://www.screener.in/screens/3529224/p3_updated_v6/",
        'p1_v7': "https://www.screener.in/screens/3530112/p1_v7/",
        'p2_v7': "https://www.screener.in/screens/3530113/p2_v7/",
        'p3_v7': "https://www.screener.in/screens/3530114/p3_v7/",
        'Pure PEG 5Y Screener': 'https://www.screener.in/screens/3582942/pure-peg-5y-screener/',
        'Temporary Slowdown Capture': 'https://www.screener.in/screens/3582944/temporary-slowdown-capture/',
        'PEG 5Y Deep Value': 'https://www.screener.in/screens/3582947/peg-5y-deep-value/',
        'Quality compounders only': 'https://www.screener.in/screens/3582950/quality-compounders-only/',
        'Dual PEG Filter':'https://www.screener.in/screens/3582951/dual-peg-filter/',
        
        "Earliest_growth_discovery_1": "https://www.screener.in/screens/3535089/earliest_growth_discovery_1/",
        "Emerging_growth_2": "https://www.screener.in/screens/3535086/emerging_growth_2/",
        "Early_fast_growers_3": "https://www.screener.in/screens/3535098/early_fast_growers_3/",
        "Very_fast_growers_4": "https://www.screener.in/screens/3534227/very_fast_growers_4/",
        "transition_growth_5": "https://www.screener.in/screens/3535085/transition_growth_5/",
        "Matured_undervalued_growth_6": "https://www.screener.in/screens/3535096/matured_undervalued_growth_6/",
        "Shankar_Nath_Stockscans": "https://www.screener.in/screens/3548140/shankar-nath_stockscans/"
    }

    # Group screens by type
    screen_groups = {
        "main": ["p1_main", "p2_main", "p3_main"],
        "v1": ["p1_v1", "p2_v1", "p3_v1"],
        "v2": ["p1_v2", "p2_v2", "p3_v2"],
        "v3": ["p1_v3", "p2_v3", "p3_v3"],
        "v4": ["p1_v4", "p2_v4", "p3_v4"],
        "v5": ["p1_v5", "p2_v5", "p3_v5"],
        "v6": ["p1_v6", "p2_v6", "p3_v6"],
        "v7": ["p1_v7", "p2_v7", "p3_v7"],
        # Add the new screens as individual groups
        "pure_peg": ["Pure PEG 5Y Screener"],
        "slowdown_capture": ["Temporary Slowdown Capture"],
        "peg_deep_value": ["PEG 5Y Deep Value"],
        "quality_compounders": ["Quality compounders only"],
        "dual_peg": ["Dual PEG Filter"],
        "earliest_growth": ["Earliest_growth_discovery_1"],
        "emerging_growth": ["Emerging_growth_2"],
        "early_fast_growers": ["Early_fast_growers_3"],
        "very_fast_growers": ["Very_fast_growers_4"],
        "transition_growth": ["transition_growth_5"],
        "matured_undervalued_growth": ["Matured_undervalued_growth_6"],
        "shankar_nath_stockscans": ["Shankar_Nath_Stockscans"]
    }

    all_companies = set()
    screen_data = {}  # Store companies for each screen with their URLs
    screen_mapping = defaultdict(set)  # Store companies by screen group
    screen_errors = {}
    name_cache = load_cache()
    page_cache = {}

    print("\n" + "="*60)
    print("CONSOLIDATING DATA FROM ALL SCREENS")
    print("="*60 + "\n")

    # Extract companies from each screen
    for sid, link in SCREENS.items():
        print(f"\n{'─'*50}")
        print(f"Processing Screen: {sid.upper()}")
        print(f"URL: {link}")
        print(f"{'─'*50}\n")
        
        screen_companies_dict = process_screen(sid, link, name_cache=name_cache, page_cache=page_cache)
        if isinstance(screen_companies_dict, dict) and "__error__" in screen_companies_dict:
            screen_errors[sid] = screen_companies_dict["__error__"]
            print(f"⚠ Skipping {sid} due to fetch failure and continuing with the remaining screens.")
            continue
        if screen_companies_dict:
            screen_data[sid] = screen_companies_dict
            companies_set = set(screen_companies_dict.keys())
            all_companies |= companies_set
            
            # Add to appropriate group
            for group_name, screen_list in screen_groups.items():
                if sid in screen_list:
                    screen_mapping[group_name] |= companies_set
                    break

    save_cache(name_cache)

    print(f"\n{'='*60}")
    print(f"Total unique companies across all screens: {len(all_companies)}")
    print(f"{'='*60}\n")

    # Build reverse mapping for all company URLs (combining from all screens)
    all_company_urls = {}
    for sid, companies_dict in screen_data.items():
        for official_name, url in companies_dict.items():
            if official_name not in all_company_urls:
                all_company_urls[official_name] = url

    # Extract Market Cap, P/E, and Current Price for all companies
    print("Extracting Market Cap, P/E, and Current Price for all companies...\n")
    metrics_cache = load_metrics_cache()
    company_metrics = {}
    fetched = resumed = 0

    for i, (company, url) in enumerate(all_company_urls.items(), 1):
        was_cached = url in metrics_cache
        if was_cached:
            resumed += 1
        else:
            fetched += 1
            print(f"  [{i}/{len(all_company_urls)}] Processing: {company}")
        metrics = get_company_metrics(url, metrics_cache, page_cache)
        company_metrics[company] = metrics
        if i % 10 == 0:  # Save cache every 10 companies
            save_metrics_cache(metrics_cache)

    save_metrics_cache(metrics_cache)
    state = "market OPEN" if RUN_MARKET_OPEN else "market CLOSED"
    print(f"\n  Metrics: {fetched} fetched live, {resumed} reused from the "
          f"{RUN_SESSION} cache ({state}).")
    if fetched == 0 and resumed:
        if RUN_MARKET_OPEN:
            print("  ⚠ Nothing fetched live — values come from earlier in this "
                  f"{INTRADAY_SLOT_MINUTES}-minute slot and may be slightly behind.")
        else:
            print("  ℹ Nothing fetched live — the market is closed, so these values "
                  "cannot have moved since the last run.")

    # Extract industry information for all companies
    print("\nExtracting Broad Industry and Industry for all companies...\n")
    industry_cache = load_industry_cache()
    company_industry_info = {}
    
    for i, (company, url) in enumerate(all_company_urls.items(), 1):
        print(f"  [{i}/{len(all_company_urls)}] Processing: {company}")
        industry_info = get_company_industry_info(url, industry_cache, page_cache)
        company_industry_info[company] = industry_info
        if i % 10 == 0:  # Save cache every 10 companies
            save_industry_cache(industry_cache)
    
    save_industry_cache(industry_cache)

    # Group companies by broad industry for separate sheets
    broad_industry_data = defaultdict(list)
    # Group companies by specific industry for separate sheets
    industry_data = defaultdict(list)
    
    for company, info in company_industry_info.items():
        # Ensure info is a dictionary
        if not isinstance(info, dict):
            info = {"broad_industry": str(info) if info else None, "industry": None}
        
        broad_industry = info.get("broad_industry")
        industry = info.get("industry")
        
        if broad_industry and broad_industry not in ["None", "-", ""]:
            broad_industry_data[broad_industry].append(company)
        else:
            broad_industry_data["Unclassified"].append(company)
        
        if industry and industry not in ["None", "-", ""]:
            industry_data[industry].append(company)
        else:
            industry_data["Unclassified"].append(company)

    # Save consolidated Excel (RUN_DATE, not today(), so a run that crosses
    # midnight still writes the file its cache was stamped for)
    fname = f"CONSOLIDATED_ALL_SCREENS_{RUN_DATE_FILE}.xlsx"

    print(f"\n{'='*60}")
    print(f"Writing consolidated results to: {fname}")
    print(f"Run session: {RUN_SESSION}  ({RUN_NOW.strftime('%d %b %Y %H:%M')} IST, "
          f"market {'open' if RUN_MARKET_OPEN else 'closed'})")
    print(f"{'='*60}\n")

    ensure_openpyxl_installed()
    with pd.ExcelWriter(fname, engine="openpyxl") as writer:
        # Summary sheet
        summary_data = []

        if screen_errors:
            summary_data.append({
                "Screen Group": "UNREACHABLE SCREENS",
                "Company Count": len(screen_errors),
                "Type": "Screen Error"
            })
        
        # Add screen group summaries
        for group_name, companies in screen_mapping.items():
            summary_data.append({
                "Screen Group": group_name.upper(),
                "Company Count": len(companies),
                "Type": "Screen Group"
            })
        
        # Add broad industry summaries
        for broad_industry, companies in sorted(broad_industry_data.items()):
            summary_data.append({
                "Screen Group": broad_industry,
                "Company Count": len(companies),
                "Type": "Broad Industry"
            })
        
        # Add specific industry summaries
        for industry, companies in sorted(industry_data.items()):
            summary_data.append({
                "Screen Group": industry,
                "Company Count": len(companies),
                "Type": "Specific Industry"
            })
        
        if summary_data:
            df_summary = pd.DataFrame(summary_data)
            df_summary = df_summary.sort_values(["Type", "Screen Group"])
            df_summary.to_excel(writer, sheet_name="Summary", index=False)
            print(f"✓ Summary sheet created")

        # Individual screen sheets with Market Cap, P/E, Current Price, and Industry info
        print("\nCreating individual screen sheets...")
        for sid, companies_dict in screen_data.items():
            if companies_dict:
                data = []
                for company_name in sorted(companies_dict.keys()):
                    metrics = company_metrics.get(company_name, {})
                    industry_info = company_industry_info.get(company_name, {})
                    
                    # Ensure industry_info is a dict
                    if not isinstance(industry_info, dict):
                        industry_info = {"broad_industry": str(industry_info) if industry_info else None, 
                                       "industry": None}
                    
                    broad_industry = industry_info.get("broad_industry")
                    industry = industry_info.get("industry")
                    
                    # Convert None to "Unclassified"
                    broad_industry = broad_industry if broad_industry and broad_industry not in ["None", "-"] else "Unclassified"
                    industry = industry if industry and industry not in ["None", "-"] else "Unclassified"
                    
                    data.append({
                        "Company Name": company_name,
                        "Broad Industry": broad_industry,
                        "Industry": industry,
                        "Market Cap (₹ Cr)": metrics.get("Market Cap"),
                        "Stock P/E": metrics.get("Stock P/E"),
                        "Current Price (₹)": metrics.get("Current Price")
                    })
                
                df = pd.DataFrame(data)
                sheet_name = sanitize_sheet_name(sid)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  ✓ {sheet_name}: {len(df)} entries")

        # Screen group sheets
        print("\nCreating screen group sheets...")
        for group_name, companies in screen_mapping.items():
            if companies:
                data = []
                for company_name in sorted(companies):
                    metrics = company_metrics.get(company_name, {})
                    industry_info = company_industry_info.get(company_name, {})
                    
                    # Ensure industry_info is a dict
                    if not isinstance(industry_info, dict):
                        industry_info = {"broad_industry": str(industry_info) if industry_info else None, 
                                       "industry": None}
                    
                    broad_industry = industry_info.get("broad_industry")
                    industry = industry_info.get("industry")
                    
                    # Convert None to "Unclassified"
                    broad_industry = broad_industry if broad_industry and broad_industry not in ["None", "-"] else "Unclassified"
                    industry = industry if industry and industry not in ["None", "-"] else "Unclassified"
                    
                    data.append({
                        "Company Name": company_name,
                        "Broad Industry": broad_industry,
                        "Industry": industry,
                        "Market Cap (₹ Cr)": metrics.get("Market Cap"),
                        "Stock P/E": metrics.get("Stock P/E"),
                        "Current Price (₹)": metrics.get("Current Price")
                    })
                
                df = pd.DataFrame(data)
                sheet_name = sanitize_sheet_name(f"{group_name}_combined")
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  ✓ {sheet_name}: {len(df)} entries")

        # All companies sheet
        print("\nCreating all companies sheet...")
        all_data = []
        for company in sorted(all_companies):
            metrics = company_metrics.get(company, {})
            industry_info = company_industry_info.get(company, {})
            
            # Ensure industry_info is a dict
            if not isinstance(industry_info, dict):
                industry_info = {"broad_industry": str(industry_info) if industry_info else None, 
                               "industry": None}
            
            broad_industry = industry_info.get("broad_industry")
            industry = industry_info.get("industry")
            
            # Convert None to "Unclassified"
            broad_industry = broad_industry if broad_industry and broad_industry not in ["None", "-"] else "Unclassified"
            industry = industry if industry and industry not in ["None", "-"] else "Unclassified"
            
            # Find which screen groups this company appears in
            appears_in = []
            for group_name, group_companies in screen_mapping.items():
                if company in group_companies:
                    appears_in.append(group_name)
            
            all_data.append({
                "Company Name": company,
                "Broad Industry": broad_industry,
                "Industry": industry,
                "Market Cap (₹ Cr)": metrics.get("Market Cap"),
                "Stock P/E": metrics.get("Stock P/E"),
                "Current Price (₹)": metrics.get("Current Price"),
                "Appears In": ", ".join(appears_in) if appears_in else "Unknown"
            })
        
        df_all = pd.DataFrame(all_data)
        df_all.to_excel(writer, sheet_name="All Companies", index=False)
        
        # Count unclassified companies
        unclassified_broad = sum(1 for row in all_data if row["Broad Industry"] == "Unclassified")
        unclassified_industry = sum(1 for row in all_data if row["Industry"] == "Unclassified")
        print(f"  ✓ All Companies sheet: {len(df_all)} entries")
        if unclassified_broad > 0:
            print(f"    ⚠ Companies with missing broad industry: {unclassified_broad}")
        if unclassified_industry > 0:
            print(f"    ⚠ Companies with missing industry: {unclassified_industry}")

        # Failed screens sheet
        if screen_errors:
            print("\nCreating failed screens sheet...")
            failed_screen_rows = [
                {"Screen ID": sid, "URL": SCREENS[sid], "Error": error}
                for sid, error in sorted(screen_errors.items())
            ]
            pd.DataFrame(failed_screen_rows).to_excel(writer, sheet_name="Failed Screens", index=False)
            print(f"  ✓ Failed Screens sheet: {len(failed_screen_rows)} entries")

        # Broad Industry sheets
        print("\nCreating broad industry sheets...")
        for broad_industry, companies in sorted(broad_industry_data.items()):
            if companies:
                data = []
                for company_name in sorted(companies):
                    metrics = company_metrics.get(company_name, {})
                    industry_info = company_industry_info.get(company_name, {})
                    
                    # Ensure industry_info is a dict
                    if not isinstance(industry_info, dict):
                        industry_info = {"broad_industry": broad_industry, "industry": None}
                    
                    industry = industry_info.get("industry")
                    industry = industry if industry and industry not in ["None", "-"] else "Unclassified"
                    
                    appears = []
                    for group_name, group_companies in screen_mapping.items():
                        if company_name in group_companies:
                            appears.append(group_name)
                    
                    data.append({
                        "Company Name": company_name,
                        "Industry": industry,
                        "Market Cap (₹ Cr)": metrics.get("Market Cap"),
                        "Stock P/E": metrics.get("Stock P/E"),
                        "Current Price (₹)": metrics.get("Current Price"),
                        "Appears In": ", ".join(appears) if appears else "Unknown"
                    })
                
                df_with_metrics = pd.DataFrame(data)
                
                # Create sanitized sheet name
                sheet_name = sanitize_sheet_name(f"Broad_{broad_industry}")
                df_with_metrics.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  ✓ {sheet_name}: {len(companies)} entries")

        # Specific Industry sheets (NEW)
        print("\nCreating specific industry sheets...")
        for industry, companies in sorted(industry_data.items()):
            if companies and industry != "Unclassified":  # Skip Unclassified for specific industry sheets
                data = []
                for company_name in sorted(companies):
                    metrics = company_metrics.get(company_name, {})
                    industry_info = company_industry_info.get(company_name, {})
                    
                    # Ensure industry_info is a dict
                    if not isinstance(industry_info, dict):
                        industry_info = {"broad_industry": None, "industry": industry}
                    
                    broad_industry = industry_info.get("broad_industry")
                    broad_industry = broad_industry if broad_industry and broad_industry not in ["None", "-"] else "Unclassified"
                    
                    appears = []
                    for group_name, group_companies in screen_mapping.items():
                        if company_name in group_companies:
                            appears.append(group_name)
                    
                    data.append({
                        "Company Name": company_name,
                        "Broad Industry": broad_industry,
                        "Market Cap (₹ Cr)": metrics.get("Market Cap"),
                        "Stock P/E": metrics.get("Stock P/E"),
                        "Current Price (₹)": metrics.get("Current Price"),
                        "Appears In": ", ".join(appears) if appears else "Unknown"
                    })
                
                df_with_metrics = pd.DataFrame(data)
                
                # Create sanitized sheet name
                sheet_name = sanitize_sheet_name(f"Industry_{industry}")
                df_with_metrics.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  ✓ {sheet_name}: {len(companies)} entries")
        
        # Numeric data sheet for analysis
        print("\nCreating numeric data sheet for analysis...")
        numeric_data = []
        for company in sorted(all_companies):
            metrics = company_metrics.get(company, {})
            industry_info = company_industry_info.get(company, {})
            
            # Ensure industry_info is a dict
            if not isinstance(industry_info, dict):
                industry_info = {"broad_industry": str(industry_info) if industry_info else None, 
                               "industry": None}
            
            broad_industry = industry_info.get("broad_industry")
            industry = industry_info.get("industry")
            
            # Convert None to "Unclassified"
            broad_industry = broad_industry if broad_industry and broad_industry not in ["None", "-"] else "Unclassified"
            industry = industry if industry and industry not in ["None", "-"] else "Unclassified"
            
            numeric_data.append({
                "Company Name": company,
                "Broad Industry": broad_industry,
                "Industry": industry,
                "Market Cap (Cr) - Numeric": metrics.get("Market Cap_Num"),
                "Stock P/E - Numeric": metrics.get("Stock P/E_Num"),
                "Current Price (₹) - Numeric": metrics.get("Current Price_Num")
            })
        
        df_numeric = pd.DataFrame(numeric_data)
        df_numeric.to_excel(writer, sheet_name="Numeric Data", index=False)
        print(f"  ✓ Numeric Data sheet: {len(df_numeric)} entries")

    print(f"\n{'='*60}")
    print(f"✓ CONSOLIDATION COMPLETE!")
    print(f"✓ File saved: {fname}")
    print(f"✓ Total unique companies: {len(all_companies)}")
    print(f"✓ Individual screens processed: {len(screen_data)}")
    print(f"✓ Screen groups created: {len(screen_mapping)}")
    print(f"✓ Broad Industry categories: {len([k for k in broad_industry_data.keys() if k != 'Unclassified'])}")
    print(f"✓ Specific Industry categories: {len([k for k in industry_data.keys() if k != 'Unclassified'])}")
    print(f"{'='*60}\n")
    
    return fname


# ==========================
# MAIN ENTRY
# ==========================
if __name__ == "__main__":
    consolidate_screens()