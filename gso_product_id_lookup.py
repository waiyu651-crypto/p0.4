"""
GSO Product-ID -> GCTS/NB lookup
=================================

Takes a list of known Product IDs (from an existing Excel export, e.g.
GSO_LV_merged.xlsx) and, for each one, searches the GSO search bar/API
using the Product ID as the search term to find its GCTS Number and
Notified Body (NB) Number.

This is a different workflow from the NB-completeness scraper
(gso_scraper.py): here we already know exactly which products we care
about (by internal Product ID), so we do a direct, targeted lookup per ID
instead of trying to reconstruct a full catalog.

Only the GCTS No is required. Two things were simplified accordingly:
  - No NB-name lookup. That was a second HTTP request per distinct NB
    number (to a public detail page) purely to resolve a human-readable
    org name -- not needed to get the GCTS No, so it's gone. NB No is
    still included in the output since it comes for free out of the same
    match (it's just the digits before the dash in the GCTS No), but no
    extra request is made for it.
  - Early exit per search. Once a product's search results contain a
    matching productId, the search stops right there instead of
    continuing to fetch further pages of that same query. Pagination
    only continues if the match hasn't shown up yet.

IMPORTANT ASSUMPTION -- please validate before a full run
-----------------------------------------------------------
This script assumes that typing a raw numeric Product ID into the GSO
search bar (the same 'term'/'q' params used by gso_scraper.py) reliably
surfaces the matching record. That's a reasonable read of what you asked
for, but it's not something I can verify myself (no network access to
gso.org.sa from where I run, and it requires your logged-in session).
Typeahead search boxes sometimes index only text fields (product name,
brand, GCTS number) and not an internal numeric ID -- if that's the case
here, searches would come back empty even for valid IDs.

That's exactly why TEST_LIMIT exists below: run with e.g. TEST_LIMIT = 10
first, look at the printed "Found"/"Not Found" results for those 10 rows,
and confirm they look right (spot-check a couple against the website
manually) before setting TEST_LIMIT = None and committing to the full
53,235-row run, which will take several hours.

Scale / runtime
----------------
With ~53,000 rows, this used to be measured in HOURS because every
request went through Playwright one at a time with a fixed delay
between each. That's been changed: Playwright is now used ONLY to log
in and obtain a valid session (cookies). Once that session exists, the
actual GCTS lookups are plain HTTP requests (via `requests`) fired
concurrently across a small thread pool (CONCURRENCY workers), which is
both much lighter per-request and lets many lookups be in flight at
once. See the printed estimate at startup for the new expected runtime.

This script is still checkpointed exactly as before: progress is saved
to PROGRESS_FILE after every SAVE_EVERY_N *completed* rows, and on the
next run any Product ID already recorded there is skipped
automatically. You can safely Ctrl+C this at any point and re-run it
later to pick up where it left off (results collected so far are also
saved on Ctrl+C / any crash, via the try/finally in main()). Note that
with concurrency, rows complete out of input order -- that's expected
and harmless, since the final output is always rebuilt from `progress`
in the original input order.
"""

from playwright.sync_api import sync_playwright
import pandas as pd
import re
import os
import json
import time
import threading
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed

# =========================
# CONFIGURATION
# =========================
INPUT_FILE = "GSO_Batch2_Input_filled.xlsx"      # Must sit next to this script, or set a full path
INPUT_SHEET = "Model Lookup"
PRODUCT_ID_COLUMN = "Product GCTS"       # Column in the input file holding the Product ID

OUTPUT_FILE = "GSO_batch2_merged_with_gcts.xlsx"
PROGRESS_FILE = "gcts_lookup_progress.json"   # Resumable checkpoint -- do not delete mid-run
STATE_FILE = "gso_state.json"                 # Login session (shared format with gso_scraper.py)

LOGIN_URL = "https://www.gso.org.sa/gcts/"
API_URL = "https://www.gso.org.sa/gcts/api/lookups/published-products"
PAGE_SIZE = 100

# Small cap: an exact-ID search should return very few genuine candidates.
# We still cap it (rather than trusting it can't misbehave) in case the
# search term is fuzzy-matching against something unexpected and returning
# many pages of noise.
SEARCH_MAX_PAGES = 3

# --- Concurrency ---
# How many lookups run in parallel. Each "worker" is just a Python thread
# doing a plain HTTP GET (via `requests`), so this is cheap -- it's not
# spinning up multiple browsers. Start conservative and raise it if the
# server holds up fine (watch for 429s in the output). 10-20 is a
# reasonable range for a typeahead-style search endpoint; going much
# higher risks the server rate-limiting or blocking you.
CONCURRENCY = 12

# --- Politeness / rate limiting ---
# Applied per-thread before each request (not globally), so actual
# throughput is roughly CONCURRENCY / (REQUEST_DELAY_SECONDS + latency).
REQUEST_DELAY_SECONDS = 0.3

# --- Retry behavior for transient failures ---
MAX_RETRIES = 3
RETRY_BACKOFF_BASE_SECONDS = 1.5

# --- Checkpointing ---
SAVE_EVERY_N = 100          # Write PROGRESS_FILE (cheap, JSON) every N completed rows
XLSX_SNAPSHOT_EVERY_N = 2000  # Rebuild the .xlsx output every N completed rows (more expensive)

# --- Test / dry-run mode ---
# Set to a small number (e.g. 10) to only process the first N not-yet-done
# rows this run, so you can sanity check results before committing to the
# full file. Set to None to process everything.
# NOTE: the initial 10-row validation batch already came back "Found" for
# all 10 (see gcts_lookup_progress.json), so this now defaults to None --
# set it back to a small number if you change anything and want to
# re-validate before a full run.
TEST_LIMIT = None

# --- Web-app overrides (used by app.py; harmless for normal CLI use) ---
if os.environ.get("WEBAPP_TEST_LIMIT"):
    TEST_LIMIT = int(os.environ["WEBAPP_TEST_LIMIT"])
if os.environ.get("WEBAPP_INPUT_FILE"):
    INPUT_FILE = os.environ["WEBAPP_INPUT_FILE"]
if os.environ.get("WEBAPP_OUTPUT_FILE"):
    OUTPUT_FILE = os.environ["WEBAPP_OUTPUT_FILE"]
if os.environ.get("WEBAPP_STATE_FILE"):
    STATE_FILE = os.environ["WEBAPP_STATE_FILE"]

# If True, rows whose last recorded status wasn't "Found" get retried this
# run (e.g. after fixing a bug, or if the server was flaky earlier). If
# False, any row already present in PROGRESS_FILE -- found or not -- is
# skipped.
RETRY_NOT_FOUND = False

DEBUG = True
DEBUG_SAMPLE_CHARS = 800
DEBUG_MAX_SAMPLES = 3  # only dump raw response bodies for the first N searches, to avoid log spam


# =========================
# Shared helpers (same behavior as gso_scraper.py)
# =========================

def get_browser_and_context(playwright):
    if os.path.exists(STATE_FILE):
        print("✅ Saved session found. Trying headless mode...")
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context(storage_state=STATE_FILE)
        try:
            check_session_valid(context)
            print("✅ Saved session is valid.")
            return browser, context
        except RuntimeError as e:
            print(f"⚠️  Saved session is no longer valid ({e}).")
            print("   🔄 Discarding it and switching to manual re-login...")
            try:
                context.close()
            except Exception:
                pass
            browser.close()
            try:
                os.remove(STATE_FILE)
            except OSError:
                pass

    return manual_login_flow(playwright)


def manual_login_flow(playwright):
    print("🔐 Manual login required.")
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto(LOGIN_URL)
    print("👉 A browser window has opened. Please log in manually (click Login,")
    print("   sign in with your GSO/GIDP account, and wait for the page to load).")

    while True:
        input("⏳ Press ENTER here after you have successfully logged in...")
        print("   🔎 Verifying login against the GSO API...")
        try:
            check_session_valid(context)
            print("   ✅ Login verified.")
            break
        except RuntimeError as e:
            print(f"   ⚠️  Not logged in yet ({e})")
            print("   Please make sure you're fully logged in in the browser window, then press ENTER again.")

    context.storage_state(path=STATE_FILE)
    print(f"💾 Session verified and saved to {STATE_FILE}. Future runs will be headless.")
    return browser, context


def check_session_valid(context, sample_term="test"):
    try:
        response = context.request.get(
            API_URL,
            params={"term": sample_term, "q": sample_term, "_type": "query",
                    "pagesize": 1, "page": 0}
        )
    except Exception as e:
        raise RuntimeError(f"Could not reach GSO API at all: {e}")

    if response.status in (401, 403):
        raise RuntimeError(f"not authenticated (HTTP {response.status})")
    if response.status != 200:
        raise RuntimeError(
            f"Unexpected status {response.status} on sanity check. "
            f"Response body: {response.text()[:500]}"
        )

    try:
        response.json()
    except Exception:
        raise RuntimeError(
            "Sanity check got HTTP 200 but the body isn't JSON "
            "(likely a login/redirect page). Session is probably expired."
        )


def extract_items_from_response(data):
    if isinstance(data, dict):
        for key in ["results", "content", "data", "items", "list"]:
            if key in data and isinstance(data[key], list):
                return data[key], data.get("total") or data.get("totalCount") or data.get("totalElements") or 0
        for value in data.values():
            if isinstance(value, list):
                return value, data.get("total") or data.get("totalCount") or data.get("totalElements") or 0
        return [], 0
    elif isinstance(data, list):
        return data, len(data)
    else:
        return [], 0


def _get_with_retries(session, url, params, max_retries=MAX_RETRIES):
    """
    Same retry/backoff behavior as before, but using a `requests.Session`
    instead of a Playwright `context.request`. Safe to call from multiple
    threads at once as long as each thread has its own Session (see
    get_thread_session below) -- requests.Session objects are not
    guaranteed thread-safe to share.
    """
    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            response = session.get(url, params=params, timeout=30)
        except Exception as e:
            last_error = e
        else:
            if response.status_code not in (429, 500, 502, 503, 504):
                return response
            last_error = Exception(f"HTTP {response.status_code} (attempt {attempt}/{max_retries})")

        if attempt < max_retries:
            backoff = RETRY_BACKOFF_BASE_SECONDS * (2 ** (attempt - 1))
            print(f"      ⚠️  Request failed ({last_error}), retrying in {backoff:.1f}s "
                  f"[{attempt}/{max_retries}]...")
            time.sleep(backoff)

    raise Exception(f"Request to {url} failed after {max_retries} attempts: {last_error}")


# =========================
# Session handling for concurrent HTTP requests
# =========================
# Playwright's sync API is single-threaded -- it can't be safely driven
# from multiple worker threads at once. So instead of keeping the browser
# open for the whole run, we use it only to log in / validate the
# session, then lift the cookies out and make plain `requests` calls from
# a thread pool. Each thread gets its own Session (built from the same
# cookies) since Session objects aren't guaranteed thread-safe to share.

_thread_local = threading.local()
_shared_cookies = None  # list of Playwright-style cookie dicts, set once in main()

DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)


def build_session_from_cookies(cookies):
    s = requests.Session()
    for c in cookies:
        domain = (c.get("domain") or "").lstrip(".")
        s.cookies.set(
            c.get("name"), c.get("value"),
            domain=domain, path=c.get("path", "/"),
        )
    s.headers.update({"User-Agent": DEFAULT_USER_AGENT})
    return s


def get_thread_session():
    """Returns this thread's own requests.Session, creating it on first use."""
    if not hasattr(_thread_local, "session"):
        _thread_local.session = build_session_from_cookies(_shared_cookies)
    return _thread_local.session


_debug_samples_shown = 0  # module-level counter so DEBUG doesn't spam across thousands of searches


def normalize(s):
    return re.sub(r"\s+", "", str(s).strip().upper())


def normalize_loose(s):
    n = normalize(s)
    if n.isdigit():
        n = n.lstrip("0") or "0"
    return n


def extract_gcts_and_nb(item):
    """
    GCTS = item's gctsNo field. NB = the 4-digit prefix before the dash in
    gctsNo (e.g. '0025-4233' -> NB '0025'). Confirmed real shape:
      {"gctsNo": "0025-4233", "productId": 77054, "product": "Steam Iron"}
    Note the numeric suffix after the dash (4233) is NOT the same as
    productId (77054) -- they're unrelated fields. Matching on productId
    below is what actually confirms an item belongs to the ID we searched
    for; the gctsNo text itself never contains the productId.
    """
    is_dict = isinstance(item, dict)
    gcts = (item.get("gctsNo") or item.get("gctsNumber") or item.get("certificateNo") or "") if is_dict else ""

    nb = ""
    if is_dict:
        candidate = item.get("nbNumber") or item.get("notifiedBody") or item.get("nbNo") or ""
        candidate_str = str(candidate).strip()
        if re.fullmatch(r"\d{2,6}", candidate_str):
            nb = candidate_str

    if not nb and gcts:
        match = re.search(r"\b(\d{4})-\d+\b", str(gcts))
        if match:
            nb = match.group(1)

    return gcts, nb


def _safe_int(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


# =========================
# Core lookup logic
# =========================

def search_product_by_id(session, product_id, max_pages=SEARCH_MAX_PAGES):
    """
    Searches the GSO search bar/API using the raw Product ID as the term,
    fetching one page at a time and validating each candidate by its
    actual 'productId' field (so a fuzzy/typeahead match on unrelated text
    can't be mistaken for a real hit).

    Stops as soon as a match is found on the current page -- it does NOT
    keep fetching further pages of that search once we have what we need.
    Only searches that genuinely have no match yet (or an empty/short
    page) continue on to the next page, up to max_pages.

    `session` is a per-thread requests.Session (see get_thread_session) --
    this function is safe to call concurrently from many threads as long
    as each thread passes its own session.

    Returns a dict:
      status: "Found" | "Not Found" | "Ambiguous"
      gcts, nb: strings (blank if not found)
      raw_candidates: total items seen across all pages actually fetched
                      for this product (stops accumulating once a match
                      is found, so this reflects only what was fetched)
      matched_candidates: how many matched productId on the page where a match was found
    """
    global _debug_samples_shown

    search_term = str(product_id)
    raw_candidates = 0
    page_num = 0
    pages_fetched = 0

    while True:
        time.sleep(REQUEST_DELAY_SECONDS)  # per-thread pacing before each request

        response = _get_with_retries(
            session,
            API_URL,
            params={
                "term": search_term,
                "q": search_term,
                "_type": "query",
                "pagesize": PAGE_SIZE,
                "page": page_num
            }
        )

        if response.status_code != 200:
            raise Exception(f"API returned {response.status_code}")

        data = response.json()
        pages_fetched += 1

        if DEBUG and page_num == 0 and _debug_samples_shown < DEBUG_MAX_SAMPLES:
            print(f"      🐛 [DEBUG] Raw response for term='{search_term}': "
                  f"{json.dumps(data, indent=2)[:DEBUG_SAMPLE_CHARS]}")
            _debug_samples_shown += 1

        items, total = extract_items_from_response(data)
        if not isinstance(items, list):
            items = []
        raw_candidates += len(items)

        # Check THIS page for a match before deciding whether to fetch another.
        matches = [item for item in items
                   if isinstance(item, dict) and _safe_int(item.get("productId")) == product_id]

        if matches:
            distinct_gcts = {item.get("gctsNo") for item in matches}
            chosen = matches[0]
            gcts, nb = extract_gcts_and_nb(chosen)
            status = "Found"
            if len(distinct_gcts) > 1:
                status = "Ambiguous"
                print(f"      ⚠️  Product ID {product_id}: {len(distinct_gcts)} distinct GCTS numbers "
                      f"returned ({distinct_gcts}) -- using the first one, but this needs a manual look.")
            return {
                "status": status,
                "gcts": gcts,
                "nb": nb,
                "raw_candidates": raw_candidates,
                "matched_candidates": len(matches),
            }

        # No match on this page -- decide whether it's worth trying another.
        if not items:
            break  # server has nothing more for this term at all
        if total and total > 0 and raw_candidates >= total:
            break  # we've now seen everything the server says exists
        if len(items) < PAGE_SIZE:
            break  # short page with no total info -- treat as the last page
        if pages_fetched >= max_pages:
            break  # hit our safety cap

        page_num += 1

    return {
        "status": "Not Found",
        "gcts": "",
        "nb": "",
        "raw_candidates": raw_candidates,
        "matched_candidates": 0,
    }


def load_progress():
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            print(f"⚠️  Could not read {PROGRESS_FILE}, starting fresh. "
                  f"(If this file was mid-write during a crash, check for a .bak.)")
    return {}


def save_progress(progress):
    # Write to a temp file then replace, so a crash mid-write can't corrupt
    # hours of accumulated progress.
    tmp_path = PROGRESS_FILE + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(progress, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, PROGRESS_FILE)


def build_output_dataframe(input_df, progress):
    """
    Merges the original input rows with whatever's in `progress` so far,
    preserving the original row order and all original columns.
    """
    records = []
    for _, row in input_df.iterrows():
        pid = int(row[PRODUCT_ID_COLUMN])
        entry = progress.get(str(pid))
        record = row.to_dict()
        if entry:
            record["GCTS No"] = entry.get("gcts", "")
            record["NB No"] = entry.get("nb", "")
            record["Match Status"] = entry.get("status", "")
            record["Raw Candidates"] = entry.get("raw_candidates", "")
        else:
            record["GCTS No"] = ""
            record["NB No"] = ""
            record["Match Status"] = "Not Processed Yet"
            record["Raw Candidates"] = ""
        records.append(record)
    return pd.DataFrame(records)


# =========================
# MAIN SCRIPT
# =========================

def main():
    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(
            f"Input file '{INPUT_FILE}' not found. Put GSO_LV_merged.xlsx next to this "
            f"script, or update INPUT_FILE at the top to the correct path."
        )

    print(f"📂 Reading '{INPUT_FILE}' (sheet '{INPUT_SHEET}')...")
    input_df = pd.read_excel(INPUT_FILE, sheet_name=INPUT_SHEET)
    if PRODUCT_ID_COLUMN not in input_df.columns:
        raise KeyError(
            f"Column '{PRODUCT_ID_COLUMN}' not found in {INPUT_FILE}. "
            f"Available columns: {list(input_df.columns)}"
        )

    total_rows = len(input_df)
    print(f"📊 {total_rows} rows loaded, column '{PRODUCT_ID_COLUMN}' as the lookup key.")

    est_hours = (total_rows * REQUEST_DELAY_SECONDS) / CONCURRENCY / 3600
    print(f"⏱️  Estimated minimum runtime (delay alone, before network latency), "
          f"with CONCURRENCY={CONCURRENCY}: ~{est_hours:.1f} hours for a full unattended run.")

    if TEST_LIMIT is not None:
        print(f"🧪 TEST_LIMIT is set to {TEST_LIMIT} -- only processing that many NEW rows this run. "
              f"Set TEST_LIMIT = None once you've confirmed the results look right.")

    progress = load_progress()
    if progress:
        already_found = sum(1 for v in progress.values() if v.get("status") == "Found")
        print(f"📇 Resuming: {len(progress)} row(s) already recorded in {PROGRESS_FILE} "
              f"({already_found} previously 'Found').")

    # Decide which rows still need work this run.
    pending_ids = []
    for _, row in input_df.iterrows():
        pid = int(row[PRODUCT_ID_COLUMN])
        key = str(pid)
        if key in progress:
            if RETRY_NOT_FOUND and progress[key].get("status") != "Found":
                pending_ids.append(pid)
            # else: already done, skip
        else:
            pending_ids.append(pid)

    if TEST_LIMIT is not None:
        pending_ids = pending_ids[:TEST_LIMIT]

    print(f"🚀 {len(pending_ids)} row(s) to process this run.\n")

    global _shared_cookies
    browser = None
    context = None
    processed_this_run = 0
    start_time = time.time()

    try:
        if pending_ids:
            # Step 1: use Playwright ONLY to obtain/validate a logged-in
            # session, then lift its cookies out. Playwright's sync API
            # can't safely be driven from multiple threads, so it never
            # touches the actual lookup loop below.
            with sync_playwright() as p:
                browser, context = get_browser_and_context(p)
                _shared_cookies = context.storage_state().get("cookies", [])
                browser.close()
                browser = None

            print(f"🚀 Running with CONCURRENCY={CONCURRENCY} worker thread(s) "
                  f"over plain HTTP requests...\n")

            # Step 2: fan the actual lookups out across a thread pool.
            # Rows complete out of order -- that's fine, since progress
            # is keyed by product ID and the final output is rebuilt in
            # original row order from build_output_dataframe().
            with ThreadPoolExecutor(max_workers=CONCURRENCY) as executor:
                future_to_id = {
                    executor.submit(search_product_by_id, get_thread_session(), pid): pid
                    for pid in pending_ids
                }

                for i, future in enumerate(as_completed(future_to_id), start=1):
                    product_id = future_to_id[future]
                    try:
                        result = future.result()
                    except Exception as e:
                        print(f"   ⚠️  Product ID {product_id}: request failed ({e}), will retry next run.")
                        continue  # don't record -- leave it pending for the next run

                    progress[str(product_id)] = {
                        "gcts": result["gcts"],
                        "nb": result["nb"],
                        "status": result["status"],
                        "raw_candidates": result["raw_candidates"],
                    }
                    processed_this_run += 1

                    status_icon = {"Found": "✅", "Not Found": "❌", "Ambiguous": "⚠️"}.get(result["status"], "?")
                    print(f"   [{i}/{len(pending_ids)}] {status_icon} Product ID {product_id}: "
                          f"status={result['status']}, GCTS='{result['gcts']}', NB='{result['nb']}'")

                    if processed_this_run % SAVE_EVERY_N == 0:
                        save_progress(progress)
                        elapsed = time.time() - start_time
                        rate = processed_this_run / elapsed if elapsed > 0 else 0
                        remaining = len(pending_ids) - i
                        eta_min = (remaining / rate / 60) if rate > 0 else float("nan")
                        print(f"      💾 Checkpoint saved ({processed_this_run} done this run, "
                              f"{rate:.2f} rows/sec, ETA ~{eta_min:.1f} min for remainder of this run)")

                    if processed_this_run % XLSX_SNAPSHOT_EVERY_N == 0:
                        build_output_dataframe(input_df, progress).to_excel(OUTPUT_FILE, index=False)
                        print(f"      📄 Snapshot written to {OUTPUT_FILE}")
        else:
            print("Nothing to do -- all rows already resolved. (Set RETRY_NOT_FOUND = True to "
                  "retry rows that weren't 'Found', or delete entries from the progress file.)")

    except KeyboardInterrupt:
        print("\n⏸️  Interrupted by user -- saving progress before exiting...")

    finally:
        if browser is not None:
            try:
                browser.close()
            except Exception as e:
                print(f"⚠️  Error closing browser: {e}")

        save_progress(progress)
        final_df = build_output_dataframe(input_df, progress)
        final_df.to_excel(OUTPUT_FILE, index=False)

        found = sum(1 for v in progress.values() if v.get("status") == "Found")
        not_found = sum(1 for v in progress.values() if v.get("status") == "Not Found")
        ambiguous = sum(1 for v in progress.values() if v.get("status") == "Ambiguous")
        not_processed = total_rows - len(progress)

        print(f"\n✅ Done for this run. Processed {processed_this_run} row(s).")
        print(f"📊 Overall totals across all runs so far: {found} Found, {not_found} Not Found, "
              f"{ambiguous} Ambiguous, {not_processed} not yet processed (of {total_rows} total).")
        print(f"📄 Output written to {OUTPUT_FILE}")
        print(f"📇 Progress checkpoint: {PROGRESS_FILE} ({len(progress)} entries)")
        if not_processed > 0:
            print(f"👉 Re-run this script (with TEST_LIMIT = None) to continue -- "
                  f"already-processed rows will be skipped automatically.")


if __name__ == "__main__":
    main()
