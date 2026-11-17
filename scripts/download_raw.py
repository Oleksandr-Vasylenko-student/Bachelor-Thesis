"""
download_raw.py  -  Step 1e: download-only script (no content decisions)

Downloads ALL pages of
    https://gnosisscan.io/api/v2/tokens?q=RealToken&type=ERC-20
and saves every API response byte-for-byte, unmodified, plus a metadata file
(manifest.json) documenting how and when the data were obtained.

Change log
  1.0  2026-11-12  Base URL gnosis.blockscout.com, default python-requests
                   User-Agent. Run 20261112T130141Z FAILED: the domain
                   redirected to gnosisscan.io, which answered HTTP 403.
  1.1  2026-11-12  Base URL changed to gnosisscan.io (the redirect target).
                   Browser User-Agent + 'Accept: application/json' added:
                   the student's manual curl test returned 403 without and
                   200 with these headers. The redirect chain of the old
                   domain and the status of earlier run folders are now
                   recorded in the manifest. Failed run folders are kept.
                   Run 20261112T130609Z FAILED: page 2 repeated page 1
                   (cursor not accepted; loop guard stopped the run).
  1.2  2026-11-12  Cursor keys with value None are now sent as the string
                   'null' instead of being omitted, so that every key of
                   next_page_params is sent back to the API.

This script deliberately does NOT filter, deduplicate, classify or convert
anything. All of that happens in later steps, which read only from the saved
files.

Requires: Python 3.9+, `pip install requests`
Run:      python download_raw.py
Output:   raw/<UTC timestamp>/page_0001.json, page_0002.json, ...
          raw/<UTC timestamp>/manifest.json
"""

import hashlib
import json
import platform
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

import requests

# --------------------------------------------------------------------------
# Configuration (every value here is recorded in manifest.json)
# --------------------------------------------------------------------------
SCRIPT_VERSION = "1.2"
BASE_URL = "https://gnosisscan.io/api/v2/tokens"
OLD_URL = "https://gnosis.blockscout.com/api/v2/tokens"  # probed once, for documentation
REQUEST_HEADERS = {
    # Same headers as the successful manual curl test (403 without, 200 with)
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)",
    "Accept": "application/json",
}
CHANGE_REASON = ("v1.0 run 20261112T130141Z failed: gnosis.blockscout.com redirects to "
                 "gnosisscan.io, which returned HTTP 403 to the default python-requests "
                 "User-Agent. Manual curl test by the student: 403 without headers, 200 with "
                 "browser User-Agent + 'Accept: application/json'. v1.1 requests "
                 "gnosisscan.io directly with these headers. "
                 "v1.1 run 20261112T130609Z failed: page 2 was identical to page 1 because "
                 "the cursor keys with null values (fiat_value, market_cap) were omitted. "
                 "v1.2 sends every cursor key, null values as the string 'null'.")
QUERY_PARAMS = {"q": "RealToken", "type": "ERC-20"}  # decision 1b (open, see log)
PAUSE_SECONDS = 1.0         # pause between pages, to be polite to the API
MAX_RETRIES = 5             # retries per page on HTTP 429 / 5xx / network errors
TIMEOUT_SECONDS = 30
MAX_PAGES = 2000            # safety stop against endless paging
OUT_ROOT = Path("raw")


def utc_now():
    return datetime.now(timezone.utc)


def encode_cursor(next_page_params):
    """Turn next_page_params into query parameters for the next request.

    Two technical conversions, recorded in the manifest:
      - keys with value None are omitted (no value to send)
      - booleans are sent as lowercase 'true'/'false' (JSON style)
    The original QUERY_PARAMS are always sent again, because the API's
    cursor does not include `type` (observed on page 1).
    """
    params = dict(QUERY_PARAMS)
    for key, value in next_page_params.items():
        if value is None:
            value = "null"   # v1.2: send every key (omitting them reset paging)
        elif isinstance(value, bool):
            value = "true" if value else "false"
        params[key] = value
    return params


def fetch(session, params, attempt_log):
    """GET one page with retries. Returns the final Response."""
    for attempt in range(1, MAX_RETRIES + 1):
        started = utc_now().isoformat()
        try:
            r = session.get(BASE_URL, params=params, timeout=TIMEOUT_SECONDS)
            attempt_log.append({"attempt": attempt, "time_utc": started,
                                "status": r.status_code, "url": r.url})
            if r.status_code == 200:
                return r
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(2 ** attempt)
                continue
            r.raise_for_status()  # other errors: stop immediately
        except requests.RequestException as e:
            attempt_log.append({"attempt": attempt, "time_utc": started,
                                "error": repr(e)})
            if attempt == MAX_RETRIES:
                raise
            time.sleep(2 ** attempt)
    raise RuntimeError(f"Giving up after {MAX_RETRIES} attempts: {params}")


def probe_old_domain(session):
    """Record (not follow for data) where the old Blockscout domain points."""
    try:
        r = session.get(OLD_URL, params=QUERY_PARAMS, timeout=TIMEOUT_SECONDS,
                        allow_redirects=True)
        return {"time_utc": utc_now().isoformat(),
                "chain": [{"url": h.url, "status": h.status_code,
                           "location": h.headers.get("Location")} for h in r.history]
                         + [{"url": r.url, "status": r.status_code}]}
    except requests.RequestException as e:
        return {"time_utc": utc_now().isoformat(), "error": repr(e)}


def earlier_runs(current_dir):
    """List earlier run folders and their recorded status. Nothing is deleted."""
    runs = []
    for d in sorted(OUT_ROOT.iterdir()):
        if not d.is_dir() or d == current_dir:
            continue
        entry = {"folder": d.name, "status": "unknown (no manifest)"}
        m = d / "manifest.json"
        if m.exists():
            old = json.loads(m.read_text(encoding="utf-8"))
            entry.update({"status": old.get("status"),
                          "stop_reason": old.get("stop_reason"),
                          "script_version": old.get("script_version"),
                          "base_url": old.get("base_url"),
                          "pages": old.get("totals", {}).get("pages")})
        runs.append(entry)
    return runs


def main():
    run_id = utc_now().strftime("%Y%m%dT%H%M%SZ")
    out_dir = OUT_ROOT / run_id
    out_dir.mkdir(parents=True, exist_ok=False)  # never overwrite an old run

    manifest = {
        "script": Path(__file__).name,
        "script_version": SCRIPT_VERSION,
        "change_reason": CHANGE_REASON,
        "earlier_runs_in_raw_folder": earlier_runs(out_dir),
        "request_headers": REQUEST_HEADERS,
        "old_domain_probe": None,
        "script_sha256": hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
        "python_version": sys.version,
        "requests_version": requests.__version__,
        "platform": platform.platform(),
        "base_url": BASE_URL,
        "query_params": QUERY_PARAMS,
        "cursor_encoding": "all cursor keys sent; None sent as 'null'; booleans as 'true'/'false'; "
                           "QUERY_PARAMS re-sent on every page",
        "pause_seconds": PAUSE_SECONDS,
        "started_utc": utc_now().isoformat(),
        "finished_utc": None,
        "status": "running",
        "stop_reason": None,
        "pages": [],
        "totals": {},
    }
    manifest_path = out_dir / "manifest.json"

    def save_manifest():
        manifest_path.write_text(json.dumps(manifest, indent=2, ensure_ascii=False),
                                 encoding="utf-8")

    session = requests.Session()
    session.headers.update(REQUEST_HEADERS)

    manifest["old_domain_probe"] = probe_old_domain(session)
    save_manifest()

    params = dict(QUERY_PARAMS)
    seen_cursors = set()
    total_items = 0
    addresses = []

    try:
        for page_no in range(1, MAX_PAGES + 1):
            attempts = []
            r = fetch(session, params, attempts)

            # Refuse silent redirects to another host (cf. the 403 episode)
            final_host = urlparse(r.url).netloc
            if final_host != urlparse(BASE_URL).netloc:
                raise RuntimeError(f"Redirected to another host: {r.url}")

            # 1) Save the response exactly as received
            file_name = f"page_{page_no:04d}.json"
            (out_dir / file_name).write_bytes(r.content)

            # 2) Read it only to find the cursor and count records
            data = json.loads(r.content)
            items = data.get("items", [])
            nxt = data.get("next_page_params")
            total_items += len(items)
            addresses += [str(it.get("address_hash", "")).lower() for it in items]

            manifest["pages"].append({
                "page": page_no,
                "file": file_name,
                "sha256": hashlib.sha256(r.content).hexdigest(),
                "bytes": len(r.content),
                "request_url": r.url,
                "redirects": [h.url for h in r.history],
                "server_date": r.headers.get("Date"),
                "attempts": attempts,
                "items": len(items),
                "next_page_params": nxt,
            })
            save_manifest()  # keep a usable manifest even if a later page fails
            print(f"page {page_no:4d}: {len(items):3d} items  (total {total_items})")

            if not nxt:
                manifest["stop_reason"] = "next_page_params is null (last page)"
                break
            key = json.dumps(nxt, sort_keys=True)
            if key in seen_cursors:
                raise RuntimeError(f"Cursor repeated, stopping to avoid a loop: {nxt}")
            seen_cursors.add(key)

            params = encode_cursor(nxt)
            time.sleep(PAUSE_SECONDS)
        else:
            manifest["stop_reason"] = f"MAX_PAGES={MAX_PAGES} reached (incomplete!)"

        manifest["status"] = ("complete" if manifest["stop_reason"].startswith("next_page")
                              else "incomplete")
    except Exception as e:
        manifest["status"] = "failed"
        manifest["stop_reason"] = repr(e)
        print("ERROR:", e)
    finally:
        manifest["finished_utc"] = utc_now().isoformat()
        # Counts only - nothing is removed. Duplicates are reported, not dropped.
        manifest["totals"] = {
            "pages": len(manifest["pages"]),
            "items": total_items,
            "unique_address_hash": len(set(addresses)),
            "address_hash_appearing_more_than_once": len(addresses) - len(set(addresses)),
        }
        save_manifest()

    # Make the raw files read-only as a guard against accidental edits
    for f in out_dir.iterdir():
        f.chmod(0o444)

    print("\nStatus:", manifest["status"], "-", manifest["stop_reason"])
    print("Totals:", manifest["totals"])
    print("Saved to:", out_dir.resolve())


if __name__ == "__main__":
    main()
