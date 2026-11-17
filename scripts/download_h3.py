"""
download_h3.py  -  H3: download-only script (no content decisions)

For each sample token, downloads from the Blockscout API v2 (gnosisscan.io):
  /api/v2/tokens/{address}            token info           -> token.json
  /api/v2/tokens/{address}/counters   counters (info only) -> counters.json
  /api/v2/tokens/{address}/holders    ALL pages            -> holders_0001.json, ...
  /api/v2/tokens/{address}/transfers  pages back to 1 Jan 2024
                                      (see TRANSFERS_SINCE) -> transfers_0001.json, ...
Every response is saved byte-for-byte, unmodified. Nothing is filtered: the
12-month window, mint/burn exclusion and holder shares are applied later, in
processing, from these saved files.

Why transfers back to 2024: decision H3a asks for monthly transfer counts
since 1 Jan 2024 (before/after comparison), in addition to the 12-month window.
The API returns transfers newest first; paging stops after the first page that
contains a transfer older than TRANSFERS_SINCE. That last page is saved
unmodified, so it may contain some transfers from before 2024; they are
ignored in processing. If a page is not sorted newest-first, the script does
not stop early (it keeps paging) and records this in the manifest.

Change log
  1.0  2026-11-12  Complete transfer history (never run).
  1.1  2026-11-12  Transfers only back to 1 Jan 2024 (author's decision). Why all holder pages: decision H3c uses the current holder
list (share of the largest and the ten largest holders).

Tokens: by default the 40 tokens of sheet Sample (draw_rank 1-40) in
output/sample_20261112T130912Z_v1.0.xlsx (decision H3e). Reserve tokens only if
a sample token fails, e.g.:  python3 download_h3.py --ranks 41
(A reserve download is a separate, later run and is recorded as such.)

Output: raw_h3/<UTC timestamp>/<rank>_<address>/...  +  raw_h3/<UTC timestamp>/manifest.json

Requires: Python 3.9+, `pip install requests openpyxl`
Run:      python3 download_h3.py
"""

import argparse
import hashlib
import json
import platform
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

import requests
from openpyxl import load_workbook

SCRIPT_VERSION = "1.1"
TRANSFERS_SINCE = "2024-01-01T00:00:00Z"   # decision H3a (author): monthly counts since 2024
BASE = "https://gnosisscan.io/api/v2/tokens"
HEADERS = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)",
           "Accept": "application/json"}   # same as download_raw.py v1.2 (decision 1g)
SAMPLE_FILE = Path("output") / "sample_20261112T130912Z_v1.0.xlsx"
OUT_ROOT = Path("raw_h3")
PAUSE_SECONDS = 0.5
MAX_RETRIES = 5
TIMEOUT_SECONDS = 30
MAX_PAGES = 5000           # per endpoint and token; safety stop


def utc_now():
    return datetime.now(timezone.utc)


def encode_cursor(nxt):
    """Same encoding as download_raw.py v1.2 (decision 1h): every key is sent,
    None as 'null', booleans as 'true'/'false'."""
    params = {}
    for k, v in nxt.items():
        if v is None:
            v = "null"
        elif isinstance(v, bool):
            v = "true" if v else "false"
        params[k] = v
    return params


def fetch(session, url, params, attempts):
    for attempt in range(1, MAX_RETRIES + 1):
        started = utc_now().isoformat()
        try:
            r = session.get(url, params=params, timeout=TIMEOUT_SECONDS)
            attempts.append({"attempt": attempt, "time_utc": started,
                             "status": r.status_code, "url": r.url})
            if urlparse(r.url).netloc != urlparse(BASE).netloc:
                raise RuntimeError(f"Redirected to another host: {r.url}")
            if r.status_code == 200:
                return r
            if r.status_code in (429, 500, 502, 503, 504):
                time.sleep(2 ** attempt)
                continue
            r.raise_for_status()
        except requests.RequestException as e:
            attempts.append({"attempt": attempt, "time_utc": started, "error": repr(e)})
            if attempt == MAX_RETRIES:
                raise
            time.sleep(2 ** attempt)
    raise RuntimeError(f"Giving up after {MAX_RETRIES} attempts: {url} {params}")


def save_page(folder, filename, r, attempts, items=None, nxt=None):
    (folder / filename).write_bytes(r.content)          # unmodified
    return {"file": filename, "sha256": hashlib.sha256(r.content).hexdigest(),
            "bytes": len(r.content), "request_url": r.url,
            "server_date": r.headers.get("Date"), "attempts": attempts,
            "items": items, "next_page_params": nxt}


def download_paged(session, url, folder, prefix, log, since=None):
    """Pages until next_page_params is null, or - if `since` is given - until a
    newest-first page contains an item with timestamp < since.
    Returns (all items, stop status)."""
    params, seen, all_items = {}, set(), []
    for page_no in range(1, MAX_PAGES + 1):
        attempts = []
        r = fetch(session, url, params, attempts)
        data = json.loads(r.content)
        items = data.get("items", [])
        nxt = data.get("next_page_params")
        log.append(save_page(folder, f"{prefix}_{page_no:04d}.json", r, attempts,
                             len(items), nxt))
        all_items += items
        if not nxt:
            return all_items, "complete"
        if since:
            ts = [i.get("timestamp") for i in items if i.get("timestamp")]
            newest_first = ts == sorted(ts, reverse=True)
            log[-1]["sorted_newest_first"] = newest_first
            # compare to the second: API format is ...T21:08:40.000000Z
            if ts and newest_first and min(t[:19] for t in ts) < since[:19]:
                return all_items, f"complete (reached transfers before {since})"
        key = json.dumps(nxt, sort_keys=True)
        if key in seen:
            raise RuntimeError(f"Cursor repeated in {prefix}: {nxt}")
        seen.add(key)
        params = encode_cursor(nxt)
        time.sleep(PAUSE_SECONDS)
    return all_items, f"MAX_PAGES={MAX_PAGES} reached (incomplete!)"


def read_tokens(ranks):
    wb = load_workbook(SAMPLE_FILE, read_only=True, data_only=True)
    tokens = []
    for sheet in ("Sample", "Reserve"):
        rows = list(wb[sheet].iter_rows(values_only=True))
        hdr = list(rows[0])
        for r in rows[1:]:
            rec = dict(zip(hdr, r))
            if rec.get("draw_rank") in ranks:
                tokens.append({"draw_rank": rec["draw_rank"], "raw_id": rec["raw_id"],
                               "address_hash": rec["address_hash"], "name": rec["name"],
                               "sheet": sheet})
    wb.close()
    missing = set(ranks) - {t["draw_rank"] for t in tokens}
    if missing:
        sys.exit(f"draw_rank not found in sample file: {sorted(missing)}")
    return sorted(tokens, key=lambda t: t["draw_rank"])


def earlier_runs(current):
    out = []
    for d in sorted(OUT_ROOT.iterdir()):
        if d.is_dir() and d != current:
            m = d / "manifest.json"
            entry = {"folder": d.name, "status": "unknown (no manifest)"}
            if m.exists():
                old = json.loads(m.read_text(encoding="utf-8"))
                entry.update(status=old.get("status"), ranks=old.get("ranks"),
                             script_version=old.get("script_version"))
            out.append(entry)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ranks", default="1-40",
                    help="draw_ranks to download, e.g. '1-40' (default) or '41' or '41,42'")
    args = ap.parse_args()
    ranks = []
    for part in args.ranks.split(","):
        a, _, b = part.partition("-")
        ranks += list(range(int(a), int(b or a) + 1))

    run_dir = OUT_ROOT / utc_now().strftime("%Y%m%dT%H%M%SZ")
    run_dir.mkdir(parents=True, exist_ok=False)
    tokens = read_tokens(ranks)

    manifest = {
        "script": Path(__file__).name, "script_version": SCRIPT_VERSION,
        "script_sha256": hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
        "python_version": sys.version, "requests_version": requests.__version__,
        "platform": platform.platform(),
        "sample_file": str(SAMPLE_FILE),
        "sample_file_sha256": hashlib.sha256(SAMPLE_FILE.read_bytes()).hexdigest(),
        "ranks": args.ranks, "base_url": BASE, "request_headers": HEADERS,
        "cursor_encoding": "all keys sent; None as 'null'; booleans as 'true'/'false'",
        "pause_seconds": PAUSE_SECONDS,
        "transfers_since": TRANSFERS_SINCE,
        "earlier_runs_in_raw_h3": earlier_runs(run_dir),
        "started_utc": utc_now().isoformat(), "finished_utc": None,
        "status": "running", "tokens": [],
    }
    mpath = run_dir / "manifest.json"

    def save_manifest():
        mpath.write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")

    session = requests.Session()
    session.headers.update(HEADERS)

    for t in tokens:
        folder = run_dir / f"{t['draw_rank']:03d}_{t['address_hash']}"
        folder.mkdir()
        entry = {**t, "folder": folder.name, "status": "running", "started_utc":
                 utc_now().isoformat(), "pages": {"token": [], "counters": [],
                                                  "holders": [], "transfers": []}}
        manifest["tokens"].append(entry)
        save_manifest()
        url = f"{BASE}/{t['address_hash']}"
        print(f"[{t['draw_rank']:3d}] {t['name']}")
        try:
            for name, suffix in (("token", ""), ("counters", "/counters")):
                attempts = []
                r = fetch(session, url + suffix, {}, attempts)
                entry["pages"][name].append(save_page(folder, f"{name}.json", r, attempts))
                time.sleep(PAUSE_SECONDS)
            counters = json.loads((folder / "counters.json").read_bytes())

            holders, h_status = download_paged(session, url + "/holders", folder, "holders",
                                               entry["pages"]["holders"])
            save_manifest()
            transfers, t_status = download_paged(session, url + "/transfers", folder,
                                                 "transfers", entry["pages"]["transfers"],
                                                 since=TRANSFERS_SINCE)

            ids = [(x.get("transaction_hash"), x.get("log_index")) for x in transfers]
            ts = [x.get("timestamp") for x in transfers if x.get("timestamp")]
            haddr = [((x.get("address") or {}).get("hash") or "").lower() for x in holders]
            # Informational checks only - nothing is removed
            entry["checks"] = {
                "holders_status": h_status, "transfers_status": t_status,
                "holders_items": len(holders), "holders_unique_addresses": len(set(haddr)),
                "counters_token_holders_count": counters.get("token_holders_count"),
                "counters_transfers_count": counters.get("transfers_count"),
                "transfers_items": len(transfers),
                "transfers_duplicate_tx_log": len(ids) - len(set(ids)),
                "oldest_transfer": min(ts) if ts else None,
                "newest_transfer": max(ts) if ts else None,
            }
            entry["checks"]["transfers_since"] = TRANSFERS_SINCE
            entry["checks"]["transfers_from_2024_on"] = sum(t[:19] >= TRANSFERS_SINCE[:19] for t in ts)
            ok = h_status == "complete" and t_status.startswith("complete")
            entry["status"] = "complete" if ok else "incomplete"
            print(f"      holders {len(holders)} | transfers {len(transfers)} "
                  f"({entry['checks']['oldest_transfer']} -> "
                  f"{entry['checks']['newest_transfer']}) | {entry['status']}")
        except Exception as e:
            entry["status"] = "failed"
            entry["error"] = repr(e)
            print("      FAILED:", e)
        entry["finished_utc"] = utc_now().isoformat()
        save_manifest()

    states = [t["status"] for t in manifest["tokens"]]
    manifest["status"] = "complete" if all(s == "complete" for s in states) else "incomplete"
    manifest["finished_utc"] = utc_now().isoformat()
    manifest["totals"] = {"tokens": len(states),
                          "complete": states.count("complete"),
                          "incomplete": states.count("incomplete"),
                          "failed": states.count("failed")}
    save_manifest()
    for f in run_dir.rglob("*"):
        if f.is_file():
            f.chmod(0o444)

    print("\nStatus:", manifest["status"], manifest["totals"])
    print("Saved to:", run_dir.resolve())
    if manifest["status"] != "complete":
        print("Tokens not complete:",
              [t["draw_rank"] for t in manifest["tokens"] if t["status"] != "complete"])


if __name__ == "__main__":
    main()
