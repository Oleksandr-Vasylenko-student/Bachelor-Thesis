"""
explore_h3.py  -  H3 step 1: exploration only (not the analysis dataset)

Looks at what the API returns for ONE token: the sample token with draw_rank 1,
read from output/sample_20261112T130912Z_v1.0.xlsx. Saves every response
unmodified in explore_h3/<UTC timestamp>/ and prints the structure.

Endpoints (Blockscout API v2, same host and headers as download_raw.py v1.2):
  /api/v2/tokens/{address}            token info
  /api/v2/tokens/{address}/counters   number of holders and transfers
  /api/v2/tokens/{address}/holders    current holders (first page only)
  /api/v2/tokens/{address}/transfers  transfers, newest first (first page only)

Requires: pip install requests openpyxl
Run:      python3 explore_h3.py
"""

import json
import time
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

import requests
from openpyxl import load_workbook

BASE = "https://gnosisscan.io/api/v2/tokens"
HEADERS = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)",
           "Accept": "application/json"}
SAMPLE_FILE = Path("output") / "sample_20261112T130912Z_v1.0.xlsx"

# Token with draw_rank 1 from the saved sample (no hard-coded address)
ws = load_workbook(SAMPLE_FILE, read_only=True, data_only=True)["Sample"]
rows = list(ws.iter_rows(values_only=True))
hdr = list(rows[0])
first = dict(zip(hdr, rows[1]))
address, name = first["address_hash"], first["name"]
print(f"Token (draw_rank {first['draw_rank']}): {name}\nAddress: {address}\n")

out = Path("explore_h3") / datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
out.mkdir(parents=True)
session = requests.Session()
session.headers.update(HEADERS)


def get(suffix, filename):
    r = session.get(f"{BASE}/{address}{suffix}", timeout=30)
    (out / filename).write_bytes(r.content)          # saved unmodified
    print(f"--- {suffix or '/'}  HTTP {r.status_code}  -> {filename}")
    time.sleep(1)
    return r.json() if r.status_code == 200 else None


info = get("", "token.json")
counters = get("/counters", "counters.json")
holders = get("/holders", "holders_page1.json")
transfers = get("/transfers", "transfers_page1.json")

print("\nCOUNTERS:", json.dumps(counters, indent=2))

for label, data in (("HOLDERS", holders), ("TRANSFERS", transfers)):
    if not data:
        continue
    items = data.get("items", [])
    print(f"\n{label}: top-level keys {list(data)}; items on page 1: {len(items)}")
    print("next_page_params:", data.get("next_page_params"))
    print("Fields of first item:", list(items[0]) if items else "-")
    print("First item (unmodified):")
    print(json.dumps(items[0] if items else None, indent=2, ensure_ascii=False)[:2500])

if transfers and transfers.get("items"):
    items = transfers["items"]
    ts = [i.get("timestamp") for i in items if i.get("timestamp")]
    print("\nTRANSFERS page 1 overview")
    print("  time span on page 1:", min(ts) if ts else "-", "->", max(ts) if ts else "-")
    print("  type:", Counter(i.get("type") for i in items))
    print("  method:", Counter(i.get("method") for i in items))
    zero = "0x0000000000000000000000000000000000000000"
    print("  from = zero address (mint):",
          sum(((i.get("from") or {}).get("hash") or "").lower() == zero for i in items))
    print("  from/to is a contract:",
          sum(bool((i.get("from") or {}).get("is_contract")) for i in items), "/",
          sum(bool((i.get("to") or {}).get("is_contract")) for i in items))

if holders and holders.get("items"):
    print("\nHOLDERS page 1: largest 5 (address, is_contract, name, value)")
    for h in holders["items"][:5]:
        a = h.get("address") or {}
        print(f"  {a.get('hash')}  contract={a.get('is_contract')}  "
              f"name={a.get('name')}  value={h.get('value')}")

print(f"\nSaved to: {out.resolve()}")
