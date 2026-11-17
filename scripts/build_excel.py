"""
build_excel.py  -  Step 4: processing of the frozen raw dataset into Excel

Reads ONLY the frozen snapshot raw/20261112T130912Z (written by download_raw.py
v1.2). Nothing is downloaded. The raw folder is never written to.

Before processing, every page file is checked against the SHA-256 checksum in
manifest.json; the script stops if anything does not match.

Output: output/RealT_tokens_20261112T130912Z_v<SCRIPT_VERSION>.xlsx
        (never overwritten - delete or rename an old file yourself)

Sheets: README, Summary, Included, Excluded, Manual_Check, Raw_Data, Rules
Every raw record ends up in exactly one of Included / Excluded / Manual_Check.

Rules implemented (IDs refer to the decision log, see sheet "Rules"):
  R2a  derivative and ecosystem tokens are excluded
  R2b  classification by symbol code (+ assumptions where marked)
  R2c  zero supply at snapshot -> excluded
  R2d  several contracts for one property -> keep the one with most holders
  R2e  no minimum number of holders

Known result of the automatic checks (snapshot 20261112T130912Z):
  Check 2 "All US tokens in Included have a valid US state code" is FALSE for
  exactly one token: "RealToken S 17562 Mansfield St Detroit"
  (symbol REALTOKEN-S-17562-MANSFIELD-ST-DETROIT). The state code is missing in
  the source data (name and symbol). It is presumably MI (Detroit), but this is
  NOT filled in, so that the data stay as delivered (decision 4a). The check is
  intentionally left FALSE and explained in the README.

Change log
  1.0  2026-11-12  First version.
  1.1  2026-11-12  README only: limitation L1 reworded (decision of the author);
                   known exception of check 2 documented (decision 4a).
                   Classification and data unchanged.

Requires: Python 3.9+, `pip install openpyxl`
Run:      python3 build_excel.py
"""

import hashlib
import json
import re
import sys
from collections import defaultdict
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SCRIPT_VERSION = "1.1"
SNAPSHOT = "20261112T130912Z"
RAW_DIR = Path("raw") / SNAPSHOT
OUT_DIR = Path("output")
MAX_ROW = 2000  # formula ranges cover rows 2..MAX_ROW (enough for 924 records)

# ---------------------------------------------------------------------------
# Rules (written to sheet "Rules"; every row references its rule IDs)
# ---------------------------------------------------------------------------
RULES = [
    # id, decision, description, rests on assumption?
    ("R2a-1", "2a", "Symbol starts with 'armm' (name 'RMM yield generating ...'): derivative "
                    "(RMM deposit token), not a property token -> Excluded.", "no"),
    ("R2a-2", "2a", "Name starts with 'Balancer' or 'BalCoW', or symbol starts with 'BCoW-': "
                    "liquidity pool token of a third-party protocol -> Excluded.", "no"),
    ("R2a-3", "2a", "Symbol REG, REUSD or starting with 'RTW-': RealT ecosystem token "
                    "(governance / stablecoin / wrapped USD) -> Excluded.", "no"),
    ("R2b-1", "2b", "Category from the symbol code REALTOKEN-<code>-...: S = US property; "
                    "SL = US loan; PA-SE / CO-SXE / MX-SXE = international property; PA-SEL = "
                    "international loan; PA-S = factoring; CH-S = holding company; D = D contract "
                    "(meaning unknown).", "yes (meaning of codes inferred from names)"),
    ("R2b-2", "2b+", "Symbol REALTOKEN-<house number>... without a code: 'US property - older "
                     "symbol scheme', main category 'US rental property'.",
     "yes ('no code = S' inferred from name and US address)"),
    ("R2b-3", "2b", "Symbol equal to the full name 'RealToken S ...': classified as US property, "
                    "flag_symbol_is_name = TRUE.", "yes (inferred from name)"),
    ("R2b-4", "2b", "Symbol matches none of the patterns above -> Manual_Check.", "no"),
    ("R2b-5", "2b", "Flags: flag_holdings ('<n> Holdings' in name), flag_name_code_conflict "
                    "('LOAN'/'Factoring' in name does not match the symbol code), "
                    "flag_symbol_whitespace (leading/trailing spaces in the raw symbol).", "no"),
    ("R2c-1", "2c", "total_supply = 0 at snapshot -> Excluded, reason 'zero supply at snapshot'.",
     "no"),
    ("R2d-1", "2d", "property_key = name without leading 'RealToken' and the series letter "
                    "S or D, lower case, letters and digits only. Contracts with the same "
                    "property_key form a duplicate group (only property/asset tokens).",
     "yes (same normalised name = same property)"),
    ("R2d-2", "2d/3c", "Within a group, the contract with supply > 0 and the most holders is kept "
                       "(Included); other contracts -> Excluded with kept_address.",
     "yes (most widely held contract = the one in use)"),
    ("R2d-3", "2d/3c", "If supplies within a group differ: flag_dup_supply_differs = TRUE "
                       "(no other consequence).", "no"),
    ("R2d-4", "2d/3c", "Tie in holders: decision rule is 'more transfers', but transfer counts "
                       "are not in the raw data -> tied contracts go to Manual_Check. "
                       "(No ties occur in snapshot 20261112T130912Z.)", "no"),
    ("R2e-1", "2e", "No minimum number of holders. holders_count is kept as a column only.", "no"),
]

REASON_TEXT = {
    "R2a-1": "derivative token (RMM)",
    "R2a-2": "derivative token (liquidity pool)",
    "R2a-3": "ecosystem token",
    "R2c-1": "zero supply at snapshot",
    "R2d-2": "duplicate contract (fewer holders than kept contract)",
}
RULE_ORDER = ["R2a-1", "R2a-2", "R2a-3", "R2c-1", "R2d-2"]  # order for primary_reason

# Symbol code -> (category_detailed, category_main)
CODE_MAP = {
    "S":      ("US property", "US rental property"),
    "SL":     ("US loan", "Loan"),
    "D":      ("D contract (meaning unknown)", "Other"),
    "PA-SE":  ("International property - Panama", "Other"),
    "CO-SXE": ("International property - Colombia", "Other"),
    "MX-SXE": ("International property - Mexico", "Other"),
    "PA-SEL": ("International loan - Panama", "Loan"),
    "PA-S":   ("Factoring", "Other"),
    "CH-S":   ("Holding company - Switzerland", "Other"),
}
CODE_RE = re.compile(r"^REALTOKEN-(PA-SEL|PA-SE|PA-S|CO-SXE|MX-SXE|CH-S|SL|S|D)-")
OLDER_SCHEME_RE = re.compile(r"^REALTOKEN-\d")

US_STATES = ["AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "DC", "FL", "GA", "HI", "ID",
             "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO",
             "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA",
             "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY"]

CATEGORY_ORDER = [
    ("US property", "US rental property"),
    ("US property - older symbol scheme", "US rental property"),
    ("US loan", "Loan"),
    ("International loan - Panama", "Loan"),
    ("International property - Panama", "Other"),
    ("International property - Colombia", "Other"),
    ("International property - Mexico", "Other"),
    ("Factoring", "Other"),
    ("Holding company - Switzerland", "Other"),
    ("D contract (meaning unknown)", "Other"),
    ("Derivative - RMM deposit token", "Other"),
    ("Derivative - liquidity pool token", "Other"),
    ("Ecosystem token", "Other"),
    ("Unclassified", "Other"),
]
MAIN_ORDER = ["US rental property", "Loan", "Other"]

LIMITATIONS = [
    "The RealT website (realt.co) was not accessible on 12 November 2026, after the platform "
    "had entered liquidation (source cited in the thesis). No official RealT token list was "
    "available and no manual check against the website was possible.",
    "Classification therefore rests on naming conventions of symbols and names "
    "(see column classification_basis and sheet Rules, column 'assumption').",
    "Single snapshot: supply and holder counts are as of the download time only.",
    "The text search q=RealToken may miss RealT tokens whose name/symbol does not "
    "contain 'RealToken'.",
    "Data were requested from gnosisscan.io (Blockscout API v2 response format); "
    "gnosis.blockscout.com redirects there. The operator of the service was not verified.",
    "Transfer counts are not part of the raw data, so the tie-break rule 'more transfers' "
    "(R2d-4) cannot be applied automatically; ties would go to Manual_Check (none occur).",
]

# Columns of Included / Excluded / Manual_Check (identical in all three sheets)
COLS = [
    ("raw_id", "Page and position in the raw data, e.g. p07-023 = page 7, item 23 (link to Raw_Data)."),
    ("address_hash", "Contract address (unique key)."),
    ("name", "Token name as delivered."),
    ("symbol", "Token symbol as delivered (incl. any spaces)."),
    ("decimals", "Number of decimals (number)."),
    ("total_supply_raw", "total_supply as delivered (text, raw integer units)."),
    ("supply_tokens", "total_supply / 10^decimals (number)."),
    ("holders_count", "Number of holders at snapshot (number)."),
    ("symbol_code", "Code from the symbol (S, SL, PA-SE, ...), 'none' for the older scheme, "
                    "'-' if not applicable."),
    ("category_detailed", "Detailed category (rules R2a/R2b)."),
    ("category_main", "Main category: US rental property / Loan / Other."),
    ("classification_basis", "What the classification rests on; 'assumption' where inferred."),
    ("us_state_code", "Last word of the name if it is two capital letters (US tokens only)."),
    ("us_state_valid", "Excel formula: TRUE if the US token's state code is in the list in "
                       "sheet Rules; empty for non-US tokens."),
    ("flag_holdings", "TRUE if the name contains '<n> Holdings'."),
    ("flag_symbol_is_name", "TRUE if the symbol equals the full name."),
    ("flag_name_code_conflict", "TRUE if 'LOAN'/'Factoring' in the name does not match the code."),
    ("flag_symbol_whitespace", "TRUE if the raw symbol has leading/trailing spaces."),
    ("property_key", "Normalised property name used for duplicate detection (R2d-1)."),
    ("dup_group_id", "ID of the duplicate group (only if group size > 1)."),
    ("dup_group_size", "Number of contracts with the same property_key."),
    ("dup_role", "unique / kept / duplicate / review (tie) / no kept contract (all zero supply)."),
    ("kept_address", "For duplicates: address of the contract that was kept."),
    ("flag_dup_supply_differs", "TRUE if the contracts in the group have different supply."),
    ("status", "Included / Excluded / Manual_Check (at creation of the file)."),
    ("primary_reason", "First applicable exclusion reason (rule order R2a, R2c, R2d)."),
    ("all_reasons", "All applicable exclusion reasons."),
    ("rule_ids", "IDs of all rules applied to this record (see sheet Rules)."),
    ("your_decision", "Empty, for your own notes / decisions."),
    ("comment", "Empty, for your own comments."),
]
COLNAMES = [c for c, _ in COLS]
L = {name: get_column_letter(i + 1) for i, name in enumerate(COLNAMES)}  # column letters


# ---------------------------------------------------------------------------
# 1. Load and verify the frozen raw data
# ---------------------------------------------------------------------------
def load_raw():
    manifest_path = RAW_DIR / "manifest.json"
    if not manifest_path.exists():
        sys.exit(f"Not found: {manifest_path.resolve()}")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("status") != "complete":
        sys.exit(f"Manifest status is '{manifest.get('status')}', expected 'complete'.")

    records, field_order = [], []
    for page in manifest["pages"]:
        content = (RAW_DIR / page["file"]).read_bytes()
        if hashlib.sha256(content).hexdigest() != page["sha256"]:
            sys.exit(f"Checksum mismatch in {page['file']} - raw data changed. Stopping.")
        items = json.loads(content)["items"]
        if len(items) != page["items"]:
            sys.exit(f"Item count mismatch in {page['file']}.")
        for pos, item in enumerate(items, start=1):
            for k in item:
                if k not in field_order:
                    field_order.append(k)
            records.append({"raw_id": f"p{page['page']:02d}-{pos:03d}",
                            "source_file": page["file"], "position": pos, "item": item})

    if len(records) != manifest["totals"]["items"]:
        sys.exit("Number of records differs from manifest totals.")
    print(f"Raw data OK: {len(manifest['pages'])} pages, checksums verified, "
          f"{len(records)} records.")
    return manifest, records, field_order


# ---------------------------------------------------------------------------
# 2. Classification (R2a, R2b)
# ---------------------------------------------------------------------------
def classify(item):
    name = item.get("name") or ""
    raw_symbol = item.get("symbol") or ""
    symbol = raw_symbol.strip()
    r = {"rule_ids": [], "reasons": [], "is_asset": True, "manual": False,
         "flag_symbol_whitespace": raw_symbol != symbol, "flag_symbol_is_name": False}

    def setcat(code, detailed, main, basis, rule):
        r.update(symbol_code=code, category_detailed=detailed, category_main=main,
                 classification_basis=basis)
        r["rule_ids"].append(rule)

    m = CODE_RE.match(symbol)
    if symbol.startswith("armm"):
        setcat("-", "Derivative - RMM deposit token", "Other", "symbol prefix 'armm'", "R2a-1")
        r["reasons"].append("R2a-1"); r["is_asset"] = False
    elif (name.startswith("Balancer") or name.startswith("BalCoW")
          or symbol.startswith("BCoW-")):
        setcat("-", "Derivative - liquidity pool token", "Other", "name / symbol 'Balancer'",
               "R2a-2")
        r["reasons"].append("R2a-2"); r["is_asset"] = False
    elif symbol in ("REG", "REUSD") or symbol.startswith("RTW-"):
        setcat("-", "Ecosystem token", "Other", "symbol REG / REUSD / RTW-", "R2a-3")
        r["reasons"].append("R2a-3"); r["is_asset"] = False
    elif m:
        code = m.group(1)
        detailed, main = CODE_MAP[code]
        basis = f"symbol code {code}" + (" (meaning unknown - assumption)" if code == "D" else "")
        setcat(code, detailed, main, basis, "R2b-1")
    elif OLDER_SCHEME_RE.match(symbol):
        setcat("none", "US property - older symbol scheme", "US rental property",
               "assumption: no code = S (name 'RealToken ...', US address)", "R2b-2")
    elif symbol.lower().startswith("realtoken s "):
        setcat("none", "US property", "US rental property",
               "assumption: name 'RealToken S ...' (symbol = full name)", "R2b-3")
        r["flag_symbol_is_name"] = True
    else:
        setcat("-", "Unclassified", "Other", "no known pattern", "R2b-4")
        r["manual"] = True

    # Flags (R2b-5)
    r["flag_holdings"] = bool(re.search(r"\b\d+ Holdings\b", name))
    conflict = False
    if r["is_asset"] and r["symbol_code"] not in ("-",):
        loan_name = "LOAN" in name.split()
        loan_code = r["symbol_code"] in ("SL", "PA-SEL")
        fact_name = "Factoring" in name
        fact_code = r["symbol_code"] == "PA-S"
        conflict = (loan_name != loan_code) or (fact_name != fact_code)
    r["flag_name_code_conflict"] = conflict
    r["rule_ids"].append("R2b-5")

    # US state code (checked by Excel formula against the list in sheet Rules)
    last = name.strip().split()[-1] if name.strip() else ""
    r["us_state_code"] = (last if r["category_detailed"].startswith("US")
                          and re.fullmatch(r"[A-Z]{2}", last) else "")
    return r


def property_key(name):
    n = re.sub(r"^(?i:realtoken)\s+", "", name.strip())
    n = re.sub(r"^(S|D)\s+", "", n)
    return re.sub(r"[^a-z0-9]", "", n.lower())


# ---------------------------------------------------------------------------
# 3. Supply (R2c) and duplicates (R2d)
# ---------------------------------------------------------------------------
def process(records):
    rows = []
    for rec in records:
        it = rec["item"]
        c = classify(it)
        dec = int(it.get("decimals") or 0)
        supply_raw = it.get("total_supply") or "0"
        row = {
            "raw_id": rec["raw_id"], "address_hash": it.get("address_hash"),
            "name": it.get("name"), "symbol": it.get("symbol"), "decimals": dec,
            "total_supply_raw": supply_raw,
            "supply_tokens": float(Decimal(int(supply_raw)) / (Decimal(10) ** dec)),
            "holders_count": int(it.get("holders_count") or 0),
            **{k: c[k] for k in ("symbol_code", "category_detailed", "category_main",
                                 "classification_basis", "us_state_code", "flag_holdings",
                                 "flag_symbol_is_name", "flag_name_code_conflict",
                                 "flag_symbol_whitespace")},
            "_rules": c["rule_ids"], "_reasons": c["reasons"], "_manual": c["manual"],
            "_asset": c["is_asset"],
            "property_key": property_key(it.get("name") or "") if c["is_asset"] else "",
            "dup_group_id": "", "dup_group_size": 1, "dup_role": "unique",
            "kept_address": "", "flag_dup_supply_differs": False,
        }
        if int(supply_raw) == 0:                                   # R2c-1
            row["_reasons"].append("R2c-1"); row["_rules"].append("R2c-1")
        rows.append(row)

    # Duplicate groups among property/asset tokens (R2d)
    groups = defaultdict(list)
    for row in rows:
        if row["_asset"]:
            groups[row["property_key"]].append(row)
    gid = 0
    for key, members in sorted(groups.items()):
        for mbr in members:
            mbr["dup_group_size"] = len(members)
            mbr["_rules"].append("R2d-1")
        if len(members) < 2:
            continue
        gid += 1
        for mbr in members:
            mbr["dup_group_id"] = f"G{gid:03d}"
        cands = [m for m in members if int(m["total_supply_raw"]) > 0]
        if len({m["total_supply_raw"] for m in cands}) > 1:      # R2d-3
            for mbr in members:
                mbr["flag_dup_supply_differs"] = True
                mbr["_rules"].append("R2d-3")
        if not cands:
            for mbr in members:
                mbr["dup_role"] = "no kept contract (all zero supply)"
            continue
        top = max(m["holders_count"] for m in cands)
        tops = [m for m in cands if m["holders_count"] == top]
        if len(tops) > 1:                                          # R2d-4 (tie)
            for mbr in members:
                mbr["dup_role"] = "review (tie)"
                mbr["_rules"].append("R2d-4")
                if mbr in tops:
                    mbr["_manual"] = True
            continue
        kept = tops[0]
        kept["dup_role"] = "kept"
        kept["_rules"].append("R2d-2")
        for mbr in members:
            if mbr is not kept:                                    # R2d-2
                mbr["dup_role"] = "duplicate"
                mbr["kept_address"] = kept["address_hash"]
                mbr["_reasons"].append("R2d-2"); mbr["_rules"].append("R2d-2")

    # Final status
    for row in rows:
        reasons = sorted(set(row["_reasons"]), key=RULE_ORDER.index)
        if reasons:
            row["status"] = "Excluded"
        elif row["_manual"]:
            row["status"] = "Manual_Check"
        else:
            row["status"] = "Included"
        row["primary_reason"] = REASON_TEXT[reasons[0]] if reasons else ""
        row["all_reasons"] = "; ".join(REASON_TEXT[r] for r in reasons)
        row["rule_ids"] = ", ".join(dict.fromkeys(row["_rules"] + ["R2e-1"]))
        row["your_decision"] = ""
        row["comment"] = ""
    return rows


# ---------------------------------------------------------------------------
# 4. Excel output
# ---------------------------------------------------------------------------
HEAD_FONT = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="305496")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
BOLD = Font(bold=True)


def header(ws, names, row=1):
    for j, n in enumerate(names, start=1):
        cell = ws.cell(row=row, column=j, value=n)
        cell.font, cell.fill = HEAD_FONT, HEAD_FILL


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def write_list_sheet(wb, title, rows):
    ws = wb.create_sheet(title)
    header(ws, COLNAMES)
    for i, row in enumerate(rows, start=2):
        for j, name in enumerate(COLNAMES, start=1):
            if name == "us_state_valid":
                cat, st = f"{L['category_detailed']}{i}", f"{L['us_state_code']}{i}"
                value = (f'=IF(LEFT({cat},2)="US",AND({st}<>"",'
                         f'COUNTIF(Rules!$H$2:$H${1 + len(US_STATES)},{st})>0),"")')
            else:
                value = row[name]
            cell = ws.cell(row=i, column=j, value=value)
            if name == "total_supply_raw":
                cell.number_format = "@"
    for name in ("your_decision", "comment"):
        ws[f"{L[name]}1"].fill = INPUT_FILL
        ws[f"{L[name]}1"].font = BOLD
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLNAMES))}{max(2, len(rows) + 1)}"
    widths(ws, {L["name"]: 45, L["symbol"]: 45, L["category_detailed"]: 32,
                L["classification_basis"]: 40, L["property_key"]: 34, L["address_hash"]: 44,
                L["all_reasons"]: 45, L["rule_ids"]: 40, L["primary_reason"]: 34,
                L["total_supply_raw"]: 28, L["kept_address"]: 44})
    if not rows:
        ws.cell(row=2, column=1, value=None)
    return ws


def write_raw_sheet(wb, records, field_order):
    ws = wb.create_sheet("Raw_Data")
    names = ["raw_id", "source_file", "position"] + field_order
    header(ws, names)
    for i, rec in enumerate(records, start=2):
        ws.cell(row=i, column=1, value=rec["raw_id"])
        ws.cell(row=i, column=2, value=rec["source_file"])
        ws.cell(row=i, column=3, value=rec["position"])
        for j, f in enumerate(field_order, start=4):
            v = rec["item"].get(f)
            # JSON strings as-is; other JSON values (null, true/false, numbers) as JSON text
            text = v if isinstance(v, str) else json.dumps(v)
            cell = ws.cell(row=i, column=j, value=text)
            cell.number_format = "@"
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(names))}{len(records) + 1}"
    return ws


def rng(sheet, col):
    return f"{sheet}!${L[col]}$2:${L[col]}${MAX_ROW}"


def write_summary(wb, manifest):
    ws = wb.create_sheet("Summary", 1)
    ws["A1"] = "Summary (all values are Excel formulas over the sheets Included / Excluded / Manual_Check)"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = f"Formula ranges cover rows 2-{MAX_ROW} of each sheet."

    # --- Automatic checks ---------------------------------------------------
    r = 4
    ws.cell(row=r, column=1, value="Automatic checks").font = BOLD
    r += 1
    header(ws, ["Check", "Result (TRUE = passed)", "Number behind the check", "Note"], row=r)
    ws.cell(row=r + 1, column=1, value="Records in manifest (fixed value from manifest.json)")
    ws.cell(row=r + 1, column=3, value=manifest["totals"]["items"])
    manifest_cell = f"$C${r + 1}"
    S = ["Included", "Excluded", "Manual_Check"]
    counta = "+".join(f"COUNTA({s}!$A$2:$A${MAX_ROW})" for s in S)
    once = "+".join(f"COUNTIF({s}!$A$2:$A${MAX_ROW},Raw_Data!$A$2:$A${MAX_ROW})" for s in S)
    pk = rng("Included", "property_key")
    checks = [
        ("1. All Included tokens have supply > 0",
         f"=C{{r}}=0",
         f'=SUMPRODUCT((Included!$A$2:$A${MAX_ROW}<>"")*({rng("Included", "supply_tokens")}<=0))',
         "Number of Included rows with supply <= 0."),
        ("2. All US tokens in Included have a valid US state code",
         f"=C{{r}}=0",
         f"=COUNTIF({rng('Included', 'us_state_valid')},FALSE)",
         "Number of Included US rows with us_state_valid = FALSE (filter that column to find them)."),
        ("3. Each property appears only once in Included",
         f"=C{{r}}=0",
         f'=SUMPRODUCT(({pk}<>"")*(COUNTIF({pk},{pk})>1))',
         "Number of Included rows whose property_key occurs more than once."),
        ("4. Included + Excluded + Manual_Check = records in manifest",
         f"=C{{r}}={manifest_cell}",
         f"={counta}",
         "Sum of rows in the three sheets."),
        ("5. Each raw record appears exactly once in the three sheets (additional check)",
         f"=C{{r}}=0",
         f'=SUMPRODUCT((Raw_Data!$A$2:$A${MAX_ROW}<>"")*(({once})<>1))',
         "Number of raw_ids appearing 0 or >1 times. Proposed by the AI in addition to checks 1-4."),
    ]
    first = r + 2
    for k, (label, result, number, note) in enumerate(checks):
        row = first + k
        ws.cell(row=row, column=1, value=label)
        ws.cell(row=row, column=2, value=result.format(r=row))
        ws.cell(row=row, column=3, value=number)
        ws.cell(row=row, column=4, value=note)
    last = first + len(checks) - 1
    ws.cell(row=last + 1, column=1, value="All checks passed").font = BOLD
    ws.cell(row=last + 1, column=2, value=f"=AND(B{first}:B{last})").font = BOLD

    # --- Main category x status ------------------------------------------------
    r = last + 4
    ws.cell(row=r, column=1, value="Records by main category and sheet").font = BOLD
    header(ws, ["category_main", "Included", "Excluded", "Manual_Check", "Total"], row=r + 1)
    r += 2
    top = r
    for main in MAIN_ORDER:
        ws.cell(row=r, column=1, value=main)
        for j, s in enumerate(S, start=2):
            ws.cell(row=r, column=j, value=f'=COUNTIF({rng(s, "category_main")},$A{r})')
        ws.cell(row=r, column=5, value=f"=SUM(B{r}:D{r})")
        r += 1
    ws.cell(row=r, column=1, value="Total").font = BOLD
    for j in range(2, 6):
        col = get_column_letter(j)
        ws.cell(row=r, column=j, value=f"=SUM({col}{top}:{col}{r - 1})").font = BOLD

    # --- Detailed category -------------------------------------------------------
    r += 3
    ws.cell(row=r, column=1, value="Records by detailed category").font = BOLD
    header(ws, ["category_detailed", "category_main", "Included", "Excluded", "Manual_Check",
                "Total", "Included: supply (tokens)", "Included: holders (sum)"], row=r + 1)
    r += 2
    top = r
    for det, main in CATEGORY_ORDER:
        ws.cell(row=r, column=1, value=det)
        ws.cell(row=r, column=2, value=main)
        for j, s in enumerate(S, start=3):
            ws.cell(row=r, column=j, value=f'=COUNTIF({rng(s, "category_detailed")},$A{r})')
        ws.cell(row=r, column=6, value=f"=SUM(C{r}:E{r})")
        ws.cell(row=r, column=7, value=f'=SUMIF({rng("Included", "category_detailed")},$A{r},'
                                       f'{rng("Included", "supply_tokens")})')
        ws.cell(row=r, column=8, value=f'=SUMIF({rng("Included", "category_detailed")},$A{r},'
                                       f'{rng("Included", "holders_count")})')
        r += 1
    ws.cell(row=r, column=1, value="Total").font = BOLD
    for j in range(3, 9):
        col = get_column_letter(j)
        ws.cell(row=r, column=j, value=f"=SUM({col}{top}:{col}{r - 1})").font = BOLD

    # --- Exclusion reasons -------------------------------------------------------
    r += 3
    ws.cell(row=r, column=1, value="Exclusion reasons (sheet Excluded)").font = BOLD
    header(ws, ["Reason", "Rule", "As primary reason", "Among all reasons"], row=r + 1)
    r += 2
    for rule in RULE_ORDER:
        ws.cell(row=r, column=1, value=REASON_TEXT[rule])
        ws.cell(row=r, column=2, value=rule)
        ws.cell(row=r, column=3, value=f'=COUNTIF({rng("Excluded", "primary_reason")},$A{r})')
        ws.cell(row=r, column=4, value=f'=COUNTIF({rng("Excluded", "all_reasons")},"*"&$A{r}&"*")')
        r += 1

    # --- Flags ---------------------------------------------------------------------
    r += 2
    ws.cell(row=r, column=1, value="Flags (number of rows with TRUE)").font = BOLD
    header(ws, ["Flag", "Included", "Excluded", "Manual_Check"], row=r + 1)
    r += 2
    for flag in ("flag_holdings", "flag_symbol_is_name", "flag_name_code_conflict",
                 "flag_symbol_whitespace", "flag_dup_supply_differs"):
        ws.cell(row=r, column=1, value=flag)
        for j, s in enumerate(S, start=2):
            ws.cell(row=r, column=j, value=f"=COUNTIF({rng(s, flag)},TRUE)")
        r += 1

    widths(ws, {"A": 62, "B": 24, "C": 24, "D": 70, "E": 14, "F": 12, "G": 24, "H": 22})
    return ws


def write_readme(wb, manifest, run_time, out_name):
    ws = wb.active
    ws.title = "README"
    pages = manifest["pages"]
    lines = [
        ("RealT token list - processed from frozen raw data", None),
        ("", None),
        ("Snapshot ID", SNAPSHOT),
        ("Source URL", manifest["base_url"]),
        ("Query parameters", json.dumps(manifest["query_params"])),
        ("Request headers", json.dumps(manifest.get("request_headers"))),
        ("Redirect of old domain (recorded at download)",
         json.dumps(manifest.get("old_domain_probe"))),
        ("Download started / finished (UTC)",
         f"{manifest['started_utc']} / {manifest['finished_utc']}"),
        ("Server date of first / last page", f"{pages[0].get('server_date')} / "
                                             f"{pages[-1].get('server_date')}"),
        ("Download script / version / SHA-256",
         f"{manifest['script']} / {manifest['script_version']} / {manifest['script_sha256']}"),
        ("Download status", f"{manifest['status']} - {manifest['stop_reason']}"),
        ("Pages / records in manifest", f"{manifest['totals']['pages']} / "
                                        f"{manifest['totals']['items']}"),
        ("Earlier (failed) download runs",
         "; ".join(f"{e['folder']}: {e['status']} (v{e.get('script_version')})"
                   for e in manifest.get("earlier_runs_in_raw_folder", []))),
        ("Checksum verification", f"All {len(pages)} page files matched manifest.json "
                                  f"(SHA-256) at {run_time}"),
        ("Processing script / version", f"{Path(__file__).name} / {SCRIPT_VERSION}"),
        ("Processing script SHA-256", hashlib.sha256(Path(__file__).read_bytes()).hexdigest()),
        ("File created (UTC)", run_time),
        ("File name", out_name),
        ("", None),
        ("Sheets", None),
        ("Summary", "Automatic checks (TRUE/FALSE) and counts; all values are Excel formulas."),
        ("Included", "Analysis population: records that pass all rules."),
        ("Excluded", "Records excluded by a rule, with primary reason, all reasons and rule IDs."),
        ("Manual_Check", "Records the rules cannot decide (unknown pattern, tie). "
                         "Empty if there are none."),
        ("Raw_Data", "All records exactly as delivered by the API, stored as text "
                     "(JSON null/true/false shown as 'null'/'true'/'false')."),
        ("Rules", "All rules with IDs, decision references and assumptions; list of US state codes."),
        ("", None),
        ("Limitations", None),
    ] + [(f"L{i}", t) for i, t in enumerate(LIMITATIONS, start=1)] + [
        ("", None),
        ("Assumptions (also marked in column classification_basis)", None),
    ] + [(rid, f"{desc} [{ass}]") for rid, _, desc, ass in RULES if ass != "no"] + [
        ("", None),
        ("Columns of Included / Excluded / Manual_Check", None),
    ] + COLS + [
        ("", None),
        ("Note", "Excel shows numbers with at most 15 significant digits. Raw integer values "
                 "(total_supply_raw, Raw_Data) are therefore stored as text."),
        ("", None),
        ("Known result of the automatic checks", None),
        ("Check 2 = FALSE (expected)",
         "Exactly one Included US token has no state code: 'RealToken S 17562 Mansfield St "
         "Detroit' (symbol REALTOKEN-S-17562-MANSFIELD-ST-DETROIT). The state code is missing "
         "in the source data (name and symbol). It is presumably MI (Detroit), but this is not "
         "filled in, so that the data remain as delivered (decision 4a). The check is "
         "intentionally left FALSE; 'All checks passed' is therefore FALSE as well."),
        ("Checks 1, 3, 4, 5", "Expected TRUE."),
    ]
    for i, (k, v) in enumerate(lines, start=1):
        ws.cell(row=i, column=1, value=k)
        if v is not None:
            ws.cell(row=i, column=2, value=v).alignment = Alignment(wrap_text=True,
                                                                    vertical="top")
        elif k:
            ws.cell(row=i, column=1).font = BOLD
    ws["A1"].font = Font(bold=True, size=13)
    widths(ws, {"A": 48, "B": 120})


def write_rules(wb):
    ws = wb.create_sheet("Rules")
    header(ws, ["rule_id", "decision", "description", "assumption"])
    for i, rule in enumerate(RULES, start=2):
        for j, v in enumerate(rule, start=1):
            ws.cell(row=i, column=j, value=v).alignment = Alignment(wrap_text=True,
                                                                    vertical="top")
    ws["H1"] = "US state codes (50 states + DC)"
    ws["H1"].font, ws["H1"].fill = HEAD_FONT, HEAD_FILL
    for i, s in enumerate(US_STATES, start=2):
        ws.cell(row=i, column=8, value=s)
    widths(ws, {"A": 10, "B": 10, "C": 100, "D": 40, "H": 30})


def main():
    out_name = f"RealT_tokens_{SNAPSHOT}_v{SCRIPT_VERSION}.xlsx"
    out_path = OUT_DIR / out_name
    if out_path.exists():
        sys.exit(f"{out_path} already exists - not overwritten. Rename or delete it first.")

    manifest, records, field_order = load_raw()
    rows = process(records)
    run_time = datetime.now(timezone.utc).isoformat(timespec="seconds")

    wb = Workbook()
    write_readme(wb, manifest, run_time, out_name)
    for status in ("Included", "Excluded", "Manual_Check"):
        write_list_sheet(wb, status, [r for r in rows if r["status"] == status])
    write_raw_sheet(wb, records, field_order)
    write_rules(wb)
    write_summary(wb, manifest)

    wb.calculation.fullCalcOnLoad = True  # Excel computes all formulas when opening
    OUT_DIR.mkdir(exist_ok=True)
    wb.save(out_path)

    # Console overview (computed in Python; the Excel file computes its own via formulas)
    print("\nRecords per sheet:")
    for s in ("Included", "Excluded", "Manual_Check"):
        print(f"  {s:13s} {sum(r['status'] == s for r in rows)}")
    print("Included by main category:")
    for m in MAIN_ORDER:
        print(f"  {m:20s} {sum(r['status'] == 'Included' and r['category_main'] == m for r in rows)}")
    print(f"\nSaved: {out_path.resolve()}")
    print("Open the file in Excel; the Summary sheet computes the checks on opening.")


if __name__ == "__main__":
    main()
