"""
process_h3.py  -  H3: measures for the 40 sample tokens (processing only)

Reads ONLY a frozen run folder raw_h3/<timestamp>/ written by download_h3.py.
Nothing is downloaded. Every page file is checked against the SHA-256 in
manifest.json first; the script stops if anything does not match or if a
token is not complete.

Output: output/H3_measures_<run>_v<SCRIPT_VERSION>.xlsx  (never overwritten)
Sheets: README, Summary, Measures, Monthly, Holders, Transfers, Python_Check

Definitions (decision log):
  H3b / 4f  Counted transfers: all transfers EXCEPT mint and burn. Mint/burn =
            type token_minting / token_burning OR sender/recipient is the zero
            address. Duplicates (same transaction hash + log index) are shown
            but not counted. Transfers before 1 Jan 2024 (from the last
            downloaded page) are shown but not counted.
  H3a       12-month window: [2025-11-12 00:00 UTC, 2026-11-12 00:00 UTC).
  4c        Activity in window: number of counted transfers and volume (tokens).
  4a        Turnover = volume in window / total supply.
  H3c / 4b  Top-1 and top-10 share of the current holder list (download time),
            denominator = total supply; check: holder balances add up to supply.
  H3d       All addresses treated the same, including contracts (main result).
  H3d-2     Additional variant: top-1 and top-10 share WITHOUT holder addresses
            that Blockscout tags as RealT (tag metadata projectName = "RealT"; in
            run 20261112T165352Z only the RealTokenWrapper v3 Proxy). Denominator
            = total supply (author's decision). Their share of supply is reported
            too. Every RealT-tagged address found is listed in sheet RealT_Tagged
            and in the console, so any address besides the wrapper is visible.
  4d / 4i   Monthly counted transfers Jan 2024 - Nov 2026. Before = Jul 2024 -
            Jun 2025 (12 months, decision 4i); Jan - Jun 2024 shown but in no
            period; event month July 2025 excluded (lawsuit filed
            2 Jul 2025); after = Aug 2025 - Oct 2026 (15 months). Nov 2026 is
            incomplete and in neither period.
  4e (A)    Descriptive: per token average monthly transfers before/after and %
            change; across tokens median of % changes and pooled totals.
  4h (C)    Only tokens with a complete 18-month before period enter the
            before/after comparison. Complete = the download contains at least
            one transfer from before 1 Jul 2024 (the token was already traded
            then). Other tokens stay in turnover and concentration (C1).
            If 15 or more tokens would be left out, the script STOPS before
            writing the Excel file and lists them (author's rule). To continue
            anyway, run again with --confirm-leftout; this is recorded in README.
  4g        Every measure is also computed in Python; sheet Python_Check
            compares it with the Excel formula (TRUE = equal).

Known result of the automatic checks (run raw_h3/20261112T165352Z, decision 4k):
  Check 2 "Holder balances add up to total supply" is FALSE for 15 tokens
  (draw_ranks 6, 7, 11, 12, 17, 18, 19, 24, 25, 27, 29, 32, 33, 37, 39). The
  balances are always BELOW the supply (at most 0.59 %, token 18: 24 of 4,100),
  mostly by round amounts. The holder lists are complete (holder count = Blockscout
  counter, no duplicate addresses) and no downloaded transfer (2024 on) goes to
  the zero address. Hypothesis, not verified: the tokens sit at an address that
  Blockscout does not list (probably the zero address) from transfers before 2024.
  The denominator stays total supply (decision 4b); the gap is shown per token in
  Measures!unlisted_tokens. Largest effect on a top-10 share: 0.27 percentage
  points. The check is intentionally left FALSE.

Requires: Python 3.9+, `pip install openpyxl`
Run:      python3 process_h3.py                (uses the only complete run)
          python3 process_h3.py --confirm-leftout   (only after the stop above)
          python3 process_h3.py --run 20261112T170000Z
"""

import argparse
import hashlib
import json
import statistics
import sys
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SCRIPT_VERSION = "1.4"   # 1.1: before period Jul 2024 - Jun 2025 (decision 4i)
                         # 1.2: column unlisted_tokens; check 2 documented as known
                         #      exception (decision 4k); no change to any measure
                         # 1.3: additional concentration variant without RealT-tagged
                         #      holder addresses (decision H3d-2); main result unchanged
                         # 1.4: H3d-2 denominator = total supply (author's decision);
                         #      every RealT-tagged address is listed (sheet RealT_Tagged)
RAW_ROOT = Path("raw_h3")
OUT_DIR = Path("output")

ZERO = "0x0000000000000000000000000000000000000000"
MINT_BURN_TYPES = {"token_minting", "token_burning"}
SINCE = "2024-01-01T00:00:00"
BEFORE_START = "2024-07-01T00:00:00"   # decision 4i
WINDOW = ("2025-11-12T00:00:00", "2026-11-12T00:00:00")
MONTHS = [f"{y}-{m:02d}" for y in (2024, 2025, 2026) for m in range(1, 13)
          if f"{y}-{m:02d}" <= "2026-11"]
BEFORE = [m for m in MONTHS if "2024-07" <= m <= "2025-06"]
EVENT = ["2025-07"]
AFTER = [m for m in MONTHS if "2025-08" <= m <= "2026-10"]
INCOMPLETE = ["2026-11"]
TOL = 1e-9
LEFTOUT_STOP = 15          # decision 4h: stop if this many tokens or more are left out

HEAD_FONT = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="305496")
BOLD = Font(bold=True)


# ---------------------------------------------------------------------------
# 1. Load and verify the frozen run
# ---------------------------------------------------------------------------
def pick_run(arg):
    if arg:
        run = RAW_ROOT / arg
    else:
        runs = []
        for d in sorted(RAW_ROOT.iterdir()):
            m = d / "manifest.json"
            if d.is_dir() and m.exists():
                if json.loads(m.read_text(encoding="utf-8")).get("status") == "complete":
                    runs.append(d)
        if len(runs) != 1:
            sys.exit(f"Found {len(runs)} complete runs in {RAW_ROOT}: "
                     f"{[r.name for r in runs]}. Choose one with --run <folder>.")
        run = runs[0]
    if not (run / "manifest.json").exists():
        sys.exit(f"No manifest.json in {run}")
    return run


def load_run(run):
    manifest = json.loads((run / "manifest.json").read_text(encoding="utf-8"))
    tokens = []
    n_files = 0
    for t in manifest["tokens"]:
        if t["status"] != "complete":
            sys.exit(f"Token draw_rank {t['draw_rank']} is '{t['status']}' in this run. "
                     "Replace it from the reserve first (decision 5c / H3e).")
        folder = run / t["folder"]
        data = {"meta": t, "holders": [], "transfers": []}
        for kind, pages in t["pages"].items():
            for p in pages:
                content = (folder / p["file"]).read_bytes()
                if hashlib.sha256(content).hexdigest() != p["sha256"]:
                    sys.exit(f"Checksum mismatch: {t['folder']}/{p['file']}. Stopping.")
                n_files += 1
                js = json.loads(content)
                if kind in ("token", "counters"):
                    data[kind] = js
                else:
                    for pos, item in enumerate(js.get("items", []), start=1):
                        data[kind].append((p["file"], pos, item))
        tokens.append(data)
    print(f"Run {run.name}: {len(tokens)} tokens, {n_files} files, checksums verified.")
    return manifest, sorted(tokens, key=lambda d: d["meta"]["draw_rank"])


# ---------------------------------------------------------------------------
# 2. Build rows (holders, transfers) and Python reference values
# ---------------------------------------------------------------------------
def to_tokens(raw, dec):
    return float(Decimal(int(raw or 0)) / (Decimal(10) ** dec))


def is_realt_tagged(a):
    """H3d-2: True if Blockscout tags the address as belonging to project RealT."""
    tags = (((a or {}).get("metadata") or {}).get("tags") or [])
    return any(((t.get("meta") or {}).get("projectName") == "RealT") for t in tags)


def addr_info(a):
    a = a or {}
    tags = ((a.get("metadata") or {}).get("tags") or [])
    return ((a.get("hash") or ""), bool(a.get("is_contract")), a.get("name") or "",
            "; ".join(t.get("name", "") for t in tags))


def period_of(month):
    if month in BEFORE:
        return "before"
    if month in AFTER:
        return "after"
    if month in EVENT:
        return "event month (excluded)"
    if month in INCOMPLETE:
        return "incomplete month (excluded)"
    return ""


def build(tokens):
    holders_rows, transfer_rows, measures = [], [], []
    for d in tokens:
        m, tok = d["meta"], d["token"]
        addr = m["address_hash"]
        dec = int(tok.get("decimals") or 0)
        supply = to_tokens(tok.get("total_supply"), dec)

        # Holders: sorted by balance (desc), ties by address -> rank
        hs = []
        for f, pos, it in d["holders"]:
            h, c, n, tg = addr_info(it.get("address"))
            hs.append({"token": addr, "draw_rank": m["draw_rank"], "holder": h,
                       "is_contract": c, "name": n, "tags": tg,
                       "is_realt_tagged": is_realt_tagged(it.get("address")),
                       "balance_raw": str(it.get("value")),
                       "balance_tokens": to_tokens(it.get("value"), dec),
                       "supply_tokens": supply, "source": f"{f}#{pos}"})
        hs.sort(key=lambda r: (-int(r["balance_raw"]), r["holder"].lower()))
        for i, r in enumerate(hs, start=1):
            r["rank"] = i
        k = 0
        for r in hs:                                   # H3d-2: rank without RealT tags
            if r["is_realt_tagged"]:
                r["rank_excl_realt"] = None
            else:
                k += 1
                r["rank_excl_realt"] = k
        holders_rows += hs
        realt_bal = sum(r["balance_tokens"] for r in hs if r["is_realt_tagged"])

        # Transfers
        seen, ts_rows = set(), []
        for f, pos, it in d["transfers"]:
            fh, fc, fn, _ = addr_info(it.get("from"))
            th, tc, tn, _ = addr_info(it.get("to"))
            ts = (it.get("timestamp") or "")
            total = it.get("total") or {}
            tdec = int(total.get("decimals") or dec)
            key = (it.get("transaction_hash"), it.get("log_index"))
            dup = key in seen
            seen.add(key)
            pre_start = ts[:19] < BEFORE_START
            mint_burn = (it.get("type") in MINT_BURN_TYPES
                         or fh.lower() == ZERO or th.lower() == ZERO)
            pre2024 = ts[:19] < SINCE
            counted = not (mint_burn or dup or pre2024)
            month = ts[:7]
            ts_rows.append({
                "token": addr, "draw_rank": m["draw_rank"], "timestamp": ts, "month": month,
                "block_number": it.get("block_number"), "log_index": it.get("log_index"),
                "transaction_hash": it.get("transaction_hash"),
                "from": fh, "from_is_contract": fc, "from_name": fn,
                "to": th, "to_is_contract": tc, "to_name": tn,
                "amount_raw": str(total.get("value")),
                "amount_tokens": to_tokens(total.get("value"), tdec),
                "type": it.get("type"), "method": it.get("method"),
                "is_mint_burn": mint_burn, "is_duplicate": dup, "before_2024": pre2024,
                "before_jul_2024": pre_start,
                "counted": counted,
                "in_window": counted and WINDOW[0] <= ts[:19] < WINDOW[1],
                "period": period_of(month) if counted else "",
                "source": f"{f}#{pos}"})
        transfer_rows += ts_rows

        # Python reference values (same definitions as the Excel formulas)
        win = [r for r in ts_rows if r["in_window"]]
        vol = sum(r["amount_tokens"] for r in win)
        n_bef = sum(r["period"] == "before" for r in ts_rows)
        n_aft = sum(r["period"] == "after" for r in ts_rows)
        full = any(r["before_jul_2024"] for r in ts_rows)
        bef, aft = n_bef / len(BEFORE), n_aft / len(AFTER)
        measures.append({
            "draw_rank": m["draw_rank"], "raw_id": m["raw_id"], "token": addr,
            "name": m["name"], "supply_tokens": supply,
            "counters_holders": (d.get("counters") or {}).get("token_holders_count"),
            "py": {
                "holders_in_list": len(hs),
                "balances_share": (sum(r["balance_tokens"] for r in hs) / supply) if supply else 0,
                "unlisted_tokens": supply - sum(r["balance_tokens"] for r in hs),
                "top1_share": (sum(r["balance_tokens"] for r in hs if r["rank"] == 1) / supply)
                if supply else 0,
                "top10_share": (sum(r["balance_tokens"] for r in hs if r["rank"] <= 10) / supply)
                if supply else 0,
                "realt_tagged_share": realt_bal / supply if supply else 0,
                "top1_excl_realt": (sum(r["balance_tokens"] for r in hs
                                        if r["rank_excl_realt"] == 1) / supply) if supply else 0,
                "top10_excl_realt": (sum(r["balance_tokens"] for r in hs
                                         if (r["rank_excl_realt"] or 99) <= 10) / supply)
                if supply else 0,
                "transfers_12m": len(win), "volume_12m": vol,
                "turnover_12m": vol / supply if supply else 0,
                "full_before_period": full,
                "transfers_before": n_bef, "transfers_after": n_aft,
                "avg_month_before": bef, "avg_month_after": aft,
                "change_pct": (aft / bef - 1) if (full and bef) else "",
            }})
    return holders_rows, transfer_rows, measures


# ---------------------------------------------------------------------------
# 3. Excel
# ---------------------------------------------------------------------------
def header(ws, names, row=1):
    for j, n in enumerate(names, start=1):
        c = ws.cell(row=row, column=j, value=n)
        c.font, c.fill = HEAD_FONT, HEAD_FILL


def write_rows(ws, cols, rows):
    header(ws, cols)
    for i, r in enumerate(rows, start=2):
        for j, c in enumerate(cols, start=1):
            cell = ws.cell(row=i, column=j, value=r[c])
            if c.endswith("_raw"):
                cell.number_format = "@"
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(2, len(rows) + 1)}"


H_COLS = ["token", "draw_rank", "rank", "rank_excl_realt", "holder", "is_contract",
          "is_realt_tagged", "name", "tags",
          "balance_raw", "balance_tokens", "supply_tokens", "share", "source"]
T_COLS = ["token", "draw_rank", "timestamp", "month", "block_number", "log_index",
          "transaction_hash", "from", "from_is_contract", "from_name", "to", "to_is_contract",
          "to_name", "amount_raw", "amount_tokens", "type", "method", "is_mint_burn",
          "is_duplicate", "before_2024", "before_jul_2024", "counted", "in_window", "period", "source"]
MEASURES = [  # (column, label, number format)
    ("holders_in_list", "Holders in list", "0"),
    ("balances_share", "Sum of balances / supply (check = 1)", "0.000000"),
    ("unlisted_tokens", "Supply minus sum of listed balances (tokens, 4k)", "#,##0.000000"),
    ("top1_share", "Top-1 share", "0.00%"),
    ("top10_share", "Top-10 share", "0.00%"),
    ("realt_tagged_share", "Share held by RealT-tagged addresses (H3d-2)", "0.00%"),
    ("top1_excl_realt", "Top-1 share without RealT-tagged (H3d-2)", "0.00%"),
    ("top10_excl_realt", "Top-10 share without RealT-tagged (H3d-2)", "0.00%"),
    ("transfers_12m", "Transfers in 12-month window", "0"),
    ("volume_12m", "Volume in window (tokens)", "#,##0.00"),
    ("turnover_12m", "Turnover (volume / supply)", "0.0000"),
    ("full_before_period", "Complete before period (transfer before 1 Jul 2024 exists)", "General"),
    ("transfers_before", "Transfers before (Jul24-Jun25)", "0"),
    ("transfers_after", "Transfers after (Aug25-Oct26)", "0"),
    ("avg_month_before", "Avg monthly transfers before (Jul24-Jun25)", "0.00"),
    ("avg_month_after", "Avg monthly transfers after (Aug25-Oct26)", "0.00"),
    ("change_pct", "Change after vs before", "0.0%"),
]


def col_letter(cols, name):
    return get_column_letter(cols.index(name) + 1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--run", help="folder name in raw_h3/")
    ap.add_argument("--confirm-leftout", action="store_true",
                    help=f"continue although >= {LEFTOUT_STOP} tokens are left out (4h)")
    args = ap.parse_args()
    run = pick_run(args.run)
    out = OUT_DIR / f"H3_measures_{run.name}_v{SCRIPT_VERSION}.xlsx"
    if out.exists():
        sys.exit(f"{out} already exists - not overwritten. Rename or delete it first.")

    manifest, tokens = load_run(run)
    holders, transfers, measures = build(tokens)

    # Decision 4h: stop before any output if too many tokens are left out
    left_out = [(m["draw_rank"], m["name"]) for m in measures
                if not m["py"]["full_before_period"]]
    if len(left_out) >= LEFTOUT_STOP and not args.confirm_leftout:
        print(f"\nSTOP (decision 4h): {len(left_out)} of {len(measures)} tokens have no "
              f"transfer before 1 Jul 2024 and would be left out of the before/after "
              f"comparison (threshold {LEFTOUT_STOP}):")
        for r_, n_ in left_out:
            print(f"  {r_:3d}  {n_}")
        sys.exit("No Excel file written. Decide how to proceed; to continue as defined, run "
                 "again with --confirm-leftout.")
    nH, nT, nM = len(holders) + 1, len(transfers) + 1, len(measures) + 1

    wb = Workbook()
    readme = wb.active
    readme.title = "README"
    summary = wb.create_sheet("Summary")
    ms = wb.create_sheet("Measures")
    mo = wb.create_sheet("Monthly")
    hsh = wb.create_sheet("Holders")
    tsh = wb.create_sheet("Transfers")
    pc = wb.create_sheet("Python_Check")

    # --- Holders & Transfers (data sheets) ----------------------------------
    for r in holders:
        r["share"] = None
    write_rows(hsh, H_COLS, holders)
    hb, hs_, hsh_c = (col_letter(H_COLS, "balance_tokens"), col_letter(H_COLS, "supply_tokens"),
                      col_letter(H_COLS, "share"))
    for i in range(2, nH + 1):
        hsh[f"{hsh_c}{i}"] = f"=IF({hs_}{i}=0,\"\",{hb}{i}/{hs_}{i})"
        hsh[f"{hsh_c}{i}"].number_format = "0.0000%"
    write_rows(tsh, T_COLS, transfers)

    def H(c):
        L = col_letter(H_COLS, c)
        return f"Holders!${L}$2:${L}${nH}"

    def T(c):
        L = col_letter(T_COLS, c)
        return f"Transfers!${L}$2:${L}${nT}"

    # --- Measures (Excel formulas) --------------------------------------------
    m_cols = ["draw_rank", "raw_id", "token", "name", "supply_tokens",
              "counters_holders"] + [c for c, _, _ in MEASURES]
    header(ms, m_cols)
    for i, m in enumerate(measures, start=2):
        for j, c in enumerate(m_cols[:6], start=1):
            ms.cell(row=i, column=j, value=m[c])
        tok, sup = f"$C{i}", f"$E{i}"
        f = {
            "holders_in_list": f"=COUNTIF({H('token')},{tok})",
            "balances_share": f"=IF({sup}=0,0,SUMIFS({H('balance_tokens')},{H('token')},{tok})/{sup})",
            "unlisted_tokens": f"={sup}-SUMIFS({H('balance_tokens')},{H('token')},{tok})",
            "top1_share": f"=IF({sup}=0,0,SUMIFS({H('balance_tokens')},{H('token')},{tok},"
                          f"{H('rank')},1)/{sup})",
            "top10_share": f"=IF({sup}=0,0,SUMIFS({H('balance_tokens')},{H('token')},{tok},"
                           f"{H('rank')},\"<=10\")/{sup})",
            "realt_tagged_share": f"=IF({sup}=0,0,SUMIFS({H('balance_tokens')},{H('token')},"
                                  f"{tok},{H('is_realt_tagged')},TRUE)/{sup})",
        }
        dr = sup          # H3d-2 denominator: total supply (decision of the author)
        f.update({
            "top1_excl_realt": f"=IF({dr}=0,0,SUMIFS({H('balance_tokens')},{H('token')},{tok},"
                               f"{H('rank_excl_realt')},1)/{dr})",
            "top10_excl_realt": f"=IF({dr}=0,0,SUMIFS({H('balance_tokens')},{H('token')},{tok},"
                                f"{H('rank_excl_realt')},\"<=10\")/{dr})",
            "transfers_12m": f"=COUNTIFS({T('token')},{tok},{T('in_window')},TRUE)",
            "volume_12m": f"=SUMIFS({T('amount_tokens')},{T('token')},{tok},{T('in_window')},TRUE)",
        })
        L = {c: col_letter(m_cols, c) for c in m_cols}
        f["turnover_12m"] = f"=IF({sup}=0,0,{L['volume_12m']}{i}/{sup})"
        f["full_before_period"] = f"=COUNTIFS({T('token')},{tok},{T('before_jul_2024')},TRUE)>0"
        f["transfers_before"] = f"=COUNTIFS({T('token')},{tok},{T('period')},\"before\")"
        f["transfers_after"] = f"=COUNTIFS({T('token')},{tok},{T('period')},\"after\")"
        f["avg_month_before"] = f"={L['transfers_before']}{i}/{len(BEFORE)}"
        f["avg_month_after"] = f"={L['transfers_after']}{i}/{len(AFTER)}"
        f["change_pct"] = (f"=IF(OR(NOT({L['full_before_period']}{i}),"
                           f"{L['avg_month_before']}{i}=0),\"\","
                           f"{L['avg_month_after']}{i}/{L['avg_month_before']}{i}-1)")
        for c, _, fmt in MEASURES:
            cell = ms[f"{L[c]}{i}"]
            cell.value = f[c]
            cell.number_format = fmt
    ms.freeze_panes = "E2"
    ms.column_dimensions["D"].width = 45
    ms.column_dimensions["C"].width = 44
    ms.auto_filter.ref = f"A1:{get_column_letter(len(m_cols))}{nM}"
    ML = {c: col_letter(m_cols, c) for c in m_cols}

    def M(c):
        return f"Measures!${ML[c]}$2:${ML[c]}${nM}"

    # --- Monthly ----------------------------------------------------------------
    header(mo, ["draw_rank", "token", "name"] + MONTHS)
    mo.cell(row=2, column=1, value="period")
    for j, mth in enumerate(MONTHS, start=4):
        mo.cell(row=2, column=j, value=period_of(mth) or "-").font = Font(italic=True)
    for i, m in enumerate(measures, start=3):
        mo.cell(row=i, column=1, value=m["draw_rank"])
        mo.cell(row=i, column=2, value=m["token"])
        mo.cell(row=i, column=3, value=m["name"])
        for j in range(4, 4 + len(MONTHS)):
            col = get_column_letter(j)
            mo.cell(row=i, column=j, value=f"=COUNTIFS({T('token')},$B{i},{T('month')},"
                                           f"{col}$1,{T('counted')},TRUE)")
    tot = len(measures) + 3
    mo.cell(row=tot, column=3, value="Total (40 tokens)").font = BOLD
    for j in range(4, 4 + len(MONTHS)):
        col = get_column_letter(j)
        mo.cell(row=tot, column=j, value=f"=SUM({col}3:{col}{tot - 1})").font = BOLD
    mo.freeze_panes = "D3"
    mo.column_dimensions["B"].width = 44
    mo.column_dimensions["C"].width = 45

    # --- Python_Check ---------------------------------------------------------
    pcols = ["draw_rank", "token"]
    for c, _, _ in MEASURES:
        pcols += [f"{c}_python", f"{c}_excel", f"{c}_match"]
    header(pc, pcols)
    for i, m in enumerate(measures, start=2):
        pc.cell(row=i, column=1, value=m["draw_rank"])
        pc.cell(row=i, column=2, value=m["token"])
        for c, _, _ in MEASURES:
            pL, eL, kL = (col_letter(pcols, f"{c}_python"), col_letter(pcols, f"{c}_excel"),
                          col_letter(pcols, f"{c}_match"))
            pc[f"{pL}{i}"] = m["py"][c]
            pc[f"{eL}{i}"] = f"=Measures!{ML[c]}{i}"
            if c == "full_before_period":
                pc[f"{kL}{i}"] = f"={eL}{i}={pL}{i}"
            else:
                pc[f"{kL}{i}"] = (f"=IF(AND({pL}{i}=\"\",{eL}{i}=\"\"),TRUE,IF(OR({pL}{i}=\"\","
                                  f"{eL}{i}=\"\"),FALSE,ABS({eL}{i}-{pL}{i})<={TOL}*"
                                  f"MAX(1,ABS({pL}{i}))))")
    match_cols = [col_letter(pcols, f"{c}_match") for c, _, _ in MEASURES]

    # --- Summary ------------------------------------------------------------------
    s = summary
    s["A1"] = "H3 summary (Excel formulas over Measures / Monthly / Holders / Transfers)"
    s["A1"].font = Font(bold=True, size=13)
    header(s, ["Automatic checks", "Result (TRUE = passed)", "Number behind the check", "Note"],
           row=3)
    mismatch = "+".join(f"COUNTIF(Python_Check!${c}$2:${c}${nM},FALSE)" for c in match_cols)
    dups = f"COUNTIFS({T('is_duplicate')},TRUE)"
    checks = [
        ("1. 40 tokens in Measures", f"=C{{r}}=40", f"=COUNTA(Measures!$C$2:$C${nM})", ""),
        ("2. Holder balances add up to total supply for every token (4b)", "=C{r}=0",
         f"=SUMPRODUCT(--(ABS({M('balances_share')}-1)>1E-9))",
         "Number of tokens where sum of balances / supply differs from 1. Known exception "
         "(4k): see README and Measures!unlisted_tokens."),
        ("3. Holders in list = Blockscout counter for every token", "=C{r}=0",
         f"=SUMPRODUCT(--({M('holders_in_list')}<>VALUE({M('counters_holders')})))",
         "Informational: the counter may lag; a difference means the list should be inspected."),
        ("4. No duplicate transfers (same transaction hash + log index)", "=C{r}=0",
         f"={dups}", "Duplicates are shown in Transfers but never counted."),
        ("5. Excel formulas = independent Python calculation (4g)", "=C{r}=0",
         f"={mismatch}", "Number of FALSE cells in sheet Python_Check."),
    ]
    for k, (lab, res, num, note) in enumerate(checks):
        r = 4 + k
        s.cell(row=r, column=1, value=lab)
        s.cell(row=r, column=2, value=res.format(r=r))
        s.cell(row=r, column=3, value=num)
        s.cell(row=r, column=4, value=note)
    last = 3 + len(checks)
    s.cell(row=last + 1, column=1, value="All checks passed").font = BOLD
    s.cell(row=last + 1, column=2, value=f"=AND(B4:B{last})").font = BOLD

    r = last + 4
    header(s, ["Results across the 40 tokens", "Value", "", "Definition"], row=r)
    full = M("full_before_period")
    pooled_b = f"SUMIFS({M('transfers_before')},{full},TRUE)"
    pooled_a = f"SUMIFS({M('transfers_after')},{full},TRUE)"
    results = [
        ("Median turnover (12-month window)", f"=MEDIAN({M('turnover_12m')})", "0.0000",
         "Volume in window / total supply (4a)."),
        ("Mean turnover", f"=AVERAGE({M('turnover_12m')})", "0.0000", ""),
        ("Median top-1 share", f"=MEDIAN({M('top1_share')})", "0.00%", "Holder list at download (H3c)."),
        ("Median top-10 share", f"=MEDIAN({M('top10_share')})", "0.00%", ""),
        ("Tokens with a RealT-tagged holder (H3d-2)",
         f"=COUNTIF({M('realt_tagged_share')},\">0\")", "0",
         "Blockscout tag projectName = RealT (RealTokenWrapper v3 Proxy)."),
        ("Median share held by RealT-tagged addresses", f"=MEDIAN({M('realt_tagged_share')})",
         "0.00%", "Of total supply."),
        ("Median top-1 share without RealT-tagged", f"=MEDIAN({M('top1_excl_realt')})", "0.00%",
         "Variant H3d-2: RealT-tagged addresses left out; denominator = total supply."),
        ("Median top-10 share without RealT-tagged", f"=MEDIAN({M('top10_excl_realt')})",
         "0.00%", "Variant H3d-2."),
        ("Median transfers in window", f"=MEDIAN({M('transfers_12m')})", "0.0", ""),
        ("Tokens in before/after comparison (complete before period, 4h)",
         f"=COUNTIF({full},TRUE)", "0", "Tokens with at least one transfer before 1 Jul 2024."),
        ("Tokens left out of the comparison (history starts after 1 Jul 2024)",
         f"=COUNTIF({full},FALSE)", "0", "Still included in turnover and concentration."),
        ("Total counted transfers before (Jul 2024 - Jun 2025)", f"={pooled_b}", "0",
         f"{len(BEFORE)} months; comparison tokens only"),
        ("Total counted transfers after (Aug 2025 - Oct 2026)", f"={pooled_a}", "0",
         f"{len(AFTER)} months; comparison tokens only"),
        ("Pooled avg per month before", f"={pooled_b}/{len(BEFORE)}", "0.00", ""),
        ("Pooled avg per month after", f"={pooled_a}/{len(AFTER)}", "0.00", ""),
        ("Pooled change after vs before",
         f"=IF({pooled_b}=0,\"\",({pooled_a}/{len(AFTER)})/({pooled_b}/{len(BEFORE)})-1)",
         "0.0%", "Monthly averages, comparison tokens together (4e, 4h)."),
        ("Median of per-token changes", f"=MEDIAN({M('change_pct')})", "0.0%",
         "Comparison tokens only; empty cells are ignored."),
        ("Comparison tokens with 0 transfers before (change not defined)",
         f"=COUNTIFS({full},TRUE,{M('avg_month_before')},0)", "0", ""),
    ]
    for k, (lab, fml, fmt, note) in enumerate(results, start=1):
        s.cell(row=r + k, column=1, value=lab)
        c = s.cell(row=r + k, column=2, value=fml)
        c.number_format = fmt
        s.cell(row=r + k, column=4, value=note)
    s.column_dimensions["A"].width = 62
    s.column_dimensions["B"].width = 24
    s.column_dimensions["C"].width = 24
    s.column_dimensions["D"].width = 80

    # --- RealT_Tagged (H3d-2): every RealT-tagged holder address found --------------
    rt = wb.create_sheet("RealT_Tagged")
    tagged = {}
    for r in holders:
        if r["is_realt_tagged"]:
            e = tagged.setdefault(r["holder"].lower(), {"holder": r["holder"], "name": r["name"],
                                                        "tags": r["tags"], "shares": []})
            e["shares"].append(r["balance_tokens"] / r["supply_tokens"] if r["supply_tokens"] else 0)
    header(rt, ["holder", "name", "tags", "tokens held in (formula)",
                "max share of supply (Python)"])
    for i, e in enumerate(sorted(tagged.values(), key=lambda e: -len(e["shares"])), start=2):
        rt.cell(row=i, column=1, value=e["holder"])
        rt.cell(row=i, column=2, value=e["name"])
        rt.cell(row=i, column=3, value=e["tags"])
        rt.cell(row=i, column=4, value=f'=COUNTIFS({H("holder")},A{i},{H("is_realt_tagged")},TRUE)')
        c = rt.cell(row=i, column=5, value=max(e["shares"]))
        c.number_format = "0.00%"
    if not tagged:
        rt.cell(row=2, column=1, value="No RealT-tagged holder address found.")
    rt.cell(row=len(tagged) + 3, column=1,
            value=f"{len(tagged)} RealT-tagged address(es) found. Rule: Blockscout tag "
                  "metadata projectName = 'RealT'.").font = BOLD
    for col, w in (("A", 46), ("B", 24), ("C", 36), ("D", 22), ("E", 26)):
        rt.column_dimensions[col].width = w

    # --- README -----------------------------------------------------------------
    run_time = datetime.now(timezone.utc).isoformat(timespec="seconds")
    lines = [
        ("H3 measures - holder concentration, transfer activity, turnover", None),
        ("", None),
        ("Raw data folder", str(run)),
        ("Download script / version", f"{manifest.get('script')} / "
                                      f"{manifest.get('script_version')}"),
        ("Download started / finished (UTC)", f"{manifest.get('started_utc')} / "
                                              f"{manifest.get('finished_utc')}"),
        ("Sample file / SHA-256", f"{manifest.get('sample_file')} / "
                                  f"{manifest.get('sample_file_sha256')}"),
        ("Checksums", f"All page files matched manifest.json at {run_time}"),
        ("Processing script / version", f"{Path(__file__).name} / {SCRIPT_VERSION}"),
        ("Processing script SHA-256", hashlib.sha256(Path(__file__).read_bytes()).hexdigest()),
        ("File created (UTC)", run_time),
        ("", None),
        ("Definitions", None),
        ("Counted transfer (H3b, 4f)", "All transfers except mint/burn (type token_minting / "
         "token_burning, or sender/recipient = zero address). Duplicates and transfers before "
         "1 Jan 2024 are shown but not counted."),
        ("12-month window (H3a)", f"[{WINDOW[0]} UTC, {WINDOW[1]} UTC)"),
        ("Activity (4c)", "Number of counted transfers and volume (tokens) in the window."),
        ("Turnover (4a)", "Volume in window / total supply."),
        ("Concentration (H3c, 4b)", "Share of the largest and of the ten largest holders in the "
         "holder list at download time; denominator = total supply (token.json)."),
        ("Holders (H3d)", "Main result: all addresses treated the same, including contracts "
         "(e.g. RealTokenWrapper, DEX pools, RMM)."),
        ("Variant without RealT (H3d-2)", "Top-1 / top-10 share without holder addresses that "
         "Blockscout tags as RealT (tag projectName = 'RealT'); denominator = total supply. "
         "Columns is_realt_tagged and rank_excl_realt in Holders show which addresses are left "
         "out; sheet RealT_Tagged lists every such address found."),
        ("Periods (4d, 4i)", "Before = Jul 2024 - Jun 2025 (12 months; Jan - Jun 2024 shown in "
         "Monthly but in no period); event month "
         "Jul 2025 excluded; after = Aug 2025 - Oct 2026 (15 months); Nov 2026 incomplete, "
         "excluded."),
        ("Comparison (4e)", "Descriptive: average monthly transfers before/after per token and "
         "pooled; median of per-token % changes."),
        ("Stop rule (4h)", f"Stop if >= {LEFTOUT_STOP} tokens are left out. Left out in this run: "
         f"{len(left_out)} (draw_ranks {[r_ for r_, _ in left_out]}). "
         f"--confirm-leftout used: {'yes' if args.confirm_leftout else 'no'}."),
        ("Comparison tokens (4h)", "Only tokens with a complete before period: the download "
         "contains at least one transfer from before 1 Jul 2024. Tokens whose history starts "
         "later are left out of the before/after comparison but stay in all other measures."),
        ("Python cross-check (4g)", "Sheet Python_Check compares every measure with an "
         "independent Python calculation."),
        ("", None),
        ("Known result of the automatic checks (decision 4k)", None),
        ("Check 2 = FALSE (expected)",
         "In run 20261112T165352Z, 15 tokens (draw_ranks 6, 7, 11, 12, 17, 18, 19, 24, 25, 27, "
         "29, 32, 33, 37, 39) have listed holder balances below the total supply (at most "
         "0.59 %, token 18: 24 of 4,100 tokens; mostly round amounts). The holder lists are "
         "complete (holder count = Blockscout counter, no duplicates) and no downloaded transfer "
         "goes to the zero address. Hypothesis (not verified): the tokens sit at an address "
         "Blockscout does not list, probably the zero address, from transfers before 2024. The "
         "denominator stays total supply (4b); the gap per token is in Measures!unlisted_tokens. "
         "The check is intentionally left FALSE, so 'All checks passed' is FALSE as well."),
        ("Checks 1, 3, 4, 5", "Expected TRUE."),
        ("", None),
        ("Events", None),
        ("2 Jul 2025", "City of Detroit files nuisance abatement lawsuit against RealT "
         "(press reports; primary source cited in the thesis)."),
        ("22 Jul 2025", "Temporary restraining order; rent into escrow from August 2025 "
         "(press reports)."),
        ("2 Jul 2026", "Liquidation announcement as reported; falls inside the after period."),
        ("", None),
        ("Limitations", None),
        ("L-H3-1", "Turnover includes DeFi movements (RMM deposits/withdrawals, wrapper, DEX "
                   "pools), because transfers are not split by type (H3b)."),
        ("L-H3-2", "The holder list reflects the download time, not the snapshot time "
                   "(12 Nov 2026, 13:09 UTC)."),
        ("L-H3-3", "Largest holders can be protocol contracts (e.g. RealTokenWrapper)."),
        ("L-H3-4", "The after period (Aug 2025 - Oct 2026) also contains the liquidation "
                   "announcement of July 2026."),
        ("L-H3-5", "The API field 'method' is not stable between requests and is not used."),
        ("L-H3-6", "The before/after comparison covers only tokens with a complete before period "
                   "(4h); tokens first transferred after 1 Jul 2024 are left out of it."),
        ("L-H3-9", "Variant H3d-2 leaves out only addresses with a RealT tag in Blockscout. "
                   "Untagged contracts with large balances (e.g. 'SoraPFPProxy') are not "
                   "classified and stay in. The owners behind the RealTokenWrapper are not "
                   "visible in the holder list."),
        ("L-H3-8", "For 15 tokens the listed holder balances are below the total supply (max. "
                   "0.59 %, see Measures!unlisted_tokens). Shares use total supply as "
                   "denominator (4b, 4k); using the sum of balances instead would raise a "
                   "top-10 share by at most 0.27 percentage points."),
        ("L-H3-7", "The before period was shortened from Jan 2024 - Jun 2025 to Jul 2024 - Jun "
                   "2025 after the download (decision 4i), because 17 of 40 tokens had no "
                   "transfer before 2024. The change rests only on issuance dates; no transfer "
                   "counts had been evaluated."),
        ("", None),
        ("Sheets", None),
        ("Summary", "Automatic checks and results across tokens (Excel formulas)."),
        ("Measures", "One row per token; all measures are Excel formulas over Holders/Transfers."),
        ("Monthly", "Counted transfers per token and month (COUNTIFS), all 40 tokens; row 2 = period. "
                    "Comparison tokens are marked in Measures, column full_before_period."),
        ("Holders", "All holder rows; rank 1 = largest balance; share = balance / supply."),
        ("Transfers", "All downloaded transfers with flags (mint/burn, duplicate, before 2024, "
                      "counted, in window, period)."),
        ("Python_Check", "Python value vs Excel value per token and measure."),
        ("RealT_Tagged", "Every holder address with a Blockscout RealT tag (H3d-2)."),
    ]
    for i, (k, v) in enumerate(lines, start=1):
        readme.cell(row=i, column=1, value=k)
        if v is None and k:
            readme.cell(row=i, column=1).font = BOLD
        if v is not None:
            readme.cell(row=i, column=2, value=v).alignment = Alignment(wrap_text=True,
                                                                        vertical="top")
    readme["A1"].font = Font(bold=True, size=13)
    readme.column_dimensions["A"].width = 38
    readme.column_dimensions["B"].width = 120

    wb.calculation.fullCalcOnLoad = True
    OUT_DIR.mkdir(exist_ok=True)
    wb.save(out)

    # Console (Python values)
    py = [m["py"] for m in measures]
    ch = [p["change_pct"] for p in py if p["change_pct"] != ""]
    out_ranks = [m["draw_rank"] for m in measures if not m["py"]["full_before_period"]]
    print(f"Before/after comparison: {len(measures) - len(out_ranks)} tokens; "
          f"left out (history starts after 1 Jul 2024): {out_ranks}")
    print(f"Holders rows: {len(holders)} | transfer rows: {len(transfers)} "
          f"(counted: {sum(t['counted'] for t in transfers)}, "
          f"in window: {sum(t['in_window'] for t in transfers)})")
    print(f"Median turnover: {statistics.median(p['turnover_12m'] for p in py):.4f}")
    print(f"Median top-1 / top-10 share: {statistics.median(p['top1_share'] for p in py):.2%} / "
          f"{statistics.median(p['top10_share'] for p in py):.2%}")
    print(f"Without RealT-tagged (H3d-2): tokens with such a holder "
          f"{sum(p['realt_tagged_share'] > 0 for p in py)}, median share held "
          f"{statistics.median(p['realt_tagged_share'] for p in py):.2%}, median top-1 / top-10 "
          f"{statistics.median(p['top1_excl_realt'] for p in py):.2%} / "
          f"{statistics.median(p['top10_excl_realt'] for p in py):.2%}")
    print(f"RealT-tagged holder addresses found: {len(tagged)}"
          + ("  <-- more than the RealTokenWrapper, check sheet RealT_Tagged" if len(tagged) > 1
             else ""))
    for e in sorted(tagged.values(), key=lambda e: -len(e["shares"])):
        print(f"  {e['holder']}  {e['tags']}  in {len(e['shares'])} tokens, "
              f"max {max(e['shares']):.2%} of supply")
    print(f"Median change after vs before: {statistics.median(ch):.1%}" if ch else
          "Median change: not defined")
    print(f"\nSaved: {out.resolve()}")


if __name__ == "__main__":
    main()
