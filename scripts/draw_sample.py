"""
draw_sample.py  -  Step 5: random sample of 40 US rental tokens

Input  (read only): output/RealT_tokens_20261112T130912Z_v1.1.xlsx, sheet "Included"
Output (new file):  output/sample_20261112T130912Z_v1.0.xlsx  (never overwritten)

Decisions implemented (see decision log):
  5a  Sampling frame = Included tokens with category_main "US rental property"
      and holders_count >= 10. The tokens below the threshold are not removed
      silently: they are listed in sheet Frame_Excluded with the reason.
      (Snapshot 20261112T130912Z: holder counts jump from 3 to 62, so the
      threshold 10 lies in a visible gap of the data.)
  5b  Simple random sample, n = 40.
  5c  Reserve list: the whole frame is put in one random order. Ranks 1-40 are
      the sample; ranks 41+ are the reserve. If a sampled token has to be
      dropped later, it is replaced by the reserve token with the lowest rank
      not used yet. Every replacement is documented in the sheets.
  5d  Frame sorted by address_hash (lower case) before drawing, so the result
      does not depend on the row order in Excel. Seed 20261112, fixed before
      drawing. Python's random.Random(seed).shuffle is deterministic for a given
      seed (Python version recorded in README).

Requires: Python 3.9+, `pip install openpyxl`
Run:      python3 draw_sample.py
"""

import hashlib
import platform
import random
import sys
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SCRIPT_VERSION = "1.0"
INPUT_FILE = Path("output") / "RealT_tokens_20261112T130912Z_v1.1.xlsx"
INPUT_SHEET = "Included"
OUTPUT_FILE = Path("output") / f"sample_20261112T130912Z_v{SCRIPT_VERSION}.xlsx"

SEED = 20261112           # decision 5d, fixed before drawing
SAMPLE_SIZE = 40          # decision 5b
FRAME_CATEGORY = "US rental property"   # main category to analyse
MIN_HOLDERS = 10          # decision 5a
DROP_COLUMNS = {"us_state_valid", "your_decision", "comment"}  # formula / input columns

HEAD_FONT = Font(bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="305496")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
BOLD = Font(bold=True)


def read_included():
    if not INPUT_FILE.exists():
        sys.exit(f"Not found: {INPUT_FILE.resolve()}")
    sha = hashlib.sha256(INPUT_FILE.read_bytes()).hexdigest()
    # data_only=True: read stored cell values (some programs store TRUE/FALSE as
    # formulas); the formula column us_state_valid is not copied anyway.
    wb = load_workbook(INPUT_FILE, read_only=True, data_only=True)
    ws = wb[INPUT_SHEET]
    rows = list(ws.iter_rows(values_only=True))
    header = list(rows[0])
    records = [dict(zip(header, r)) for r in rows[1:] if r and r[0] is not None]
    wb.close()
    keep = [h for h in header if h not in DROP_COLUMNS]
    for rec in records:  # safety: the columns used for the frame must be filled
        for h in ("raw_id", "address_hash", "category_main", "holders_count"):
            if rec.get(h) in (None, ""):
                sys.exit(f"Empty value in column {h}, row {rec.get('raw_id')}.")
    return sha, keep, records


def header_row(ws, names, row=1):
    for j, n in enumerate(names, start=1):
        c = ws.cell(row=row, column=j, value=n)
        c.font, c.fill = HEAD_FONT, HEAD_FILL


def write_table(wb, title, front_cols, rows, cols, input_cols=()):
    ws = wb.create_sheet(title)
    names = list(front_cols) + list(input_cols) + cols
    header_row(ws, names)
    for i, r in enumerate(rows, start=2):
        for j, n in enumerate(names, start=1):
            ws.cell(row=i, column=j, value=r.get(n, ""))
    for n in input_cols:
        cell = ws.cell(row=1, column=names.index(n) + 1)
        cell.fill, cell.font = INPUT_FILL, BOLD
    ws.freeze_panes = ws.cell(row=2, column=len(front_cols) + 1)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(names))}{max(2, len(rows) + 1)}"
    for n, w in (("name", 45), ("symbol", 45), ("address_hash", 44), ("drop_reason", 30),
                 ("replaces_rank", 14), ("frame_exclusion_reason", 40)):
        if n in names:
            ws.column_dimensions[get_column_letter(names.index(n) + 1)].width = w
    return ws, names


def main():
    if OUTPUT_FILE.exists():
        sys.exit(f"{OUTPUT_FILE} already exists - not overwritten. Rename or delete it first.")

    sha, cols, included = read_included()

    # 5a: frame
    us = [r for r in included if r["category_main"] == FRAME_CATEGORY]
    frame = [r for r in us if int(r["holders_count"]) >= MIN_HOLDERS]
    below = [r for r in us if int(r["holders_count"]) < MIN_HOLDERS]
    for r in below:
        r["frame_exclusion_reason"] = f"holders_count < {MIN_HOLDERS} (decision 5a)"
    if len(frame) < SAMPLE_SIZE:
        sys.exit("Frame smaller than sample size.")

    # 5d: deterministic order, then 5b/5c: one random order for the whole frame
    frame.sort(key=lambda r: str(r["address_hash"]).lower())
    rng = random.Random(SEED)
    order = frame[:]
    rng.shuffle(order)
    for rank, r in enumerate(order, start=1):
        r["draw_rank"] = rank
    sample, reserve = order[:SAMPLE_SIZE], order[SAMPLE_SIZE:]

    run_time = datetime.now(timezone.utc).isoformat(timespec="seconds")
    wb = Workbook()

    # README
    ws = wb.active
    ws.title = "README"
    lines = [
        ("Random sample of US rental tokens", None),
        ("", None),
        ("Input file", str(INPUT_FILE)),
        ("Input file SHA-256", sha),
        ("Input sheet", INPUT_SHEET),
        ("Records in input sheet", len(included)),
        ("Of which main category 'US rental property'", len(us)),
        ("Sampling frame (holders_count >= 10)", len(frame)),
        ("Excluded from frame (holders_count < 10), see sheet Frame_Excluded", len(below)),
        ("Sample size", SAMPLE_SIZE),
        ("Reserve (frame minus sample), see sheet Reserve", len(reserve)),
        ("Seed", SEED),
        ("Method", "Frame sorted by address_hash (lower case); random.Random(seed).shuffle "
                   "over the whole frame; ranks 1-40 = sample, ranks 41+ = reserve."),
        ("Replacement rule (5c)", "If a sampled token must be dropped, enter the reason in "
                                  "Sample!drop_reason and use the reserve token with the lowest "
                                  "draw_rank not yet used; enter the replaced rank in "
                                  "Reserve!replaces_rank. Every replacement stays visible."),
        ("Unit of analysis (H3)", "Holder concentration and transfer activity over the last "
                                  "12 months, per token."),
        ("Python version", sys.version.split()[0]),
        ("Platform", platform.platform()),
        ("Sampling script / version", f"{Path(__file__).name} / {SCRIPT_VERSION}"),
        ("Sampling script SHA-256", hashlib.sha256(Path(__file__).read_bytes()).hexdigest()),
        ("File created (UTC)", run_time),
        ("", None),
        ("Note", "Columns are copied unchanged from the input sheet. Not copied: "
                 "us_state_valid (Excel formula referring to the input file), your_decision, "
                 "comment (empty input columns)."),
    ]
    for i, (k, v) in enumerate(lines, start=1):
        ws.cell(row=i, column=1, value=k).font = BOLD if v is None and k else Font()
        if v is not None:
            ws.cell(row=i, column=2, value=v).alignment = Alignment(wrap_text=True,
                                                                    vertical="top")
    ws["A1"].font = Font(bold=True, size=13)
    ws.column_dimensions["A"].width = 62
    ws.column_dimensions["B"].width = 110

    s_ws, s_names = write_table(wb, "Sample", ["draw_rank"], sample, cols,
                                input_cols=["drop_reason"])
    r_ws, r_names = write_table(wb, "Reserve", ["draw_rank"], reserve, cols,
                                input_cols=["replaces_rank"])
    write_table(wb, "Frame_Excluded", ["frame_exclusion_reason"], below, cols)

    # Checks (Excel formulas)
    def col(names, n):
        return get_column_letter(names.index(n) + 1)
    M = 2000
    sa, ra = col(s_names, "address_hash"), col(r_names, "address_hash")
    sh, rh = col(s_names, "holders_count"), col(r_names, "holders_count")
    sc, rc = col(s_names, "category_main"), col(r_names, "category_main")
    ws = wb.create_sheet("Checks", 1)
    ws["A1"] = "Automatic checks (Excel formulas)"
    ws["A1"].font = Font(bold=True, size=13)
    header_row(ws, ["Check", "Result (TRUE = passed)", "Number behind the check"], row=3)
    ws["A4"] = "Frame size (fixed value written by the script)"
    ws["C4"] = len(frame)
    checks = [
        ("1. Sample contains 40 tokens", "=C{r}=40",
         f"=COUNTA(Sample!${sa}$2:${sa}${M})"),
        ("2. Sample + Reserve = frame size", "=C{r}=$C$4",
         f"=COUNTA(Sample!${sa}$2:${sa}${M})+COUNTA(Reserve!${ra}$2:${ra}${M})"),
        ("3. No address appears in both Sample and Reserve", "=C{r}=0",
         f"=SUMPRODUCT((Sample!${sa}$2:${sa}${M}<>\"\")*"
         f"(COUNTIF(Reserve!${ra}$2:${ra}${M},Sample!${sa}$2:${sa}${M})>0))"),
        ("4. All Sample and Reserve tokens have >= 10 holders", "=C{r}=0",
         f"=SUMPRODUCT((Sample!${sa}$2:${sa}${M}<>\"\")*(Sample!${sh}$2:${sh}${M}<10))"
         f"+SUMPRODUCT((Reserve!${ra}$2:${ra}${M}<>\"\")*(Reserve!${rh}$2:${rh}${M}<10))"),
        ("5. All Sample tokens are 'US rental property'", "=C{r}=0",
         f"=SUMPRODUCT((Sample!${sa}$2:${sa}${M}<>\"\")*"
         f"(Sample!${sc}$2:${sc}${M}<>\"US rental property\"))"),
    ]
    for k, (label, res, num) in enumerate(checks):
        r = 5 + k
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=res.format(r=r))
        ws.cell(row=r, column=3, value=num)
    last = 4 + len(checks)
    ws.cell(row=last + 1, column=1, value="All checks passed").font = BOLD
    ws.cell(row=last + 1, column=2, value=f"=AND(B5:B{last})").font = BOLD
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 26

    wb.calculation.fullCalcOnLoad = True
    wb.save(OUTPUT_FILE)

    print(f"Input: {INPUT_FILE} (SHA-256 {sha[:16]}...)")
    print(f"US rental property: {len(us)} | frame (>= {MIN_HOLDERS} holders): {len(frame)} | "
          f"excluded from frame: {len(below)}")
    print(f"Seed {SEED}: sample {len(sample)}, reserve {len(reserve)}")
    print("First 5 of the sample (draw_rank, raw_id, name):")
    for r in sample[:5]:
        print(f"  {r['draw_rank']:3d}  {r['raw_id']}  {r['name']}")
    print(f"\nSaved: {OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()
