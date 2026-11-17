# RealT tokens on Gnosis – data and code

Data collection and preparation for the bachelor's thesis
*Tokenized Real Estate and the RWA Movement: A Mixed-Methods Study of
Expectations and Perceived Benefits*, University of St. Gallen (HSG),
Oleksandr Vasylenko, 2026.

> **Note on labels:** files and steps labelled "H3" refer to the liquidity analysis, which is hypothesis H1 in the final thesis.

**Case study:** RealT, whose real estate tokens (ERC-20) are issued on the
Gnosis blockchain. This repository covers the on-chain part of the thesis: the
list of RealT tokens, a random sample of 40 US rental property tokens, and
hypothesis H3 (holder concentration, transfer activity and turnover).

## Contents

| Folder / file | Contents |
|---|---|
| `scripts/` | The six Python scripts, in the versions that were run. Each file header gives its version and change log |
| `data/` | Raw API responses as zip archives, unmodified. The two download runs (token list, H3) each include a `manifest.json` (URLs, server times, attempts, SHA-256 per file); `explorations.zip` holds the single-token exploration only |
| `output/` | The three Excel files produced by the scripts |
| `docs/Decision_log_RealT.xlsx` | Every methodological decision: options, decision, rationale |
| `docs/AI_use.md` | Declaration of AI use |
| `SHA256SUMS.txt` | Checksums of all archives, scripts and output files |

## Data source

Blockscout API v2 for the Gnosis chain at `https://gnosisscan.io/api/v2/`. The
old address `gnosis.blockscout.com` redirects there. The API is public and needs
no key. All data were retrieved on **12 November 2026**. The RealT website
(realt.co) was not accessible at that time.

The data are publicly available on-chain data. Wallet addresses are
pseudonymous; no attempt was made to identify individuals.

## From 924 tokens to the sample

| Stage | Tokens | Rule |
|---|---|---|
| Search result for "RealToken" (ERC-20) | 924 | – |
| Included | 763 | Without derivatives, ecosystem tokens, zero-supply tokens and duplicate contracts |
| US rental property | 712 | Classified by symbol code |
| Sampling frame | 702 | At least 10 holders |
| Random sample | 40 | Seed 20261112; ranks 41–702 kept as a reserve (none needed) |
| H3 before/after comparison | 27 | Traded before 1 July 2024 |

## Steps

| # | Script | Reads | Writes |
|---|---|---|---|
| 1 | `download_raw.py` v1.2 | API | `raw/20261112T130912Z/` (19 pages) |
| 2 | `build_excel.py` v1.1 | step 1 | `output/RealT_tokens_20261112T130912Z_v1.1.xlsx` |
| 3 | `draw_sample.py` v1.0 | step 2 | `output/sample_20261112T130912Z_v1.0.xlsx` |
| 4 | `explore_h3.py` | step 3 + API | `explore_h3/` (one token, exploration only) |
| 5 | `download_h3.py` v1.1 | step 3 + API | `raw_h3/20261112T165352Z/` (holder lists, transfers since 1 Jan 2024) |
| 6 | `process_h3.py` v1.4 | step 5 | `output/H3_measures_20261112T165352Z_v1.4.xlsx` |

Two earlier download attempts are kept as documented failures in
`data/raw_failed_runs.zip`: run `20261112T130141Z` (HTTP 403) and run
`20261112T130609Z` (page 1 returned again). The causes and fixes are in the
decision log (1g, 1h) and in the change log of `download_raw.py`.

## Reproducing

1. Install Python 3.9 or later, then run `pip install -r requirements.txt`.
2. Unzip the archives in `data/` into the repository root, so that `raw/`,
   `raw_h3/` and `explore_h3/` lie next to `scripts/`.
3. Run the scripts from the root, e.g. `python3 scripts/build_excel.py`.

Existing output files are never overwritten: delete or rename them first.

- **Steps 2, 3 and 6 are fully reproducible.**
  - Steps 2 and 6 check every raw file against its SHA-256 before processing and stop if anything differs.
  - Step 3 reads the Excel file from step 2 and records that file's checksum. With the fixed seed, it draws the same 40 tokens.
  - A rerun of step 2 gives the same content but a different file checksum, because the file contains its creation time.
- **Steps 1, 4 and 5 cannot be repeated identically,** because the blockchain and the API keep changing. The stored raw data are therefore the basis of all results.

To check the files in this repository, run `shasum -a 256 -c SHA256SUMS.txt` in
the root.

## Logs and checks

- **Download logs:** the `manifest.json` in each raw folder.
- **Processing logs:** each Excel file has a `README` sheet (input files, checksums, script version, definitions, limitations) and a sheet of automatic checks (TRUE/FALSE).
- **Known check results** (documented and intentional):
  - `RealT_tokens_…_v1.1.xlsx`, check 2 FALSE: one token has no US state code in the source.
  - `H3_measures_…_v1.4.xlsx`, check 2 FALSE: for 15 tokens the listed holder balances are up to 0.59 % below total supply.
- **H3 holder concentration is reported in two variants:**
  - all addresses (main result);
  - without the address tagged as RealT in Blockscout, the RealTokenWrapper v3.
- **Decisions:** `docs/Decision_log_RealT.xlsx`.

## Use of AI

The scripts were written with the help of Claude (Anthropic). All
methodological decisions, all script runs and all checks were made by the
author. See `docs/AI_use.md`.
