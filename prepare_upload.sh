#!/bin/bash
# prepare_upload.sh - fills this folder with the scripts, raw data and outputs
# from the thesis working folder. Copies only: the working folder is not changed.
#
# Run from inside this folder:
#   bash prepare_upload.sh "/Users/oleksandr/Downloads/Thesis Raw Data"
set -euo pipefail

SRC="${1:-$HOME/Downloads/Thesis Raw Data}"
DST="$(cd "$(dirname "$0")" && pwd)"
[ -d "$SRC/raw" ] || { echo "Not found: $SRC/raw - give the thesis folder as argument."; exit 1; }
mkdir -p "$DST/scripts" "$DST/data" "$DST/output"
missing=()

need() {   # copy a required file
  if [ -f "$SRC/$1" ]; then cp "$SRC/$1" "$DST/$2/"; else missing+=("$1"); fi
}
pack() {   # zip folders (relative to SRC) into data/<name>.zip; skip folders that don't exist
  local name="$1"; shift; local parts=()
  for p in "$@"; do if [ -e "$SRC/$p" ]; then parts+=("$p"); else missing+=("$p"); fi; done
  [ ${#parts[@]} -eq 0 ] && return
  [ -e "$DST/data/$name.zip" ] && { echo "Exists, not overwritten: data/$name.zip"; return; }
  (cd "$SRC" && zip -rqX "$DST/data/$name.zip" "${parts[@]}" -x "*.DS_Store")
  echo "data/$name.zip  <- ${parts[*]}"
}

# Scripts: the versions that were run in the working folder
for f in download_raw.py build_excel.py draw_sample.py explore_h3.py download_h3.py process_h3.py; do
  need "$f" scripts
done

# Raw data, unmodified, as zip archives
pack raw_tokenlist_20261112T130912Z raw/20261112T130912Z
pack raw_failed_runs                raw/20261112T130141Z raw/20261112T130609Z
pack raw_h3_20261112T165352Z        raw_h3/20261112T165352Z
pack explorations                   explore explore_h3

# Output files produced by the scripts
for f in RealT_tokens_20261112T130912Z_v1.1.xlsx sample_20261112T130912Z_v1.0.xlsx \
         H3_measures_20261112T165352Z_v1.4.xlsx; do
  need "output/$f" output
done

# Checksums of everything that goes up
cd "$DST"
shasum -a 256 data/*.zip scripts/*.py output/*.xlsx > SHA256SUMS.txt
echo; echo "SHA256SUMS.txt written:"; cat SHA256SUMS.txt
if [ ${#missing[@]} -gt 0 ]; then
  echo; echo "NOT FOUND (check names, or leave out on purpose):"; printf '  %s\n' "${missing[@]}"
fi
