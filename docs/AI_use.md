# Use of AI tools (code and data)

This file covers the use of AI for the code and data in this repository.
The use of AI tools for the thesis as a whole is listed in the thesis
(list of aids / Hilfsmittelverzeichnis).

**Tool:** Claude (Anthropic), model Claude Opus 5.5, used through the Claude
app in November 2026. Claude Code was used for one early test.

**Purpose:** support with designing and documenting the data collection and
data preparation for the on-chain analysis (token list, sampling, liquidity
measures). Steps and files labelled "H3" refer to the liquidity analysis,
which is hypothesis H1 in the final thesis.

**What the AI did:**

- Explained the Blockscout API: its fields, units and how paging works.
- Proposed options with pros, cons and a recommendation at every
  methodological step (data source, selection rules, classification,
  duplicate contracts, zero supply, Excel structure, sampling frame and method,
  liquidity measures, before/after periods).
- Wrote the Python scripts to the author's decisions and tested them on
  simulated data and on files the author uploaded.
- Analyzed uploaded outputs, e.g. the causes of failed downloads and of a
  failed data check.
- Drafted text passages for the methods section, which the author reviewed
  and revised.
- Kept the decision log (`Decision_log_RealT.xlsx`).

**What the author did:**

- Made all methodological decisions (see the decision log).
- Ran all scripts on their own computer. The AI had no direct access to the
  API; all data come from the author's own downloads.
- Checked the results: reproduction of the outputs and the automatic checks
  in the Excel files.
- Interpreted the results and wrote the thesis.
