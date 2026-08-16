---
name: chinese-firm-panel-cleaning
description: Clean and merge Chinese listed-firm annual panel Excel data from CSMAR, CNRDS, or similar sources. Use when the user provides a main Excel path and a folder of other Excel files and wants date normalization, year-end filtering, enterprise-year keys, one-to-one panel merging, missing-value deletion, automatic exclusion of binary 0/1 variables from winsorization, 1%-99% winsorization, and audit logs.
---

# Chinese Firm Panel Cleaning

Use this skill for repeatable CSMAR/CNRDS-style enterprise panel preparation. The user normally supplies only the main-file path and the folder containing other Excel files.

## Default contract

- Input files are Excel workbooks; the main file is the sample spine and all other files are left-joined to it.
- The user manually names the stock-code column std and the date column date. Do not invent or silently rename substantive variables.
- Parse common Chinese and ISO date forms, keep only observations dated December 31, and insert integer year immediately after date.
- Normalize std by removing only leading zeros (000001 -> 1, 000300 -> 300, 600000 -> 600000); never pad to six digits.
- Construct string key yearstd = string(year) + string(std) with no separator (2017 + 1 -> 20171).
- Require yearstd to be unique in every file. Never silently drop duplicates; export them and stop for review.
- Merge one-to-one on yearstd, retaining every main-file observation. Preserve same-named using variables with a source suffix and report them.
- After merging, drop rows with missing values in explicitly configured analysis variables.
- Winsorize configured numeric variables at the 1st and 99th percentiles. Preserve originals and create variables with suffix _w.
- Skip pure binary variables whose nonmissing support is a subset of {0, 1}; do not create redundant _w columns.
- Never winsorize identifiers, dates, years, keys, industry/region codes, or other categorical variables.
- Save cleaned individual files, the final panel, merge/missing/winsorization logs, and duplicate records.

## Required interaction

Accept a request such as:

    使用企业面板数据清洗技能。
    主表：D:\\project\\main.xlsx
    其他Excel文件夹：D:\\project\\raw

Ask only for missing information that cannot be safely inferred, especially the analysis-variable list. Do not guess which substantive variables are central to the paper.

## Execution workflow

1. Validate the main path and scan the other folder for Excel files, excluding the main file and generated outputs.
2. Read workbooks without modifying raw files. Clean only whitespace and newline artifacts in column labels.
3. Require std and date; parse dates after replacing Chinese date characters, slash, and dot separators. Report unparseable rows.
4. Keep only month 12 and day 31. Insert year after date.
5. Clean stock codes by extracting the numeric code and applying lstrip("0"); use "0" if the result is empty. Build yearstd as a string.
6. Check key uniqueness. If duplicates exist, export a duplicate-record file and stop rather than deduplicating.
7. Save each cleaned source. Left-merge all sources into the main data using one-to-one validation and an indicator. Record matched and unmatched counts and rates.
8. Delete missing observations for configured analysis variables and record before and after counts.
9. For each configured winsorization variable, coerce to numeric, skip pure 0/1 variables, otherwise clip to global 1% and 99% quantiles. Keep the original and add suffix _w.
10. Validate final key uniqueness, missingness, date/year range, row count, and output columns. Save Excel, CSV, and audit logs.

Prefer the bundled scripts/process_panel.py for deterministic processing. It accepts --main, --others-dir, --analysis-vars, and --winsor-vars; use --header-row for leading notes. If the user asks for Stata, generate an equivalent do-file with the same rules and checks.
