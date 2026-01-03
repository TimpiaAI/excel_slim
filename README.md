
**excel_slim** is a lightweight, open-source **Excel file optimizer** that reduces workbook size safely and transparently. It is designed to be used **like a normal Python library** (similar to `openpyxl`, but for compression), with a fast Rust core under the hood.

It works by analyzing how an Excel file is built (ZIP container + XML + optional VBA + embedded media) and applying **modular, lossless-first optimizations** to shrink the file without breaking it.

## Why it exists

Most “Excel compressor” repos solve only one slice of the problem (ZIP repack, or VBA, or data compression). excel_slim aims to cover the full stack, with clear reporting and safe defaults.

## Features

* ✅ **Python-first API**

  * `import excel_slim`
  * `excel_slim.optimize("input.xlsx", output="output.xlsx")`
  * `excel_slim.analyze("input.xlsx")`
* ✅ **Lossless by default** (no visual or data changes)
* ✅ **Format support**

  * `.xlsx` and `.xlsm` (full support)
  * `.xls` (planned / partial)
* ✅ **Modular optimization pipeline**

  * ZIP container recompression (deterministic repack)
  * XML optimization (shared strings and styles dedupe, minify)
  * VBA stream compression for `.xlsm` (optional)
  * Embedded media optimization (lossless; lossy optional)
* ✅ **Transparent reporting**

  * Total savings plus per-module byte breakdown
  * JSON report support for automation pipelines
* ✅ **CLI + library**

  * Use it in scripts, CI, ETL pipelines, or from terminal

## Safety principles

* Lossless-first defaults
* Lossy steps require explicit opt-in
* Never executes macros
* Offline-only (no network calls)
* Deterministic output whenever possible

## Quick start (Python)

```python
import excel_slim

report = excel_slim.optimize("input.xlsx", output="output.xlsx", profile="safe", report=True)
print(report)
```

## Quick start (CLI)

```bash
excel_slim input.xlsx --auto --report
```

## Project structure

* Rust core for speed and reliability
* Python bindings (PyO3) for `pip install` and `import excel_slim`
* CLI that mirrors the Python options to avoid drift
