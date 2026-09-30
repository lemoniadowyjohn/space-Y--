# Python Excel Data Reconciliation Demo

A compact, reproducible OpenPyXL project for comparing two structured Excel workbooks by business key, normalizing values, detecting missing rows and value mismatches, rejecting duplicate keys and generating a separate reconciliation report.

![Demo](docs/demo.svg)

## Why this exists

Operational and quality workflows often involve the same component or record appearing in multiple spreadsheets. Manual comparison is slow and error-prone. This project demonstrates a deterministic reconciliation pipeline that separates data-quality rules from reporting.

## Features

- workbook comparison by configurable key column;
- whitespace/case normalization for text;
- stable numeric normalization;
- missing-left / missing-right detection;
- field-level mismatch reporting;
- duplicate-key rejection;
- Excel report with detail and summary sheets;
- synthetic sample data only;
- pytest + GitHub Actions CI.

## Run

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -e ".[dev]"
pytest -q
python examples/build_demo.py
```

The demo script generates two synthetic workbooks and `reconciliation_report.xlsx`.

## Claim boundary

This is public portfolio code designed to demonstrate Python/OpenPyXL reconciliation patterns. It contains no employer workbooks, production paths, customer data or confidential business rules.
