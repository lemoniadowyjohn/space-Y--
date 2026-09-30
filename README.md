# Python Excel Data Reconciliation Demo

[![CI](https://github.com/lemoniadowyjohn/space-Y--/actions/workflows/reconciliation-ci.yml/badge.svg?branch=master)](https://github.com/lemoniadowyjohn/space-Y--/actions/workflows/reconciliation-ci.yml)

> **Repository identity:** the slug `space-Y--` is historical. The active branch now contains the **Python Excel Data Reconciliation Demo**. Earlier coursework is preserved on the `legacy-space-y-coursework-2023` branch and is no longer part of the active portfolio surface.

A compact OpenPyXL project for reconciling structured workbooks by business key, normalizing values, detecting missing records and field-level mismatches, rejecting duplicate keys and generating a separate Excel evidence report.

![Synthetic reconciliation demo](docs/demo.svg)

## Problem

Operational and quality workflows often contain the same component or record in multiple spreadsheets. Manual comparison is slow, inconsistent and difficult to audit.

## Solution

```text
Workbook A ──┐
             ├─► key-indexed read
Workbook B ──┘
                  ↓
          duplicate-key validation
                  ↓
           value normalization
                  ↓
     missing-row + field comparison
                  ↓
      structured Difference records
                  ↓
       Excel detail + summary report
```

## Technology

- Python 3.11+
- OpenPyXL
- pytest
- GitHub Actions
- synthetic CSV/Excel fixtures

## Demonstrated behavior

- configurable key-column reconciliation;
- whitespace/case normalization for text;
- stable numeric normalization;
- missing-left and missing-right detection;
- field-level mismatch reporting;
- duplicate-key rejection;
- structured reconciliation result;
- Excel detail and summary report generation.

## Quick start

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -e ".[dev]"
pytest -q
python examples/build_demo.py
```

The demo generates synthetic workbooks and `reconciliation_report.xlsx`.

## Tests and CI

GitHub Actions performs:

1. editable package installation;
2. pytest regression tests;
3. reproducible demo/report generation.

The current suite checks mismatch/missing-row detection, duplicate-key failure and generated report content.

## Repository layout

```text
src/recon/          reconciliation logic
tests/              automated regression tests
data/               synthetic source records
examples/           workbook/report demo
docs/demo.svg       recruiter-facing visual
ARCHITECTURE.md      design overview
LIMITATIONS.md       known boundaries
SECURITY.md          public-data boundary
LICENSE              MIT license
pyproject.toml       packaging/dependencies
```

## Limitations

This is a portfolio implementation for structured worksheet reconciliation, not a general Excel ETL platform. It assumes a stable header row and cached formula values. Large workbooks would benefit from streaming/database-backed processing.

## Data and claim boundary

- all sample records are synthetic;
- no employer workbooks, customer data, production file paths or confidential reconciliation rules are included;
- the project demonstrates public portfolio code for Python/OpenPyXL reconciliation;
- the historical repository name is not part of the technical claim.

## Related portfolio

- [Governed Agent Workflow Demo](https://github.com/lemoniadowyjohn/space-Y-)
- [Industrial Quality Documentation Assistant](https://github.com/lemoniadowyjohn/hermes)
- [CARLA Map Quality Toolkit](https://github.com/lemoniadowyjohn/carla-control-suite)
- [Power Platform Quality App reference design](https://github.com/lemoniadowyjohn/watson)
