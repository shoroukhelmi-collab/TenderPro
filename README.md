# TenderPro

TenderPro is a basic Streamlit MVP for supplier BOQ comparison.

## What the MVP does

- Upload multiple supplier Excel files (`.xlsx`).
- Detect common BOQ columns automatically and let users review supplier names, worksheets, header rows, and mappings before processing.
- Normalize uploaded BOQs into a master query.
- Compare suppliers item by item, including lowest prices, missing prices, variance, and visible outlier flags.
- Show dashboard KPIs, supplier summaries, package summaries, missing items, and outliers.
- Exclude or restore detected outliers from totals without deleting raw data.
- Export one Excel workbook with dashboard, master query, summaries, pivots, missing items, outliers, and raw data.

The app intentionally does **not** include AI, login, templates, databases, or advanced architecture. Supplier names are generic (`Supplier 1`, `Supplier 2`, etc.) so the comparison stays simple and anonymized.

## Supported columns

TenderPro detects common variants of these BOQ fields:

| Field | Example headers |
| --- | --- |
| Item number | `Item No`, `Item`, `Item Number`, `Item Code`, `Code` |
| Description | `Description`, `Item Description`, `Desc`, `Scope`, `Particulars` |
| Unit | `Unit`, `UOM`, `Measure`, `Unit of Measure` |
| Quantity | `Quantity`, `Qty`, `Q'ty`, `BOQ Qty` |
| Unit rate | `Unit Rate`, `Rate`, `Price`, `Unit Price` |
| Total amount | `Total Amount`, `Amount`, `Total`, `Line Total` |

If total amount is blank and quantity plus unit rate are present, TenderPro calculates total amount automatically.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud

This repository includes the required `requirements.txt`. Deploy the app with:

- Main file path: `app.py`
- Python dependencies: installed from `requirements.txt`

## Project structure

```text
TenderPro/
├── app.py
├── requirements.txt
├── README.md
├── modules/
│   ├── comparison_engine.py
│   ├── dashboard.py
│   ├── excel_reader.py
│   ├── export.py
│   └── mapping.py
└── tests/
    └── test_comparison_engine.py
```
