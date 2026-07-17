# TenderPro

TenderPro is a production-oriented Streamlit web application for Procurement Engineers who compare commercial BOQ offers from multiple EPC/MEP vendors.

## Phase 1 Features

- Modern Streamlit interface with a clear TenderPro landing page.
- New Project, Upload Vendor Files, and Generate Comparison workflow.
- Multiple Excel upload support using Pandas and OpenPyXL.
- Automatic sample BOQ generation for Enext and Valtria when no upload files exist.
- Flexible column detection for vendor spreadsheets with inconsistent headers.
- Normalization of all vendor BOQs into one master dataframe.
- Commercial analysis showing:
  - Lowest offer by line item.
  - Missing prices.
  - Vendor ranking.
  - Amount and percentage variance against the lowest offer.
- Excel export with normalized data, side-by-side comparison, and vendor ranking sheets.

## Project Structure

```text
TenderPro/
├── app.py
├── requirements.txt
├── README.md
├── assets/
├── templates/
├── uploads/
├── exports/
├── modules/
│   ├── excel_reader.py
│   ├── mapping.py
│   ├── comparison_engine.py
│   ├── dashboard.py
│   └── export.py
```

## Supported Column Mappings

TenderPro automatically detects common BOQ column names and maps them to a canonical schema.

| Canonical Field | Accepted Headers |
| --- | --- |
| Item No | `Item No`, `Item`, `Code`, `Item Number`, `Item Code` |
| Description | `Description`, `Item Description`, `Desc`, `Scope` |
| Unit | `Unit`, `UOM`, `Measure`, `Unit of Measure` |
| Quantity | `Quantity`, `Qty`, `Q'ty`, `Quant` |
| Unit Rate | `Unit Rate`, `Price`, `Rate`, `Unit Price` |
| Total Amount | `Total Amount`, `Amount`, `Total`, `Line Total` |
| Vendor | `Vendor`, `Supplier`, `Bidder`, `Company` |

## How to Run

1. Create and activate a Python virtual environment:

   ```bash
   python -m venv .venv
   source .venv/bin/activate
   ```

2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Start the Streamlit app:

   ```bash
   streamlit run app.py
   ```

4. Open the local Streamlit URL shown in your terminal, usually <http://localhost:8501>.

## Typical Workflow

1. Enter a project name in the sidebar.
2. Upload multiple vendor Excel BOQ files, or use the generated Enext and Valtria samples.
3. Click **Generate Comparison**.
4. Review the KPI cards, vendor ranking, side-by-side comparison, and detailed variance analysis.
5. Download the generated Excel comparison workbook.

## Development Notes

- Uploaded vendor files are stored in `uploads/`.
- Generated comparison workbooks are stored in `exports/`.
- Sample files are created only when no `.xlsx` files are present in `uploads/`.
- The app uses OpenPyXL as the Excel engine for reading and writing `.xlsx` workbooks.
