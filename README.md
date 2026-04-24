# Australia Leading Index

Python scripts for building Australian business-cycle and leading-index dashboards from local Australian Bureau of Statistics (ABS) workbooks.

The repository currently contains three dashboard builders:

- `build_business_cycle_indices.py` builds fixed-basket leading, coincident, and lagging quarterly business-cycle indices.
- `build_cyclical_dashboard.py` builds a cyclical GDP dashboard using current-price contribution weights and real activity signals.
- `build_high_frequency_leading_index.py` builds a monthly high-frequency leading index from household spending categories.

Each script writes a self-contained Plotly HTML dashboard.

## Requirements

- Python 3.10 or newer
- `pandas`
- `numpy`
- `openpyxl`
- `plotly`

The scripts include inline dependency metadata, so the easiest way to run them is with `uv`:

```bash
uv run build_business_cycle_indices.py
```

Alternatively, create a virtual environment and install the dependencies yourself:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install pandas numpy openpyxl plotly
```

## Data Layout

Raw ABS workbooks are expected locally and are intentionally ignored by git. Put them under:

```text
abs_workbooks/
  australian_national_accounts/
    5206001_Key_Aggregates.xlsx
    5206002_Expenditure_Volume_Measures.xlsx
    5206003_Expenditure_Current_Price.xlsx
    5206006_Industry_GVA.xlsx
    5206008_Household_Final_Consumption_Expenditure.xlsx
    5206045_Industry_GVA_Current_Price.xlsx
  monthly_household_spending/
    5682002.xlsx
```

You can also point the scripts at another folder with their command-line options.

## Usage

Build the quarterly business-cycle indices:

```bash
uv run build_business_cycle_indices.py \
  --national-accounts-dir abs_workbooks/australian_national_accounts \
  --output business_cycle_indices.html
```

Build the cyclical GDP dashboard:

```bash
uv run build_cyclical_dashboard.py \
  --national-accounts-dir abs_workbooks/australian_national_accounts \
  --output cyclical_dashboard.html
```

Build the monthly high-frequency leading index:

```bash
uv run build_high_frequency_leading_index.py \
  --household-spending-dir abs_workbooks/monthly_household_spending \
  --output high_frequency_leading_index.html
```

Open the generated `.html` files in a browser to inspect the dashboards.

## Notes

- The scripts validate that required ABS series are present in the workbook `Data1` sheets.
- The generated dashboards are ignored by git because they can be recreated from the scripts and local workbooks.
- ABS workbook filenames and labels must match the script definitions; if ABS changes workbook structure or labels, update the relevant `SERIES_SPECS` or component definitions.
