# Excel Value Getter

**Version**: `v3.0.0`  
**License**: Proprietary / Molded Precision Components (MPC)

An intelligent, resilient Excel quote variable extraction engine, batch processor, and interactive Streamlit review dashboard for legacy manufacturing quotation workbooks (`.xlsx`, `.xlsm`).

---

## Key Features

- **Pure-Python OpenXML Reader**: Fast, zero-dependency streaming XML archive parser with lazy sheet loading. Bypasses multi-thousand row administrative and lookup tabs until accessed.
- **Two-Tier Search Engine**:
  - **Tier 1 (Light Scan)**: Rapidly sweeps down worksheets (up to 250 rows) to detect table headers and anchor labels in milliseconds.
  - **Tier 2 (Deep Search)**: Deeply classifies columns, verifies data row density, and handles multi-band tables (revision tables, disclaimers, etc.).
- **Multi-Item & Hybrid Layout Support**:
  - **2D Tabular Grid Matrix**: Resolves line items intersecting with Part #, Cycle Time, Resin Grade, Description, Labour Rate, #Ops, Weight, and Volume.
  - **Multi-Tab Workbooks**: Extracts individual quote items across separate tabs.
  - **Single-Item Dedicated Forms**: Resolves key-value label offsets with merged cell boundary handling.
- **Data Validation & Sanity Highlighting**: Flags out-of-bounds cycle times, missing fields, or questionable part numbers with interactive modal drill-downs showing exact Excel cell coordinates.
- **Job# to Quote# Linker Workspace**: Matches manufacturing jobs with quote line items using drawing number trimming and partial match resolution.
- **Batch CLI & Web UI**: Full command-line directory processing and interactive Streamlit review dashboard.

---

## Installation

Ensure Python 3.10+ is installed:

```bash
git clone https://github.com/cox-ben/Excel_Value_Getter.git
cd Excel_Value_Getter
pip install -r requirements.txt
```

---

## Usage

### 1. Interactive Streamlit Dashboard

Run the interactive dashboard:

```bash
streamlit run app.py
```

Then open `http://localhost:8501` in your browser.

- Upload `.xlsx` or `.xlsm` quote workbooks.
- Review and edit extracted line items in an interactive dark-mode table.
- Click any cell to inspect its exact Excel origin coordinates.
- Export clean results to JSON, CSV, or Excel format.
- Switch to the **Job# - Quote# Linker** tab to link ERP jobs with quote numbers.

### 2. Command Line Interface (CLI)

Process individual workbooks or entire directories in batch:

```bash
# Process a single quote file
python cli.py quote.xlsx

# Process a directory of quotes recursively
python cli.py "C:\Quotes\2024" --recursive --output-dir "C:\Quotes\Export"

# Export formatted CSV and JSON summaries
python cli.py "C:\Quotes" --format csv
```

---

## Running Automated Tests

Run the full automated test suite (277 tests covering all extraction layouts and adversarial cases):

```bash
pytest
```
