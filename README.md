# ExcelExtractor

A high-performance Python application, batch processor, and interactive Streamlit web dashboard for extracting structured quotation data, part specifications, and manufacturing metrics from legacy Excel spreadsheets (`.xlsx`, `.xlsm`).

---

## Overview

Industrial quoting spreadsheets often span multiple revisions, complex multi-tab layouts, or 2D matrix grids with merged cells and formula lookups. **ExcelExtractor** automatically analyzes and extracts quote metadata, line item numbers, part numbers, descriptions, resin grades, cycle times, labor rates, and volumes with zero manual copy-pasting.

---

## Key Features

- **Pure-Python OpenXML Reader**: Zero external binary dependencies. Uses a fast streaming XML archive parser with lazy sheet evaluation, bypassing heavy administrative or lookup tabs until accessed.
- **Two-Tier Search Engine**:
  - **Tier 1 (Light Scan)**: Rapidly sweeps down sheets (up to 250 rows) to detect candidate table headers and anchor labels in milliseconds.
  - **Tier 2 (Deep Search)**: Deeply classifies columns, verifies data row density, and filters out non-quote tables (e.g. revision blocks, checklists).
- **Multi-Layout Extraction Support**:
  - **2D Tabular Grid Matrix**: Extracts rows intersecting with Part #, Cycle Time, Resin Grade, Description, Labor Rate, #Ops, Weight, and Volume columns.
  - **Multi-Tab Workbooks**: Detects individual quote items across separate tabs.
  - **Single-Item Dedicated Forms**: Resolves key-value label offsets with merged cell boundary handling.
- **Data Validation & Sanity Highlighting**: Flags out-of-bounds cycle times, missing fields, or questionable part numbers with an interactive drill-down modal showing exact Excel cell coordinates.
- **Job# - Quote# Linker Workspace**: Matches manufacturing jobs with quote line items using drawing number trimming and partial match resolution.
- **Interactive Web UI & Batch CLI**: Review and edit results in a modern Streamlit interface, or process thousands of files automatically via the command line.

---

## Installation

Ensure Python 3.10+ is installed:

```bash
git clone https://github.com/cox-ben/ExcelExtractor.git
cd ExcelExtractor
pip install -r requirements.txt
```

---

## Usage

### 1. Interactive Web Dashboard

Launch the Streamlit app:

```bash
streamlit run app.py
```

Open `http://localhost:8501` in your browser to:
- Drag and drop `.xlsx` or `.xlsm` quotation spreadsheets.
- Inspect extracted line items in a dark-mode table.
- Click any value to view its source worksheet and exact cell coordinate (e.g. `B18`, `Quote!C4`).
- Toggle sanity checks to highlight warnings or errors.
- Export cleaned datasets to JSON, CSV, or formatted Excel files.
- Switch to the **Job# - Quote# Linker** tab to link ERP jobs with quote numbers.

### 2. Command Line Interface (CLI)

Process individual files or entire directory trees:

```bash
# Extract a single quote file
python cli.py quote.xlsx

# Process a directory of spreadsheets recursively
python cli.py "C:\Quotes\2024" --recursive --output-dir "C:\Quotes\Export"

# Export formatted CSV and JSON summaries
python cli.py "C:\Quotes" --format csv
```

---

## Project Structure

```text
ExcelExtractor/
├── app.py                      # Interactive Streamlit application
├── cli.py                      # Command-line batch processor
├── config.yaml                 # Configurable keywords, anchors, and bounds
├── requirements.txt            # Project dependencies
├── src/
│   └── excel_extractor/
│       ├── __init__.py         # Package exports & version
│       ├── config.py           # Configuration schema & Pydantic models
│       ├── engine.py           # Extraction orchestrator
│       ├── exporters.py        # JSON, CSV, and Excel exporters
│       ├── linker.py           # Job-to-Quote matching engine
│       ├── locator.py          # Spatial anchor resolver & cell cleaner
│       ├── models.py           # Data contracts & QuoteDocument models
│       ├── reader.py           # Resilient Pure-Python OpenXML reader
│       └── strategies/         # Extraction strategies
│           ├── base.py         # Strategy base class
│           ├── grid_matrix.py  # 2D tabular row-column matrix strategy
│           ├── multi_tab.py    # Multi-tab quote strategy
│           └── single_item.py  # Single-item form strategy
└── tests/                      # Automated test suite (278 tests)
```

---

## Running Automated Tests

Run the full automated test suite covering unit tests, adversarial layouts, and edge cases:

```bash
pytest
```
