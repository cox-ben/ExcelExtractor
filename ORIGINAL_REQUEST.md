# Original User Request

## 2026-09-18T16:59:07Z

Build a Python-based Excel Quote Extractor utility that parses legacy `.xlsx` and `.xlsm` quoting files, isolates multi-item quotes (supporting both key-value label offsets and 2D row-column grid intersection), and provides an interactive Streamlit testing dashboard and a CLI batch processor.

Working directory: c:\Users\bcox\OneDrive - Molded Precision Components\3_Code\Excel_Value_Getter
Integrity mode: development

## Verification Resources
* Real legacy Excel quote files (`.xlsx`, `.xlsm`) provided in the project workspace.
* Test suite generating structured sample workbooks covering multi-tab and 2D table grid layouts.

## Requirements

### R1. Extraction Engine & Multi-Item Parsing
Parse `.xlsx` and `.xlsm` files to extract `quote_id`, `quote_item_number` (format `QXXXX-X`), `part_number`, and `cycle_time_sec` (strictly in seconds, range 5.0s–300.0s). Support two multi-item patterns:
- Multi-tab quotes where each tab corresponds to an item.
- 2D grid matrix quotes where quote item rows intersect with Part # and Cycle Time columns.
Support configurable anchor keywords and search offsets via an external YAML configuration file (`config.yaml`).

### R2. Data Contract & Validation
Enforce structured data output adhering to the target schema (`source_file`, `quote_id`, `total_items`, and item list with `source_cell_coords` and confidence ratings). Flag missing fields or out-of-bounds cycle times.

### R3. Interactive Streamlit Beta Testing Dashboard
Provide a Streamlit web app (`app.py`) allowing users to upload `.xlsx`/`.xlsm` files, immediately trigger parsing, view all extracted variables in an interactive and editable table, and export results to JSON and CSV/Excel.

### R4. CLI Batch Processor
Provide a command-line script (`cli.py`) to process directories of `.xlsx`/`.xlsm` files in batch mode and output structured JSON and CSV files.

## Acceptance Criteria

### Extraction Accuracy & Matrix Intersection
- [ ] Correctly extracts quote items from multi-tab workbooks.
- [ ] Correctly resolves cell values at row-column intersections in 2D tabular workbooks.
- [ ] Correctly records cell coordinates (`source_cell_coords`) for each extracted field.
- [ ] Correctly flags cycle times outside 5.0–300.0s as warnings.

### User Interface & Beta Testing
- [ ] Streamlit app accepts `.xlsx` and `.xlsm` uploads, runs extraction, and displays all extracted variables in an editable table.
- [ ] Streamlit app provides functional download buttons for JSON and CSV exports.

### Automated Testing & CLI
- [ ] Automated test suite verifies extraction logic across single-item, multi-tab, merged-cell, and tabular grid fixtures.
- [ ] CLI command executes successfully against quote files and writes valid JSON and CSV outputs.
