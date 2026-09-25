# Quote Extractor Prototype: Architecture & Agent Execution Plan

## 1. Prototype Scope & Target Schema
Build a lightweight extraction utility that ingests legacy Excel quoting files, isolates multi-item quotes, and extracts three primary fields per item into a standardized data model.

### Target Data Contract
```json
{
  "source_file": "legacy_quote_example.xlsx",
  "quote_id": "Q1001",
  "total_items": 2,
  "items": [
    {
      "quote_item_number": "Q1001-1",
      "part_number": "2202046213000",
      "cycle_time_sec": 24.5,
      "confidence": "high",
      "source_cell_coords": {
        "quote_num": "B2",
        "part_num": "C10",
        "cycle_time": "F18"
      }
    }
  ]
}
```

---

## 2. Technical Stack
* **Language & Runtime:** Python 3.11+
* **Spreadsheet Engine:** `openpyxl` (data-only mode for calculated values) and `pandas`
* **Parsing & Logic:** Regex (`re`), fuzzy matching (`rapidfuzz`), and fallback Pydantic validation
* **Prototype UI / Testing Rig:** Streamlit (fast drag-and-drop file upload, item preview, and manual correction interface)

---

## 3. Extraction Methodology

```
┌─────────────────┐     ┌───────────────────────┐     ┌────────────────────────┐
│  Upload Excel   │ ──> │ Layout Identification │ ──> │ Value Parsing Engine   │
│  (.xlsx / .xls) │     │ (Multi-tab vs. Rows)  │     │ (Regex + Anchor Offsets│
└─────────────────┘     └───────────────────────┘     └────────────────────────┘
                                                                  │
                                                                  ▼
                                                      ┌────────────────────────┐
                                                      │ Staging & Verification │
                                                      │ (Portal Ingest Format) │
                                                      └────────────────────────┘
```

### A. Quote Number & Item Discovery (`QXXXX-X`)
1. **Global Regex Scan:** Scan all sheet names and cell values using pattern:
   $$\text{Regex: } \texttt{\bQ[0-9]{4,5}-[0-9]+\b}$$
2. **Item Count Heuristics:**
   * **Pattern A (Tab per Item):** Workbook contains distinct tabs named or tagged by item (e.g., `Q1001-1`, `Q1001-2`, or `Item 1`, `Item 2`).
   * **Pattern B (Tabular Multi-Row):** A single master sheet has items listed sequentially down rows or across distinct column clusters.
   * Total item count is determined by deduplicating unique `QXXXX-X` matches.

### B. Part Number Extraction
* **Primary Search:** Anchor-based scan matching labels: `["Part No", "Part #", "P/N", "Customer Part", "Item Number"]`.
* **Coordinate Offset:** Read adjacent cells (offset: `+1 col`, `+2 col`, or `+1 row`).
* **Format Cleansing:**
  * Extract alphanumeric string.
  * Discard Drawing Numbers (DWG) if listed separately.
  * Split revision suffixes into metadata where applicable (e.g., separating `3042605C` from internal drawing codes).

### C. Cycle Time Extraction (Seconds)
* **Anchor Search:** Match labels: `["Cycle Time", "Cycle Time (sec)", "C/T", "Sec/Shot", "Molding Cycle"]`.
* **Value Rules:**
  * Must parse to a positive float or integer.
  * Sanity filter: $5.0 \le t \le 300.0\text{ s}$ (flags outliers for manual review).

---

## 4. Phased Agent Build Steps

### Phase 1: Ingestion & Anchor Scanner (`extractor/scanner.py`)
* Load workbook using `openpyxl.load_workbook(filename, data_only=True)`.
* Build coordinate map of text values across all non-hidden sheets.
* Locate primary anchors for Quote #, Part #, and Cycle Time.

### Phase 2: Item Grouping & Disambiguation (`extractor/parser.py`)
* If single-item layout: Extract base Quote ID and assign item suffix `-1`.
* If multi-item layout:
  * Check for repeating coordinate blocks.
  * Associate each `part_number` and `cycle_time_sec` with its respective `QXXXX-X` item tag.

### Phase 3: Validation & Anomaly Flagging (`models/schemas.py`)
* Use Pydantic to enforce data typing:
  * Fail if `quote_item_number` does not strictly follow `Q\d+-\d+`.
  * Warn if `part_number` contains non-standard characters or matches drawing labels.
  * Flag `cycle_time_sec` if zero, blank, or exceeding production thresholds.

### Phase 4: Streamlit Verification Dashboard (`app.py`)
* Drag-and-drop file upload zone.
* Display parsed results in an editable data table:
  * **Quote Item #** | **Part #** | **Cycle Time (s)** | **Source Tab / Cell**
* "Export to JSON / Portal Staging" button.

---

## 5. Agent Verification & Test Plan
* **Test Case 1 (Standard Modern):** Single-item quote, explicit label `Q1001-1`.
* **Test Case 2 (Multi-Item Tabs):** Single workbook with 3 tabs (`Q2005-1`, `Q2005-2`, `Q2005-3`).
* **Test Case 3 (Shifted Grid / Merged Cells):** Quote where "Cycle Time" is merged across cells `D14:E14` and value is in `F14`.
* **Test Case 4 (Legacy Raw):** Old sheet format where only base quote `Q3050` is written in header; verify agent infers `Q3050-1`.