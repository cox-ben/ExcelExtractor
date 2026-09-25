# Project: Excel Quote Extractor

## Architecture

The Excel Quote Extractor is a modular Python package providing automated extraction of legacy quoting workbooks (.xlsx, .xlsm), structured schema validation, a Streamlit interactive dashboard, and a CLI batch processing tool.

```
┌─────────────────────────────────────────────────────────────┐
│                       Entry Points                          │
│         cli.py (CLI Batch)       app.py (Streamlit UI)      │
└──────────────────────────────┬──────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────┐
│               Extraction & Orchestration Layer              │
│       QuoteExtractorEngine (src/excel_extractor/engine.py)  │
└──────────────┬──────────────────────────────┬───────────────┘
               │                              │
               ▼                              ▼
┌──────────────────────────────┐┌─────────────────────────────┐
│        Reader Layer          ││      Strategy Layer         │
│  src/excel_extractor/reader  ││ src/excel_extractor/strategy│
│  - OpenPyXLReader            ││  - MultiTabStrategy         │
│  - NativeXMLReader (fallback)││  - GridMatrixStrategy       │
│  - MergedCellResolver        ││  - SingleItemStrategy       │
└──────────────────────────────┘└─────────────────────────────┘
                               │
                               ▼
┌─────────────────────────────────────────────────────────────┐
│           Validation, Models & Configuration Layer          │
│    src/excel_extractor/models.py (Pydantic V2 / Contracts)  │
│    src/excel_extractor/config.py (YAML Loader + Defaults)   │
└─────────────────────────────────────────────────────────────┘
```

## Code Layout

```
c:\Users\bcox\OneDrive - Molded Precision Components\3_Code\Excel_Value_Getter\
├── config.yaml                     # External YAML configuration for anchors & bounds
├── app.py                          # Streamlit UI entrypoint
├── cli.py                          # CLI Batch processor entrypoint
├── src/
│   └── excel_extractor/
│       ├── __init__.py             # Package exports
│       ├── models.py               # Data contracts (QuoteItem, QuoteDocument, etc.)
│       ├── config.py               # Config parser & default anchor dictionary
│       ├── reader.py               # Resilient workbook reader (OpenPyXL + pure XML fallback)
│       ├── locator.py              # Anchor scanner, fuzzy matching, merged cell traversal
│       ├── strategies/             # Layout extraction strategies
│       │   ├── __init__.py
│       │   ├── base.py             # Base extraction strategy
│       │   ├── single_item.py      # Single-item form parser
│       │   ├── multi_tab.py        # Multi-tab sheet-per-item parser
│       │   └── grid_matrix.py      # 2D tabular row-column intersection parser
│       ├── engine.py               # Main QuoteExtractorEngine coordinator
│       └── exporters.py            # JSON, CSV, and Excel exporters
├── tests/
│   ├── conftest.py                 # Pytest fixtures and paths
│   ├── fixtures/
│   │   ├── __init__.py
│   │   └── generator.py            # Synthetic reference workbook generator
│   ├── test_models.py              # Unit tests for data contracts & validation
│   ├── test_reader.py              # Unit tests for workbook readers & merged cells
│   ├── test_extraction.py          # Unit tests for multi-tab & grid strategies
│   ├── test_cli.py                 # CLI end-to-end batch tests
│   ├── test_streamlit.py          # Headless Streamlit AppTest tests
│   └── test_e2e_tiers.py           # Comprehensive Tier 1-4 requirement tests
└── .agents/                        # Agent metadata, logs, and handoffs only
```

## Feature Inventory

| # | Feature | Description | Milestone | Source |
|---|---------|-------------|-----------|--------|
| 1 | Quote ID & Item Extraction | Extract `quote_id` (QXXXX) and `quote_item_number` (QXXXX-X), infer `-1` if base only | M2 | ORIGINAL_REQUEST §R1 |
| 2 | Multi-Tab Item Isolation | Extract items across tabs, filter non-quote sheets (Summary, TOC) | M2 | ORIGINAL_REQUEST §R1 |
| 3 | 2D Tabular Grid Intersection | Identify header row, map Part # / Cycle Time columns, extract row items | M2 | ORIGINAL_REQUEST §R1 |
| 4 | Part Number Parsing & Cleaning | Locate Part # anchor/col, strip DWG tags, normalize whitespace & notation | M2 | ORIGINAL_REQUEST §R1 |
| 5 | Cycle Time Extraction | Extract cycle time in seconds, strictly evaluate manufacturing bounds (5.0s–300.0s) | M2 | ORIGINAL_REQUEST §R1 |
| 6 | Merged Cell Resolution | Traverse merged cell ranges so offsets skip merged anchor cells (e.g. D14:E14 -> F14) | M2 | ORIGINAL_REQUEST §R1 |
| 7 | External YAML Configuration | Load anchors, offsets, bounds, and ignore lists from `config.yaml` with safe defaults | M1 | ORIGINAL_REQUEST §R1 |
| 8 | Target Schema Enforcement | Validated `QuoteDocument` and `QuoteItem` schema adhering to specification | M1 | ORIGINAL_REQUEST §R2 |
| 9 | Cell Coordinate Tracking | Record exact cell coordinates (`source_cell_coords`) for all extracted variables | M1, M2 | ORIGINAL_REQUEST §R2 |
| 10 | Deterministic Confidence Scoring | Assign `high`, `medium`, `low` confidence based on anchor distance & validation | M1 | ORIGINAL_REQUEST §R2 |
| 11 | Cycle Time Bounds Warning | Warn on cycle times outside 5.0s–300.0s without discarding value | M1 | ORIGINAL_REQUEST §R2 |
| 12 | Streamlit Drag-and-Drop Upload | Web file uploader accepting `.xlsx` and `.xlsm` files | M4 | ORIGINAL_REQUEST §R3 |
| 13 | Immediate Parsing & Summary Cards | Instant parse on upload, summary KPI cards (Quote ID, Item Count, Status) | M4 | ORIGINAL_REQUEST §R3 |
| 14 | Editable Data Table UI | Interactive table (`st.data_editor`) for user review and manual edits | M4 | ORIGINAL_REQUEST §R3 |
| 15 | Multi-Format Export Buttons | Immediate export of extracted/edited items to JSON, CSV, and Excel | M4 | ORIGINAL_REQUEST §R3 |
| 16 | CLI Batch Directory Scanning | Scan directories recursively for `.xlsx`/`.xlsm`, filtering temp files (`~$*`) | M3 | ORIGINAL_REQUEST §R4 |
| 17 | Batch Extraction & Consolidated Export | Process discovered files into consolidated `extracted_quotes.json` and `.csv` | M3 | ORIGINAL_REQUEST §R4 |
| 18 | CLI Flags & Exit Codes | Support `-i`, `-o`, `-c`, `-f`, `--strict` and standard exit codes (0, 1, 2) | M3 | ORIGINAL_REQUEST §R4 |
| 19 | Synthetic Fixture Generation | Generate valid `.xlsx`/`.xlsm` workbooks representing all 4 layout archetypes | E2E Track | ORIGINAL_REQUEST §Verification |
| 20 | Tier 1-4 Opaque-Box Test Suite | Category-partition, BVA, pairwise combinations, real-world application workloads | E2E Track | ORIGINAL_REQUEST §Verification |

*Cross-Check Verification*: All 20 features are assigned to a milestone or the E2E test track. No unassigned features.

## Milestones

| # | Name | Scope | Dependencies | Status |
|---|------|-------|-------------|--------|
| E2E | E2E Testing Track | Synthetic fixture generator (`tests/fixtures/generator.py`), 4-tier E2E test suite, `TEST_INFRA.md`, publish `TEST_READY.md` | None | DONE |
| M1 | Data Contract, Validation & Config | `src/excel_extractor/models.py`, `src/excel_extractor/config.py`, `config.yaml`, validation rules, confidence calculation | None | DONE |
| M2 | Extraction Engine & Strategies | `reader.py`, `locator.py`, `strategies/`, `engine.py`, multi-tab + 2D grid matrix, merged cell handling | M1 | DONE |
| M3 | CLI Batch Processor | `cli.py`, directory batch scanning, lock file exclusion, batch export, return codes | M1, M2 | DONE |
| M4 | Streamlit Testing Dashboard | `app.py`, upload, instant parse, summary cards, editable table, export buttons | M1, M2 | DONE |
| M5 | Final Milestone: 100% E2E Pass & Hardening | Run 100% of E2E test suite (230 tests passed), adversarial hardening, Forensic Audit CLEAN | E2E, M1-M4 | DONE |

## Interface Contracts

### 1. `config.yaml` ↔ `ConfigManager`
```yaml
extraction:
  quote_id:
    patterns: ["\\b(Q[0-9]{4,5})\\b"]
    item_patterns: ["\\b(Q[0-9]{4,5}-[0-9]+)\\b"]
  anchors:
    part_number:
      keywords: ["part number", "part no", "part #", "p/n", "customer part", "item number"]
      search_offsets: [[0, 1], [0, 2], [1, 0]]
    cycle_time:
      keywords: ["cycle time", "cycle time (sec)", "c/t", "sec/shot", "molding cycle"]
      search_offsets: [[0, 1], [0, 2], [1, 0]]
  bounds:
    cycle_time_sec:
      min: 5.0
      max: 300.0
  sheets:
    ignore: ["summary", "cover", "toc", "notes", "lookup", "instructions"]
```

### 2. `Data Contract`: `QuoteItem` and `QuoteDocument`
```python
from pydantic import BaseModel, Field
from typing import Optional, List, Dict

class SourceCellCoords(BaseModel):
    quote_num: Optional[str] = None
    part_num: Optional[str] = None
    cycle_time: Optional[str] = None

class QuoteItem(BaseModel):
    quote_item_number: str            # e.g. "Q1001-1"
    part_number: Optional[str] = None # e.g. "3042605C"
    cycle_time_sec: Optional[float] = None # e.g. 24.5
    source_cell_coords: SourceCellCoords = Field(default_factory=SourceCellCoords)
    confidence: str = "high"          # "high" | "medium" | "low"
    warnings: List[str] = Field(default_factory=list)

class QuoteDocument(BaseModel):
    source_file: str
    quote_id: Optional[str] = None    # e.g. "Q1001"
    total_items: int = 0
    items: List[QuoteItem] = Field(default_factory=list)
    warnings: List[str] = Field(default_factory=list)
```

### 3. `QuoteExtractorEngine` API
```python
class QuoteExtractorEngine:
    def __init__(self, config_path: Optional[str] = None): ...
    def extract_file(self, file_path_or_bytes: Union[str, bytes, BinaryIO], filename: Optional[str] = None) -> QuoteDocument: ...
    def extract_directory(self, dir_path: str, recursive: bool = True) -> List[QuoteDocument]: ...
```

### 4. `cli.py` Interface
```bash
python cli.py -i <input_path> -o <output_dir> [-c config.yaml] [-f json|csv|both] [--strict]
# Exit codes:
# 0: Clean run, all files parsed successfully
# 1: Fatal invocation / argument error
# 2: Files parsed with warnings or partial failures
```

### 5. `app.py` Streamlit Interface
- Accepts `.xlsx` and `.xlsm` uploads.
- Displays Quote ID, Total Items, Warnings metrics.
- Presents editable table with columns: `Quote Item`, `Part Number`, `Cycle Time (s)`, `Confidence`, `Cell Coordinates`, `Warnings`.
- Provides Download buttons for JSON, CSV, and Excel.
