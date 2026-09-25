# Test Infrastructure & Synthetic Fixture Architecture

## 1. Overview & Purpose

The Excel Quote Extractor operates on complex legacy Excel workbooks (`.xlsx`, `.xlsm`) from Molded Precision Components (MPC). Real manufacturing quote workbooks contain proprietary customer data, multi-tab quoting sheets, 2D tabular row-column matrix grids, merged cells, formulas, and formatting anomalies.

To guarantee comprehensive, repeatable, and automated testing across all environments **without relying on external proprietary files or Microsoft Excel desktop installations**, the test infrastructure provides:
1. **Pure-Python Synthetic Reference Workbook Generator** (`tests/fixtures/generator.py`): Generates valid OpenXML `.xlsx` and `.xlsm` zip archives containing standard XML structures (`[Content_Types].xml`, `xl/workbook.xml`, `xl/worksheets/sheetN.xml`, `xl/sharedStrings.xml`, `xl/styles.xml`, and merged cell ranges).
2. **Pytest Fixtures & Environment Configuration** (`tests/conftest.py`): Scoped temporary directory fixtures that isolate test workbooks and batch directory trees.
3. **Comprehensive 4-Tier Opaque-Box E2E Test Suite** (`tests/test_e2e_tiers.py`): Strict specification-based tests spanning Category-Partition feature coverage, Boundary Value Analysis (BVA), Cross-Feature Pairwise Combinations, and Real-World Workload Scenarios.

---

## 2. Synthetic Workbook Generator (`tests/fixtures/generator.py`)

### 2.1 Zero-Dependency OpenXML Architecture
The generator implements `PureXmlWorkbook`, an in-memory OpenXML spreadsheet builder using Python's standard `zipfile` and `xml` formatting:
- **Zero Binary Dependencies:** Works in any standard Python 3.12+ environment without requiring `openpyxl`, LibreOffice, or Excel.
- **Valid OpenXML Archive:** Every generated file adheres to the ECMA-376 OpenXML standard, ensuring that `openpyxl`, `pandas`, Excel, and native XML readers can open them without warnings.
- **Merged Cell Range Tracking:** Encodes `<mergeCells><mergeCell ref="..."/></mergeCells>` tags to test spatial offset traversal.
- **Formula & Cached Value Encoding:** Supports formula syntax `<f>...</f>` alongside cached calculated values `<v>...</v>`.
- **Macro-Enabled Workbooks (`.xlsm`):** Produces valid `.xlsm` archives with the appropriate MIME content type (`application/vnd.ms-excel.sheet.macroEnabled.main+xml`).

### 2.2 Synthetic Fixture Inventory

| Fixture Name | File Name | Layout Archetype | Key Characteristics | Target Output |
|:---|:---|:---|:---|:---|
| `single_modern` | `quote_single_modern.xlsx` | Single-Item Modern Form | Key-value label offsets, DWG disambiguation | `Q1001-1`, `2202046213000`, 24.5s |
| `multitab` | `quote_multitab_q2005.xlsx` | Multi-Tab Form (Sheet-per-Item) | 5 sheets (`Summary`, 3 item tabs, `Lookups`) | `Q2005-1`, `Q2005-2`, `Q2005-3` |
| `grid_matrix` | `quote_grid_matrix_q4000.xlsx` | 2D Tabular Grid Matrix | Header row at row 7, data rows 8-10, footer row 11 | `Q4000-1`, `Q4000-2`, `Q4000-3` |
| `merged_cells` | `quote_merged_cells_q5000.xlsx` | Merged Cells & Spacer Offsets | Merged labels `B2:D2`, `B6:C6`, `B14:C14`, value at `E14` | `Q5000-1`, `3042605C`, 28.0s |
| `legacy_raw` | `quote_legacy_raw_q3050.xlsm` | Legacy Macro Workbook (.xlsm) | Base `Q3050` with no suffix, formula `=D14+D15` (27.5s) | `Q3050-1` (inferred), `994021A-REV2`, 27.5s |
| `edge_cases` | `quote_edge_cases_q6000.xlsx` | Anomaly & Boundary Testing | 4 sheets: 2.4s (<5s), 420.0s (>300s), text units, missing fields | Out-of-bounds warnings, clean text float parsing |
| `empty` | `quote_empty.xlsx` | Edge Case | Empty workbook with single blank sheet | `total_items=0`, warning recorded |
| `non_quote` | `non_quote_spreadsheet.xlsx` | Negative Control | Warehouse inventory sheet without quote markers | `total_items=0`, `items=[]` |
| `scientific_notation`| `quote_scientific_notation.xlsx` | Edge Case | 13-digit numeric part number (`2202046213000`) | Clean string without `e+12` or `.0` |
| `whitespace_variation`| `quote_whitespace_variation.xlsx`| Edge Case | Mixed casing, `\u00a0` non-breaking spaces, trailing colons | Normalized labels, clean part & cycle time |
| `dwg_disambiguation`| `quote_dwg_disambiguation.xlsx` | Edge Case | Adjacent `Drawing Number:` and `Part Number:` | Part number extracted, DWG excluded |
| `hybrid` | `quote_hybrid.xlsx` | Composite Hybrid Layout | Master summary 2D grid + child item detail tabs | Deduplicated items `Q10050-1`, `Q10050-2` |

---

## 3. 4-Tier Opaque-Box E2E Test Suite Structure

The test suite in `tests/test_e2e_tiers.py` implements a 4-tier testing pyramid plus self-contained fixture verification:

```
┌─────────────────────────────────────────────────────────────┐
│             Tier 4: Real-World Workload Scenarios           │
│   (Batch folder scanning, lock file exclusion, CSV/JSON)    │
├─────────────────────────────────────────────────────────────┤
│         Tier 3: Cross-Feature Pairwise Combinations         │
│  (Multi-tab + Merged, Grid + Out-of-Bounds, Hybrid layouts) │
├─────────────────────────────────────────────────────────────┤
│          Tier 2: Boundary & Corner Cases (BVA)              │
│    (Empty files, <5s / >300s, spacer offsets, formulas)     │
├─────────────────────────────────────────────────────────────┤
│               Tier 1: Feature Coverage (Equivalence)        │
│  (Single-item, Multi-tab, Grid matrix, Part #, Cycle time)  │
├─────────────────────────────────────────────────────────────┤
│         Tier 0: Generator Integrity Verification            │
│   (Validates OpenXML zip structure, sheets & strings)       │
└─────────────────────────────────────────────────────────────┘
```

### 3.1 Tier 0: Generator Integrity Verification
- **Purpose**: Verify that all 12 synthetic fixtures generate valid, corrupt-free OpenXML zip archives.
- **Pass Status**: 100% passing immediately in standard Python without external dependencies.
- **Classes**: `TestTier0FixtureGeneratorIntegrity`

### 3.2 Tier 1: Feature Coverage (>=5 tests per core feature)
- **Feature 1: Single-Item Modern Quote Extraction**
  - Base Quote ID (`Q1001`), Item Number (`Q1001-1`), Part Number (`2202046213000`), Cycle Time (`24.5s`), Coordinate tracking (`C2`, `C10`, `C18`), Confidence (`high`), Warnings (`[]`).
- **Feature 2: Multi-Tab Quote Isolation**
  - Item count (`3`), Exclusion of `Summary` and `Lookups` auxiliary tabs, Sheet-specific part numbers (`MPC-7011-A`, `MPC-7012-B`, `MPC-7013-C`), Sheet-specific cycle times (`18.2s`, `32.0s`, `45.5s`), Sheet-scoped coordinates (`Q2005-1!C2`, etc.).
- **Feature 3: 2D Tabular Grid Matrix Intersection**
  - Header detection (Row 7), Item count (`3`), Column coordinate intersections (`A8`/`B8`/`E8`, etc.), Subtotal/footer exclusion (Row 11).
- **Feature 4: Part Number Cleansing**
  - Standard alphanumeric, Scientific notation float restoration, Whitespace & non-breaking space (`\u00a0`) stripping, Trailing punctuation removal, Drawing number (`DWG-*`) exclusion.
- **Feature 5: Cycle Time Extraction**
  - Direct float extraction, String unit cleansing (`" 26.5 sec "` -> `26.5`), Bounds validation (`[5.0, 300.0]`), Coordinate recording, Confidence score assignment.

### 3.3 Tier 2: Boundary & Corner Cases (>=5 tests per feature)
- **Feature 1: Empty & Non-Quote Files**
  - Empty workbook handling (`total_items=0`), Non-quote inventory sheet handling, Missing file `FileNotFoundError`, Empty sheets in multi-tab ignored, Corrupt byte stream error handling.
- **Feature 2: Negative & Excessive Cycle Times**
  - Fast cycle boundary (`2.4s < 5.0s` warning), Slow cycle boundary (`420.0s > 300.0s` warning), Exact minimum boundary (`5.0s` valid), Exact maximum boundary (`300.0s` valid), Zero/negative cycle times.
- **Feature 3: Merged Cells and Offset Shifts**
  - Merged header `B2:D2` resolution, Merged label `B6:C6` skipping, Merged value `D6:F6` reading from top-left anchor `D6`, Spacer column `D14` traversal to `E14`, Precise coordinate attribution.
- **Feature 4: Formulas and Legacy Workbooks**
  - Cached formula evaluation (`=D14+D15` -> `27.5`), Suffix inference (`Q3050` -> `Q3050-1` with warning), Error cell handling (`#N/A`), Missing cached value warning, Macro-enabled `.xlsm` container support.
- **Feature 5: Whitespace and Formatting Variations**
  - Non-breaking spaces (`\u00a0`), Irregular case matching (`qUoTe # :`), Trailing delimiters (`;`, `-`), Multiline cells, Indented cells.

### 3.4 Tier 3: Cross-Feature Pairwise Combinations
- Multi-tab workbook with merged cells across all sheets.
- 2D grid matrix containing out-of-bounds cycle times (normal, low, high in same table).
- 2D grid matrix with formula-calculated cycle times.
- Legacy `.xlsm` with multiple quote items.
- Hybrid workbook containing both a master summary 2D grid and child detail tabs (deduplication check).
- 2D grid containing 13-digit numeric part numbers.
- Merged cell layout with missing part number.
- Legacy `.xlsm` with noisy whitespace and drawing number anchors.
- Multi-tab quote where one tab has a missing cycle time and other tabs are valid.
- Multi-tab quote with mixed confidence distribution (`high` vs `low`).

### 3.5 Tier 4: Real-World Workload Scenarios
- Directory scanning strictly ignores Excel temporary lock files (`~$*.xlsx`, `~$*.xlsm`).
- Directory scanning recurses through nested subdirectories.
- Consolidated processing of mixed quote archetypes (single, multitab, grid, legacy).
- Full schema validation against `QuoteDocument` contract.
- Serialization and validity of batch JSON export payload.
- Representation and validity of batch CSV export rows.
- Complete isolation between part numbers and drawing numbers across the workload.
- Consistent exclusion of auxiliary administrative tabs.
- Full coordinate traceability (`source_cell_coords`) across all extracted items.
- High-volume batch performance benchmark (<10.0s for 12 complex workbooks).

---

## 4. How to Run the Tests

### 4.1 Run Generator Integrity Tests (Tier 0)
```powershell
pytest tests/test_e2e_tiers.py -k "TestTier0" -v
```

### 4.2 Run Full E2E Test Suite (All Tiers)
```powershell
pytest tests/test_e2e_tiers.py -v
```

### 4.3 Run Specific Tiers
```powershell
# Tier 1: Feature Coverage
pytest tests/test_e2e_tiers.py -k "Tier1" -v

# Tier 2: Boundary & Corner Cases
pytest tests/test_e2e_tiers.py -k "Tier2" -v

# Tier 3: Cross-Feature Combinations
pytest tests/test_e2e_tiers.py -k "Tier3" -v

# Tier 4: Real-World Workload Scenarios
pytest tests/test_e2e_tiers.py -k "Tier4" -v
```

---

## 5. Progressive Testability & Milestone Integration

- **Current State (Milestone E2E / M1 Pre-implementation):**
  - Tier 0 tests execute and pass immediately.
  - Tiers 1–4 are fully implemented with authoritative assertions derived from `PROJECT.md` and `workbooks_report.md`.
  - In `tests/conftest.py`, the `extractor_engine` fixture cleanly checks for the presence of `QuoteExtractorEngine`. If not yet implemented, tests gracefully mark as skipped with a clear milestone dependency message (`"QuoteExtractorEngine is not implemented yet (Milestone 2 dependency)"`).
- **Milestone 1 Completion:**
  - `QuoteDocument`, `QuoteItem`, and `SourceCellCoords` models are available; schema tests validate data structures.
- **Milestone 2 Completion:**
  - `QuoteExtractorEngine`, `reader.py`, `locator.py`, and `strategies/` are implemented; all Tier 1 through Tier 4 tests execute and validate the live extraction engine.
