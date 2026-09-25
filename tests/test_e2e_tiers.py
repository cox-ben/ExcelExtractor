"""Comprehensive 4-Tier Opaque-Box E2E Test Suite for Excel Quote Extractor.

Tiers:
- Tier 1: Feature Coverage (>=5 tests per core feature: single-item, multi-tab,
          2D grid matrix, part number cleansing, cycle time extraction).
- Tier 2: Boundary & Corner Cases (>=5 tests per feature: empty files, zero items,
          out-of-bounds cycle times, merged labels, formula cells, whitespace variations).
- Tier 3: Cross-Feature Pairwise Combinations (multi-tab + merged cells, 2D grid +
          out-of-bounds, legacy xlsm + formula, hybrid layouts).
- Tier 4: Real-World Workload Scenarios (realistic MPC packages, batch scanning,
          lock-file filtering, export payload verification).

Also includes Tier 0: Synthetic Fixture Generator Integrity Verification (runs & passes).
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List
import pytest

from tests.fixtures import generator
from tests.fixtures.generator import inspect_openxml_zip


# ===========================================================================
# Tier 0: Synthetic Fixture Generator Integrity Verification (Self-Contained)
# ===========================================================================

class TestTier0FixtureGeneratorIntegrity:
    """Verify that the synthetic fixture generator creates 100% valid OpenXML workbooks."""

    def test_generator_single_modern_openxml_validity(self, single_modern_quote_file: Path):
        """Single modern quote produces valid zip with Content_Types, workbook, sheets."""
        info = inspect_openxml_zip(single_modern_quote_file)
        assert info["is_valid_openxml"]
        assert "Quote Summary" in info["sheets"]
        assert "xl/worksheets/sheet1.xml" in info["namelist"]

    def test_generator_multitab_sheet_structure(self, multitab_quote_file: Path):
        """Multi-tab quote produces exactly 5 sheets matching specification."""
        info = inspect_openxml_zip(multitab_quote_file)
        assert info["is_valid_openxml"]
        assert info["sheets"] == ["Summary", "Q2005-1", "Q2005-2", "Q2005-3", "Lookups"]

    def test_generator_grid_matrix_sheet_structure(self, grid_matrix_quote_file: Path):
        """Grid matrix quote produces QuoteMatrix sheet with shared strings."""
        info = inspect_openxml_zip(grid_matrix_quote_file)
        assert info["is_valid_openxml"]
        assert "QuoteMatrix" in info["sheets"]
        # Must contain headers in shared strings
        assert "Item #" in info["shared_strings"]
        assert "Customer Part #" in info["shared_strings"]
        assert "Cycle Time (sec)" in info["shared_strings"]

    def test_generator_merged_cells_structure(self, merged_cells_quote_file: Path):
        """Merged-cell quote produces valid openxml with Quote Detail sheet."""
        info = inspect_openxml_zip(merged_cells_quote_file)
        assert info["is_valid_openxml"]
        assert "Quote Detail" in info["sheets"]
        assert any("Q5000-1" in s for s in info["shared_strings"])

    def test_generator_legacy_raw_xlsm_content_type(self, legacy_raw_quote_file: Path):
        """Legacy raw workbook produces valid macro-enabled .xlsm zip archive."""
        assert legacy_raw_quote_file.suffix.lower() == ".xlsm"
        info = inspect_openxml_zip(legacy_raw_quote_file)
        assert info["is_valid_openxml"]
        assert "Sheet1" in info["sheets"]

    def test_generator_edge_cases_all_four_sheets(self, edge_cases_quote_file: Path):
        """Edge cases workbook creates all 4 anomaly sheets."""
        info = inspect_openxml_zip(edge_cases_quote_file)
        assert info["is_valid_openxml"]
        assert info["sheets"] == [
            "Out_Of_Bounds_Low",
            "Out_Of_Bounds_High",
            "Text_Formatted_Units",
            "Missing_Fields",
        ]

    def test_generator_all_fixtures_batch_creation(self, all_synthetic_fixtures: Dict[str, Path]):
        """create_all_fixtures generates all 12 distinct fixtures cleanly."""
        assert len(all_synthetic_fixtures) == 12
        for name, path in all_synthetic_fixtures.items():
            assert path.exists(), f"Fixture {name} does not exist at {path}"
            assert path.stat().st_size > 0, f"Fixture {name} is 0 bytes"
            info = inspect_openxml_zip(path)
            assert info["is_valid_openxml"], f"Fixture {name} is not valid OpenXML"


# ===========================================================================
# Tier 1: Feature Coverage (>=5 tests per core feature)
# ===========================================================================

class TestTier1Feature1SingleItemExtraction:
    """Core Feature 1: Single-Item Modern Quote Extraction (PROJECT.md Feature 1)."""

    def test_tier1_single_item_quote_id(self, single_modern_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(single_modern_quote_file)
        assert doc.quote_id == "Q1001"

    def test_tier1_single_item_number(self, single_modern_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(single_modern_quote_file)
        assert doc.total_items == 1
        assert doc.items[0].quote_item_number == "Q1001-1"

    def test_tier1_single_item_part_number(self, single_modern_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(single_modern_quote_file)
        assert doc.items[0].part_number == "2202046213000"

    def test_tier1_single_item_cycle_time(self, single_modern_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(single_modern_quote_file)
        assert doc.items[0].cycle_time_sec == 24.5

    def test_tier1_single_item_coordinates_and_confidence(self, single_modern_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(single_modern_quote_file)
        item = doc.items[0]
        assert "C2" in (item.source_cell_coords.quote_num or "")
        assert "C10" in (item.source_cell_coords.part_num or "")
        assert "C18" in (item.source_cell_coords.cycle_time or "")
        assert item.confidence == "high"
        assert len(item.warnings) == 0


class TestTier1Feature2MultiTabExtraction:
    """Core Feature 2: Multi-Tab Quote Isolation (PROJECT.md Feature 2)."""

    def test_tier1_multitab_total_items_count(self, multitab_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(multitab_quote_file)
        assert doc.quote_id == "Q2005"
        assert doc.total_items == 3
        assert len(doc.items) == 3

    def test_tier1_multitab_sheet_filtering_ignores_summary_and_lookups(self, multitab_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(multitab_quote_file)
        item_nums = [it.quote_item_number for it in doc.items]
        assert "Summary" not in item_nums
        assert "Lookups" not in item_nums
        assert item_nums == ["Q2005-1", "Q2005-2", "Q2005-3"]

    def test_tier1_multitab_part_numbers(self, multitab_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(multitab_quote_file)
        parts = [it.part_number for it in doc.items]
        assert parts == ["MPC-7011-A", "MPC-7012-B", "MPC-7013-C"]

    def test_tier1_multitab_cycle_times(self, multitab_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(multitab_quote_file)
        cts = [it.cycle_time_sec for it in doc.items]
        assert cts == [18.2, 32.0, 45.5]

    def test_tier1_multitab_sheet_scoped_coordinates(self, multitab_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(multitab_quote_file)
        for idx, item in enumerate(doc.items, start=1):
            sheet_tag = f"Q2005-{idx}"
            assert sheet_tag in (item.source_cell_coords.quote_num or "") or "C2" in (item.source_cell_coords.quote_num or "")
            assert item.confidence == "high"


class TestTier1Feature3GridMatrixExtraction:
    """Core Feature 3: 2D Tabular Grid Matrix Intersection (PROJECT.md Feature 3)."""

    def test_tier1_grid_matrix_header_detection(self, grid_matrix_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(grid_matrix_quote_file)
        assert doc.quote_id == "Q4000"
        assert doc.total_items == 3

    def test_tier1_grid_matrix_item_numbers(self, grid_matrix_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(grid_matrix_quote_file)
        item_nums = [it.quote_item_number for it in doc.items]
        assert item_nums == ["Q4000-1", "Q4000-2", "Q4000-3"]

    def test_tier1_grid_matrix_part_numbers(self, grid_matrix_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(grid_matrix_quote_file)
        parts = [it.part_number for it in doc.items]
        assert parts == ["MED-10492", "MED-10493", "MED-10494"]

    def test_tier1_grid_matrix_cycle_times(self, grid_matrix_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(grid_matrix_quote_file)
        cts = [it.cycle_time_sec for it in doc.items]
        assert cts == [14.5, 22.0, 9.8]

    def test_tier1_grid_matrix_row_coordinates_and_footer_exclusion(self, grid_matrix_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(grid_matrix_quote_file)
        # Verify footer row 11 ("Total") was not extracted as an item
        item_nums = [it.quote_item_number for it in doc.items]
        assert not any("Total" in n for n in item_nums)
        # Verify row coordinates
        assert "A8" in (doc.items[0].source_cell_coords.quote_num or "")
        assert "B8" in (doc.items[0].source_cell_coords.part_num or "")
        assert "E8" in (doc.items[0].source_cell_coords.cycle_time or "")


class TestTier1Feature4PartNumberCleansing:
    """Core Feature 4: Part Number Cleansing and Disambiguation (PROJECT.md Feature 4)."""

    def test_tier1_part_cleansing_standard(self, single_modern_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(single_modern_quote_file)
        assert doc.items[0].part_number == "2202046213000"

    def test_tier1_part_cleansing_scientific_notation(self, scientific_notation_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(scientific_notation_quote_file)
        assert doc.items[0].part_number == "2202046213000"
        assert not doc.items[0].part_number.endswith(".0")
        assert "e+" not in doc.items[0].part_number.lower()

    def test_tier1_part_cleansing_whitespace_and_nbsp(self, whitespace_variation_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(whitespace_variation_quote_file)
        part = doc.items[0].part_number
        assert "\u00a0" not in part
        assert part == "3042605C-REV1"

    def test_tier1_part_cleansing_drawing_number_exclusion(self, dwg_disambiguation_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(dwg_disambiguation_quote_file)
        # Must extract REAL-PART-1234, NOT DWG-9999-X
        assert doc.items[0].part_number == "REAL-PART-1234"
        assert "DWG" not in doc.items[0].part_number

    def test_tier1_part_cleansing_alphanumeric_with_revision(self, legacy_raw_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(legacy_raw_quote_file)
        assert doc.items[0].part_number == "994021A-REV2"


class TestTier1Feature5CycleTimeExtraction:
    """Core Feature 5: Cycle Time Extraction & Bounds Validation (PROJECT.md Feature 5)."""

    def test_tier1_cycle_time_numeric_float(self, single_modern_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(single_modern_quote_file)
        assert isinstance(doc.items[0].cycle_time_sec, float)
        assert doc.items[0].cycle_time_sec == 24.5

    def test_tier1_cycle_time_string_units_parsed(self, edge_cases_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(edge_cases_quote_file)
        txt_item = next(it for it in doc.items if it.quote_item_number == "Q6000-3")
        assert txt_item.cycle_time_sec == 26.5

    def test_tier1_cycle_time_valid_bounds_zero_warnings(self, single_modern_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(single_modern_quote_file)
        assert len(doc.items[0].warnings) == 0

    def test_tier1_cycle_time_coordinate_recording(self, single_modern_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(single_modern_quote_file)
        assert "C18" in (doc.items[0].source_cell_coords.cycle_time or "")

    def test_tier1_cycle_time_high_confidence(self, single_modern_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(single_modern_quote_file)
        assert doc.items[0].confidence == "high"


# ===========================================================================
# Tier 2: Boundary & Corner Cases (>=5 tests per feature)
# ===========================================================================

class TestTier2Feature1EmptyAndNonQuoteFiles:
    """Boundary Feature 1: Empty and Non-Quote Workbooks (PROJECT.md Feature 8)."""

    def test_tier2_empty_workbook_graceful_handling(self, empty_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(empty_quote_file)
        assert doc.total_items == 0
        assert len(doc.items) == 0
        assert len(doc.warnings) >= 1

    def test_tier2_non_quote_spreadsheet_zero_items(self, non_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(non_quote_file)
        assert doc.total_items == 0
        assert doc.items == []

    def test_tier2_nonexistent_file_raises_not_found(self, tmp_path: Path, extractor_engine):
        missing = tmp_path / "non_existent_file.xlsx"
        with pytest.raises(FileNotFoundError):
            extractor_engine.extract_file(missing)

    def test_tier2_empty_sheet_in_multitab_ignored(self, tmp_path: Path, extractor_engine):
        path = tmp_path / "quote_empty_sheet.xlsx"
        wb = generator.PureXmlWorkbook()
        ws1 = wb.active
        ws1.title = "Q1001-1"
        ws1["B2"] = "Quote Item #:"
        ws1["C2"] = "Q1001-1"
        ws1["B5"] = "Part #:"
        ws1["C5"] = "PART-A"
        ws1["B8"] = "Cycle Time:"
        ws1["C8"] = 20.0
        ws_blank = wb.create_sheet("BlankSheet")
        wb.save(path)

        doc = extractor_engine.extract_file(path)
        assert doc.total_items == 1
        assert doc.items[0].quote_item_number == "Q1001-1"

    def test_tier2_corrupted_file_raises_or_warns(self, tmp_path: Path, extractor_engine):
        corrupt = tmp_path / "corrupt.xlsx"
        corrupt.write_bytes(b"Not an excel file!")
        with pytest.raises(Exception):
            extractor_engine.extract_file(corrupt)


class TestTier2Feature2CycleTimeOutofBounds:
    """Boundary Feature 2: Out of Bounds Cycle Times (<5.0s, >300.0s) (PROJECT.md Feature 11)."""

    def test_tier2_cycle_time_low_boundary_violation(self, edge_cases_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(edge_cases_quote_file)
        low_item = next(it for it in doc.items if it.quote_item_number == "Q6000-1")
        assert low_item.cycle_time_sec == 2.4
        assert any("5.0" in w for w in low_item.warnings)
        assert low_item.confidence in ("medium", "low")

    def test_tier2_cycle_time_high_boundary_violation(self, edge_cases_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(edge_cases_quote_file)
        high_item = next(it for it in doc.items if it.quote_item_number == "Q6000-2")
        assert high_item.cycle_time_sec == 420.0
        assert any("300.0" in w for w in high_item.warnings)
        assert high_item.confidence in ("medium", "low")

    def test_tier2_cycle_time_exact_min_boundary_5_0s(self, tmp_path: Path, extractor_engine):
        path = tmp_path / "quote_boundary_5s.xlsx"
        wb = generator.PureXmlWorkbook()
        ws = wb.active
        ws.title = "Quote"
        ws["B2"] = "Quote #:"
        ws["C2"] = "Q7001-1"
        ws["B5"] = "Part #:"
        ws["C5"] = "BOUND-5S"
        ws["B8"] = "Cycle Time:"
        ws["C8"] = 5.0
        wb.save(path)

        doc = extractor_engine.extract_file(path)
        assert doc.items[0].cycle_time_sec == 5.0
        assert not any("out of bounds" in w.lower() for w in doc.items[0].warnings)

    def test_tier2_cycle_time_exact_max_boundary_300_0s(self, tmp_path: Path, extractor_engine):
        path = tmp_path / "quote_boundary_300s.xlsx"
        wb = generator.PureXmlWorkbook()
        ws = wb.active
        ws.title = "Quote"
        ws["B2"] = "Quote #:"
        ws["C2"] = "Q7002-1"
        ws["B5"] = "Part #:"
        ws["C5"] = "BOUND-300S"
        ws["B8"] = "Cycle Time:"
        ws["C8"] = 300.0
        wb.save(path)

        doc = extractor_engine.extract_file(path)
        assert doc.items[0].cycle_time_sec == 300.0
        assert not any("out of bounds" in w.lower() for w in doc.items[0].warnings)

    def test_tier2_cycle_time_negative_or_zero(self, tmp_path: Path, extractor_engine):
        path = tmp_path / "quote_zero_ct.xlsx"
        wb = generator.PureXmlWorkbook()
        ws = wb.active
        ws.title = "Quote"
        ws["B2"] = "Quote #:"
        ws["C2"] = "Q7003-1"
        ws["B5"] = "Part #:"
        ws["C5"] = "ZERO-CT"
        ws["B8"] = "Cycle Time:"
        ws["C8"] = 0.0
        wb.save(path)

        doc = extractor_engine.extract_file(path)
        assert doc.items[0].cycle_time_sec == 0.0 or doc.items[0].cycle_time_sec is None
        assert len(doc.items[0].warnings) >= 1


class TestTier2Feature3MergedCellResolution:
    """Boundary Feature 3: Merged Cells and Label Offset Traversal (PROJECT.md Feature 6)."""

    def test_tier2_merged_header_range(self, merged_cells_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(merged_cells_quote_file)
        assert doc.quote_id == "Q5000"
        assert doc.items[0].quote_item_number == "Q5000-1"

    def test_tier2_merged_part_number_label_offset(self, merged_cells_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(merged_cells_quote_file)
        assert doc.items[0].part_number == "3042605C"

    def test_tier2_merged_cycle_time_spacer_traversal(self, merged_cells_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(merged_cells_quote_file)
        assert doc.items[0].cycle_time_sec == 28.0

    def test_tier2_merged_cell_coordinates_precision(self, merged_cells_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(merged_cells_quote_file)
        assert "E14" in (doc.items[0].source_cell_coords.cycle_time or "")
        assert "D6" in (doc.items[0].source_cell_coords.part_num or "")

    def test_tier2_merged_cells_confidence(self, merged_cells_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(merged_cells_quote_file)
        assert doc.items[0].confidence in ("high", "medium")


class TestTier2Feature4FormulaAndLegacyFiles:
    """Boundary Feature 4: Formula Evaluation & Legacy Macro Workbooks (PROJECT.md Feature 1)."""

    def test_tier2_legacy_raw_quote_id(self, legacy_raw_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(legacy_raw_quote_file)
        assert doc.quote_id == "Q3050"

    def test_tier2_legacy_raw_inferred_item_suffix(self, legacy_raw_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(legacy_raw_quote_file)
        assert doc.items[0].quote_item_number == "Q3050-1"
        assert doc.items[0].confidence in ("medium", "high")
        assert any("inferred" in w.lower() for w in doc.items[0].warnings)

    def test_tier2_legacy_formula_cached_cycle_time(self, legacy_raw_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(legacy_raw_quote_file)
        assert doc.items[0].cycle_time_sec == 27.5

    def test_tier2_formula_error_na_handling(self, edge_cases_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(edge_cases_quote_file)
        missing_item = next(it for it in doc.items if it.quote_item_number == "Q6000-4")
        assert missing_item.cycle_time_sec is None
        assert missing_item.confidence == "low"
        assert len(missing_item.warnings) >= 1

    def test_tier2_macro_xlsm_opened_cleanly(self, legacy_raw_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(legacy_raw_quote_file)
        assert doc.total_items == 1
        assert doc.items[0].part_number == "994021A-REV2"


class TestTier2Feature5WhitespaceAndFormatVariations:
    """Boundary Feature 5: Whitespace, Delimiters & Case Variations (PROJECT.md Feature 4)."""

    def test_tier2_whitespace_quote_id_normalization(self, whitespace_variation_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(whitespace_variation_quote_file)
        assert doc.quote_id == "Q8000"
        assert doc.items[0].quote_item_number == "Q8000-1"

    def test_tier2_whitespace_part_number_strip(self, whitespace_variation_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(whitespace_variation_quote_file)
        assert doc.items[0].part_number == "3042605C-REV1"

    def test_tier2_whitespace_cycle_time_extracted(self, whitespace_variation_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(whitespace_variation_quote_file)
        assert doc.items[0].cycle_time_sec == 25.0

    def test_tier2_whitespace_anchor_casing_resilience(self, whitespace_variation_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(whitespace_variation_quote_file)
        assert doc.items[0].source_cell_coords.cycle_time is not None

    def test_tier2_whitespace_clean_output_no_control_chars(self, whitespace_variation_quote_file: Path, extractor_engine):
        doc = extractor_engine.extract_file(whitespace_variation_quote_file)
        for it in doc.items:
            assert "\n" not in (it.part_number or "")
            assert "\t" not in (it.part_number or "")


# ===========================================================================
# Tier 3: Cross-Feature Combinations (>=10 tests)
# ===========================================================================

class TestTier3CrossFeatureCombinations:
    """Pairwise Cross-Feature Tests combining multiple complex layouts and anomalies."""

    def test_tier3_multitab_with_merged_cells(self, tmp_path: Path, extractor_engine):
        """Multi-tab workbook where each item tab uses merged cells for labels and values."""
        path = tmp_path / "multitab_merged.xlsx"
        wb = generator.PureXmlWorkbook()
        ws1 = wb.active
        ws1.title = "Q1100-1"
        ws1.merge_cells("B2:C2")
        ws1["B2"] = "Quote Item:"
        ws1["D2"] = "Q1100-1"
        ws1.merge_cells("B5:C5")
        ws1["B5"] = "Customer P/N:"
        ws1["D5"] = "PART-MERGE-A"
        ws1.merge_cells("B8:C8")
        ws1["B8"] = "Cycle Time (s):"
        ws1["D8"] = 17.5

        ws2 = wb.create_sheet("Q1100-2")
        ws2.merge_cells("B2:C2")
        ws2["B2"] = "Quote Item:"
        ws2["D2"] = "Q1100-2"
        ws2.merge_cells("B5:C5")
        ws2["B5"] = "Part #:"
        ws2["D5"] = "PART-MERGE-B"
        ws2.merge_cells("B8:C8")
        ws2["B8"] = "Cycle Time:"
        ws2["D8"] = 29.0
        wb.save(path)

        doc = extractor_engine.extract_file(path)
        assert doc.quote_id == "Q1100"
        assert doc.total_items == 2
        assert doc.items[0].part_number == "PART-MERGE-A"
        assert doc.items[1].cycle_time_sec == 29.0

    def test_tier3_grid_matrix_with_out_of_bounds_cycle_times(self, tmp_path: Path, extractor_engine):
        """2D Grid where item 1 is normal, item 2 is <5.0s, and item 3 is >300s."""
        path = tmp_path / "grid_oob.xlsx"
        wb = generator.PureXmlWorkbook()
        ws = wb.active
        ws.title = "GridOOB"
        ws["B2"] = "Quote Reference:"
        ws["C2"] = "Q1200"
        headers = ["Item #", "Part #", "Cycle Time"]
        for c_idx, h in enumerate(headers, start=1):
            ws.cell(row=5, column=c_idx, value=h)
        # Rows
        ws.cell(row=6, column=1, value="Q1200-1")
        ws.cell(row=6, column=2, value="GRID-NORMAL")
        ws.cell(row=6, column=3, value=25.0)

        ws.cell(row=7, column=1, value="Q1200-2")
        ws.cell(row=7, column=2, value="GRID-FAST")
        ws.cell(row=7, column=3, value=3.2)

        ws.cell(row=8, column=1, value="Q1200-3")
        ws.cell(row=8, column=2, value="GRID-SLOW")
        ws.cell(row=8, column=3, value=550.0)
        wb.save(path)

        doc = extractor_engine.extract_file(path)
        assert doc.total_items == 3
        assert len(doc.items[0].warnings) == 0
        assert any("5.0" in w for w in doc.items[1].warnings)
        assert any("300.0" in w for w in doc.items[2].warnings)

    def test_tier3_grid_matrix_with_formula_cached_evaluations(self, tmp_path: Path, extractor_engine):
        """2D Grid where cycle times are calculated by formula with cached evaluation."""
        path = tmp_path / "grid_formula.xlsx"
        wb = generator.PureXmlWorkbook()
        ws = wb.active
        ws.title = "GridFormula"
        ws["B2"] = "Quote #:"
        ws["C2"] = "Q1300"
        for c_idx, h in enumerate(["Item #", "Part #", "Inj Time", "Cool Time", "Cycle Time"], start=1):
            ws.cell(row=5, column=c_idx, value=h)
        ws.cell(row=6, column=1, value="Q1300-1")
        ws.cell(row=6, column=2, value="FORMULA-PART")
        ws.cell(row=6, column=3, value=10.0)
        ws.cell(row=6, column=4, value=15.0)
        ws.cell(row=6, column=5, value=25.0, formula="=C6+D6")
        wb.save(path)

        doc = extractor_engine.extract_file(path)
        assert doc.items[0].cycle_time_sec == 25.0

    def test_tier3_legacy_xlsm_with_multiple_items(self, tmp_path: Path, extractor_engine):
        """Macro .xlsm with multiple item tabs."""
        path = tmp_path / "legacy_multitab.xlsm"
        wb = generator.PureXmlWorkbook()
        ws1 = wb.active
        ws1.title = "Item 1"
        ws1["B2"] = "Quote ID:"
        ws1["C2"] = "Q1400"
        ws1["B5"] = "Part No:"
        ws1["C5"] = "LEGACY-A"
        ws1["B8"] = "Cycle Time:"
        ws1["C8"] = 22.5

        ws2 = wb.create_sheet("Item 2")
        ws2["B2"] = "Quote ID:"
        ws2["C2"] = "Q1400"
        ws2["B5"] = "Part No:"
        ws2["C5"] = "LEGACY-B"
        ws2["B8"] = "Cycle Time:"
        ws2["C8"] = 35.0
        wb.save(path)

        doc = extractor_engine.extract_file(path)
        assert doc.quote_id == "Q1400"
        assert doc.total_items == 2

    def test_tier3_hybrid_matrix_and_child_tabs(self, hybrid_quote_file: Path, extractor_engine):
        """Hybrid workbook with 2D master summary and child detail sheets."""
        doc = extractor_engine.extract_file(hybrid_quote_file)
        assert doc.quote_id == "Q10050"
        # Must isolate 2 unique items without duplicate inflation
        assert doc.total_items == 2
        item_nums = [it.quote_item_number for it in doc.items]
        assert item_nums == ["Q10050-1", "Q10050-2"]

    def test_tier3_grid_matrix_with_scientific_notation(self, tmp_path: Path, extractor_engine):
        """2D Grid containing 13-digit numeric part numbers."""
        path = tmp_path / "grid_scinot.xlsx"
        wb = generator.PureXmlWorkbook()
        ws = wb.active
        ws.title = "GridSci"
        ws["B2"] = "Quote #:"
        ws["C2"] = "Q1500"
        ws.cell(row=5, column=1, value="Item #")
        ws.cell(row=5, column=2, value="Part Number")
        ws.cell(row=5, column=3, value="Cycle Time")
        ws.cell(row=6, column=1, value="Q1500-1")
        ws.cell(row=6, column=2, value=2202046213000)
        ws.cell(row=6, column=3, value=21.0)
        wb.save(path)

        doc = extractor_engine.extract_file(path)
        assert doc.items[0].part_number == "2202046213000"

    def test_tier3_merged_cells_missing_part_number(self, tmp_path: Path, extractor_engine):
        """Merged layout where part number is missing but cycle time is valid."""
        path = tmp_path / "merged_missing_pn.xlsx"
        wb = generator.PureXmlWorkbook()
        ws = wb.active
        ws.title = "MissingPN"
        ws.merge_cells("B2:D2")
        ws["B2"] = "QUOTE Q1600-1"
        ws.merge_cells("B5:C5")
        ws["B5"] = "Part Number:"
        ws["D5"] = ""  # empty
        ws.merge_cells("B8:C8")
        ws["B8"] = "Cycle Time:"
        ws["D8"] = 23.0
        wb.save(path)

        doc = extractor_engine.extract_file(path)
        assert doc.items[0].part_number is None or doc.items[0].part_number == ""
        assert doc.items[0].confidence == "low"
        assert len(doc.items[0].warnings) >= 1

    def test_tier3_legacy_xlsm_with_whitespace_and_dwg(self, tmp_path: Path, extractor_engine):
        """Macro workbook with noisy whitespace and drawing number."""
        path = tmp_path / "legacy_noisy.xlsm"
        wb = generator.PureXmlWorkbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["C2"] = "  Quote ID :  "
        ws["D2"] = "  Q1700  "
        ws["C5"] = "DWG #:"
        ws["D5"] = "DWG-12345"
        ws["C6"] = "Customer Part:"
        ws["D6"] = "  MPC-5544-B \u00a0"
        ws["C8"] = "Cycle Time:"
        ws["D8"] = 30.5
        wb.save(path)

        doc = extractor_engine.extract_file(path)
        assert doc.quote_id == "Q1700"
        assert doc.items[0].quote_item_number == "Q1700-1"
        assert doc.items[0].part_number == "MPC-5544-B"

    def test_tier3_multitab_with_one_missing_cycle_time(self, tmp_path: Path, extractor_engine):
        """Multi-tab where tab 1 is valid, tab 2 has missing cycle time, tab 3 is valid."""
        path = tmp_path / "multitab_missing_ct.xlsx"
        wb = generator.PureXmlWorkbook()
        ws1 = wb.active
        ws1.title = "Q1800-1"
        ws1["B2"] = "Quote Item #:"
        ws1["C2"] = "Q1800-1"
        ws1["B5"] = "Part #:"
        ws1["C5"] = "PART-1"
        ws1["B8"] = "Cycle Time:"
        ws1["C8"] = 18.0

        ws2 = wb.create_sheet("Q1800-2")
        ws2["B2"] = "Quote Item #:"
        ws2["C2"] = "Q1800-2"
        ws2["B5"] = "Part #:"
        ws2["C5"] = "PART-2"
        ws2["B8"] = "Cycle Time:"
        ws2["C8"] = ""  # missing

        ws3 = wb.create_sheet("Q1800-3")
        ws3["B2"] = "Quote Item #:"
        ws3["C2"] = "Q1800-3"
        ws3["B5"] = "Part #:"
        ws3["C5"] = "PART-3"
        ws3["B8"] = "Cycle Time:"
        ws3["C8"] = 28.0
        wb.save(path)

        doc = extractor_engine.extract_file(path)
        assert doc.total_items == 3
        assert doc.items[0].confidence == "high"
        assert doc.items[1].confidence == "low"
        assert doc.items[1].cycle_time_sec is None
        assert doc.items[2].confidence == "high"

    def test_tier3_multitab_mixed_confidence_distribution(self, tmp_path: Path, extractor_engine):
        """Verify that document items accurately reflect varied confidence scores."""
        path = tmp_path / "mixed_conf.xlsx"
        wb = generator.PureXmlWorkbook()
        # Item 1: High
        ws1 = wb.active
        ws1.title = "Q1900-1"
        ws1["B2"] = "Quote Item #:"
        ws1["C2"] = "Q1900-1"
        ws1["B5"] = "Part #:"
        ws1["C5"] = "HIGH-CONF"
        ws1["B8"] = "Cycle Time:"
        ws1["C8"] = 20.0
        # Item 2: Low (cycle time out of bounds)
        ws2 = wb.create_sheet("Q1900-2")
        ws2["B2"] = "Quote Item #:"
        ws2["C2"] = "Q1900-2"
        ws2["B5"] = "Part #:"
        ws2["C5"] = "LOW-CONF"
        ws2["B8"] = "Cycle Time:"
        ws2["C8"] = 600.0
        wb.save(path)

        doc = extractor_engine.extract_file(path)
        assert doc.items[0].confidence == "high"
        assert doc.items[1].confidence == "low"


# ===========================================================================
# Tier 4: Real-World Workload Scenarios (>=10 tests)
# ===========================================================================

class TestTier4RealWorldWorkloadScenarios:
    """End-to-End realistic manufacturing workloads, directory scanning, and exports."""

    def test_tier4_batch_directory_scanning_excludes_lock_files(self, batch_directory_fixture: Path, extractor_engine):
        """Batch scanner ignores ~$* temporary lock files."""
        results = extractor_engine.extract_directory(str(batch_directory_fixture), recursive=True)
        filenames = [Path(r.source_file).name for r in results]
        assert not any(fn.startswith("~$") for fn in filenames)

    def test_tier4_batch_directory_scanning_finds_nested_quotes(self, batch_directory_fixture: Path, extractor_engine):
        """Batch scanner discovers quotes in nested subdirectories."""
        results = extractor_engine.extract_directory(str(batch_directory_fixture), recursive=True)
        quote_ids = [r.quote_id for r in results if r.quote_id]
        assert "Q3050" in quote_ids  # Nested in subfolder_legacy/

    def test_tier4_batch_directory_extracts_all_archetypes(self, batch_directory_fixture: Path, extractor_engine):
        """Batch run processes single, multitab, grid, and macro workbooks together."""
        results = extractor_engine.extract_directory(str(batch_directory_fixture), recursive=True)
        quote_ids = set(r.quote_id for r in results if r.quote_id)
        assert {"Q1001", "Q2005", "Q4000", "Q3050"}.issubset(quote_ids)

    def test_tier4_all_fixtures_schema_conformance(self, all_synthetic_fixtures: Dict[str, Path], extractor_engine):
        """Every generated fixture parses into a schema-compliant QuoteDocument."""
        for name, path in all_synthetic_fixtures.items():
            doc = extractor_engine.extract_file(path)
            assert hasattr(doc, "source_file")
            assert hasattr(doc, "quote_id")
            assert hasattr(doc, "total_items")
            assert hasattr(doc, "items")
            assert hasattr(doc, "warnings")
            assert len(doc.items) == doc.total_items

    def test_tier4_batch_export_json_validity(self, single_modern_quote_file: Path, extractor_engine):
        """Extracted QuoteDocument serializes to valid JSON matching R2 schema."""
        doc = extractor_engine.extract_file(single_modern_quote_file)
        # Verify JSON dump
        if hasattr(doc, "model_dump_json"):
            json_str = doc.model_dump_json()
        else:
            json_str = json.dumps(doc.__dict__, default=str)
        data = json.loads(json_str)
        assert data["quote_id"] == "Q1001"
        assert len(data["items"]) == 1
        assert "source_cell_coords" in data["items"][0]

    def test_tier4_batch_export_csv_representation(self, multitab_quote_file: Path, extractor_engine):
        """Extracted items can be formatted into clean tabular records for CSV."""
        doc = extractor_engine.extract_file(multitab_quote_file)
        rows = []
        for item in doc.items:
            rows.append({
                "source_file": Path(doc.source_file).name,
                "quote_id": doc.quote_id,
                "quote_item_number": item.quote_item_number,
                "part_number": item.part_number,
                "cycle_time_sec": item.cycle_time_sec,
                "confidence": item.confidence,
                "warnings": "; ".join(item.warnings),
            })
        assert len(rows) == 3
        assert rows[0]["quote_item_number"] == "Q2005-1"
        assert rows[1]["part_number"] == "MPC-7012-B"

    def test_tier4_drawing_number_never_overwrites_part_number(self, single_modern_quote_file: Path, extractor_engine):
        """DWG-8841-B in modern quote is never extracted as the part number."""
        doc = extractor_engine.extract_file(single_modern_quote_file)
        assert doc.items[0].part_number == "2202046213000"
        assert "DWG" not in doc.items[0].part_number

    def test_tier4_auxiliary_tabs_excluded_across_workload(self, multitab_quote_file: Path, extractor_engine):
        """TOC, Summary, and Lookups sheets are consistently excluded across runs."""
        doc = extractor_engine.extract_file(multitab_quote_file)
        for it in doc.items:
            assert "Summary" not in it.quote_item_number
            assert "Lookups" not in it.quote_item_number

    def test_tier4_source_coordinates_traceability_across_items(self, grid_matrix_quote_file: Path, extractor_engine):
        """All items extracted have non-null source cell coordinates."""
        doc = extractor_engine.extract_file(grid_matrix_quote_file)
        for it in doc.items:
            assert it.source_cell_coords.quote_num is not None
            assert it.source_cell_coords.part_num is not None
            assert it.source_cell_coords.cycle_time is not None

    def test_tier4_batch_performance_under_threshold(self, all_synthetic_fixtures: Dict[str, Path], extractor_engine):
        """Batch extraction of 12 distinct workbooks executes rapidly without hanging."""
        import time
        start = time.time()
        for path in all_synthetic_fixtures.values():
            extractor_engine.extract_file(path)
        duration = time.time() - start
        # 12 synthetic files should parse well under 10 seconds in Python
        assert duration < 10.0, f"Batch parse took {duration:.2f}s, expected < 10.0s"
