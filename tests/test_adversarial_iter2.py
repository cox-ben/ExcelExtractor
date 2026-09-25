"""
Adversarial Stress Test Suite - Iteration 2 Challenger
Tests core extraction engine against boundary conditions, adversarial inputs,
grid matrix termination edge cases, and multi-tab administrative sheet filtering.
"""

import os
from pathlib import Path
import pytest

from excel_extractor.config import ConfigManager, ExtractorConfig
from excel_extractor.engine import QuoteExtractorEngine
from excel_extractor.models import QuoteDocument, QuoteItem, SourceCellCoords
from excel_extractor.reader import Workbook, Worksheet, load_workbook
from excel_extractor.strategies.grid_matrix import GridMatrixStrategy
from excel_extractor.strategies.multi_tab import MultiTabStrategy
from tests.fixtures.generator import PureXmlWorkbook


# ===========================================================================
# 1. Boundary Cycle Time Adversarial Tests
# ===========================================================================

class TestAdversarialCycleTimeBoundaries:
    """Stress-tests cycle time boundaries (4.99s, 5.0s, 300.0s, 300.01s, 2.4s, 450.0s)."""

    @pytest.mark.parametrize(
        "ct_input,expected_val,is_out_of_bounds",
        [
            (2.4, 2.4, True),
            ("2.4s", 2.4, True),
            (4.99, 4.99, True),
            ("4.99 sec", 4.99, True),
            ("4.99 seconds", 4.99, True),
            (5.0, 5.0, False),
            ("5.0 s", 5.0, False),
            (5.00, 5.0, False),
            (15.0, 15.0, False),
            (24.5, 24.5, False),
            (150.0, 150.0, False),
            (300.0, 300.0, False),
            ("300.0 sec", 300.0, False),
            (300.01, 300.01, True),
            ("300.01s", 300.01, True),
            (450.0, 450.0, True),
            ("450.0 seconds", 450.0, True),
        ],
    )
    def test_cycle_time_boundary_values_model_contract(
        self, ct_input, expected_val, is_out_of_bounds
    ):
        """Verify cycle times at boundary thresholds preserve float values and flag warnings."""
        item = QuoteItem(
            quote_item_number="Q1001-1",
            part_number="ABC-100",
            cycle_time_sec=ct_input,
        )

        assert item.cycle_time_sec == pytest.approx(expected_val, rel=1e-3)

        if is_out_of_bounds:
            assert item.confidence == "low"
            assert any("outside expected range (5.0s–300.0s)" in w for w in item.warnings)
        else:
            assert item.confidence == "high"
            assert not any("outside expected range" in w for w in item.warnings)

    def test_cycle_time_mutation_does_not_recurse(self):
        """Verify mutating cycle time on an existing item does not trigger infinite recursion."""
        item = QuoteItem(
            quote_item_number="Q1001-1",
            part_number="ABC-100",
            cycle_time_sec=25.0,
        )
        assert item.confidence == "high"

        # Mutate to out-of-bounds low
        item.cycle_time_sec = 4.99
        assert item.confidence == "low"
        assert any("outside expected range" in w for w in item.warnings)

        # Mutate to out-of-bounds high
        item.cycle_time_sec = 300.01
        assert item.confidence == "low"

        # Mutate to valid boundary
        item.cycle_time_sec = 5.0
        # Once warnings are in the list, confidence stays medium/low unless warnings are cleared
        assert item.cycle_time_sec == 5.0

    @pytest.mark.parametrize(
        "boundary_ct,should_warn",
        [
            (2.4, True),
            (4.99, True),
            (5.0, False),
            (300.0, False),
            (300.01, True),
            (450.0, True),
        ],
    )
    def test_cycle_time_boundaries_end_to_end_workbook(self, tmp_path, boundary_ct, should_warn):
        """Generate actual Excel workbooks with boundary cycle times and run QuoteExtractorEngine."""
        wb = PureXmlWorkbook()
        ws = wb.active
        ws["A1"] = "Quote #"
        ws["B1"] = "Q2000-1"
        ws["A3"] = "Part Number"
        ws["B3"] = "PART-BOUND"
        ws["A5"] = "Cycle Time"
        ws["B5"] = boundary_ct

        file_path = tmp_path / f"quote_ct_{boundary_ct}.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 1
        item = doc.items[0]
        assert item.part_number == "PART-BOUND"
        assert item.cycle_time_sec == pytest.approx(boundary_ct, rel=1e-2)

        has_oob_warn = any("outside expected range" in w for w in item.warnings)
        if should_warn:
            assert has_oob_warn, f"Expected warning for cycle time {boundary_ct}"
            assert item.confidence == "low"
        else:
            assert not has_oob_warn, f"Did not expect warning for boundary cycle time {boundary_ct}"
            assert item.confidence == "high"

    def test_cycle_time_extreme_anomalies(self):
        """Verify extreme anomalies: 0.0, negative, None, non-numeric."""
        item_zero = QuoteItem(quote_item_number="Q1001-1", part_number="P1", cycle_time_sec=0.0)
        assert item_zero.confidence == "low"
        assert any("outside expected range" in w for w in item_zero.warnings)

        item_neg = QuoteItem(quote_item_number="Q1001-1", part_number="P1", cycle_time_sec=-10.5)
        assert item_neg.confidence == "low"
        assert any("outside expected range" in w for w in item_neg.warnings)

        item_none = QuoteItem(quote_item_number="Q1001-1", part_number="P1", cycle_time_sec=None)
        assert item_none.confidence == "low"
        assert any("Missing cycle_time_sec" in w for w in item_none.warnings)


# ===========================================================================
# 2. 2D Grid Matrix Row Traversal Termination Adversarial Tests
# ===========================================================================

class TestAdversarialGridMatrixTermination:
    """Stress-tests 2D grid row traversal termination on stop words, footers, blank lines."""

    def test_grid_matrix_terminates_on_blank_row_and_ignores_footer(self, tmp_path):
        """Verify extraction terminates on the first empty row after data and does not bleed into footers."""
        wb = PureXmlWorkbook()
        ws = wb.active
        ws["A1"] = "Quote # Q3001"

        # Grid Header at row 5
        ws["A5"] = "Item #"
        ws["B5"] = "Part Number"
        ws["C5"] = "Cycle Time (s)"

        # Data Rows 6-8
        ws["A6"] = "1"
        ws["B6"] = "PART-A"
        ws["C6"] = 15.0

        ws["A7"] = "2"
        ws["B7"] = "PART-B"
        ws["C7"] = 25.0

        ws["A8"] = "3"
        ws["B8"] = "PART-C"
        ws["C8"] = 35.0

        # Row 9 is completely empty (blank)

        # Row 10-15: Footer junk that should NOT be extracted as items
        ws["A10"] = "Footer Notes: All parts subject to review"
        ws["B10"] = "PART-JUNK"
        ws["C10"] = 50.0

        ws["A12"] = "Approval Signoff"
        ws["B12"] = "Engineering"
        ws["C12"] = 100.0

        file_path = tmp_path / "grid_blank_break.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 3
        part_nums = [it.part_number for it in doc.items]
        assert part_nums == ["PART-A", "PART-B", "PART-C"]
        assert "PART-JUNK" not in part_nums
        assert "Engineering" not in part_nums

    @pytest.mark.parametrize(
        "stop_label",
        [
            "Total",
            "TOTAL",
            "Subtotal",
            "SUBTOTAL",
            "Totals",
            "Grand Total",
            "Total Amount",
            "Notes",
            "Notes:",
            "Approved by",
            "Approved",
        ],
    )
    def test_grid_matrix_terminates_on_stop_keyword_rows(self, tmp_path, stop_label):
        """Verify extraction terminates immediately when encountering any stop keyword row."""
        wb = PureXmlWorkbook()
        ws = wb.active
        ws["A1"] = "Quote Ref: Q3002"

        ws["A4"] = "Item"
        ws["B4"] = "Part #"
        ws["C4"] = "Cycle Time"

        ws["A5"] = "1"
        ws["B5"] = "P-101"
        ws["C5"] = 12.0

        ws["A6"] = "2"
        ws["B6"] = "P-102"
        ws["C6"] = 18.5

        # Stop keyword row at row 7
        ws["A7"] = stop_label
        ws["B7"] = "Summary Row"
        ws["C7"] = 30.5  # sum of cycle times or dummy number

        # Trailing row after stop keyword at row 8
        ws["A8"] = "3"
        ws["B8"] = "P-103-SHOULD-NOT-EXTRACT"
        ws["C8"] = 22.0

        file_path = tmp_path / f"grid_stop_{stop_label.replace(' ', '_').replace(':', '')}.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 2, f"Failed on stop keyword: '{stop_label}', extracted {doc.total_items} items"
        part_nums = [it.part_number for it in doc.items]
        assert part_nums == ["P-101", "P-102"]

    def test_grid_matrix_spacer_blank_row_below_header(self, tmp_path):
        """Verify grid matrix tolerates a single blank spacer row between header and first data item."""
        wb = PureXmlWorkbook()
        ws = wb.active
        ws["A1"] = "Quote ID: Q3003"

        ws["A3"] = "Item #"
        ws["B3"] = "Part Number"
        ws["C3"] = "Cycle Time"

        # Row 4 is a blank spacer row!

        # Data Rows 5-6
        ws["A5"] = "1"
        ws["B5"] = "P-SPACER-1"
        ws["C5"] = 20.0

        ws["A6"] = "2"
        ws["B6"] = "P-SPACER-2"
        ws["C6"] = 30.0

        file_path = tmp_path / "grid_spacer_row.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 2
        assert [it.part_number for it in doc.items] == ["P-SPACER-1", "P-SPACER-2"]

    def test_grid_matrix_stops_after_two_consecutive_blank_rows_before_data(self, tmp_path):
        """Verify grid terminates cleanly without extracting if table area has 2 consecutive blank rows."""
        wb = PureXmlWorkbook()
        ws = wb.active
        ws["A1"] = "Quote ID: Q3004"

        ws["A3"] = "Item #"
        ws["B3"] = "Part Number"
        ws["C3"] = "Cycle Time"

        # Rows 4 and 5 are blank
        # Row 6 has some orphan text
        ws["A6"] = "Orphan Note"
        ws["B6"] = "Should Not Extract"
        ws["C6"] = 55.0

        file_path = tmp_path / "grid_two_blanks.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 0

    def test_grid_matrix_large_50_item_stress(self, tmp_path):
        """Stress-test a large 50-item table with stop keyword and ensure exact boundary row termination."""
        wb = PureXmlWorkbook()
        ws = wb.active
        ws["A1"] = "Quote Q3050"

        ws["A5"] = "Line #"
        ws["B5"] = "Customer Part"
        ws["C5"] = "Cycle Time (sec)"

        for i in range(1, 51):
            row = 5 + i
            ws[f"A{row}"] = str(i)
            ws[f"B{row}"] = f"PART-{i:03d}"
            ws[f"C{row}"] = 10.0 + (i % 30)

        # Row 56 is Total
        ws["A56"] = "Grand Total"
        ws["B56"] = "50 parts"
        ws["C56"] = 750.0

        # Rows 57-60 junk
        ws["A57"] = "Payment Terms: Net 30 Days"
        ws["B57"] = "FOB Destination"

        file_path = tmp_path / "grid_50_items.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 50
        assert doc.items[0].quote_item_number == "Q3050-1"
        assert doc.items[0].part_number == "PART-001"
        assert doc.items[0].source_cell_coords.part_num == "B6"

        assert doc.items[-1].quote_item_number == "Q3050-50"
        assert doc.items[-1].part_number == "PART-050"
        assert doc.items[-1].source_cell_coords.part_num == "B55"


# ===========================================================================
# 3. Multi-Tab Administrative Sheet Filtering Adversarial Tests
# ===========================================================================

class TestAdversarialMultiTabAdministrativeFiltering:
    """Stress-tests multi-tab extraction with ignored administrative tabs."""

    def test_multi_tab_complete_administrative_suite_filtering(self, tmp_path):
        """Verify all configured administrative tabs are ignored and quote ID is resolved across tabs."""
        wb = PureXmlWorkbook()

        # Sheet 1: Summary (Administrative - holds global quote ID)
        ws_summary = wb.active
        ws_summary.title = "Summary"
        ws_summary["A1"] = "Master Quote Summary"
        ws_summary["B1"] = "Q4001"
        ws_summary["A3"] = "Total Projects: 2"

        # Sheet 2: Cover Page (Administrative)
        ws_cover = wb.create_sheet("Cover")
        ws_cover["A1"] = "Confidential Quotation Document"

        # Sheet 3: TOC (Administrative)
        ws_toc = wb.create_sheet("TOC")
        ws_toc["A1"] = "Table of Contents"

        # Sheet 4: Item 1 (Valid Quote Item)
        ws_item1 = wb.create_sheet("Item 1")
        ws_item1["A2"] = "Quote #"
        ws_item1["B2"] = "Q4001-1"
        ws_item1["A4"] = "Part Number"
        ws_item1["B4"] = "MOLD-ALPHA"
        ws_item1["A6"] = "Cycle Time"
        ws_item1["B6"] = 18.2

        # Sheet 5: Terms & Conditions (Administrative)
        ws_terms = wb.create_sheet("Terms & Conditions")
        ws_terms["A1"] = "Standard Sale Terms"

        # Sheet 6: Item 2 (Valid Quote Item)
        ws_item2 = wb.create_sheet("Item 2")
        ws_item2["A2"] = "Quote #"
        ws_item2["B2"] = "Q4001-2"
        ws_item2["A4"] = "Part Number"
        ws_item2["B4"] = "MOLD-BETA"
        ws_item2["A6"] = "Cycle Time"
        ws_item2["B6"] = 42.0

        # Sheet 7: Notes (Administrative)
        ws_notes = wb.create_sheet("Notes")
        ws_notes["A1"] = "Tooling lead time: 8 weeks"

        # Sheet 8: Lookup (Administrative)
        ws_lookup = wb.create_sheet("Lookup")
        ws_lookup["A1"] = "Resin Code"
        ws_lookup["B1"] = "Nylon 6/6"

        # Sheet 9: Instructions (Administrative)
        ws_instr = wb.create_sheet("Instructions")
        ws_instr["A1"] = "Do not edit green cells"

        file_path = tmp_path / "multitab_admin_suite.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        # Only Item 1 and Item 2 should be extracted
        assert doc.quote_id == "Q4001"
        assert doc.total_items == 2

        item_nums = [it.quote_item_number for it in doc.items]
        part_nums = [it.part_number for it in doc.items]

        assert item_nums == ["Q4001-1", "Q4001-2"]
        assert part_nums == ["MOLD-ALPHA", "MOLD-BETA"]

        # Verify source cell coordinates have sheet prefixes
        assert doc.items[0].source_cell_coords.quote_num == "Item 1!B2"
        assert doc.items[0].source_cell_coords.part_num == "Item 1!B4"
        assert doc.items[0].source_cell_coords.cycle_time == "Item 1!B6"

        assert doc.items[1].source_cell_coords.quote_num == "Item 2!B2"
        assert doc.items[1].source_cell_coords.part_num == "Item 2!B4"
        assert doc.items[1].source_cell_coords.cycle_time == "Item 2!B6"

    @pytest.mark.parametrize(
        "admin_tab_name",
        [
            "Summary",
            "SUMMARY",
            " summary ",
            "Cover Page",
            "Project Cover",
            "TOC",
            "Terms",
            "Notes",
            "Note",
            "Lookup",
            "Lookups",
            "Lookup Data",
            "Instructions",
            "Instruction",
        ],
    )
    def test_administrative_tab_name_variations_ignored(self, tmp_path, admin_tab_name):
        """Verify variations of administrative sheet names are properly caught by the filter."""
        wb = PureXmlWorkbook()
        ws_admin = wb.active
        ws_admin.title = admin_tab_name
        ws_admin["A1"] = "Administrative Content"
        ws_admin["B1"] = "Q5000"
        ws_admin["A3"] = "Part Number"
        ws_admin["B3"] = "IGNORE-ME"
        ws_admin["A5"] = "Cycle Time"
        ws_admin["B5"] = 25.0

        # Valid Item Sheet 1
        ws_item = wb.create_sheet("Item 1")
        ws_item["A1"] = "Quote #"
        ws_item["B1"] = "Q5000-1"
        ws_item["A3"] = "Part #"
        ws_item["B3"] = "REAL-PART-1"
        ws_item["A5"] = "Cycle Time"
        ws_item["B5"] = 33.0

        # Valid Item Sheet 2
        ws_item2 = wb.create_sheet("Item 2")
        ws_item2["A1"] = "Quote #"
        ws_item2["B1"] = "Q5000-2"
        ws_item2["A3"] = "Part #"
        ws_item2["B3"] = "REAL-PART-2"
        ws_item2["A5"] = "Cycle Time"
        ws_item2["B5"] = 44.0

        file_path = tmp_path / f"admin_var_{admin_tab_name.strip().replace(' ', '_')}.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 2, f"Failed to filter administrative sheet named '{admin_tab_name}' in multi-tab mode"
        part_nums = [it.part_number for it in doc.items]
        assert part_nums == ["REAL-PART-1", "REAL-PART-2"]
        assert "IGNORE-ME" not in part_nums

    def test_single_item_with_preceding_administrative_sheet(self, tmp_path):
        """EXPOSES DEFECT: A workbook with 1 administrative sheet (e.g. Cover or Summary)
        and 1 quote item sheet is routed to SingleItemStrategy, which blindly reads worksheets[0]
        and fails to filter the administrative sheet.
        """
        wb = PureXmlWorkbook()
        ws_cover = wb.active
        ws_cover.title = "Cover"
        ws_cover["A1"] = "Project Cover Sheet"
        ws_cover["B1"] = "Q5000"
        # Even if Cover has some notes
        ws_cover["A3"] = "Notes"
        ws_cover["B3"] = "Confidential"

        # Item 1 is the actual quote sheet
        ws_item = wb.create_sheet("Quote Item")
        ws_item["A1"] = "Quote #"
        ws_item["B1"] = "Q5000-1"
        ws_item["A3"] = "Part #"
        ws_item["B3"] = "TRUE-PART-100"
        ws_item["A5"] = "Cycle Time"
        ws_item["B5"] = 28.5

        file_path = tmp_path / "single_item_with_cover.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        # The engine should extract TRUE-PART-100 from 'Quote Item', not empty/notes from 'Cover'
        assert doc.total_items == 1
        assert doc.items[0].part_number == "TRUE-PART-100"

    def test_single_item_with_cover_sheet_no_quote_id_yields_zero_items(self, tmp_path):
        """EXPOSES DEFECT VARIANT: When Cover sheet has no quote ID and no parts,
        SingleItemStrategy checks only worksheets[0], finds nothing, and returns 0 items,
        completely ignoring the valid quote on worksheets[1].
        """
        wb = PureXmlWorkbook()
        ws_cover = wb.active
        ws_cover.title = "Cover"
        ws_cover["A1"] = "Company Confidential Overview"

        ws_item = wb.create_sheet("Quote Item")
        ws_item["A1"] = "Quote #"
        ws_item["B1"] = "Q5002-1"
        ws_item["A3"] = "Part #"
        ws_item["B3"] = "TRUE-PART-200"
        ws_item["A5"] = "Cycle Time"
        ws_item["B5"] = 19.5

        file_path = tmp_path / "single_item_cover_no_qid.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        # Expected: Engine extracts TRUE-PART-200 from 'Quote Item'
        # Actual Defect: doc.total_items == 0 because worksheets[0] was blank and worksheets[1] ignored
        assert doc.total_items == 1
        assert doc.items[0].part_number == "TRUE-PART-200"

    def test_multi_tab_hybrid_with_grid_sheet(self, tmp_path):
        """Verify multi-tab workbook where one sheet is a 2D grid matrix and another is a single-item form."""
        wb = PureXmlWorkbook()

        # Summary Tab
        ws_sum = wb.active
        ws_sum.title = "Summary"
        ws_sum["A1"] = "Quote: Q6001"

        # Sheet 2: 2D Grid Matrix tab
        ws_grid = wb.create_sheet("Molded Parts")
        ws_grid["A2"] = "Item #"
        ws_grid["B2"] = "Part #"
        ws_grid["C2"] = "Cycle Time"

        ws_grid["A3"] = "1"
        ws_grid["B3"] = "GRID-01"
        ws_grid["C3"] = 14.5

        ws_grid["A4"] = "2"
        ws_grid["B4"] = "GRID-02"
        ws_grid["C4"] = 28.0

        # Sheet 3: Form tab
        ws_form = wb.create_sheet("Assembly")
        ws_form["A2"] = "Quote #"
        ws_form["B2"] = "Q6001-3"
        ws_form["A4"] = "Part Number"
        ws_form["B4"] = "ASSY-01"
        ws_form["A6"] = "Cycle Time"
        ws_form["B6"] = 60.0

        file_path = tmp_path / "hybrid_multitab_grid.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 3
        part_nums = [it.part_number for it in doc.items]
        assert "GRID-01" in part_nums
        assert "GRID-02" in part_nums
        assert "ASSY-01" in part_nums

    def test_multi_tab_item_number_inference_hierarchies(self, tmp_path):
        """Verify item number inference hierarchy:
        Priority A: QID-X in cell
        Priority B: QID-X in tab title
        Priority C: 'Item X' in tab title
        Priority D: Sequence index fallback
        """
        wb = PureXmlWorkbook()

        # Tab A: Cell has explicit Q7001-1
        ws_a = wb.active
        ws_a.title = "RandomSheetName"
        ws_a["A1"] = "Quote #"
        ws_a["B1"] = "Q7001-1"
        ws_a["A3"] = "Part #"
        ws_a["B3"] = "PART-A"
        ws_a["A5"] = "Cycle Time"
        ws_a["B5"] = 15.0

        # Tab B: Tab title has Q7001-2, cell has base Q7001
        ws_b = wb.create_sheet("Q7001-2")
        ws_b["A1"] = "Quote #"
        ws_b["B1"] = "Q7001"
        ws_b["A3"] = "Part #"
        ws_b["B3"] = "PART-B"
        ws_b["A5"] = "Cycle Time"
        ws_b["B5"] = 25.0

        # Tab C: Tab title has 'Item 3', cell has base Q7001
        ws_c = wb.create_sheet("Item 3")
        ws_c["A1"] = "Quote #"
        ws_c["B1"] = "Q7001"
        ws_c["A3"] = "Part #"
        ws_c["B3"] = "PART-C"
        ws_c["A5"] = "Cycle Time"
        ws_c["B5"] = 35.0

        # Tab D: Tab title has no number, cell has base Q7001 -> sequence fallback
        ws_d = wb.create_sheet("Custom Housing")
        ws_d["A1"] = "Quote #"
        ws_d["B1"] = "Q7001"
        ws_d["A3"] = "Part #"
        ws_d["B3"] = "PART-D"
        ws_d["A5"] = "Cycle Time"
        ws_d["B5"] = 45.0

        file_path = tmp_path / "multitab_naming_hierarchy.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 4
        item_nums = [it.quote_item_number for it in doc.items]
        assert item_nums == ["Q7001-1", "Q7001-2", "Q7001-3", "Q7001-4"]
