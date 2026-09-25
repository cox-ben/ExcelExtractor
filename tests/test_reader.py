"""
Unit tests for the resilient OpenXML Excel reader and merged cell resolver.
"""

from __future__ import annotations

import io
from pathlib import Path
import pytest

from excel_extractor.reader import (
    Cell,
    MergedCellCollection,
    MergedCellRange,
    Workbook,
    Worksheet,
    col_index,
    col_letter,
    load_workbook,
    make_coord,
    parse_coord,
)
from tests.fixtures import generator


class TestCoordinateHelpers:
    """Tests for Excel column and coordinate conversion routines."""

    def test_col_letter_standard_and_multi_letter(self):
        assert col_letter(1) == "A"
        assert col_letter(2) == "B"
        assert col_letter(26) == "Z"
        assert col_letter(27) == "AA"
        assert col_letter(28) == "AB"
        assert col_letter(52) == "AZ"
        assert col_letter(53) == "BA"

    def test_col_letter_invalid_raises_error(self):
        with pytest.raises(ValueError):
            col_letter(0)
        with pytest.raises(ValueError):
            col_letter(-1)

    def test_col_index_standard_and_multi_letter(self):
        assert col_index("A") == 1
        assert col_index("b") == 2
        assert col_index("Z") == 26
        assert col_index("AA") == 27
        assert col_index("AB") == 28
        assert col_index("AZ") == 52
        assert col_index("BA") == 53

    def test_col_index_invalid_raises_error(self):
        with pytest.raises(ValueError):
            col_index("1")
        with pytest.raises(ValueError):
            col_index("A1")

    def test_parse_coord_and_make_coord_roundtrip(self):
        assert parse_coord("B18") == (18, 2)
        assert parse_coord("c2") == (2, 3)
        assert parse_coord("AA105") == (105, 27)

        assert make_coord(18, 2) == "B18"
        assert make_coord(2, 3) == "C2"
        assert make_coord(105, 27) == "AA105"

    def test_parse_coord_invalid_format(self):
        with pytest.raises(ValueError):
            parse_coord("INVALID")
        with pytest.raises(ValueError):
            parse_coord("123")


class TestMergedCells:
    """Tests for MergedCellRange and MergedCellCollection."""

    def test_merged_cell_range_properties(self):
        m = MergedCellRange("B2:D2")
        assert m.ref == "B2:D2"
        assert m.min_row == 2
        assert m.max_row == 2
        assert m.min_col == 2
        assert m.max_col == 4
        assert m.top_left_coord == "B2"
        assert m.bottom_right_coord == "D2"
        assert str(m) == "B2:D2"
        assert m == "B2:D2"

    def test_merged_cell_range_contains(self):
        m = MergedCellRange("B6:C10")
        assert m.contains(6, 2)  # B6
        assert m.contains(10, 3)  # C10
        assert m.contains(8, 2)
        assert not m.contains(5, 2)
        assert not m.contains(11, 2)
        assert not m.contains(6, 4)
        # String coord check
        assert "B6" in m
        assert "C10" in m
        assert "D6" not in m

    def test_merged_cell_collection(self):
        coll = MergedCellCollection()
        coll.append(MergedCellRange("B2:D2"))
        coll.append(MergedCellRange("B6:C6"))

        assert len(coll.ranges) == 2
        assert "B2:D2" in coll
        assert "B6:C6" in coll
        assert "A1:A2" not in coll

        found = coll.get_range_for_cell(2, 3)  # C2
        assert found is not None
        assert found.ref == "B2:D2"

        none_found = coll.get_range_for_cell(1, 1)
        assert none_found is None


class TestWorksheetAndWorkbook:
    """Tests for in-memory Worksheet and Workbook manipulation."""

    def test_worksheet_cell_access_and_assignment(self):
        ws = Worksheet("TestSheet")
        ws["B2"] = "Hello"
        ws["C2"] = 42
        ws["D2"] = 24.5
        ws["E2"] = True
        ws["F2"] = "#N/A"

        assert ws["B2"].value == "Hello"
        assert ws["B2"].data_type == "s"
        assert ws["C2"].value == 42
        assert ws["C2"].data_type == "n"
        assert ws["D2"].value == 24.5
        assert ws["D2"].data_type == "n"
        assert ws["E2"].value is True
        assert ws["E2"].data_type == "b"
        assert ws["F2"].value == "#N/A"
        assert ws["F2"].data_type == "e"

    def test_worksheet_max_dimensions(self):
        ws = Worksheet("Dimensions")
        ws["A1"] = 1
        ws["E10"] = 10
        assert ws.max_row == 10
        assert ws.max_column == 5

    def test_worksheet_iter_rows(self):
        ws = Worksheet("IterSheet")
        ws["A1"] = "A1"
        ws["B1"] = "B1"
        ws["A2"] = "A2"
        ws["B2"] = "B2"

        rows_cells = list(ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=2))
        assert len(rows_cells) == 2
        assert rows_cells[0][0].value == "A1"
        assert rows_cells[1][1].value == "B2"

        rows_vals = list(ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=2, values_only=True))
        assert rows_vals == [("A1", "B1"), ("A2", "B2")]

    def test_workbook_sheet_management(self):
        wb = Workbook()
        assert wb.sheetnames == ["Sheet1"]
        assert wb.active.title == "Sheet1"

        ws2 = wb.create_sheet("Sheet2")
        assert wb.sheetnames == ["Sheet1", "Sheet2"]
        assert wb["Sheet2"] is ws2

        with pytest.raises(KeyError):
            _ = wb["NonExistent"]


class TestOpenXmlFileLoading:
    """Tests for pure-Python OpenXML reader against synthetic workbooks."""

    def test_load_workbook_file_not_found(self, tmp_path: Path):
        missing = tmp_path / "missing.xlsx"
        with pytest.raises(FileNotFoundError):
            load_workbook(missing)

    def test_load_workbook_invalid_zip(self, tmp_path: Path):
        corrupt = tmp_path / "corrupt.xlsx"
        corrupt.write_bytes(b"NOT A ZIP")
        with pytest.raises(ValueError):
            load_workbook(corrupt)

    def test_load_workbook_from_bytes(self, single_modern_quote_file: Path):
        content = single_modern_quote_file.read_bytes()
        wb = load_workbook(content)
        assert len(wb.worksheets) == 1
        ws = wb.active
        assert ws["B2"].value == "Quote Number:"
        assert ws["C2"].value == "Q1001-1"
        assert ws["C10"].value == "2202046213000"
        assert ws["C18"].value == 24.5

    def test_load_workbook_from_stream(self, single_modern_quote_file: Path):
        with open(single_modern_quote_file, "rb") as f:
            wb = load_workbook(f)
            assert wb["Quote Summary"]["C2"].value == "Q1001-1"

    def test_load_workbook_merged_cells_read(self, merged_cells_quote_file: Path):
        wb = load_workbook(merged_cells_quote_file)
        ws = wb.active
        assert len(ws.merged_cells) >= 3
        assert "B2:D2" in ws.merged_cells
        assert "B6:C6" in ws.merged_cells
        assert "B14:C14" in ws.merged_cells
        # Value at top-left
        assert ws["B2"].value == "MPC MANUFACTURING QUOTE Q5000-1"
        assert ws["D6"].value == "3042605C"
        assert ws["E14"].value == 28.0

    def test_load_workbook_cached_formulas(self, legacy_raw_quote_file: Path):
        wb = load_workbook(legacy_raw_quote_file, data_only=True)
        ws = wb.active
        assert ws["D2"].value == "Q3050"
        assert ws["D8"].value == "994021A-REV2"
        # Formula cached value in D16 is 27.5
        assert ws["D16"].value == 27.5
        assert ws["D16"].formula == "=D14+D15"
