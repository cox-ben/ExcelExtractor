"""
Tests for hidden sheet filtering and MPC / QInfo target sheet recognition.

Validates:
1. Hidden sheets (state='hidden' or 'veryHidden') are strictly ignored and do not generate ghost items.
2. Target sheet filtering accurately prioritizes tabs named 'MPC', 'mpc', 'MPC 1', 'mpc-2',
   'QInfo', 'Quote Info', 'Quote_Info', 'qinfo', etc.
3. Non-quoting auxiliary tabs (e.g. 'MachineRates', 'ToolingCosts') are filtered out when MPC/QInfo tabs exist.
4. Item numbering correctly maps for 'MPC 1' / 'mpc-2' / 'QInfo 1' / 'Quote Info'.
"""

from pathlib import Path
import pytest

from tests.fixtures.generator import PureXmlWorkbook
from excel_extractor.engine import QuoteExtractorEngine
from excel_extractor.reader import load_workbook


class TestHiddenSheetsFiltering:
    """Test suite verifying hidden sheets are strictly ignored."""

    def test_hidden_sheets_openxml_reader_state(self, tmp_path: Path):
        """Verify OpenXML reader parses sheet_state and is_visible/is_hidden accurately."""
        wb = PureXmlWorkbook()
        ws_vis = wb.active
        ws_vis.title = "VisibleQuote"

        ws_hid = wb.create_sheet(title="OldHiddenQuote", state="hidden")
        ws_very_hid = wb.create_sheet(title="VeryHiddenBackup", state="veryHidden")

        file_path = tmp_path / "test_hidden_reader.xlsx"
        wb.save(file_path)

        loaded_wb = load_workbook(file_path)
        assert len(loaded_wb.worksheets) == 3

        s_vis = loaded_wb.get_sheet_by_name("VisibleQuote")
        s_hid = loaded_wb.get_sheet_by_name("OldHiddenQuote")
        s_very_hid = loaded_wb.get_sheet_by_name("VeryHiddenBackup")

        assert s_vis.is_visible is True
        assert s_vis.is_hidden is False
        assert s_vis.sheet_state == "visible"

        assert s_hid.is_visible is False
        assert s_hid.is_hidden is True
        assert s_hid.sheet_state == "hidden"

        assert s_very_hid.is_visible is False
        assert s_very_hid.is_hidden is True
        assert s_very_hid.sheet_state == "veryhidden"

    def test_hidden_sheet_with_old_quote_is_ignored(self, tmp_path: Path):
        """Verify that an old hidden tab with quote numbers does NOT generate ghost items."""
        wb = PureXmlWorkbook()

        # Visible active quote tab
        ws_vis = wb.active
        ws_vis.title = "MPC"
        ws_vis["A1"] = "Quote #"
        ws_vis["B1"] = "Q1001-1"
        ws_vis["A3"] = "Part #"
        ws_vis["B3"] = "PART-ACTIVE"
        ws_vis["A5"] = "Cycle Time"
        ws_vis["B5"] = 24.5

        # Hidden old quote revision tab
        ws_hid = wb.create_sheet(title="OldRevision_MPC", state="hidden")
        ws_hid["A1"] = "Quote #"
        ws_hid["B1"] = "Q9999-1"
        ws_hid["A3"] = "Part #"
        ws_hid["B3"] = "PART-OLD-GHOST"
        ws_hid["A5"] = "Cycle Time"
        ws_hid["B5"] = 99.0

        file_path = tmp_path / "quote_with_hidden_rev.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 1
        assert doc.items[0].quote_item_number == "Q1001-1"
        assert doc.items[0].part_number == "PART-ACTIVE"
        assert doc.items[0].cycle_time_sec == 24.5
        assert doc.quote_id == "Q1001"


class TestMPCTargetSheetRecognition:
    """Test suite verifying 'MPC' and 'mpc' target sheet recognition."""

    @pytest.mark.parametrize("mpc_title", ["MPC", "mpc", "Mpc", "MPC 1", "mpc-1", "MPC_1", "mpc (1)"])
    def test_single_mpc_tab_with_auxiliary_tabs(self, tmp_path: Path, mpc_title: str):
        """Verify that an MPC tab is isolated and auxiliary background tabs are ignored."""
        wb = PureXmlWorkbook()

        # Auxiliary tab 1
        ws_cover = wb.active
        ws_cover.title = "Summary"
        ws_cover["A1"] = "Customer"
        ws_cover["B1"] = "General Motors"

        # The actual MPC quoting tab
        ws_mpc = wb.create_sheet(title=mpc_title)
        ws_mpc["A1"] = "Quote #"
        ws_mpc["B1"] = "Q2050-1"
        ws_mpc["A3"] = "Part #"
        ws_mpc["B3"] = "2202046213000"
        ws_mpc["A5"] = "Cycle Time"
        ws_mpc["B5"] = 18.5

        # Auxiliary non-quote calculations tab
        ws_rates = wb.create_sheet(title="MachineRates")
        ws_rates["A1"] = "Press 150T Rate"
        ws_rates["B1"] = 65.00

        # Hidden scratchpad tab
        ws_scratch = wb.create_sheet(title="Scratchpad", state="hidden")
        ws_scratch["A1"] = "Quote #"
        ws_scratch["B1"] = "Q8888-1"
        ws_scratch["A3"] = "Part #"
        ws_scratch["B3"] = "GHOST"

        file_path = tmp_path / f"test_{mpc_title.replace(' ', '_')}.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 1
        assert doc.quote_id == "Q2050"
        assert doc.items[0].quote_item_number == "Q2050-1"
        assert doc.items[0].part_number == "2202046213000"
        assert doc.items[0].cycle_time_sec == 18.5

    def test_multi_tab_mpc_sheets(self, tmp_path: Path):
        """Verify multiple MPC tabs ('MPC 1', 'mpc 2') extract all items accurately."""
        wb = PureXmlWorkbook()

        # Summary sheet
        ws_sum = wb.active
        ws_sum.title = "Summary Page"
        ws_sum["A1"] = "Project Ref"
        ws_sum["B1"] = "Q3000"

        # MPC Tab 1
        ws1 = wb.create_sheet(title="MPC 1")
        ws1["A1"] = "Quote #"
        ws1["B1"] = "Q3000"
        ws1["A3"] = "Part #"
        ws1["B3"] = "HOUSING-A"
        ws1["A5"] = "Cycle Time"
        ws1["B5"] = 22.0

        # MPC Tab 2
        ws2 = wb.create_sheet(title="mpc 2")
        ws2["A1"] = "Quote #"
        ws2["B1"] = "Q3000"
        ws2["A3"] = "Part #"
        ws2["B3"] = "BEZEL-B"
        ws2["A5"] = "Cycle Time"
        ws2["B5"] = 16.5

        # Extra auxiliary tab that should NOT be extracted
        ws_calc = wb.create_sheet(title="ToolingCalculations")
        ws_calc["A1"] = "Tool Cost"
        ws_calc["B1"] = 55000

        file_path = tmp_path / "multi_mpc_quote.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 2
        assert doc.quote_id == "Q3000"
        assert [it.quote_item_number for it in doc.items] == ["Q3000-1", "Q3000-2"]
        assert [it.part_number for it in doc.items] == ["HOUSING-A", "BEZEL-B"]
        assert [it.cycle_time_sec for it in doc.items] == [22.0, 16.5]


class TestQInfoTargetSheetRecognition:
    """Test suite verifying 'QInfo' and 'Quote Info' target sheet recognition."""

    @pytest.mark.parametrize("qinfo_title", ["QInfo", "qinfo", "Quote Info", "QuoteInfo", "Q-Info", "Quote_Info", "quote info"])
    def test_qinfo_tab_isolation(self, tmp_path: Path, qinfo_title: str):
        """Verify that QInfo / Quote Info tabs are cleanly extracted while ignoring other tabs."""
        wb = PureXmlWorkbook()

        ws_cov = wb.active
        ws_cov.title = "Cover"
        ws_cov["A1"] = "MPC Quoting Portal"

        ws_qinfo = wb.create_sheet(title=qinfo_title)
        ws_qinfo["A1"] = "Quote #"
        ws_qinfo["B1"] = "Q4050-1"
        ws_qinfo["A3"] = "Part #"
        ws_qinfo["B3"] = "BRACKET-99"
        ws_qinfo["A5"] = "Cycle Time"
        ws_qinfo["B5"] = 31.0

        ws_lookup = wb.create_sheet(title="MaterialData")
        ws_lookup["A1"] = "Resin"
        ws_lookup["B1"] = "PA66-GF30"

        file_path = tmp_path / f"test_{qinfo_title.replace(' ', '_')}.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 1
        assert doc.quote_id == "Q4050"
        assert doc.items[0].quote_item_number == "Q4050-1"
        assert doc.items[0].part_number == "BRACKET-99"
        assert doc.items[0].cycle_time_sec == 31.0

    def test_qinfo_with_grid_matrix(self, tmp_path: Path):
        """Verify that a 2D grid matrix table on a 'Quote Info' tab is accurately parsed."""
        wb = PureXmlWorkbook()

        ws_cover = wb.active
        ws_cover.title = "Terms & Conditions"
        ws_cover["A1"] = "Standard MPC Terms"

        ws_qinfo = wb.create_sheet(title="Quote Info")
        ws_qinfo["B2"] = "Quote #:"
        ws_qinfo["C2"] = "Q5500"

        # Grid Table Headers
        ws_qinfo["A6"] = "Item #"
        ws_qinfo["B6"] = "Part #"
        ws_qinfo["C6"] = "Cycle Time (s)"

        # Rows
        ws_qinfo["A7"] = "Q5500-1"
        ws_qinfo["B7"] = "PART-101"
        ws_qinfo["C7"] = 14.5

        ws_qinfo["A8"] = "Q5500-2"
        ws_qinfo["B8"] = "PART-102"
        ws_qinfo["C8"] = 28.0

        # Footer
        ws_qinfo["A9"] = "Total"
        ws_qinfo["B9"] = "2 Items"

        file_path = tmp_path / "quote_info_grid.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 2
        assert doc.quote_id == "Q5500"
        assert [it.quote_item_number for it in doc.items] == ["Q5500-1", "Q5500-2"]
        assert [it.part_number for it in doc.items] == ["PART-101", "PART-102"]
        assert [it.cycle_time_sec for it in doc.items] == [14.5, 28.0]

    def test_qinfo_table_with_hidden_rows_and_zero_values_disregarded(self, tmp_path: Path):
        """
        Exact user scenario:
        - QInfo sheet has 6 real quote items.
        - Rows 7-18 have formula quote IDs (Q4476-R1-7 to Q4476-R1-18) but part numbers are 0/blank,
          cycle times are 0.0/blank, and some rows are hidden.
        - A 'Quote Revision' table appears below item data.
        - Other auxiliary sheets exist in the workbook (Job Values vs Quote, Huff US, Shipping, Multimatic, Litens).
        
        Expected:
        - Disregards all unfilled rows (items 7-18).
        - Disregards all hidden rows.
        - Extracts strictly 6 items from QInfo.
        - Does NOT scrape other auxiliary sheets.
        """
        wb = PureXmlWorkbook()

        # Sheet 1: Master QInfo sheet
        ws_qinfo = wb.active
        ws_qinfo.title = "QInfo"

        # Table Header at row 33
        ws_qinfo["A33"] = "Line Item"
        ws_qinfo["B33"] = "MPC RFQ #"
        ws_qinfo["C33"] = "Part No."
        ws_qinfo["D33"] = "Description"
        ws_qinfo["S33"] = "Cycle Time (sec)"

        # 6 Real Active Quote Items
        active_items = [
            (34, "1", "Q4476-R1-1", "40601-BBA001", 30.0),
            (35, "2", "Q4476-R1-2", "40601-COL0112_TRM01", 5.0),
            (36, "3", "Q4476-R1-3", "40601-COL0112_TRM02", 5.0),
            (37, "4", "Q4476-R1-4", "40601-BBA001", 30.0),
            (38, "5", "Q4476-R1-5", "40601-COL0112_TRM01", 5.0),
            (39, "6", "Q4476-R1-6", "40601-COL0112_TRM02", 5.0),
        ]
        for row_idx, line_it, rfq_num, pn, ct in active_items:
            ws_qinfo[f"A{row_idx}"] = line_it
            ws_qinfo[f"B{row_idx}"] = rfq_num
            ws_qinfo[f"C{row_idx}"] = pn
            ws_qinfo[f"S{row_idx}"] = ct

        # Unfilled template rows 40 to 50:
        # Some with part number = 0, cycle time = 0.00
        # Some completely blank
        # Rows 43 to 50 marked as hidden
        for row_idx in range(40, 51):
            line_idx = row_idx - 33
            ws_qinfo[f"A{row_idx}"] = str(line_idx)
            ws_qinfo[f"B{row_idx}"] = f"Q4476-R1-{line_idx}"
            if row_idx % 2 == 0:
                ws_qinfo[f"C{row_idx}"] = 0  # evaluates to 0 in Excel formula
                ws_qinfo[f"S{row_idx}"] = 0.0
            if row_idx >= 43:
                ws_qinfo.hidden_rows.add(row_idx)

        # Row 70: Quote Revision footer table
        ws_qinfo["B70"] = "Quote Revision"
        ws_qinfo["B71"] = "R1"
        ws_qinfo["C71"] = "Initial quote release"

        # Other auxiliary sheets that must NOT be extracted as quote items
        ws_jv = wb.create_sheet(title="Job Values vs Quote")
        ws_jv["A21"] = "Q4476-19"
        ws_jv["E21"] = 0

        ws_huff = wb.create_sheet(title="Huff US")
        ws_huff["B2"] = "Q4476"
        ws_huff["C11"] = "40601-BBA001"

        ws_ship = wb.create_sheet(title="Shipping")
        ws_ship["B2"] = "Q4476"
        ws_ship["C2"] = "Volume/Year"

        ws_multi = wb.create_sheet(title="Multimatic QAF 1-2020")
        ws_multi["G4"] = "Q4476"
        ws_multi["C4"] = "40601-BBA001"
        ws_multi["K4"] = 30.0

        ws_litens = wb.create_sheet(title="Litens US (OLD)", state="hidden")
        ws_litens["K1"] = "Q4476"
        ws_litens["J8"] = "40601-BBA001"

        file_path = tmp_path / "q4476_complete_repro.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        # Strictly 6 items from QInfo, none from auxiliary sheets, none from rows 7-18
        assert doc.total_items == 6
        assert doc.quote_id == "Q4476"
        assert [it.part_number for it in doc.items] == [
            "40601-BBA001",
            "40601-COL0112_TRM01",
            "40601-COL0112_TRM02",
            "40601-BBA001",
            "40601-COL0112_TRM01",
            "40601-COL0112_TRM02",
        ]
        assert [it.cycle_time_sec for it in doc.items] == [30.0, 5.0, 5.0, 30.0, 5.0, 5.0]
        # Coordinates must all originate from QInfo
        for it in doc.items:
            assert "QInfo!" in it.source_cell_coords.part_num
