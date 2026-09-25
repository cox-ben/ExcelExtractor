"""
Targeted tests verifying fixes for the 5 issues specified in task.md:
1. Quote Item # extracted from Quote Number column (e.g. 'Q4378-4 (2140)' -> 'Q4378-4'), never Tool #.
2. Workbook with multiple sheets matches MPC the closest, uses QInfo for validation, and disregards other sheets.
3. Description column called 'Desc' is accurately captured, while 'Generic Resin Description' is strictly rejected.
4. Material captured from 'Resin Grade', strictly rejecting 'Resin Supplier' and 'Generic Resin Description'.
5. Static labour rate found on MPC sheet is correctly resolved and applied.
"""

from pathlib import Path
import pytest

from tests.fixtures.generator import PureXmlWorkbook
from excel_extractor.engine import QuoteExtractorEngine


class TestTaskIssues:
    """Validates the 5 specific domain issues from task.md."""

    def test_quote_item_from_quote_number_not_tool_number(self, tmp_path: Path):
        """
        Issue 1:
        Excel sheet layout matching user's screenshot:
        Col A: Tool # ((3082-3) 3161, 3083-1, 3083-2)
        Col B: Item# (3, 4, 5)
        Col C: Quote Number (Q4378-4 (2140), Q4378-5, Q4231-R15-7)
        Col D: Drawing # (M00894900, M00894901, M00894902)
        Col E: Part # (PART-A, PART-B, PART-C)
        Col F: Desc (Cover, Bezel, Housing)
        Col G: Resin Supplier (DuPont, BASF, Sabic)
        Col H: Resin Grade (Zytel 70G33L, Ultramid A3WG6, Lexan 141R)
        Col I: Cycle Sec (18.5, 22.0, 30.0)
        Col J: Labour (0.3, 0.3, 0.3)
        Col K: Weight (g) (15.2, 28.4, 45.0)
        Col L: Ann. Volume (EAU) (50000, 100000, 25000)
        """
        wb = PureXmlWorkbook()
        ws = wb.active
        ws.title = "MPC"

        # Global labour rate header & value
        ws["N2"] = "Labour"
        ws["N3"] = "Rate $CAD/Hr"
        ws["N4"] = 48.50

        # Grid headers at row 10
        ws["A10"] = "Tool #"
        ws["B10"] = "Item#"
        ws["C10"] = "Quote Number"
        ws["D10"] = "Drawing #"
        ws["E10"] = "Part #"
        ws["F10"] = "Desc"
        ws["G10"] = "Resin Supplier"
        ws["H10"] = "Resin Grade"
        ws["I10"] = "Cycle Sec"
        ws["J10"] = "Labour"
        ws["K10"] = "Weight (g)"
        ws["L10"] = "Ann. Volume (EAU)"

        # Data Row 1
        ws["A11"] = "(3082-3) 3161"
        ws["B11"] = "3"
        ws["C11"] = "Q4378-4 (2140)"
        ws["D11"] = "M00894900"
        ws["E11"] = "PART-A"
        ws["F11"] = "Upper Housing"
        ws["G11"] = "DuPont"
        ws["H11"] = "Zytel 70G33L"
        ws["I11"] = 18.5
        ws["J11"] = 0.3
        ws["K11"] = 15.2
        ws["L11"] = 50000

        # Data Row 2
        ws["A12"] = "3083-1"
        ws["B12"] = "4"
        ws["C12"] = "Q4378-5"
        ws["D12"] = "M00894901"
        ws["E12"] = "PART-B"
        ws["F12"] = "Lower Bezel"
        ws["G12"] = "BASF"
        ws["H12"] = "Ultramid A3WG6"
        ws["I12"] = 22.0
        ws["J12"] = 0.3
        ws["K12"] = 28.4
        ws["L12"] = 100000

        # Data Row 3
        ws["A13"] = "3083-2"
        ws["B13"] = "5"
        ws["C13"] = "Q4231-R15-7"
        ws["D13"] = "M00894902"
        ws["E13"] = "PART-C"
        ws["F13"] = "Bracket"
        ws["G13"] = "Sabic"
        ws["H13"] = "Lexan 141R"
        ws["I13"] = 30.0
        ws["J13"] = 0.3
        ws["K13"] = 45.0
        ws["L13"] = 25000

        # Row 14: Total stop
        ws["A14"] = "Total"

        file_path = tmp_path / "quote_issue_repro.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 3
        # Issue 1: Quote Item Numbers extracted from Quote Number column, NOT Tool #
        assert doc.items[0].quote_item_number == "Q4378-4"
        assert doc.items[1].quote_item_number == "Q4378-5"
        assert doc.items[2].quote_item_number == "Q4231-R15-7"

        # Issue 3: Description captured from 'Desc'
        assert doc.items[0].description == "Upper Housing"
        assert doc.items[1].description == "Lower Bezel"
        assert doc.items[2].description == "Bracket"

        # Issue 4: Material captured from 'Resin Grade', NOT 'Resin Supplier'
        assert doc.items[0].material == "Zytel 70G33L"
        assert doc.items[1].material == "Ultramid A3WG6"
        assert doc.items[2].material == "Lexan 141R"

        # Issue 5: Static Labour rate resolved
        assert doc.items[0].labour_rate == 48.50
        assert doc.items[1].labour_rate == 48.50
        assert doc.items[2].labour_rate == 48.50

        # Cycle time & part numbers
        assert [it.part_number for it in doc.items] == ["PART-A", "PART-B", "PART-C"]
        assert [it.cycle_time_sec for it in doc.items] == [18.5, 22.0, 30.0]

    def test_description_desc_vs_generic_resin_description(self, tmp_path: Path):
        """
        Issue 3 & 4:
        Header has 'Desc', 'Generic Resin Description', 'Resin Supplier', 'Resin Grade'.
        Must pick 'Desc' for Description, and 'Resin Grade' for Material.
        """
        wb = PureXmlWorkbook()
        ws = wb.active
        ws.title = "MPC"

        ws["A1"] = "Rate $CAD/Hr"
        ws["A2"] = 52.00

        # Headers at row 5
        ws["A5"] = "Quote Item #"
        ws["B5"] = "Part #"
        ws["C5"] = "Cycle Sec"
        ws["D5"] = "Desc"
        ws["E5"] = "Generic Resin Description"
        ws["F5"] = "Resin Supplier"
        ws["G5"] = "Resin Grade"

        ws["A6"] = "Q7000-1"
        ws["B6"] = "PN-7001"
        ws["C6"] = 15.0
        ws["D6"] = "Handle Assembly"
        ws["E6"] = "Polyamide 66 30% Glass Filled"
        ws["F6"] = "DuPont Canada"
        ws["G6"] = "Zytel 101L"

        ws["A7"] = "Total"

        file_path = tmp_path / "desc_resin_test.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 1
        it = doc.items[0]
        assert it.description == "Handle Assembly"
        assert it.material == "Zytel 101L"
        assert it.labour_rate == 52.00

    def test_closest_mpc_sheet_selection_and_qinfo_validation(self, tmp_path: Path):
        """
        Issue 2:
        Workbook has:
        - 'MPC' (exact match -> primary)
        - 'MPC Old' (should be disregarded)
        - 'MPC (Copy)' (should be disregarded)
        - 'QInfo' (closest QInfo -> used for validation)
        - 'ToolingCosts' (auxiliary -> disregarded)
        """
        wb = PureXmlWorkbook()

        # Primary sheet: MPC
        ws_mpc = wb.active
        ws_mpc.title = "MPC"
        ws_mpc["H1"] = "Rate $CAD/Hr"
        ws_mpc["H2"] = 55.00

        ws_mpc["A5"] = "Quote Number"
        ws_mpc["B5"] = "Part #"
        ws_mpc["C5"] = "Cycle Sec"
        ws_mpc["D5"] = "Desc"
        ws_mpc["E5"] = "Resin Grade"

        ws_mpc["A6"] = "Q8000-1"
        ws_mpc["B6"] = "PART-ACTIVE-1"
        ws_mpc["C6"] = 20.0
        ws_mpc["D6"] = "Lens"
        ws_mpc["E6"] = "PMMA 8N"

        ws_mpc["A7"] = "Q8000-2"
        ws_mpc["B7"] = "PART-ACTIVE-2"
        ws_mpc["C7"] = 25.0
        ws_mpc["D7"] = "Housing"
        ws_mpc["E6"] = "ABS HI-121"

        ws_mpc["A8"] = "Total"

        # Disregarded sheet: MPC Old
        ws_old = wb.create_sheet(title="MPC Old")
        ws_old["A1"] = "Quote Number"
        ws_old["B1"] = "Part #"
        ws_old["C1"] = "Cycle Sec"
        ws_old["A2"] = "Q8000-99"
        ws_old["B2"] = "PART-OLD-WRONG"
        ws_old["C2"] = 99.0

        # Disregarded sheet: MPC (Copy)
        ws_copy = wb.create_sheet(title="MPC (Copy)")
        ws_copy["A1"] = "Quote Number"
        ws_copy["B1"] = "Part #"
        ws_copy["C1"] = "Cycle Sec"
        ws_copy["A2"] = "Q8000-88"
        ws_copy["B2"] = "PART-COPY-WRONG"
        ws_copy["C2"] = 88.0

        # Closest QInfo sheet: used for validation
        ws_qinfo = wb.create_sheet(title="QInfo")
        ws_qinfo["A5"] = "Quote Number"
        ws_qinfo["B5"] = "Part #"
        ws_qinfo["C5"] = "Cycle Sec"
        ws_qinfo["A6"] = "Q8000-1"
        ws_qinfo["B6"] = "PART-ACTIVE-1"
        ws_qinfo["C6"] = 20.0
        ws_qinfo["A7"] = "Q8000-2"
        ws_qinfo["B7"] = "PART-ACTIVE-2"
        ws_qinfo["C7"] = 25.0

        # Auxiliary sheet: ToolingCosts
        ws_aux = wb.create_sheet(title="ToolingCosts")
        ws_aux["A1"] = "Total Tool Cost"
        ws_aux["B1"] = 120000

        file_path = tmp_path / "closest_mpc_test.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        # Strictly 2 items, none from MPC Old or MPC (Copy) or ToolingCosts
        assert doc.total_items == 2
        assert [it.part_number for it in doc.items] == ["PART-ACTIVE-1", "PART-ACTIVE-2"]
        assert [it.quote_item_number for it in doc.items] == ["Q8000-1", "Q8000-2"]
        assert doc.items[0].labour_rate == 55.00
        assert doc.items[1].labour_rate == 55.00

    def test_qinfo_with_gap_rows_and_mpc_labour_rate(self, tmp_path: Path):
        """
        Exact user screenshot scenario:
        - Sheet 'QInfo':
          Row 31: Line Item | MPC RFQ # | Part No. | Description | Annual Volume
          Row 32: [blank spacer row]
          Row 33: [empty]   | 2         | 3        | 4           | 5  [column index row]
          Row 34: 1         | Q4476-R1-1| 40601-BBA001 | Bobbin  | 30,000
          Row 35: 2         | Q4476-R1-2| 40601-COL0112_TRM01 | Short Bobbin Terminal - TRM01 | 30,000
          Row 36: 3         | Q4476-R1-3| 40601-COL0112_TRM02 | Long Bobbin Terminal - TRM01  | 30,000
        - Sheet 'MPC':
          Row 1: Labour
          Row 2: Rate $CAD/Hr
          Row 3: 48.50
        
        Expected:
        - Accurately extracts all 3 items from QInfo despite gap at rows 32-33.
        - Accurately grabs Labour Rate (48.50) from MPC sheet.
        - Does NOT get fooled by column index row 33.
        """
        wb = PureXmlWorkbook()

        # Sheet 1: QInfo
        ws_qinfo = wb.active
        ws_qinfo.title = "QInfo"

        # Row 31: Table Header
        ws_qinfo["A31"] = "Line Item"
        ws_qinfo["B31"] = "MPC RFQ #"
        ws_qinfo["C31"] = "Part No."
        ws_qinfo["D31"] = "Description"
        ws_qinfo["E31"] = "Annual Volume"

        # Row 32: Blank row

        # Row 33: Column index numbers
        ws_qinfo["B33"] = 2
        ws_qinfo["C33"] = 3
        ws_qinfo["D33"] = 4
        ws_qinfo["E33"] = 5

        # Row 34: Item 1
        ws_qinfo["A34"] = "1"
        ws_qinfo["B34"] = "Q4476-R1-1"
        ws_qinfo["C34"] = "40601-BBA001"
        ws_qinfo["D34"] = "Bobbin"
        ws_qinfo["E34"] = 30000

        # Row 35: Item 2
        ws_qinfo["A35"] = "2"
        ws_qinfo["B35"] = "Q4476-R1-2"
        ws_qinfo["C35"] = "40601-COL0112_TRM01"
        ws_qinfo["D35"] = "Short Bobbin Terminal - TRM01"
        ws_qinfo["E35"] = 30000

        # Row 36: Item 3
        ws_qinfo["A36"] = "3"
        ws_qinfo["B36"] = "Q4476-R1-3"
        ws_qinfo["C36"] = "40601-COL0112_TRM02"
        ws_qinfo["D36"] = "Long Bobbin Terminal - TRM01"
        ws_qinfo["E36"] = 30000

        # Sheet 2: MPC (Labour Rate)
        ws_mpc = wb.create_sheet(title="MPC")
        ws_mpc["B2"] = "Labour"
        ws_mpc["B3"] = "Rate $CAD/Hr"
        ws_mpc["B4"] = 48.50

        file_path = tmp_path / "qinfo_gap_rows_mpc_labour.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 3
        assert doc.quote_id == "Q4476"
        assert [it.quote_item_number for it in doc.items] == ["Q4476-R1-1", "Q4476-R1-2", "Q4476-R1-3"]
        assert [it.part_number for it in doc.items] == [
            "40601-BBA001",
            "40601-COL0112_TRM01",
            "40601-COL0112_TRM02",
        ]
        assert [it.description for it in doc.items] == [
            "Bobbin",
            "Short Bobbin Terminal - TRM01",
            "Long Bobbin Terminal - TRM01",
        ]
        assert [it.annual_volume for it in doc.items] == [30000, 30000, 30000]
        # Labour rate resolved from MPC sheet for all items
        assert all(it.labour_rate == 48.50 for it in doc.items)

    def test_separate_quote_number_and_item_number_resolution(self, tmp_path: Path):
        """
        Verify resolution when Quote Number and Item# are in separate columns:
        Col A: Tool # (1, 2)
        Col B: Item# (1, 2)
        Col C: Quote Number (Q3428, Q3428)
        Col D: Drawing # (1014437, 403058)
        Col E: Part Number (1014437, 403058)
        """
        from excel_extractor.strategies.grid_matrix import (
            quote_number_has_item_suffix,
            resolve_quote_item_number,
        )

        # 1. Direct helper unit tests
        assert not quote_number_has_item_suffix("Q3428")
        assert not quote_number_has_item_suffix("3428")
        assert quote_number_has_item_suffix("Q3428-1")
        assert quote_number_has_item_suffix("Q3428-01")
        assert quote_number_has_item_suffix("Q4378-4 (2140)")
        assert quote_number_has_item_suffix("Q4231-R15-7")

        # When quote number is bare Q3428 and Item# is 1 -> Q3428-1
        item_num, _ = resolve_quote_item_number(raw_quote_val="Q3428", raw_item_val=1, quote_id="Q3428", item_seq=1)
        assert item_num == "Q3428-1"

        # When quote number already contains item suffix (Q3428-1) and Item# is 1 -> Q3428-1 (not Q3428-1-1!)
        item_num, _ = resolve_quote_item_number(raw_quote_val="Q3428-1", raw_item_val=1, quote_id="Q3428", item_seq=1)
        assert item_num == "Q3428-1"

        # 2. End-to-end workbook extraction matching user's screenshot
        wb = PureXmlWorkbook()
        ws = wb.active
        ws.title = "MPC"

        # Headers at row 1
        ws["A1"] = "Tool #"
        ws["B1"] = "Item#"
        ws["C1"] = "Quote Number"
        ws["D1"] = "Drawing #"
        ws["E1"] = "Part Number"
        ws["F1"] = "Cycle Sec"

        # Row 2 (Item 1)
        ws["A2"] = 1
        ws["B2"] = 1
        ws["C2"] = "Q3428"
        ws["D2"] = "1014437"
        ws["E2"] = "1014437"
        ws["F2"] = 25.0

        # Row 3 (Item 2)
        ws["A3"] = 2
        ws["B3"] = 2
        ws["C3"] = "Q3428"
        ws["D3"] = "403058"
        ws["E3"] = "403058"
        ws["F3"] = 32.5

        file_path = tmp_path / "separate_quote_and_item_cols.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 2
        assert doc.quote_id == "Q3428"
        assert [it.quote_item_number for it in doc.items] == ["Q3428-1", "Q3428-2"]
        assert [it.part_number for it in doc.items] == ["1014437", "403058"]
        assert [it.cycle_time_sec for it in doc.items] == [25.0, 32.5]

    def test_quote_with_revision_suffix_and_separate_item(self, tmp_path: Path):
        """
        Verify that revision suffixes like -R3 or -Rev3 are NOT mistaken for item suffixes,
        and combine correctly with separate Item# columns.
        """
        from excel_extractor.strategies.grid_matrix import (
            quote_number_has_item_suffix,
            resolve_quote_item_number,
        )

        assert not quote_number_has_item_suffix("Q3428-R3")
        assert not quote_number_has_item_suffix("Q3428-Rev3")

        # When quote has revision Q3428-R3 and item is 1 -> Q3428-1
        item_num, _ = resolve_quote_item_number(raw_quote_val="Q3428-R3", raw_item_val=1, quote_id="Q3428", item_seq=1)
        assert item_num == "Q3428-1"

        item_num2, _ = resolve_quote_item_number(raw_quote_val="Q3428-R3", raw_item_val=2, quote_id="Q3428", item_seq=2)
        assert item_num2 == "Q3428-2"

    def test_quote_without_item_column_infers_sequential(self, tmp_path: Path):
        """
        Verify that when a sheet only has 'Quote Number' with bare 'Q3428' and NO 'Item#' column,
        items receive sequential suffixes (Q3428-1, Q3428-2) instead of duplicate bare 'Q3428'.
        """
        wb = PureXmlWorkbook()
        ws = wb.active
        ws.title = "MPC"

        ws["A1"] = "Quote Number"
        ws["B1"] = "Part Number"
        ws["C1"] = "Cycle Sec"

        ws["A2"] = "Q3428"
        ws["B2"] = "1014437"
        ws["C2"] = 12.6

        ws["A3"] = "Q3428"
        ws["B3"] = "403058"
        ws["C3"] = 25.0

        file_path = tmp_path / "no_item_col_quote.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 2
        assert [it.quote_item_number for it in doc.items] == ["Q3428-1", "Q3428-2"]
        assert [it.part_number for it in doc.items] == ["1014437", "403058"]

    def test_quote_qinfo_cross_reference_item_mapping(self, tmp_path: Path):
        """
        Verify that when MPC sheet only has bare 'Q3428', but QInfo sheet defines
        Item# 1 and 2 for parts 1014437 and 403058, the extractor cross-references
        QInfo to assign Q3428-1 and Q3428-2.
        """
        wb = PureXmlWorkbook()
        ws_qinfo = wb.active
        ws_qinfo.title = "QInfo"

        ws_qinfo["A1"] = "Tool #"
        ws_qinfo["B1"] = "Item#"
        ws_qinfo["C1"] = "Quote Number"
        ws_qinfo["D1"] = "Part Number"

        ws_qinfo["A2"] = 1
        ws_qinfo["B2"] = 1
        ws_qinfo["C2"] = "Q3428"
        ws_qinfo["D2"] = "1014437"

        ws_qinfo["A3"] = 2
        ws_qinfo["B3"] = 2
        ws_qinfo["C3"] = "Q3428"
        ws_qinfo["D3"] = "403058"

        ws_mpc = wb.create_sheet("MPC")
        ws_mpc["A1"] = "Quote Number"
        ws_mpc["B1"] = "Part Number"
        ws_mpc["C1"] = "Description"
        ws_mpc["D1"] = "Cycle Sec"

        ws_mpc["A2"] = "Q3428"
        ws_mpc["B2"] = "1014437"
        ws_mpc["C2"] = "Spring Support"
        ws_mpc["D2"] = 12.6

        ws_mpc["A3"] = "Q3428"
        ws_mpc["B3"] = "403058"
        ws_mpc["C3"] = "Spring Support"
        ws_mpc["D3"] = 25.0

        file_path = tmp_path / "qinfo_cross_ref.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 2
        assert [it.quote_item_number for it in doc.items] == ["Q3428-1", "Q3428-2"]
        assert [it.part_number for it in doc.items] == ["1014437", "403058"]
        assert [it.cycle_time_sec for it in doc.items] == [12.6, 25.0]



