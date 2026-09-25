"""
Targeted tests for fast streaming XML loader, lazy sheet evaluation,
and two-tier search down to row 150+.
"""

from pathlib import Path
import pytest

from tests.fixtures.generator import PureXmlWorkbook
from excel_extractor.engine import QuoteExtractorEngine
from excel_extractor.reader import load_workbook
from excel_extractor.strategies.grid_matrix import GridMatrixStrategy


class TestFastLoaderAndDeepSearch:
    """Verifies fast OpenXML loading and two-tier search enhancements."""

    def test_lazy_loading_of_ignored_sheets(self, tmp_path: Path):
        """
        Verifies that ignored sheets (e.g. 'Lookup', 'Terms', 'Notes')
        remain unparsed (_loaded=False) upon load_workbook until their cells are accessed.
        """
        wb_gen = PureXmlWorkbook()
        ws_mpc = wb_gen.active
        ws_mpc.title = "MPC"
        ws_mpc["A1"] = "Part Number"
        ws_mpc["B1"] = "Cycle Time"
        ws_mpc["A2"] = "PN-100"
        ws_mpc["B2"] = 25.0

        # Add ignored sheets
        ws_lookup = wb_gen.create_sheet("Lookup")
        ws_lookup["A1"] = "Lookup Data"
        ws_terms = wb_gen.create_sheet("Terms")
        ws_terms["A1"] = "Legal Terms"

        file_p = tmp_path / "test_lazy.xlsx"
        wb_gen.save(file_p)

        # Load workbook with reader
        wb = load_workbook(file_p)
        assert len(wb.worksheets) == 3

        ws_mpc_read = wb["MPC"]
        ws_lookup_read = wb["Lookup"]
        ws_terms_read = wb["Terms"]

        # MPC was eagerly loaded
        assert ws_mpc_read._loaded is True
        # Lookup and Terms were lazily loaded
        assert ws_lookup_read._loaded is False
        assert ws_terms_read._loaded is False

        # Accessing Lookup sheet triggers transparent on-demand loading
        assert ws_lookup_read["A1"].value == "Lookup Data"
        assert ws_lookup_read._loaded is True
        # Terms is still unparsed
        assert ws_terms_read._loaded is False

    def test_two_tier_search_finds_header_at_row_31(self, tmp_path: Path):
        """
        Verifies that Tier 1 light sweep down the sheet finds table headers
        at row 31 with a blank row, column index row (1, 2, 3...), and quote data.
        """
        wb_gen = PureXmlWorkbook()
        ws = wb_gen.active
        ws.title = "MPC"

        # Global header metadata at top
        ws["B2"] = "Quote Number:"
        ws["C2"] = "Q7000"

        # Row 31 has table headers
        ws["A31"] = "Quote Number"
        ws["B31"] = "Item#"
        ws["C31"] = "Part #"
        ws["D31"] = "Desc"
        ws["E31"] = "Resin Grade"
        ws["F31"] = "Cycle Sec"
        ws["G31"] = "Weight (g)"
        ws["H31"] = "Ann. Volume (EAU)"

        # Row 32 is blank
        # Row 33 has column index row
        ws["A33"] = 1
        ws["B33"] = 2
        ws["C33"] = 3
        ws["D33"] = 4
        ws["E33"] = 5
        ws["F33"] = 6
        ws["G33"] = 7
        ws["H33"] = 8

        # Row 34+ has quote data
        ws["A34"] = "Q7000-1"
        ws["B34"] = 1
        ws["C34"] = "858007801000"
        ws["D34"] = "Sensor Cover"
        ws["E34"] = "PA66-GF30"
        ws["F34"] = 24.5
        ws["G34"] = 12.0
        ws["H34"] = 50000

        ws["A35"] = "Q7000-2"
        ws["B35"] = 2
        ws["C35"] = "858007802000"
        ws["D35"] = "Sensor Housing"
        ws["E35"] = "PBT-GF20"
        ws["F35"] = 28.0
        ws["G35"] = 18.5
        ws["H35"] = 25000

        file_p = tmp_path / "test_deep_row31.xlsx"
        wb_gen.save(file_p)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_p)

        assert doc.quote_id == "Q7000"
        assert doc.total_items == 2
        assert doc.items[0].quote_item_number == "Q7000-1"
        assert doc.items[0].part_number == "858007801000"
        assert doc.items[0].cycle_time_sec == 24.5
        assert doc.items[1].quote_item_number == "Q7000-2"
        assert doc.items[1].part_number == "858007802000"
        assert doc.items[1].cycle_time_sec == 28.0

    def test_multi_band_resolution_picks_genuine_quote_table(self, tmp_path: Path):
        """
        Verifies that when multiple candidate bands exist (e.g. an upper checklist/metadata table
        and a lower full quote table at row 45), the multi-band evaluation picks the true quote table.
        """
        wb_gen = PureXmlWorkbook()
        ws = wb_gen.active
        ws.title = "MPC"

        # Candidate table at row 10: Revision history table
        ws["A10"] = "Item #"
        ws["B10"] = "Part Revision"
        ws["C10"] = "Description"
        ws["A11"] = 1
        ws["B11"] = "Rev A"
        ws["C11"] = "Initial release"

        # Real quote table at row 45
        ws["A45"] = "Quote Number"
        ws["B45"] = "Part Number"
        ws["C45"] = "Cycle Time (sec)"
        ws["D45"] = "Resin Grade"

        ws["A46"] = "Q9500-1"
        ws["B46"] = "994021A"
        ws["C46"] = 19.5
        ws["D46"] = "Lexan 141R"

        ws["A47"] = "Q9500-2"
        ws["B47"] = "994022B"
        ws["C47"] = 23.0
        ws["D47"] = "Valox 357"

        file_p = tmp_path / "test_multiband.xlsx"
        wb_gen.save(file_p)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_p)

        assert doc.quote_id == "Q9500"
        assert doc.total_items == 2
        assert doc.items[0].quote_item_number == "Q9500-1"
        assert doc.items[0].part_number == "994021A"
        assert doc.items[0].cycle_time_sec == 19.5
        assert doc.items[1].quote_item_number == "Q9500-2"
        assert doc.items[1].part_number == "994022B"
        assert doc.items[1].cycle_time_sec == 23.0

    def test_benchmark_33_item_large_quote_instantaneous(self, tmp_path: Path):
        """
        Benchmarks a 33-item quote with multiple ignored lookup tabs to verify
        near-instantaneous parsing and extraction time (< 250ms total).
        """
        import time

        wb_gen = PureXmlWorkbook()
        ws = wb_gen.active
        ws.title = "MPC"

        ws["A1"] = "Quote Number"
        ws["B1"] = "Part Number"
        ws["C1"] = "Desc"
        ws["D1"] = "Resin Grade"
        ws["E1"] = "Cycle Sec"
        ws["F1"] = "#Ops"
        ws["G1"] = "Weight (g)"
        ws["H1"] = "Ann. Volume (EAU)"

        for i in range(2, 35):
            ws[f"A{i}"] = f"Q4500-{i - 1}"
            ws[f"B{i}"] = f"PART-{1000 + i}"
            ws[f"C{i}"] = f"Component Description {i}"
            ws[f"D{i}"] = "Nylon 66"
            ws[f"E{i}"] = 18.0 + (i * 0.5)
            ws[f"F{i}"] = 0.5
            ws[f"G{i}"] = 15.0 + i
            ws[f"H{i}"] = 20000 + (i * 1000)

        # Add heavy ignored sheets
        for sheet_name in ["Lookup", "Master Data", "Terms", "ERP Dump"]:
            ign_ws = wb_gen.create_sheet(sheet_name)
            for r in range(1, 50):
                ign_ws[f"A{r}"] = f"Ignored Data {r}"

        file_p = tmp_path / "test_33_items.xlsx"
        wb_gen.save(file_p)

        t0 = time.perf_counter()
        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_p)
        elapsed_ms = (time.perf_counter() - t0) * 1000.0

        assert doc.quote_id == "Q4500"
        assert doc.total_items == 33
        assert len(doc.items) == 33
        # Strict latency threshold: under 250ms (typically ~50ms)
        assert elapsed_ms < 250.0, f"Extraction took {elapsed_ms:.2f}ms, expected < 250ms"
