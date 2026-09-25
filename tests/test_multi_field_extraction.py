"""
Tests for comprehensive multi-field manufacturing variable extraction.

Validates:
1. Full 9-field manufacturing extraction from MPC quote sheets (MPC, QInfo):
   - Quote Item (from MPC RFQ #, preserving revision strings like Q4476-R1-1)
   - Part No.
   - Description / Discription
   - Material
   - Cycle Sec
   - #Ops / Labour (e.g. 0.3)
   - Static Labour Rate ($CAD/Hr) located directly below Rate $CAD/Hr
   - Weight (g)
   - Ann. Volume (EAU)
2. Static Labour Rate resolution:
   - Positioned above the table below "Rate $CAD/Hr" and applied statically to all quote items.
3. Validation and Confidence:
   - #Ops strictly between 0.0 and 1.0 (flagged if > 1.0).
   - Missing required fields downgrade confidence to 'low' and add informative warnings.
   - Out-of-bounds metrics flag warnings and downgrade confidence.
4. Export correctness:
   - CSV export columns match the canonical schema and contain all variables.
   - DataFrame and JSON exports preserve all values and coordinates.
"""

import json
from pathlib import Path
import pytest

from excel_extractor.engine import QuoteExtractorEngine
from excel_extractor.exporters import export_to_csv, export_to_json
from excel_extractor.models import QuoteItem, SourceCellCoords
from tests.fixtures.generator import PureXmlWorkbook


class TestMultiFieldExtraction:
    """Test suite for full 9-field MPC manufacturing variable extraction."""

    def test_mpc_grid_matrix_all_9_fields(self, tmp_path: Path):
        """
        Simulates an authentic MPC quote workbook (e.g. Q4476-R1):
        - Sheet named 'MPC'
        - Static Labour Rate block with 'Rate $CAD/Hr' and '$32.50' directly below it
        - Grid matrix starting at row 32 with headers:
          MPC RFQ #, Part No., Discription, Material, Cycle Sec, Labour, Weight (g), Ann. Volume (EAU)
        - 3 filled item rows with revision strings (Q4476-R1-1, Q4476-R1-2, Q4476-R1-3)
        - Disregards formula placeholder zeros or empty rows below items
        """
        wb = PureXmlWorkbook()
        ws = wb.active
        ws.title = "MPC"

        # Global metadata block
        ws["B2"] = "MOLDED PRECISION COMPONENTS QUOTE"
        ws["B3"] = "MPC RFQ #"
        ws["C3"] = "Q4476-R1"

        # Static Labour Rate block (Rate $CAD/Hr with rate in cell below)
        ws["E10"] = "Rate $CAD/Hr"
        ws["E11"] = 32.50

        # Table header row at row 32
        header_row = 32
        headers = [
            ("A", "MPC RFQ #"),
            ("B", "Part No."),
            ("C", "Discription"),
            ("D", "Material"),
            ("E", "Cycle Sec"),
            ("F", "Labour"),
            ("G", "Weight (g)"),
            ("H", "Ann. Volume (EAU)"),
        ]
        for col_letter, title in headers:
            ws[f"{col_letter}{header_row}"] = title

        # Item 1: nominal complete row
        ws["A33"] = "Q4476-R1-1"
        ws["B33"] = "100-2490-01"
        ws["C33"] = "Housing Bezel Lower"
        ws["D33"] = "Polypropylene 30% GF"
        ws["E33"] = 28.5
        ws["F33"] = 0.3
        ws["G33"] = 45.2
        ws["H33"] = 150000

        # Item 2: nominal complete row with integer part number
        ws["A34"] = "Q4476-R1-2"
        ws["B34"] = 884012
        ws["C34"] = "Actuator Pin"
        ws["D34"] = "POM Delrin 500P"
        ws["E34"] = 16.0
        ws["F34"] = 0.25
        ws["G34"] = 3.8
        ws["H34"] = 500000

        # Item 3: nominal complete row with string units
        ws["A35"] = "Q4476-R1-3"
        ws["B35"] = "CAP-TOP-A"
        ws["C35"] = "Top Sealing Cap"
        ws["D35"] = "Nylon PA66"
        ws["E35"] = "22.5s"
        ws["F35"] = "0.5 ops"
        ws["G35"] = "12.0g"
        ws["H35"] = "25,000 EAU"

        # Row 36: Unfilled formula placeholder row (must be disregarded)
        ws["A36"] = 0
        ws["B36"] = 0
        ws["C36"] = "-"
        ws["D36"] = 0
        ws["E36"] = 0.0
        ws["F36"] = 0.0
        ws["G36"] = 0.0
        ws["H36"] = 0

        file_path = tmp_path / "Q4476_MPC_Full.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert doc.total_items == 3
        assert len(doc.items) == 3

        # Verify Item 1
        it1 = doc.items[0]
        assert it1.quote_item_number == "Q4476-R1-1"
        assert it1.part_number == "100-2490-01"
        assert it1.description == "Housing Bezel Lower"
        assert it1.material == "Polypropylene 30% GF"
        assert it1.cycle_time_sec == 28.5
        assert it1.ops_labour == 0.3
        assert it1.labour_rate == 32.50
        assert it1.weight_g == 45.2
        assert it1.annual_volume == 150000
        assert it1.confidence == "high"
        assert len(it1.warnings) == 0

        # Verify coordinates recorded
        assert it1.source_cell_coords.quote_num == "A33"
        assert it1.source_cell_coords.part_num == "B33"
        assert it1.source_cell_coords.description == "C33"
        assert it1.source_cell_coords.material == "D33"
        assert it1.source_cell_coords.cycle_time == "E33"
        assert it1.source_cell_coords.ops_labour == "F33"
        assert it1.source_cell_coords.labour_rate == "E11"
        assert it1.source_cell_coords.weight_g == "G33"
        assert it1.source_cell_coords.annual_volume == "H33"

        # Verify Item 2 (Integer PN normalization)
        it2 = doc.items[1]
        assert it2.quote_item_number == "Q4476-R1-2"
        assert it2.part_number == "884012"
        assert it2.ops_labour == 0.25
        assert it2.labour_rate == 32.50
        assert it2.confidence == "high"

        # Verify Item 3 (String units parsing)
        it3 = doc.items[2]
        assert it3.quote_item_number == "Q4476-R1-3"
        assert it3.cycle_time_sec == 22.5
        assert it3.ops_labour == 0.5
        assert it3.weight_g == 12.0
        assert it3.annual_volume == 25000
        assert it3.labour_rate == 32.50
        assert it3.confidence == "high"

    def test_ops_labour_boundary_validation(self, tmp_path: Path):
        """Verify #Ops strictly validated between 0.0 and 1.0 (flags warning if > 1.0)."""
        wb = PureXmlWorkbook()
        ws = wb.active
        ws.title = "QInfo"

        ws["A1"] = "MPC RFQ #"
        ws["B1"] = "Part No."
        ws["C1"] = "Cycle Sec"
        ws["D1"] = "#Ops"

        # Row 2: valid 0.3 ops
        ws["A2"] = "Q5000-1"
        ws["B2"] = "PN-VALID"
        ws["C2"] = 15.0
        ws["D2"] = 0.3

        # Row 3: invalid 1.5 ops (exceeds 1.0)
        ws["A3"] = "Q5000-2"
        ws["B3"] = "PN-INVALID-OPS"
        ws["C3"] = 18.0
        ws["D3"] = 1.5

        file_path = tmp_path / "test_ops_bounds.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert len(doc.items) == 2
        # Item 1 is valid
        assert doc.items[0].ops_labour == 0.3

        # Item 2 has out-of-bounds warning and low confidence
        it2 = doc.items[1]
        assert it2.ops_labour == 1.5
        assert any("outside expected range (0.0–1.0)" in w for w in it2.warnings)
        assert it2.confidence == "low"

    def test_missing_required_fields_warning(self, tmp_path: Path):
        """Verify that in an MPC multi-field sheet, any missing field flags a warning and downgrades confidence to low."""
        wb = PureXmlWorkbook()
        ws = wb.active
        ws.title = "MPC"

        ws["F5"] = "Rate $CAD/Hr"
        ws["F6"] = 30.0

        # Header with all columns
        ws["A10"] = "MPC RFQ #"
        ws["B10"] = "Part No."
        ws["C10"] = "Description"
        ws["D10"] = "Material"
        ws["E10"] = "Cycle Sec"
        ws["F10"] = "Labour"
        ws["G10"] = "Weight (g)"
        ws["H10"] = "Ann. Volume (EAU)"

        # Item row missing Material and Weight
        ws["A11"] = "Q4476-R1-1"
        ws["B11"] = "100-2490-01"
        ws["C11"] = "Housing Bezel Lower"
        ws["D11"] = None  # Missing Material
        ws["E11"] = 28.5
        ws["F11"] = 0.3
        ws["G11"] = None  # Missing Weight
        ws["H11"] = 150000

        file_path = tmp_path / "test_missing_fields.xlsx"
        wb.save(file_path)

        engine = QuoteExtractorEngine()
        doc = engine.extract_file(file_path)

        assert len(doc.items) == 1
        item = doc.items[0]
        assert "Missing material" in item.warnings
        assert "Missing weight_g" in item.warnings
        assert item.confidence == "low"

    def test_multi_field_exports_contain_all_variables(self, tmp_path: Path):
        """Verify CSV, JSON, and DataFrame exports include all 9 manufacturing variables."""
        coords = SourceCellCoords(
            quote_num="A10",
            part_num="B10",
            description="C10",
            material="D10",
            cycle_time="E10",
            ops_labour="F10",
            labour_rate="F2",
            weight_g="G10",
            annual_volume="H10",
        )
        item = QuoteItem(
            quote_item_number="Q4476-R1-1",
            part_number="PN-TEST-123",
            description="Test Bracket",
            material="ABS Black",
            cycle_time_sec=22.0,
            ops_labour=0.3,
            labour_rate=35.0,
            weight_g=15.5,
            annual_volume=100000,
            source_cell_coords=coords,
        )
        doc = engine = QuoteExtractorEngine()
        from excel_extractor.models import QuoteDocument
        doc_obj = QuoteDocument(
            source_file="test_export.xlsx",
            quote_id="Q4476",
            items=[item],
        )

        # 1. Test CSV export
        csv_str = export_to_csv(doc_obj)
        assert "quote_item_number" in csv_str
        assert "description" in csv_str
        assert "material" in csv_str
        assert "ops_labour" in csv_str
        assert "labour_rate" in csv_str
        assert "weight_g" in csv_str
        assert "annual_volume" in csv_str
        assert "Test Bracket" in csv_str
        assert "ABS Black" in csv_str
        assert "35.0" in csv_str

        # 2. Test JSON export
        json_str = export_to_json(doc_obj)
        parsed = json.loads(json_str)
        item_dict = parsed["items"][0]
        assert item_dict["description"] == "Test Bracket"
        assert item_dict["material"] == "ABS Black"
        assert item_dict["ops_labour"] == 0.3
        assert item_dict["labour_rate"] == 35.0
        assert item_dict["weight_g"] == 15.5
        assert item_dict["annual_volume"] == 100000

        # 3. Test DataFrame conversion
        df = doc_obj.to_dataframe()
        assert df.iloc[0]["description"] == "Test Bracket"
        assert df.iloc[0]["material"] == "ABS Black"
        assert df.iloc[0]["ops_labour"] == 0.3
        assert df.iloc[0]["labour_rate"] == 35.0
        assert df.iloc[0]["weight_g"] == 15.5
        assert df.iloc[0]["annual_volume"] == 100000
