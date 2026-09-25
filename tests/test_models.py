"""
Unit tests for data contracts, validation rules, and configuration management.

Covers:
- SourceCellCoords instantiation and dictionary conversion
- QuoteItem validation, bounds checking, and confidence calculation
- Cycle time bounds edge cases (strictly 5.0s–300.0s: 2.4s, 5.0s, 24.5s, 300.0s, 450.0s)
- Part number formatting and cleaning
- QuoteDocument synchronization, JSON/dict round-tripping
- Flat row tabular serialization and Pandas DataFrame conversion
- ConfigManager YAML loading, domain defaults, and corrupted/missing file fallbacks
"""

import sys
from pathlib import Path
import pytest

# Ensure src/ is on sys.path for direct pytest invocation
SRC_DIR = Path(__file__).resolve().parent.parent / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from excel_extractor.models import (
    QuoteDocument,
    QuoteItem,
    SourceCellCoords,
)
from excel_extractor.config import (
    AnchorConfig,
    BoundsConfig,
    ConfigManager,
    ExtractorConfig,
)


# ==============================================================================
# 1. SourceCellCoords Tests
# ==============================================================================

def test_source_cell_coords_defaults():
    """Verify default SourceCellCoords initializes all coordinates to None."""
    coords = SourceCellCoords()
    assert coords.quote_num is None
    assert coords.part_num is None
    assert coords.cycle_time is None
    assert coords.description is None
    assert coords.material is None
    assert coords.ops_labour is None
    assert coords.labour_rate is None
    assert coords.weight_g is None
    assert coords.annual_volume is None
    d = coords.to_dict()
    assert d["quote_num"] is None
    assert d["part_num"] is None
    assert d["cycle_time"] is None
    assert d["description"] is None
    assert d["material"] is None
    assert d["ops_labour"] is None
    assert d["labour_rate"] is None
    assert d["weight_g"] is None
    assert d["annual_volume"] is None


def test_source_cell_coords_custom():
    """Verify coordinate tracking records exact cell references."""
    coords = SourceCellCoords(quote_num="B2", part_num="C10", cycle_time="F18")
    assert coords.quote_num == "B2"
    assert coords.part_num == "C10"
    assert coords.cycle_time == "F18"
    assert coords.to_dict()["cycle_time"] == "F18"


# ==============================================================================
# 2. QuoteItem Instantiation & Field Cleaning Tests
# ==============================================================================

def test_valid_quote_item():
    """Verify normal QuoteItem instantiation with valid fields."""
    coords = SourceCellCoords(quote_num="B2", part_num="C10", cycle_time="F18")
    item = QuoteItem(
        quote_item_number="Q1001-1",
        part_number="3042605C",
        cycle_time_sec=24.5,
        source_cell_coords=coords,
    )
    assert item.quote_item_number == "Q1001-1"
    assert item.part_number == "3042605C"
    assert item.cycle_time_sec == 24.5
    assert item.confidence == "high"
    assert len(item.warnings) == 0
    assert item.source_cell_coords.quote_num == "B2"


def test_part_number_float_cleaning():
    """Verify integer-equivalent floats (e.g. 2202046213000.0) are normalized to integer strings."""
    item = QuoteItem(
        quote_item_number="Q1001-1",
        part_number=2202046213000.0,
        cycle_time_sec=25.0,
    )
    assert item.part_number == "2202046213000"


def test_part_number_colon_and_whitespace_cleaning():
    """Verify trailing colons and extra whitespace are stripped from part number."""
    item = QuoteItem(
        quote_item_number="Q1001-1",
        part_number=" : 3042605C : ",
        cycle_time_sec=25.0,
    )
    assert item.part_number == "3042605C"


def test_part_number_dwg_cleaning():
    """Verify drawing number suffixes are stripped from part numbers."""
    item1 = QuoteItem(
        quote_item_number="Q1001-1",
        part_number="3042605C (DWG 4022-A)",
        cycle_time_sec=25.0,
    )
    assert item1.part_number == "3042605C"

    item2 = QuoteItem(
        quote_item_number="Q1001-2",
        part_number="ABC-999 DWG #100",
        cycle_time_sec=25.0,
    )
    assert item2.part_number == "ABC-999"


def test_cycle_time_string_with_units_parsing():
    """Verify cycle time strings with units like 's', 'sec', 'seconds' are parsed to float."""
    item1 = QuoteItem(quote_item_number="Q1001-1", part_number="PART-A", cycle_time_sec="24.5s")
    assert item1.cycle_time_sec == 24.5

    item2 = QuoteItem(quote_item_number="Q1001-2", part_number="PART-B", cycle_time_sec="15.2 sec")
    assert item2.cycle_time_sec == 15.2

    item3 = QuoteItem(quote_item_number="Q1001-3", part_number="PART-C", cycle_time_sec="30.0 seconds")
    assert item3.cycle_time_sec == 30.0


# ==============================================================================
# 3. Cycle Time Bounds Warning Tests (2.4s, 5.0s, 24.5s, 300.0s, 450.0s)
# ==============================================================================

def test_cycle_time_bounds_2_4s_under_bounds():
    """
    Test 2.4s (under min 5.0s):
    Must NOT discard value, must append warning, and must downgrade confidence to 'low'.
    """
    item = QuoteItem(
        quote_item_number="Q1001-1",
        part_number="PART-X",
        cycle_time_sec=2.4,
    )
    assert item.cycle_time_sec == 2.4  # Value preserved
    assert item.confidence == "low"
    assert any("2.4" in w and "5.0s–300.0s" in w for w in item.warnings)


def test_cycle_time_bounds_5_0s_min_boundary():
    """
    Test 5.0s (exact min boundary):
    Must be accepted within bounds, no warning, confidence 'high'.
    """
    item = QuoteItem(
        quote_item_number="Q1001-1",
        part_number="PART-X",
        cycle_time_sec=5.0,
    )
    assert item.cycle_time_sec == 5.0
    assert item.confidence == "high"
    assert not any("outside expected range" in w for w in item.warnings)


def test_cycle_time_bounds_24_5s_nominal():
    """
    Test 24.5s (nominal mid-range):
    Must be accepted within bounds, no warning, confidence 'high'.
    """
    item = QuoteItem(
        quote_item_number="Q1001-1",
        part_number="PART-X",
        cycle_time_sec=24.5,
    )
    assert item.cycle_time_sec == 24.5
    assert item.confidence == "high"
    assert not any("outside expected range" in w for w in item.warnings)


def test_cycle_time_bounds_300_0s_max_boundary():
    """
    Test 300.0s (exact max boundary):
    Must be accepted within bounds, no warning, confidence 'high'.
    """
    item = QuoteItem(
        quote_item_number="Q1001-1",
        part_number="PART-X",
        cycle_time_sec=300.0,
    )
    assert item.cycle_time_sec == 300.0
    assert item.confidence == "high"
    assert not any("outside expected range" in w for w in item.warnings)


def test_cycle_time_bounds_450_0s_over_bounds():
    """
    Test 450.0s (over max 300.0s):
    Must NOT discard value, must append warning, and must downgrade confidence to 'low'.
    """
    item = QuoteItem(
        quote_item_number="Q1001-1",
        part_number="PART-X",
        cycle_time_sec=450.0,
    )
    assert item.cycle_time_sec == 450.0  # Value preserved
    assert item.confidence == "low"
    assert any("450.0" in w and "5.0s–300.0s" in w for w in item.warnings)


# ==============================================================================
# 4. Missing Fields & Confidence Scoring Heuristics
# ==============================================================================

def test_missing_cycle_time_sec():
    """Verify missing cycle_time_sec appends warning and sets confidence to 'low'."""
    item = QuoteItem(
        quote_item_number="Q1001-1",
        part_number="PART-A",
        cycle_time_sec=None,
    )
    assert item.cycle_time_sec is None
    assert item.confidence == "low"
    assert "Missing cycle_time_sec" in item.warnings


def test_missing_part_number():
    """Verify missing part_number appends warning and sets confidence to 'low'."""
    item = QuoteItem(
        quote_item_number="Q1001-1",
        part_number=None,
        cycle_time_sec=25.0,
    )
    assert item.part_number is None
    assert item.confidence == "low"
    assert "Missing part_number" in item.warnings


def test_confidence_medium_on_inferred_item():
    """Verify non-critical warnings (e.g. inferred suffix) result in 'medium' confidence."""
    item = QuoteItem(
        quote_item_number="Q1001-1",
        part_number="PART-A",
        cycle_time_sec=25.0,
        warnings=["Inferred item suffix -1 from base quote Q1001"],
    )
    assert item.confidence == "medium"


def test_confidence_medium_on_non_standard_quote_number():
    """Verify non-standard item number format degrades confidence to 'medium'."""
    item = QuoteItem(
        quote_item_number="Q1001",  # Missing item suffix
        part_number="PART-A",
        cycle_time_sec=25.0,
    )
    assert item.confidence == "medium"
    assert any("Non-standard quote item number format" in w for w in item.warnings)


def test_confidence_low_takes_precedence_over_medium():
    """Verify that if both a non-critical warning and an out-of-bounds error occur, confidence is 'low'."""
    item = QuoteItem(
        quote_item_number="Q1001",
        part_number="PART-A",
        cycle_time_sec=1.5,
        warnings=["Inferred item suffix -1 from base quote Q1001"],
    )
    assert item.confidence == "low"


# ==============================================================================
# 5. QuoteDocument & Serialization Tests
# ==============================================================================

def test_quote_document_total_items_and_quote_id_sync():
    """Verify QuoteDocument auto-syncs total_items and infers quote_id from first item."""
    item1 = QuoteItem(
        quote_item_number="Q2005-1",
        part_number="PART-1",
        cycle_time_sec=18.0,
        source_cell_coords=SourceCellCoords(quote_num="B2", part_num="C5", cycle_time="F10"),
    )
    item2 = QuoteItem(
        quote_item_number="Q2005-2",
        part_number="PART-2",
        cycle_time_sec=22.5,
        source_cell_coords=SourceCellCoords(quote_num="B2", part_num="C5", cycle_time="F10"),
    )
    doc = QuoteDocument(
        source_file="Quote_Q2005.xlsx",
        items=[item1, item2],
    )
    assert doc.total_items == 2
    assert doc.quote_id == "Q2005"


def test_quote_document_add_item():
    """Verify add_item method dynamically updates total_items and quote_id."""
    doc = QuoteDocument(source_file="Quote_Q3050.xlsx")
    assert doc.total_items == 0
    assert doc.quote_id is None

    item = QuoteItem(
        quote_item_number="Q3050-1",
        part_number="PART-3050",
        cycle_time_sec=35.0,
    )
    doc.add_item(item)
    assert doc.total_items == 1
    assert doc.quote_id == "Q3050"


def test_quote_document_dict_and_json_roundtrip():
    """Verify QuoteDocument serializes to dict and JSON and round-trips cleanly."""
    item = QuoteItem(
        quote_item_number="Q1001-1",
        part_number="PART-ROUNDTRIP",
        cycle_time_sec=14.0,
        source_cell_coords=SourceCellCoords(quote_num="A1", part_num="B2", cycle_time="C3"),
    )
    doc = QuoteDocument(
        source_file="Roundtrip.xlsx",
        items=[item],
        warnings=["Test doc warning"],
    )

    # To Dict and back
    data_dict = doc.to_dict()
    assert data_dict["source_file"] == "Roundtrip.xlsx"
    assert data_dict["items"][0]["part_number"] == "PART-ROUNDTRIP"
    doc_from_dict = QuoteDocument.from_dict(data_dict)
    assert doc_from_dict.quote_id == "Q1001"
    assert doc_from_dict.total_items == 1

    # To JSON and back
    json_str = doc.to_json()
    assert "PART-ROUNDTRIP" in json_str
    doc_from_json = QuoteDocument.from_json(json_str)
    assert doc_from_json.source_file == "Roundtrip.xlsx"
    assert doc_from_json.items[0].cycle_time_sec == 14.0


def test_flat_rows_serialization():
    """Verify to_flat_rows() produces flat dictionaries matching CSV/tabular schema."""
    item = QuoteItem(
        quote_item_number="Q5001-1",
        part_number="PN-5001",
        cycle_time_sec=19.5,
        source_cell_coords=SourceCellCoords(quote_num="B2", part_num="D4", cycle_time="E8"),
        warnings=["Minor note"],
    )
    doc = QuoteDocument(
        source_file="Q5001.xlsx",
        items=[item],
    )
    rows = doc.to_flat_rows()
    assert len(rows) == 1
    row = rows[0]

    expected_keys = [
        "source_file",
        "quote_id",
        "quote_item_number",
        "part_number",
        "cycle_time_sec",
        "confidence",
        "quote_num_coord",
        "part_num_coord",
        "cycle_time_coord",
        "warnings",
    ]
    for key in expected_keys:
        assert key in row

    assert row["source_file"] == "Q5001.xlsx"
    assert row["quote_id"] == "Q5001"
    assert row["quote_item_number"] == "Q5001-1"
    assert row["part_number"] == "PN-5001"
    assert row["cycle_time_sec"] == 19.5
    assert row["confidence"] == "medium"
    assert row["quote_num_coord"] == "B2"
    assert row["part_num_coord"] == "D4"
    assert row["cycle_time_coord"] == "E8"
    assert row["warnings"] == "Minor note"


def test_to_dataframe():
    """Verify to_dataframe() exports a valid pandas DataFrame with correct columns and data."""
    item1 = QuoteItem(
        quote_item_number="Q6001-1",
        part_number="PN-6001A",
        cycle_time_sec=12.0,
        source_cell_coords=SourceCellCoords(quote_num="A1", part_num="B2", cycle_time="C3"),
    )
    item2 = QuoteItem(
        quote_item_number="Q6001-2",
        part_number="PN-6001B",
        cycle_time_sec=25.0,
        source_cell_coords=SourceCellCoords(quote_num="A1", part_num="B5", cycle_time="C6"),
    )
    doc = QuoteDocument(
        source_file="Q6001.xlsx",
        items=[item1, item2],
    )
    df = doc.to_dataframe()
    assert df.shape == (2, 22)
    assert list(df["quote_item_number"]) == ["Q6001-1", "Q6001-2"]
    assert list(df["part_number"]) == ["PN-6001A", "PN-6001B"]
    assert list(df["cycle_time_sec"]) == [12.0, 25.0]


# ==============================================================================
# 6. ConfigManager & YAML Loading Tests
# ==============================================================================

def test_config_manager_loads_root_yaml():
    """Verify ConfigManager loads config.yaml at project root correctly."""
    root_config_path = Path(__file__).resolve().parent.parent / "config.yaml"
    mgr = ConfigManager(root_config_path)
    cfg = mgr.config

    # Regex patterns
    assert r"\b(Q[0-9]{4,5})\b" in cfg.quote_id_patterns
    assert r"\b(Q[0-9]{4,5}-[0-9]+)\b" in cfg.quote_item_patterns

    # Part number anchors
    assert "part number" in cfg.part_number_anchors
    assert "p/n" in cfg.part_number_anchors
    assert [0, 1] in cfg.part_number_offsets

    # Cycle time anchors and bounds
    assert "cycle time" in cfg.cycle_time_anchors
    assert "c/t" in cfg.cycle_time_anchors
    assert cfg.min_cycle_time == 5.0
    assert cfg.max_cycle_time == 300.0

    # Ignore sheets
    assert "summary" in cfg.ignore_sheets
    assert "cover" in cfg.ignore_sheets
    assert "toc" in cfg.ignore_sheets
    assert "notes" in cfg.ignore_sheets
    assert "lookup" in cfg.ignore_sheets
    assert "instructions" in cfg.ignore_sheets


def test_config_manager_missing_file_fallback():
    """Verify ConfigManager falls back to defaults if file does not exist."""
    mgr = ConfigManager("non_existent_config_file_xyz.yaml")
    cfg = mgr.config

    assert cfg.config_source == "defaults"
    assert cfg.min_cycle_time == 5.0
    assert cfg.max_cycle_time == 300.0
    assert len(cfg.part_number_anchors) > 0
    assert len(cfg.cycle_time_anchors) > 0
    assert "summary" in cfg.ignore_sheets


def test_config_manager_corrupted_yaml_fallback(tmp_path):
    """Verify ConfigManager falls back to defaults if YAML file is corrupted/invalid."""
    bad_yaml = tmp_path / "corrupted.yaml"
    bad_yaml.write_text("invalid: yaml: [syntax: error: :::", encoding="utf-8")

    mgr = ConfigManager(bad_yaml)
    cfg = mgr.config

    assert cfg.config_source == "defaults"
    assert cfg.min_cycle_time == 5.0
    assert cfg.max_cycle_time == 300.0


def test_config_manager_partial_yaml(tmp_path):
    """Verify ConfigManager safely merges partial custom YAML configurations."""
    partial_yaml = tmp_path / "partial.yaml"
    partial_yaml.write_text(
        """
extraction:
  bounds:
    cycle_time_sec:
      min: 10.0
      max: 180.0
""",
        encoding="utf-8",
    )

    mgr = ConfigManager(partial_yaml)
    cfg = mgr.config

    assert cfg.min_cycle_time == 10.0
    assert cfg.max_cycle_time == 180.0
    # Other sections should safely fall back to domain defaults
    assert "part number" in cfg.part_number_anchors
    assert "cycle time" in cfg.cycle_time_anchors
    assert "summary" in cfg.ignore_sheets
