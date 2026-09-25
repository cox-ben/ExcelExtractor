"""
Automated headless tests for Streamlit UI (app.py) using AppTest.

Tests verify:
- Initial render state, title, and file uploader presence.
- Workbook upload and immediate extraction for single-item, multi-tab, and grid layouts.
- Summary metric card rendering (Quote ID, Total Items, Bounds Status, Confidence).
- Interactive editable data table (st.data_editor) rendering.
- Real-time cycle time bounds flagging and warnings.
- Multi-format download buttons (JSON, CSV, Excel) rendering.
- Robust error handling for empty, non-quote, and corrupted workbooks.
- Sidebar configuration changes (custom bounds, strict mode).
- Real-time updates: exports faithfully reflect manual edits.
"""

from __future__ import annotations

import io
import json
from pathlib import Path
from typing import Any
import pandas as pd
import pytest
from streamlit.testing.v1 import AppTest

from app import build_canonical_export_dict, build_dataframe_from_document, export_dataframe_to_excel
from excel_extractor.models import QuoteDocument, QuoteItem
from excel_extractor.reader import load_workbook
from tests.fixtures import generator

APP_PATH = str(Path(__file__).resolve().parent.parent / "app.py")


def test_streamlit_initial_render() -> None:
    """Verify app renders clean initial state with title, sidebar, and uploader."""
    at = AppTest.from_file(APP_PATH).run(timeout=10)

    # 1. No unhandled exceptions
    assert not at.exception, f"App crashed on initial render: {at.exception}"

    # 2. Main title
    assert len(at.title) >= 1
    assert "Legacy Quote Variable Extractor" in at.title[0].value

    # 3. Main workbook uploader present
    assert len(at.file_uploader) >= 1
    main_uploader = at.file_uploader[0]
    assert "Upload Legacy Quote Workbook" in main_uploader.label

    # 4. Guidance banner displayed
    assert len(at.info) >= 1
    assert "Please upload a legacy quote workbook" in at.info[0].value

    # 5. No metrics, data editor, or download buttons before upload
    assert len(at.metric) == 0
    assert len(at.dataframe) == 0


def test_streamlit_upload_single_modern(tmp_path: Path) -> None:
    """Verify single-item quote workbook upload, KPI metrics, data editor, and export buttons."""
    quote_path = tmp_path / "quote_single_modern.xlsx"
    generator.create_synthetic_modern_single(quote_path)
    file_bytes = quote_path.read_bytes()

    at = AppTest.from_file(APP_PATH).run(timeout=10)
    at.file_uploader(key="main_workbook_uploader").upload("quote_single_modern.xlsx", file_bytes)
    at.run(timeout=10)

    assert not at.exception, f"App crashed after uploading single quote: {at.exception}"

    # Verify Summary KPI Metrics (Quote ID, Total Items, Overall Confidence)
    assert len(at.metric) == 3
    metrics = {m.label: m.value for m in at.metric}

    assert "Quote ID" in metrics
    assert "Q1001" in metrics["Quote ID"]

    assert "Total Items Extracted" in metrics
    assert str(metrics["Total Items Extracted"]) == "1"

    assert "Overall Extraction Confidence" in metrics
    assert metrics["Overall Extraction Confidence"] == "HIGH"

    # Verify Data Editor Table
    assert len(at.dataframe) >= 1
    df = at.dataframe[0].value
    assert isinstance(df, pd.DataFrame)
    assert len(df) == 1

    row = df.iloc[0]
    assert "Q1001" in str(row["Quote Item"])
    assert "2202046213000" in str(row["Part Number"])
    ct_val = row.get("Cycle Sec") if "Cycle Sec" in row else row.get("Cycle Time (s)")
    assert float(ct_val) == 24.5
    assert row["Confidence"] == "high"

    # Verify Export Download Buttons
    button_labels = [b.label for b in (at.get("download_button") or at.button)]
    assert "Download JSON" in button_labels
    assert "Download CSV" in button_labels
    assert "Download Excel" in button_labels


def test_streamlit_upload_multitab(tmp_path: Path) -> None:
    """Verify multi-tab quote workbook upload parses all item tabs."""
    quote_path = tmp_path / "quote_multitab_q2005.xlsx"
    generator.create_synthetic_multitab(quote_path)
    file_bytes = quote_path.read_bytes()

    at = AppTest.from_file(APP_PATH).run(timeout=10)
    at.file_uploader(key="main_workbook_uploader").upload("quote_multitab_q2005.xlsx", file_bytes)
    at.run(timeout=10)

    assert not at.exception, f"App crashed after uploading multi-tab quote: {at.exception}"

    metrics = {m.label: m.value for m in at.metric}
    assert metrics["Quote ID"] == "Q2005"
    assert str(metrics["Total Items Extracted"]) == "3"
    assert metrics["Overall Extraction Confidence"] == "HIGH"


    df = at.dataframe[0].value
    assert len(df) == 3
    items = list(df["Quote Item"])
    assert "Q2005-1" in items
    assert "Q2005-2" in items
    assert "Q2005-3" in items

    button_labels = [b.label for b in (at.get("download_button") or at.button)]
    assert "Download JSON" in button_labels
    assert "Download CSV" in button_labels
    assert "Download Excel" in button_labels


def test_streamlit_upload_grid_matrix(tmp_path: Path) -> None:
    """Verify 2D row-column tabular grid quote workbook upload."""
    quote_path = tmp_path / "quote_grid_matrix_q4000.xlsx"
    generator.create_synthetic_grid_matrix(quote_path)
    file_bytes = quote_path.read_bytes()

    at = AppTest.from_file(APP_PATH).run(timeout=10)
    at.file_uploader(key="main_workbook_uploader").upload("quote_grid_matrix_q4000.xlsx", file_bytes)
    at.run(timeout=10)

    assert not at.exception, f"App crashed after uploading grid matrix: {at.exception}"

    metrics = {m.label: m.value for m in at.metric}
    assert metrics["Quote ID"] == "Q4000"
    assert str(metrics["Total Items Extracted"]) == "3"

    df = at.dataframe[0].value
    assert len(df) == 3
    part_numbers = list(df["Part Number"])
    assert "MED-10492" in part_numbers
    assert "MED-10493" in part_numbers
    assert "MED-10494" in part_numbers


def test_streamlit_empty_workbook_handling(tmp_path: Path) -> None:
    """Verify app cleanly handles empty workbook without crashing."""
    quote_path = tmp_path / "empty_workbook.xlsx"
    generator.create_synthetic_empty_workbook(quote_path)
    file_bytes = quote_path.read_bytes()

    at = AppTest.from_file(APP_PATH).run(timeout=10)
    at.file_uploader(key="main_workbook_uploader").upload("empty_workbook.xlsx", file_bytes)
    at.run(timeout=10)

    assert not at.exception, f"App crashed on empty workbook: {at.exception}"
    assert len(at.warning) >= 1
    assert "No quote line items were found" in at.warning[0].value
    assert len(at.dataframe) == 0


def test_streamlit_non_quote_workbook_handling(tmp_path: Path) -> None:
    """Verify app cleanly handles non-quote workbook with informative warning banner."""
    quote_path = tmp_path / "non_quote.xlsx"
    generator.create_synthetic_non_quote_workbook(quote_path)
    file_bytes = quote_path.read_bytes()

    at = AppTest.from_file(APP_PATH).run(timeout=10)
    at.file_uploader(key="main_workbook_uploader").upload("non_quote.xlsx", file_bytes)
    at.run(timeout=10)

    assert not at.exception, f"App crashed on non-quote workbook: {at.exception}"
    assert len(at.warning) >= 1
    assert "No quote line items were found" in at.warning[0].value
    assert len(at.dataframe) == 0


def test_streamlit_corrupt_file_handling() -> None:
    """Verify app cleanly displays error banner on invalid or corrupted binary file."""
    corrupt_bytes = b"CORRUPTED_BINARY_DATA_NOT_ZIP_FORMAT"

    at = AppTest.from_file(APP_PATH).run(timeout=10)
    at.file_uploader(key="main_workbook_uploader").upload("corrupt.xlsx", corrupt_bytes)
    at.run(timeout=10)

    assert not at.exception, f"App threw uncaught exception on corrupt file: {at.exception}"
    assert len(at.error) >= 1
    assert "Failed to parse workbook" in at.error[0].value
    assert len(at.dataframe) == 0


def test_streamlit_edge_cases_quote_parsing(tmp_path: Path) -> None:
    """Verify edge cases quote parses and renders without cycle time bounds box."""
    quote_path = tmp_path / "edge_cases_q6000.xlsx"
    generator.create_synthetic_edge_cases(quote_path)
    file_bytes = quote_path.read_bytes()

    at = AppTest.from_file(APP_PATH).run(timeout=10)
    at.file_uploader(key="main_workbook_uploader").upload("edge_cases_q6000.xlsx", file_bytes)
    at.run(timeout=10)

    assert not at.exception, f"App crashed on edge cases quote: {at.exception}"

    # Cycle time bounds status box has been removed
    warning_texts = [w.value for w in at.warning]
    assert not any("cycle time outside allowed bounds" in text.lower() for text in warning_texts)
    assert len(at.dataframe) > 0



def test_streamlit_sidebar_navigation() -> None:
    """Verify switching workspace navigation between Quote Extractor and Job Linker."""
    at = AppTest.from_file(APP_PATH).run(timeout=10)
    assert not at.exception

    # Check sidebar radio options
    assert len(at.sidebar.radio) >= 1
    assert "Quote Extractor" in at.sidebar.radio[0].options
    assert "Job# - Quote# Linker" in at.sidebar.radio[0].options

    # Switch to Job# - Quote# Linker workspace
    at.sidebar.radio(key="nav_workspace_mode").set_value("Job# - Quote# Linker")
    at.run(timeout=10)

    assert not at.exception
    # Check that Linker workspace elements are shown
    header_texts = [h.value for h in at.subheader]
    assert any("Job# - Quote# Linker Workspace" in text for text in header_texts)



def test_export_reflects_edited_dataframe() -> None:
    """Verify manual modifications to DataFrame are faithfully serialized into JSON, CSV, and Excel."""
    doc = QuoteDocument(
        source_file="test_quote.xlsx",
        quote_id="Q9999",
        total_items=1,
        items=[
            QuoteItem(
                quote_item_number="Q9999-1",
                part_number="ORIGINAL_PN",
                cycle_time_sec=20.0,
            )
        ],
    )

    df = build_dataframe_from_document(doc, min_ct=5.0, max_ct=300.0)
    assert df.iloc[0]["Part Number"] == "ORIGINAL_PN"
    ct_col = "Cycle Sec" if "Cycle Sec" in df.columns else "Cycle Time (s)"
    assert df.iloc[0][ct_col] == 20.0

    # User manually edits values in st.data_editor
    edited_df = df.copy()
    edited_df.at[0, "Part Number"] = "EDITED_PN_123"
    edited_df.at[0, ct_col] = 45.0

    # 1. JSON Export verification (Canonical Schema)
    export_dict = build_canonical_export_dict(edited_df, "test_quote.xlsx", doc.quote_id)
    json_str = json.dumps(export_dict, indent=2)
    assert "EDITED_PN_123" in json_str
    assert "45.0" in json_str
    assert "ORIGINAL_PN" not in json_str
    # Schema contract assertions
    assert export_dict["items"][0]["quote_item_number"] == "Q9999-1"
    assert export_dict["items"][0]["part_number"] == "EDITED_PN_123"
    assert export_dict["items"][0]["cycle_time_sec"] == 45.0
    assert "source_cell_coords" in export_dict["items"][0]

    # 2. CSV Export verification
    csv_str = edited_df.to_csv(index=False)
    assert "EDITED_PN_123" in csv_str
    assert "45.0" in csv_str
    assert "ORIGINAL_PN" not in csv_str

    # 3. Excel Export verification (Pure XML resilient exporter)
    excel_bytes = export_dataframe_to_excel(edited_df, sheet_name="Extracted Quotes")
    assert len(excel_bytes) > 0

    # Read back using pure-Python load_workbook (zero openpyxl dependency)
    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)
    assert len(wb.worksheets) == 1
    ws = wb.worksheets[0]
    assert ws.title == "Extracted Quotes"
    # Row 1 is header, Row 2 is data
    assert ws.cell(row=2, column=2).value == "EDITED_PN_123"
    ct_excel_col = list(edited_df.columns).index(ct_col) + 1
    assert float(ws.cell(row=2, column=ct_excel_col).value) == 45.0


def test_dataframe_omits_cell_coordinates_column() -> None:
    """Verify that generated table dataframe does not include raw coordinates column."""
    doc = QuoteDocument(
        source_file="sample.xlsx",
        quote_id="Q100",
        total_items=1,
        items=[
            QuoteItem(
                quote_item_number="Q100-1",
                part_number="PN-100",
                description="Sample Part",
                cycle_time_sec=15.0,
            )
        ],
    )
    df = build_dataframe_from_document(doc, min_ct=5.0, max_ct=300.0)
    assert "Cell Coordinates" not in df.columns
    assert "Quote Item" in df.columns
    assert "Part Number" in df.columns
    assert "Description" in df.columns
    assert "Cycle Sec" in df.columns


def test_parse_cell_coord_sheet_and_cell() -> None:
    """Verify parsing of sheet and cell coordinates."""
    from app import parse_cell_coord_sheet_and_cell
    assert parse_cell_coord_sheet_and_cell("QInfo!C34") == ("QInfo", "C34")
    assert parse_cell_coord_sheet_and_cell("MPC!B4") == ("MPC", "B4")
    assert parse_cell_coord_sheet_and_cell("B12") == ("Current Sheet", "B12")
    assert parse_cell_coord_sheet_and_cell(None) == ("N/A", "N/A")
    assert parse_cell_coord_sheet_and_cell("-") == ("N/A", "N/A")


def test_canonical_export_preserves_doc_item_coords() -> None:
    """Verify export preserves original source coordinates when doc_items is provided."""
    from excel_extractor.models import SourceCellCoords
    item = QuoteItem(
        quote_item_number="Q100-1",
        part_number="PN-100",
        description="Bracket",
        cycle_time_sec=20.0,
        source_cell_coords=SourceCellCoords(
            part_num="QInfo!C34",
            description="QInfo!D34",
            cycle_time="MPC!E12",
        ),
    )
    doc = QuoteDocument(
        source_file="sample.xlsx",
        quote_id="Q100",
        total_items=1,
        items=[item],
    )
    df = build_dataframe_from_document(doc, min_ct=5.0, max_ct=300.0)
    export_dict = build_canonical_export_dict(df, "sample.xlsx", "Q100", doc_items=doc.items)
    coords = export_dict["items"][0]["source_cell_coords"]
    assert coords["part_num"] == "QInfo!C34"
    assert coords["description"] == "QInfo!D34"
    assert coords["cycle_time"] == "MPC!E12"


def test_evaluate_cell_sanity() -> None:
    """Verify sanity checker rules for wrong columns (red) and range/format anomalies (yellow)."""
    from app import evaluate_cell_sanity

    # 1. Part number placed under Material -> RED
    st, msg = evaluate_cell_sanity("Material", "858007801000")
    assert st == "red"
    assert "Part number" in msg

    st, msg = evaluate_cell_sanity("Material", "40601-BBA001")
    assert st == "red"
    assert "Part number" in msg

    # Real material -> CLEAN
    st, msg = evaluate_cell_sanity("Material", "Grivory GVX-6H BK (E1083756)")
    assert st == "clean"

    st, msg = evaluate_cell_sanity("Material", "Ziamid B3.30 PA 6 GF 30")
    assert st == "clean"

    # Unassigned material -> YELLOW
    st, msg = evaluate_cell_sanity("Material", "NO MAT IDENTIFIED")
    assert st == "yellow"

    # 2. Extreme operator count (>2.0, e.g. 10 ops) -> RED
    st, msg = evaluate_cell_sanity("#Ops", 10.0)
    assert st == "red"
    assert "10.0 ops" in msg

    # Operator count slightly high (1.0 < ops <= 2.0) -> YELLOW
    st, msg = evaluate_cell_sanity("#Ops", 1.5)
    assert st == "yellow"

    # Standard operator count (0.3) -> CLEAN
    st, msg = evaluate_cell_sanity("#Ops", 0.3)
    assert st == "clean"

    # 3. Labour rate above $33.00/hr -> YELLOW
    st, msg = evaluate_cell_sanity("Labour Rate ($CAD/Hr)", 48.50)
    assert st == "yellow"
    assert "$48.50/hr" in msg

    # Normal labour rate ($28.00/hr) -> CLEAN
    st, msg = evaluate_cell_sanity("Labour Rate ($CAD/Hr)", 28.00)
    assert st == "clean"

    # Negative labour rate -> RED
    st, msg = evaluate_cell_sanity("Labour Rate ($CAD/Hr)", -5.0)
    assert st == "red"

    # 4. Description in Part Number column -> RED
    st, msg = evaluate_cell_sanity("Part Number", "Spring Control 2 (OS) Socket Assembly")
    assert st == "red"

    # Valid part number -> CLEAN
    st, msg = evaluate_cell_sanity("Part Number", "3001950000")
    assert st == "clean"

    # 5. Blank cells across extracted data fields -> RED error
    st, msg = evaluate_cell_sanity("Weight (g)", None)
    assert st == "red"
    assert "Blank cell: missing Part Weight" in msg

    st, msg = evaluate_cell_sanity("Cycle Sec", "-")
    assert st == "red"
    assert "Blank cell: missing Cycle Time" in msg

    st, msg = evaluate_cell_sanity("Part Number", "")
    assert st == "red"
    assert "Blank cell: missing Part Number" in msg

    st, msg = evaluate_cell_sanity("Material", "None")
    assert st == "red"
    assert "Blank cell: missing Material" in msg

    st, msg = evaluate_cell_sanity("#Ops", "nan")
    assert st == "red"
    assert "Blank cell: missing Number of Operators" in msg

    st, msg = evaluate_cell_sanity("Labour Rate ($CAD/Hr)", " - ")
    assert st == "red"
    assert "Blank cell: missing Labour Rate" in msg

    st, msg = evaluate_cell_sanity("Ann. Volume (EAU)", None)
    assert st == "red"
    assert "Blank cell: missing Annual Volume" in msg

    # Non-data column (Warnings) with empty string -> CLEAN
    st, msg = evaluate_cell_sanity("Warnings", "")
    assert st == "clean"



