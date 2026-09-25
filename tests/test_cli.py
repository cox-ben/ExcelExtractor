"""
Unit and integration tests for CLI batch processor and exporters.

Tests cover:
- exporters.py: JSON, CSV, and Excel exports (single doc, batch docs, empty docs, read-back)
- cli.py argument parsing: missing args, invalid args, non-existent files, unsupported extensions
- cli.py single-file extraction: --format json, --format csv, --format both, --format excel
- cli.py directory batch processing: recursive scanning, temp lock file filtering (~$*)
- cli.py exit codes: 0 (clean run), 1 (argument/invocation error), 2 (strict mode warnings / errors)
"""

from __future__ import annotations

import csv
import json
from pathlib import Path
import sys
from typing import Dict, List
import pytest

# Ensure project root and src/ are in sys.path
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_SRC_DIR = _PROJECT_ROOT / "src"
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))
if str(_SRC_DIR) not in sys.path:
    sys.path.insert(0, str(_SRC_DIR))

import cli
from excel_extractor.exporters import (
    CSV_COLUMNS,
    export_csv,
    export_excel,
    export_json,
    export_to_csv,
    export_to_excel,
    export_to_json,
)
from excel_extractor.models import QuoteDocument, QuoteItem, SourceCellCoords
from excel_extractor.reader import load_workbook
from tests.fixtures import generator


# ==============================================================================
# Helper Functions & Fixtures
# ==============================================================================

def create_sample_quote_doc(
    quote_id: str = "Q1001",
    source_file: str = "Q1001_sample.xlsx",
    with_warning: bool = False,
) -> QuoteDocument:
    """Helper to create a populated QuoteDocument."""
    coords = SourceCellCoords(quote_num="B2", part_num="C10", cycle_time="F18")
    warnings = ["Cycle time 450.0s outside expected range (5.0s–300.0s)"] if with_warning else []
    item = QuoteItem(
        quote_item_number=f"{quote_id}-1",
        part_number="PN-9901-A",
        cycle_time_sec=450.0 if with_warning else 24.5,
        source_cell_coords=coords,
        confidence="low" if with_warning else "high",
        warnings=warnings,
    )
    return QuoteDocument(
        source_file=source_file,
        quote_id=quote_id,
        items=[item],
        warnings=[],
    )


def create_out_of_bounds_workbook(path: Path) -> Path:
    """Generate a valid quote workbook where cycle time is out of bounds (450s)."""
    wb = generator.PureXmlWorkbook()
    ws = wb.active
    ws.title = "Quote Summary"
    ws["B2"] = "Quote Number:"
    ws["C2"] = "Q9999-1"
    ws["B10"] = "Part Number:"
    ws["C10"] = "PN-OOB-9999"
    ws["B18"] = "Cycle Time (sec):"
    ws["C18"] = 450.0
    wb.save(path)
    return path


# ==============================================================================
# 1. Exporter Unit Tests
# ==============================================================================

class TestExporters:
    """Unit tests for JSON, CSV, and Excel exporters."""

    def test_export_to_json_single_document(self, tmp_path: Path):
        """Verify single QuoteDocument exports to valid JSON dict with contract structure."""
        doc = create_sample_quote_doc("Q1001")
        target_file = tmp_path / "quote.json"

        result_str = export_to_json(doc, file_path=target_file)

        # Verify return value matches file content
        assert target_file.exists()
        assert target_file.read_text(encoding="utf-8") == result_str

        # Verify parsed JSON structure
        data = json.loads(result_str)
        assert isinstance(data, dict)
        assert data["source_file"] == "Q1001_sample.xlsx"
        assert data["quote_id"] == "Q1001"
        assert data["total_items"] == 1
        assert len(data["items"]) == 1

        item = data["items"][0]
        assert item["quote_item_number"] == "Q1001-1"
        assert item["part_number"] == "PN-9901-A"
        assert item["cycle_time_sec"] == 24.5
        assert item["confidence"] == "high"
        assert item["source_cell_coords"]["quote_num"] == "B2"
        assert item["source_cell_coords"]["part_num"] == "C10"
        assert item["source_cell_coords"]["cycle_time"] == "F18"

    def test_export_to_json_as_list(self):
        """Verify as_list parameter controls whether output is a list or single dict."""
        doc = create_sample_quote_doc("Q1001")

        # as_list=True forces array wrapping
        res_list = export_to_json(doc, as_list=True)
        data_list = json.loads(res_list)
        assert isinstance(data_list, list)
        assert len(data_list) == 1
        assert data_list[0]["quote_id"] == "Q1001"

        # List of documents defaults to array
        doc2 = create_sample_quote_doc("Q1002")
        res_multi = export_to_json([doc, doc2])
        data_multi = json.loads(res_multi)
        assert isinstance(data_multi, list)
        assert len(data_multi) == 2

    def test_export_to_csv_columns_and_data(self, tmp_path: Path):
        """Verify CSV export produces exact header columns and properly formatted rows."""
        doc1 = create_sample_quote_doc("Q1001", "file1.xlsx")
        doc2 = create_sample_quote_doc("Q1002", "file2.xlsx", with_warning=True)
        target_file = tmp_path / "quotes.csv"

        csv_str = export_to_csv([doc1, doc2], file_path=target_file)

        assert target_file.exists()
        assert target_file.read_text(encoding="utf-8") == csv_str

        reader = csv.DictReader(csv_str.splitlines())
        assert reader.fieldnames == CSV_COLUMNS

        rows = list(reader)
        assert len(rows) == 2

        assert rows[0]["source_file"] == "file1.xlsx"
        assert rows[0]["quote_id"] == "Q1001"
        assert rows[0]["quote_item_number"] == "Q1001-1"
        assert rows[0]["part_number"] == "PN-9901-A"
        assert rows[0]["cycle_time_sec"] == "24.5"
        assert rows[0]["confidence"] == "high"
        assert rows[0]["quote_num_coord"] == "B2"
        assert rows[0]["part_num_coord"] == "C10"
        assert rows[0]["cycle_time_coord"] == "F18"
        assert rows[0]["warnings"] == ""

        assert rows[1]["source_file"] == "file2.xlsx"
        assert rows[1]["quote_id"] == "Q1002"
        assert rows[1]["cycle_time_sec"] == "450.0"
        assert rows[1]["confidence"] == "low"
        assert "450.0s outside expected range" in rows[1]["warnings"]

    def test_export_to_csv_empty(self):
        """Verify CSV export for an empty document produces the header row without errors."""
        doc = QuoteDocument(source_file="empty.xlsx", items=[])
        csv_str = export_to_csv(doc)
        lines = [line.strip() for line in csv_str.strip().splitlines()]
        assert len(lines) == 1
        assert lines[0] == ",".join(CSV_COLUMNS)

    def test_export_to_excel_and_readback(self, tmp_path: Path):
        """Verify Excel export creates a valid .xlsx workbook that can be re-read."""
        doc = create_sample_quote_doc("Q3000", "sample.xlsx")
        xlsx_file = tmp_path / "output.xlsx"

        data_bytes = export_to_excel([doc], file_path=xlsx_file)
        assert len(data_bytes) > 0
        assert xlsx_file.exists()
        assert xlsx_file.stat().st_size > 0

        # Read back with our resilient reader
        wb = load_workbook(xlsx_file, data_only=True)
        assert len(wb.worksheets) == 1
        ws = wb.worksheets[0]
        assert ws.title == "ExtractedQuotes"

        # Check header in Row 1
        for c_idx, col_name in enumerate(CSV_COLUMNS, start=1):
            cell = ws.cell(row=1, column=c_idx)
            assert cell.value == col_name

        # Check data in Row 2
        ct_col_idx = CSV_COLUMNS.index("cycle_time_sec") + 1
        assert ws.cell(row=2, column=1).value == "sample.xlsx"
        assert ws.cell(row=2, column=2).value == "Q3000"
        assert ws.cell(row=2, column=3).value == "Q3000-1"
        assert ws.cell(row=2, column=4).value == "PN-9901-A"
        assert ws.cell(row=2, column=ct_col_idx).value == 24.5

    def test_exporter_aliases(self):
        """Verify convenience aliases exist and behave identically."""
        assert export_json is export_to_json
        assert export_csv is export_to_csv
        assert export_excel is export_to_excel


# ==============================================================================
# 2. CLI Argument Validation Tests (Exit Code 1)
# ==============================================================================

class TestCLIArgumentValidation:
    """Tests verifying CLI argument validation and error exit codes."""

    def test_cli_missing_input_exits_code_1(self, capsys):
        """Verify calling CLI with no arguments exits with code 1."""
        code = cli.main([])
        assert code == 1
        captured = capsys.readouterr()
        assert "Missing required input path" in captured.err

    def test_cli_non_existent_input_exits_code_1(self, capsys):
        """Verify calling CLI with a non-existent path exits with code 1."""
        code = cli.main(["-i", "non_existent_quote_file.xlsx"])
        assert code == 1
        captured = capsys.readouterr()
        assert "Input path does not exist" in captured.err

    def test_cli_invalid_flag_exits_code_1(self):
        """Verify unrecognized flag triggers parser error and exits with code 1."""
        with pytest.raises(SystemExit) as exc_info:
            cli.main(["--invalid-custom-flag"])
        assert exc_info.value.code == 1

    def test_cli_non_existent_config_exits_code_1(self, single_modern_quote_file: Path, capsys):
        """Verify specifying a non-existent config file exits with code 1."""
        code = cli.main(["-i", str(single_modern_quote_file), "-c", "missing_config.yaml"])
        assert code == 1
        captured = capsys.readouterr()
        assert "Custom config file does not exist" in captured.err

    def test_cli_unsupported_file_extension_exits_code_1(self, tmp_path: Path, capsys):
        """Verify passing a non-Excel file (.txt) directly exits with code 1."""
        txt_file = tmp_path / "notes.txt"
        txt_file.write_text("Not an excel file", encoding="utf-8")

        code = cli.main(["-i", str(txt_file)])
        assert code == 1
        captured = capsys.readouterr()
        assert "Unsupported file format" in captured.err


# ==============================================================================
# 3. CLI Single File Execution Tests
# ==============================================================================

class TestCLISingleFile:
    """Tests verifying CLI execution on single Excel quote files."""

    def test_cli_single_file_both_formats(
        self, single_modern_quote_file: Path, tmp_path: Path, capsys
    ):
        """Verify CLI single file run generates both JSON and CSV files and returns 0."""
        out_dir = tmp_path / "output"
        code = cli.main(["-i", str(single_modern_quote_file), "-o", str(out_dir), "-f", "both"])
        assert code == 0

        # Check outputs
        json_file = out_dir / "extracted_quotes.json"
        csv_file = out_dir / "extracted_quotes.csv"
        indiv_json = out_dir / "Q1001.json"

        assert json_file.exists()
        assert csv_file.exists()
        assert indiv_json.exists()

        # Validate JSON content
        data = json.loads(json_file.read_text(encoding="utf-8"))
        assert isinstance(data, list)
        assert len(data) == 1
        assert data[0]["quote_id"] == "Q1001"
        assert len(data[0]["items"]) == 1
        assert data[0]["items"][0]["quote_item_number"] == "Q1001-1"
        assert data[0]["items"][0]["part_number"] == "2202046213000"
        assert data[0]["items"][0]["cycle_time_sec"] == 24.5

        # Validate CSV content
        lines = csv_file.read_text(encoding="utf-8").strip().splitlines()
        assert len(lines) == 2  # Header + 1 row
        assert "2202046213000" in lines[1]
        assert "24.5" in lines[1]

        # Verify console summary
        captured = capsys.readouterr()
        assert "Batch Extraction Complete" in captured.out
        assert "Files Scanned:       1" in captured.out
        assert "Quotes Extracted:    1 (Total Items: 1)" in captured.out
        assert "Errors / Failed:     0" in captured.out

    def test_cli_single_file_json_only(
        self, single_modern_quote_file: Path, tmp_path: Path
    ):
        """Verify --format json exports only JSON and skips CSV."""
        out_dir = tmp_path / "output_json"
        code = cli.main(["-i", str(single_modern_quote_file), "-o", str(out_dir), "-f", "json"])
        assert code == 0

        assert (out_dir / "extracted_quotes.json").exists()
        assert not (out_dir / "extracted_quotes.csv").exists()

    def test_cli_single_file_csv_only(
        self, single_modern_quote_file: Path, tmp_path: Path
    ):
        """Verify --format csv exports only CSV and skips JSON."""
        out_dir = tmp_path / "output_csv"
        code = cli.main(["-i", str(single_modern_quote_file), "-o", str(out_dir), "-f", "csv"])
        assert code == 0

        assert (out_dir / "extracted_quotes.csv").exists()
        assert not (out_dir / "extracted_quotes.json").exists()

    def test_cli_single_file_excel_format(
        self, single_modern_quote_file: Path, tmp_path: Path
    ):
        """Verify --format excel exports .xlsx workbook."""
        out_dir = tmp_path / "output_xlsx"
        code = cli.main(["-i", str(single_modern_quote_file), "-o", str(out_dir), "-f", "excel"])
        assert code == 0

        xlsx_out = out_dir / "extracted_quotes.xlsx"
        assert xlsx_out.exists()
        wb = load_workbook(xlsx_out, data_only=True)
        assert len(wb.worksheets) == 1
        assert wb.worksheets[0].cell(row=2, column=2).value == "Q1001"


# ==============================================================================
# 4. CLI Directory Batch Processing Tests
# ==============================================================================

class TestCLIDirectoryBatch:
    """Tests verifying CLI directory scanning, recursive search, and lock file filtering."""

    def test_cli_batch_directory_filtering_lock_files(
        self, batch_directory_fixture: Path, tmp_path: Path, capsys
    ):
        """
        Verify CLI directory scanning filters out temporary lock files (~$*)
        and non-Excel files, correctly aggregating quotes across multiple workbooks.
        """
        out_dir = tmp_path / "output_batch"
        code = cli.main(["-i", str(batch_directory_fixture), "-o", str(out_dir), "-f", "both"])
        assert code == 0

        # Check consolidated files
        json_file = out_dir / "extracted_quotes.json"
        csv_file = out_dir / "extracted_quotes.csv"
        assert json_file.exists()
        assert csv_file.exists()

        data = json.loads(json_file.read_text(encoding="utf-8"))
        # batch_directory_fixture contains:
        # - Q1001_single.xlsx (1 item)
        # - Q2005_multitab.xlsx (3 items)
        # - Q4000_grid.xlsx (3 items)
        # - subfolder_legacy/Q3050_macro.xlsm (1 item)
        # - subfolder_auxiliary/inventory.xlsx (0 quote items)
        # - ~$Q1001_single.xlsx (IGNORED)
        # - subfolder_legacy/~$Q3050_macro.xlsm (IGNORED)
        # Total valid files: 5. Total quote items: 8.
        source_files = [d["source_file"] for d in data]

        # Strictly ensure no lock file was processed
        assert not any(f.startswith("~$") for f in source_files)

        total_extracted_items = sum(d["total_items"] for d in data)
        assert total_extracted_items == 8

        # Verify CSV has 1 header + 8 data rows = 9 lines
        csv_lines = csv_file.read_text(encoding="utf-8").strip().splitlines()
        assert len(csv_lines) == 9

        # Verify console output
        captured = capsys.readouterr()
        assert "Batch Extraction Complete" in captured.out
        assert "Quotes Extracted:    4 (Total Items: 8)" in captured.out
        assert "Errors / Failed:     0" in captured.out

    def test_cli_directory_no_recursive(
        self, batch_directory_fixture: Path, tmp_path: Path
    ):
        """Verify --no-recursive scans only the top-level directory."""
        out_dir = tmp_path / "output_flat"
        code = cli.main([
            "-i", str(batch_directory_fixture),
            "-o", str(out_dir),
            "--no-recursive",
        ])
        assert code == 0

        data = json.loads((out_dir / "extracted_quotes.json").read_text(encoding="utf-8"))
        # Top level has Q1001 (1), Q2005 (3), Q4000 (3) = 7 items
        # Subfolder Q3050 should NOT be included
        q_ids = [d["quote_id"] for d in data if d["quote_id"]]
        assert "Q1001" in q_ids
        assert "Q2005" in q_ids
        assert "Q4000" in q_ids
        assert "Q3050" not in q_ids

    def test_cli_positional_and_dir_alias(
        self, single_modern_quote_file: Path, tmp_path: Path
    ):
        """Verify positional argument and -d alias both resolve input path correctly."""
        # Positional test
        out_pos = tmp_path / "out_pos"
        code_pos = cli.main([str(single_modern_quote_file), "-o", str(out_pos)])
        assert code_pos == 0
        assert (out_pos / "extracted_quotes.json").exists()

        # -d / --dir test
        out_dir_flag = tmp_path / "out_dir_flag"
        code_dir = cli.main(["-d", str(single_modern_quote_file.parent), "-o", str(out_dir_flag)])
        assert code_dir == 0
        assert (out_dir_flag / "extracted_quotes.json").exists()


# ==============================================================================
# 5. CLI Exit Codes & Strict Mode Tests
# ==============================================================================

class TestCLIExitCodesAndStrict:
    """Tests verifying exit code behavior and --strict flag operation."""

    def test_cli_clean_file_strict_mode_exits_0(
        self, single_modern_quote_file: Path, tmp_path: Path
    ):
        """Verify clean file with no warnings exits with 0 even when --strict is enabled."""
        out_dir = tmp_path / "output_clean_strict"
        code = cli.main(["-i", str(single_modern_quote_file), "-o", str(out_dir), "--strict"])
        assert code == 0

    def test_cli_warning_file_non_strict_exits_0(self, tmp_path: Path):
        """Verify workbook with out-of-bounds warning exits with 0 when NOT in strict mode."""
        wb_path = tmp_path / "oob_quote.xlsx"
        create_out_of_bounds_workbook(wb_path)

        out_dir = tmp_path / "output_non_strict"
        code = cli.main(["-i", str(wb_path), "-o", str(out_dir)])
        assert code == 0

    def test_cli_warning_file_strict_mode_exits_2(self, tmp_path: Path):
        """Verify workbook with out-of-bounds warning exits with code 2 when --strict is enabled."""
        wb_path = tmp_path / "oob_quote.xlsx"
        create_out_of_bounds_workbook(wb_path)

        out_dir = tmp_path / "output_strict"
        code = cli.main(["-i", str(wb_path), "-o", str(out_dir), "--strict"])
        assert code == 2

    def test_cli_verbose_flag_prints_warnings(self, tmp_path: Path, capsys):
        """Verify -v / --verbose prints detailed warning messages to stdout."""
        wb_path = tmp_path / "oob_quote_verbose.xlsx"
        create_out_of_bounds_workbook(wb_path)

        out_dir = tmp_path / "output_verbose"
        code = cli.main(["-i", str(wb_path), "-o", str(out_dir), "-v"])
        assert code == 0

        captured = capsys.readouterr()
        assert "Warnings Details:" in captured.out
        assert "450.0s outside expected range" in captured.out

    def test_cli_help_flag_exits_code_0(self, capsys):
        """Verify -h / --help prints usage information and exits with code 0."""
        with pytest.raises(SystemExit) as exc_info:
            cli.main(["--help"])
        assert exc_info.value.code == 0
        captured = capsys.readouterr()
        assert "Batch Quote Extractor CLI" in captured.out

    def test_cli_empty_directory_exits_code_0(self, tmp_path: Path, capsys):
        """Verify running CLI against an empty directory processes 0 files and exits with code 0."""
        empty_dir = tmp_path / "empty_dir"
        empty_dir.mkdir()
        out_dir = tmp_path / "empty_out"

        code = cli.main(["-i", str(empty_dir), "-o", str(out_dir)])
        assert code == 0

        captured = capsys.readouterr()
        assert "Files Scanned:       0" in captured.out
        assert "Quotes Extracted:    0 (Total Items: 0)" in captured.out
        assert (out_dir / "extracted_quotes.json").exists()
        assert (out_dir / "extracted_quotes.csv").exists()

