"""
Unit tests for Job# to Quote# Linker module (src/excel_extractor/linker.py).
"""

from __future__ import annotations

import io
from pathlib import Path
import pandas as pd
import pytest

from excel_extractor.linker import (
    JobItem,
    build_linked_dataframe,
    build_unlinked_quotes_dataframe,
    export_linked_data_to_multi_sheet_excel,
    link_jobs_to_quotes,
    normalize_part_number,
    parse_job_input_file,
    parse_job_input_text,
)
from excel_extractor.models import QuoteItem
from excel_extractor.reader import load_workbook


def test_normalize_part_number() -> None:
    """Verify part number normalization removes dashes, spaces, and revision suffixes."""
    assert normalize_part_number("2202046213000") == "2202046213000"
    assert normalize_part_number("2202046213000-01") == "2202046213000"
    assert normalize_part_number("MED-10492") == "med10492"
    assert normalize_part_number("MED10492") == "med10492"
    assert normalize_part_number("40601-BBA001 (LW)") == "40601bba001"
    assert normalize_part_number("DWG_9901_REV_A") == "dwg9901"
    assert normalize_part_number(None) == ""


def test_parse_job_input_text_tab_delimited() -> None:
    """Verify parsing tab-separated clipboard data copied directly from Excel."""
    pasted = """Job #\tPart Number
J1001\t2202046213000
J1002\t3001950000
J1003\tMED-10492
"""
    jobs = parse_job_input_text(pasted)
    assert len(jobs) == 3
    assert jobs[0].job_number == "J1001"
    assert jobs[0].part_number == "2202046213000"
    assert jobs[1].job_number == "J1002"
    assert jobs[1].part_number == "3001950000"
    assert jobs[2].job_number == "J1003"
    assert jobs[2].part_number == "MED-10492"


def test_parse_job_input_text_comma_delimited() -> None:
    """Verify parsing CSV formatted text."""
    pasted = """Job,Part
55401,858007801000
55402,858007802000
"""
    jobs = parse_job_input_text(pasted)
    assert len(jobs) == 2
    assert jobs[0].job_number == "55401"
    assert jobs[0].part_number == "858007801000"


def test_link_jobs_to_quotes_exact_order_and_blanks() -> None:
    """
    Verify linking:
    1. Preserves exact order of input jobs.
    2. Populates matched quote details.
    3. Leaves blank quote details for unquoted jobs.
    4. Separates unlinked quotes.
    """
    quotes = [
        QuoteItem(
            quote_item_number="Q100-1",
            part_number="2202046213000",
            description="Socket",
            cycle_time_sec=24.5,
            ops_labour=0.3,
            labour_rate=28.50,
        ),
        QuoteItem(
            quote_item_number="Q100-2",
            part_number="MED-10492",
            description="Medical Cap",
            cycle_time_sec=18.0,
            ops_labour=0.5,
            labour_rate=28.50,
        ),
        QuoteItem(
            quote_item_number="Q100-3",
            part_number="UNLINKED_PN_999",
            description="Spare Bracket",
            cycle_time_sec=40.0,
        ),
    ]

    jobs = [
        JobItem(job_number="JOB-A", part_number="MED-10492"),  # Matches Q100-2
        JobItem(job_number="JOB-B", part_number="UNKNOWN_PN"), # No match
        JobItem(job_number="JOB-C", part_number="2202046213000-01"), # Normalized match to Q100-1
    ]

    res = link_jobs_to_quotes(jobs, quotes)

    assert res.total_jobs == 3
    assert res.matched_count == 2
    assert res.unmatched_jobs_count == 1
    assert res.unmatched_quotes_count == 1

    # Exact order preserved
    recs = res.linked_records
    assert recs[0].job_number == "JOB-A"
    assert recs[0].quote_item_number == "Q100-2"
    assert recs[0].match_status == "matched"
    assert recs[0].cycle_time_sec == 18.0

    assert recs[1].job_number == "JOB-B"
    assert recs[1].quote_item_number is None
    assert recs[1].match_status == "job_without_quote"
    assert recs[1].cycle_time_sec is None

    assert recs[2].job_number == "JOB-C"
    assert recs[2].quote_item_number == "Q100-1"
    assert recs[2].match_status == "matched"
    assert recs[2].match_confidence == "normalized"

    # Unlinked quote
    assert len(res.unlinked_quotes) == 1
    assert res.unlinked_quotes[0].quote_item_number == "Q100-3"


def test_export_multi_sheet_excel() -> None:
    """Verify multi-sheet Excel export contains 'Linked Jobs' and 'Quotes Missing Jobs' tabs."""
    linked_df = pd.DataFrame([
        {"Job Number": "J1", "Part Number": "P1", "Quote Item": "Q1", "Cycle Sec": 20.0},
        {"Job Number": "J2", "Part Number": "P2", "Quote Item": "", "Cycle Sec": None},
    ])
    unlinked_df = pd.DataFrame([
        {"Quote Item": "Q3", "Part Number": "P3", "Description": "Extra Part"},
    ])

    excel_bytes = export_linked_data_to_multi_sheet_excel(linked_df, unlinked_df)
    assert len(excel_bytes) > 0

    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)
    sheet_names = [ws.title for ws in wb.worksheets]
    assert "Linked Jobs" in sheet_names
    assert "Quotes Missing Jobs" in sheet_names


def test_link_jobs_to_quotes_partial_match_and_brackets() -> None:
    """Verify linking supports drawing numbers, bracketed numbers, and partial matches after spaces."""
    quotes = [
        QuoteItem(
            quote_item_number="Q500-1",
            part_number="858007801000",
            drawing_number="P0900612",
            cycle_time_sec=15.0,
        ),
        QuoteItem(
            quote_item_number="Q500-2",
            part_number="2202046201000",
            drawing_number="P09000410",
            cycle_time_sec=22.5,
        ),
    ]

    jobs = [
        # Match by drawing number
        JobItem(job_number="J1", part_number="P0900612"),
        # Match by base part number
        JobItem(job_number="J2", part_number="858007801000"),
        # Match by part number with drawing in brackets
        JobItem(job_number="J3", part_number="858007801000 (P0900612)"),
        # Match by part number with drawing after space
        JobItem(job_number="J4", part_number="858007801000 P0900612"),
        # Match by partial match after space (e.g. prefix before part number)
        JobItem(job_number="J5", part_number="JOB 2202046201000"),
    ]

    res = link_jobs_to_quotes(jobs, quotes)
    assert res.matched_count == 5
    assert res.unmatched_jobs_count == 0
    assert res.linked_records[0].quote_item_number == "Q500-1"
    assert res.linked_records[1].quote_item_number == "Q500-1"
    assert res.linked_records[2].quote_item_number == "Q500-1"
    assert res.linked_records[3].quote_item_number == "Q500-1"
    assert res.linked_records[4].quote_item_number == "Q500-2"

