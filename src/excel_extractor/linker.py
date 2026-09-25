"""
Job Number to Quote Number Linker Module.

Provides:
- Data models for Job items, Linked records, and LinkResult summary.
- Robust parsing for pasted text (tab-separated, comma-separated, space-aligned) and uploaded files (.xlsx, .csv).
- Smart normalized part number matching (exact match, case/dash/space-insensitive, suffix-tolerant).
- Strict order preservation of input jobs (blanks for unquoted jobs).
- Extraction of unlinked quotes (quotes missing jobs).
- Multi-sheet Excel workbook exporter.
"""

from __future__ import annotations

import csv
import io
import re
from typing import Any, Dict, List, Literal, Optional, Tuple
import pandas as pd
from pydantic import BaseModel, Field

from excel_extractor.exporters import export_dataframe_to_excel
from excel_extractor.models import QuoteItem
from excel_extractor.reader import load_workbook


class JobItem(BaseModel):
    """Represents a job record provided by the user."""
    job_number: str = Field(..., description="Job identifier (e.g., J10023, 55420)")
    part_number: str = Field(..., description="Part number associated with the job")
    original_index: int = Field(default=0, description="Original input row order")


class LinkedRecord(BaseModel):
    """Represents a joined Job + Quote record."""
    job_number: str
    part_number_job: str
    part_number_quote: Optional[str] = None
    quote_item_number: Optional[str] = None
    description: Optional[str] = None
    material: Optional[str] = None
    cycle_time_sec: Optional[float] = None
    ops_labour: Optional[float] = None
    labour_rate: Optional[float] = None
    weight_g: Optional[float] = None
    annual_volume: Optional[int] = None
    match_status: Literal["matched", "job_without_quote"] = "job_without_quote"
    match_confidence: Literal["exact", "normalized", "none"] = "none"


class LinkResult(BaseModel):
    """Result of linking a list of jobs to quote line items."""
    total_jobs: int
    matched_count: int
    unmatched_jobs_count: int
    unmatched_quotes_count: int
    linked_records: List[LinkedRecord]
    unlinked_quotes: List[QuoteItem]


def normalize_part_number(pn: Optional[str]) -> str:
    """
    Normalize part number for resilient cross-system matching.
    - Strips leading/trailing whitespace
    - Lowercases
    - Strips spaces, hyphens, underscores, dots
    - Removes common revision/drawing trailing tags like '-r1', '-01', '(lw)', '(os)', '_rev_a'
    """
    if not pn:
        return ""
    text = str(pn).strip().lower()
    # Strip parenthetical/bracketed annotations e.g. (LW), (OS), (P0900612), [P0900712]
    text = re.sub(r"\s*[\(\[][a-z0-9\s_-]+[\)\]]$", "", text)
    # Strip revision tags e.g. _rev_a, -rev01, _r1, -01
    text = re.sub(r"[-_]rev(?:ision)?[-_\s]*[a-z0-9]*$", "", text)
    text = re.sub(r"[-_](?:r)?[a-z0-9]{1,2}$", "", text)
    # Strip non-alphanumeric
    text = re.sub(r"[\s\-_.:/]+", "", text)
    return text


def extract_matching_tokens(text: Optional[str]) -> List[str]:
    """
    Extracts all candidate matching keys from a part number or drawing identifier:
    - Full normalized string
    - Base string before brackets
    - Content inside brackets / parentheticals (e.g. '(P0900612)' -> 'p0900612')
    - Substrings and tokens after spaces (e.g. '858007801000 (P0900612)' -> 'p0900612', '858007801000')
    - Substrings after the first space, and individual whitespace-split words
    """
    if not text:
        return []
    s = str(text).strip()
    if not s:
        return []

    tokens: List[str] = []

    def _add_token(tok: str) -> None:
        norm = normalize_part_number(tok)
        if norm and len(norm) >= 2 and norm not in ("dwg", "drawing", "rev", "revision", "part", "item"):
            if norm not in tokens:
                tokens.append(norm)

    # 1. Full normalized
    _add_token(s)

    # 2. Content inside brackets/parentheses e.g. (P0900612), [P0900712]
    bracket_matches = re.findall(r"[\(\[]\s*([a-zA-Z0-9\s_-]+)\s*[\)\]]", s)
    for bm in bracket_matches:
        _add_token(bm)

    # 3. Base part before brackets
    before_bracket = re.sub(r"\s*[\(\[].*?[\)\]].*$", "", s).strip()
    if before_bracket and before_bracket != s:
        _add_token(before_bracket)

    # 4. Partial matches after space (and individual tokens)
    # Split by whitespace, newline, or tab
    parts = [p.strip(" \t\r\n:;-()[]") for p in re.split(r"[\s\t\n\r]+", s) if p.strip(" \t\r\n:;-()[]")]
    if len(parts) >= 2:
        # First part (before space)
        _add_token(parts[0])

        # Every token after the space
        for p in parts[1:]:
            _add_token(p)

        # Entire substring after first space
        after_space = s.split(None, 1)[1]
        _add_token(after_space)

    return tokens




def parse_job_input_text(raw_text: str) -> List[JobItem]:
    """
    Parse pasted text (from Excel clipboard, CSV, or whitespace-separated columns)
    into a list of JobItem records, preserving exact row order.
    """
    if not raw_text or not raw_text.strip():
        return []

    lines = [line.strip() for line in raw_text.strip().splitlines() if line.strip()]
    if not lines:
        return []

    jobs: List[JobItem] = []
    job_idx = 0

    for line_idx, line in enumerate(lines):
        # Determine delimiter: tab, comma, semicolon, pipe, or multiple spaces
        if "\t" in line:
            parts = [p.strip() for p in line.split("\t") if p.strip()]
        elif "," in line:
            reader = csv.reader([line])
            parts = [p.strip() for p in next(reader) if p.strip()]
        elif ";" in line:
            parts = [p.strip() for p in line.split(";") if p.strip()]
        elif "|" in line:
            parts = [p.strip() for p in line.split("|") if p.strip()]
        else:
            parts = [p.strip() for p in re.split(r"\s{2,}|\s+", line) if p.strip()]

        if not parts:
            continue

        # Check if this line is a header row (e.g. "Job #", "Part Number")
        first_token = parts[0].lower()
        if line_idx == 0 and ("job" in first_token or "order" in first_token or "wo" in first_token):
            continue

        if len(parts) >= 2:
            job_no = parts[0]
            part_no = parts[1]
        elif len(parts) == 1:
            job_no = parts[0]
            part_no = ""
        else:
            continue

        jobs.append(JobItem(
            job_number=job_no,
            part_number=part_no,
            original_index=job_idx,
        ))
        job_idx += 1

    return jobs


def parse_job_input_file(file_bytes: bytes, filename: str) -> List[JobItem]:
    """
    Parse an uploaded spreadsheet (.xlsx, .xlsm, .csv) into a list of JobItem records.
    """
    filename_lower = filename.lower()
    jobs: List[JobItem] = []

    if filename_lower.endswith(".csv"):
        text = file_bytes.decode("utf-8", errors="replace")
        return parse_job_input_text(text)

    # Excel file
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb.worksheets[0]

        # Scan first 5 rows to locate header columns
        job_col = 1
        part_col = 2
        start_row = 1

        for r in range(1, min(6, ws.max_row + 1)):
            row_vals = [str(ws.cell(row=r, column=c).value or "").strip().lower() for c in range(1, min(15, ws.max_column + 1))]
            for c_idx, val in enumerate(row_vals, start=1):
                if "job" in val or "order" in val or "wo" in val:
                    job_col = c_idx
                    start_row = r + 1
                elif "part" in val or "item" in val or "pn" in val:
                    part_col = c_idx
                    start_row = r + 1

        job_idx = 0
        for r in range(start_row, ws.max_row + 1):
            j_val = str(ws.cell(row=r, column=job_col).value or "").strip()
            p_val = str(ws.cell(row=r, column=part_col).value or "").strip()
            if not j_val and not p_val:
                continue
            if j_val.lower() in ("none", "-", "nan"):
                j_val = ""
            if p_val.lower() in ("none", "-", "nan"):
                p_val = ""

            if j_val or p_val:
                jobs.append(JobItem(
                    job_number=j_val if j_val else f"JOB-{job_idx+1}",
                    part_number=p_val,
                    original_index=job_idx,
                ))
                job_idx += 1
    except Exception as exc:
        raise ValueError(f"Failed to parse Excel job list: {exc}") from exc

    return jobs


def link_jobs_to_quotes(jobs: List[JobItem], quote_items: List[QuoteItem]) -> LinkResult:
    """
    Link a list of Job records to extracted Quote line items via Part Number.

    Contract:
    1. Preserves the exact input order of jobs.
    2. Uses exact Part Number match first, falling back to smart normalized matching.
    3. If a job has no matching quote number, leaves quote details as None / blank.
    4. Tracks quotes that were not linked to any job.
    """
    # Build lookup dictionaries for quote items
    # 1. Exact part number map (case-insensitive) -> list of QuoteItem
    exact_map: Dict[str, List[QuoteItem]] = {}
    # 2. Normalized part number map -> list of QuoteItem
    norm_map: Dict[str, List[QuoteItem]] = {}
    # 3. Token & partial match map (after-space tokens, bracket contents, base numbers)
    token_map: Dict[str, List[QuoteItem]] = {}

    for q in quote_items:
        raw_pn = (q.part_number or "").strip()
        if raw_pn:
            k_exact = raw_pn.lower()
            exact_map.setdefault(k_exact, []).append(q)

            k_norm = normalize_part_number(raw_pn)
            if k_norm:
                norm_map.setdefault(k_norm, []).append(q)

            for t in extract_matching_tokens(raw_pn):
                token_map.setdefault(t, []).append(q)

        # Index drawing_number if present
        raw_dwg = (getattr(q, "drawing_number", None) or "").strip()
        if raw_dwg:
            exact_map.setdefault(raw_dwg.lower(), []).append(q)
            k_dwg_norm = normalize_part_number(raw_dwg)
            if k_dwg_norm:
                norm_map.setdefault(k_dwg_norm, []).append(q)
            for t in extract_matching_tokens(raw_dwg):
                token_map.setdefault(t, []).append(q)

    used_quote_ids: set[str] = set()
    linked_records: List[LinkedRecord] = []
    matched_count = 0
    unmatched_jobs_count = 0

    for job in jobs:
        raw_job_pn = (job.part_number or "").strip()
        matched_quote: Optional[QuoteItem] = None
        match_conf: Literal["exact", "normalized", "none"] = "none"

        if raw_job_pn:
            # 1. Try exact match
            k_exact = raw_job_pn.lower()
            candidates = exact_map.get(k_exact, [])
            avail = [c for c in candidates if (c.quote_item_number or id(c)) not in used_quote_ids]
            if avail:
                matched_quote = avail[0]
                match_conf = "exact"
            elif candidates:
                matched_quote = candidates[0]
                match_conf = "exact"

            # 2. Try normalized match on full string if no exact match found
            if not matched_quote:
                k_norm = normalize_part_number(raw_job_pn)
                if k_norm:
                    norm_candidates = norm_map.get(k_norm, [])
                    avail_norm = [c for c in norm_candidates if (c.quote_item_number or id(c)) not in used_quote_ids]
                    if avail_norm:
                        matched_quote = avail_norm[0]
                        match_conf = "normalized"
                    elif norm_candidates:
                        matched_quote = norm_candidates[0]
                        match_conf = "normalized"

            # 3. Try partial matches after the space / bracket tokens / drawing references
            if not matched_quote:
                job_tokens = extract_matching_tokens(raw_job_pn)
                for t in job_tokens:
                    tok_candidates = token_map.get(t, [])
                    avail_tok = [c for c in tok_candidates if (c.quote_item_number or id(c)) not in used_quote_ids]
                    if avail_tok:
                        matched_quote = avail_tok[0]
                        match_conf = "normalized"
                        break
                    elif tok_candidates:
                        matched_quote = tok_candidates[0]
                        match_conf = "normalized"
                        break

        # 4. Fallback: If part number was empty but job_number contained space-separated tokens
        if not matched_quote and job.job_number and (" " in job.job_number or "\t" in job.job_number):
            job_no_tokens = extract_matching_tokens(job.job_number)
            for t in job_no_tokens:
                tok_candidates = token_map.get(t, [])
                avail_tok = [c for c in tok_candidates if (c.quote_item_number or id(c)) not in used_quote_ids]
                if avail_tok:
                    matched_quote = avail_tok[0]
                    match_conf = "normalized"
                    break
                elif tok_candidates:
                    matched_quote = tok_candidates[0]
                    match_conf = "normalized"
                    break

        if matched_quote:
            used_quote_ids.add(matched_quote.quote_item_number or id(matched_quote))
            matched_count += 1
            linked_records.append(LinkedRecord(
                job_number=job.job_number,
                part_number_job=job.part_number,
                part_number_quote=matched_quote.part_number,
                quote_item_number=matched_quote.quote_item_number,
                description=matched_quote.description,
                material=matched_quote.material,
                cycle_time_sec=matched_quote.cycle_time_sec,
                ops_labour=matched_quote.ops_labour,
                labour_rate=matched_quote.labour_rate,
                weight_g=matched_quote.weight_g,
                annual_volume=matched_quote.annual_volume,
                match_status="matched",
                match_confidence=match_conf,
            ))
        else:
            unmatched_jobs_count += 1
            linked_records.append(LinkedRecord(
                job_number=job.job_number,
                part_number_job=job.part_number,
                part_number_quote=None,
                quote_item_number=None,
                description=None,
                material=None,
                cycle_time_sec=None,
                ops_labour=None,
                labour_rate=None,
                weight_g=None,
                annual_volume=None,
                match_status="job_without_quote",
                match_confidence="none",
            ))

    # Identify unlinked quotes
    unlinked_quotes = [
        q for q in quote_items
        if (q.quote_item_number or id(q)) not in used_quote_ids
    ]

    return LinkResult(
        total_jobs=len(jobs),
        matched_count=matched_count,
        unmatched_jobs_count=unmatched_jobs_count,
        unmatched_quotes_count=len(unlinked_quotes),
        linked_records=linked_records,
        unlinked_quotes=unlinked_quotes,
    )


def build_linked_dataframe(linked_records: List[LinkedRecord]) -> pd.DataFrame:
    """Transform linked records into a structured tabular DataFrame."""
    rows: List[Dict[str, Any]] = []
    for rec in linked_records:
        status_label = "🟢 Matched" if rec.match_status == "matched" else "🔴 Job without Quote"
        rows.append({
            "Job Number": rec.job_number,
            "Part Number": rec.part_number_job or (rec.part_number_quote or ""),
            "Quote Item": rec.quote_item_number or "",
            "Description": rec.description or "",
            "Material": rec.material or "",
            "Cycle Sec": rec.cycle_time_sec,
            "#Ops": rec.ops_labour,
            "Labour Rate ($CAD/Hr)": rec.labour_rate,
            "Weight (g)": rec.weight_g,
            "Ann. Volume (EAU)": rec.annual_volume,
            "Link Status": status_label,
            "Match Confidence": rec.match_confidence.upper(),
        })

    cols = [
        "Job Number",
        "Part Number",
        "Quote Item",
        "Description",
        "Material",
        "Cycle Sec",
        "#Ops",
        "Labour Rate ($CAD/Hr)",
        "Weight (g)",
        "Ann. Volume (EAU)",
        "Link Status",
        "Match Confidence",
    ]
    return pd.DataFrame(rows, columns=cols)


def build_unlinked_quotes_dataframe(unlinked_quotes: List[QuoteItem]) -> pd.DataFrame:
    """Transform unlinked quotes into a structured tabular DataFrame."""
    rows: List[Dict[str, Any]] = []
    for item in unlinked_quotes:
        rows.append({
            "Quote Item": item.quote_item_number or "",
            "Part Number": item.part_number or "",
            "Description": item.description or "",
            "Material": item.material or "",
            "Cycle Sec": item.cycle_time_sec,
            "#Ops": item.ops_labour,
            "Labour Rate ($CAD/Hr)": item.labour_rate,
            "Weight (g)": item.weight_g,
            "Ann. Volume (EAU)": item.annual_volume,
            "Confidence": item.confidence,
            "Link Status": "🟡 Quote without Job",
        })

    cols = [
        "Quote Item",
        "Part Number",
        "Description",
        "Material",
        "Cycle Sec",
        "#Ops",
        "Labour Rate ($CAD/Hr)",
        "Weight (g)",
        "Ann. Volume (EAU)",
        "Confidence",
        "Link Status",
    ]
    return pd.DataFrame(rows, columns=cols)


def export_linked_data_to_multi_sheet_excel(
    linked_df: pd.DataFrame,
    unlinked_quotes_df: pd.DataFrame,
) -> bytes:
    """
    Export multi-sheet Excel workbook:
    - Sheet 1: 'Linked Jobs' (Exact Job order, matched quote details, blanks for missing quotes)
    - Sheet 2: 'Quotes Missing Jobs' (Quote items without a corresponding Job)
    """
    try:
        import openpyxl
        wb = openpyxl.Workbook()
        # Sheet 1
        ws1 = wb.active
        ws1.title = "Linked Jobs"

        # Headers
        ws1.append(list(linked_df.columns))
        for _, row in linked_df.iterrows():
            row_vals = []
            for col in linked_df.columns:
                val = row[col]
                row_vals.append(val if pd.notna(val) else "")
            ws1.append(row_vals)

        # Sheet 2
        ws2 = wb.create_sheet(title="Quotes Missing Jobs")
        ws2.append(list(unlinked_quotes_df.columns))
        for _, row in unlinked_quotes_df.iterrows():
            row_vals = []
            for col in unlinked_quotes_df.columns:
                val = row[col]
                row_vals.append(val if pd.notna(val) else "")
            ws2.append(row_vals)

        buffer = io.BytesIO()
        wb.save(buffer)
        return buffer.getvalue()
    except Exception:
        # Fallback to pure XML single sheet exporter
        return export_dataframe_to_excel(linked_df, sheet_name="Linked Jobs")
