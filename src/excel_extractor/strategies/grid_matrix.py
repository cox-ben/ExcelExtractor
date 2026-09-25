"""
2D Tabular Grid Matrix extraction strategy.

Identifies 2D tabular row-column matrix headers (Item #, Part #, Cycle Time)
and extracts quote item rows at the column intersections until end-of-table
or subtotal/total rows.
"""

from __future__ import annotations

from dataclasses import dataclass
import functools
import re
from typing import Any, List, Optional, Tuple

from excel_extractor.config import ExtractorConfig
from excel_extractor.locator import (
    AnchorLocator,
    clean_cycle_time_value,
    clean_ops_value,
    clean_part_number_string,
    clean_text_value,
    clean_volume_value,
    clean_weight_value,
    extract_bracketed_drawing,
    is_drawing_anchor,
    normalize_text,
)
from excel_extractor.models import QuoteItem, SourceCellCoords
from excel_extractor.reader import Cell, Workbook, Worksheet, make_coord
from excel_extractor.strategies.base import BaseStrategy

QUOTE_COL_KEYWORDS = [
    "quote number",
    "quote #",
    "quote#",
    "quote no",
    "quote no.",
    "mpc rfq #",
    "mpc rfq#",
    "mpc rfq no",
    "mpc rfq",
    "quote item #",
    "quote item#",
    "quote item no",
    "quote item",
    "rfq #",
    "rfq#",
    "rfq no",
]

ITEM_COL_KEYWORDS = [
    "item #",
    "item#",
    "item no",
    "item no.",
    "item number",
    "line item",
    "line #",
    "line#",
    "line",
    "item",
    "pos #",
    "pos#",
    "pos",
]

TOOL_COL_KEYWORDS = [
    "tool #",
    "tool#",
    "tool no",
    "tool no.",
    "tool number",
    "tooling #",
    "tooling#",
    "tooling",
    "tool",
    "mold #",
    "mold#",
    "mold no",
    "die #",
    "die#",
]

PART_SPECIFIC_KEYWORDS = [
    "customer part #",
    "customer part#",
    "customer part number",
    "customer part",
    "customer p/n",
    "part number",
    "part no",
    "part no.",
    "part #",
    "part#",
    "p/n",
]

PART_GENERIC_KEYWORDS = [
    "part id",
    "part",
]

PART_COL_KEYWORDS = PART_SPECIFIC_KEYWORDS + PART_GENERIC_KEYWORDS


CYCLE_TIME_COL_KEYWORDS = [
    "cycle time (sec)",
    "cycle time (s)",
    "cycle time",
    "cycle sec",
    "cycle (sec)",
    "c/t (s)",
    "c/t (sec)",
    "c/t",
    "sec/shot",
    "molding cycle",
]

DESCRIPTION_COL_KEYWORDS = [
    "part description",
    "part desc",
    "part desc.",
    "part discription",
    "description",
    "discription",
    "desc",
    "desc.",
]

DESCRIPTION_EXCLUDE_KEYWORDS = [
    "resin",
    "material",
    "generic",
    "raw",
    "tool",
    "machine",
    "press",
    "supplier",
    "vendor",
    "grade",
]

MATERIAL_GRADE_KEYWORDS = [
    "resin grade",
    "material grade",
    "resin grade description",
    "grade",
]

MATERIAL_GENERIC_KEYWORDS = [
    "material",
    "resin",
    "raw material",
    "material description",
    "raw material description",
]

MATERIAL_EXCLUDE_KEYWORDS = [
    "supplier",
    "vendor",
    "distributor",
    "generic",
    "mfr",
    "manufacturer",
    "tool",
]

OPS_COL_KEYWORDS = [
    "#ops",
    "# ops",
    "ops",
    "labour",
    "labor",
    "operators",
    "operator",
    "# operators",
]

WEIGHT_COL_KEYWORDS = [
    "weight (g)",
    "weight(g)",
    "part weight (g)",
    "part wt (g)",
    "part weight",
    "weight",
    "wt (g)",
]

VOLUME_COL_KEYWORDS = [
    "ann. volume (eau)",
    "ann. volume",
    "annual volume (eau)",
    "annual volume",
    "volume (eau)",
    "volume",
    "eau",
    "ann volume",
]


# Precomputed sets and prefix/suffix tuples for high-performance classification
QUOTE_EXACT = set(QUOTE_COL_KEYWORDS)
QUOTE_PREFIXES = tuple(f"{kw} " for kw in QUOTE_COL_KEYWORDS)
QUOTE_SUFFIXES = tuple(f" {kw}" for kw in QUOTE_COL_KEYWORDS)

ITEM_EXACT = set(ITEM_COL_KEYWORDS)
ITEM_PREFIXES = tuple(f"{kw} " for kw in ITEM_COL_KEYWORDS)
ITEM_SUFFIXES = tuple(f" {kw}" for kw in ITEM_COL_KEYWORDS)

TOOL_EXACT = set(TOOL_COL_KEYWORDS)
TOOL_PREFIXES = tuple(f"{kw} " for kw in TOOL_COL_KEYWORDS)
TOOL_SUFFIXES = tuple(f" {kw}" for kw in TOOL_COL_KEYWORDS)

PART_SPECIFIC_EXACT = set(PART_SPECIFIC_KEYWORDS)
PART_SPECIFIC_PREFIXES = tuple(f"{kw} " for kw in PART_SPECIFIC_KEYWORDS)
PART_SPECIFIC_SUFFIXES = tuple(f" {kw}" for kw in PART_SPECIFIC_KEYWORDS)

PART_GENERIC_EXACT = set(PART_GENERIC_KEYWORDS)
PART_GENERIC_PREFIXES = tuple(f"{kw} " for kw in PART_GENERIC_KEYWORDS)
PART_GENERIC_SUFFIXES = tuple(f" {kw}" for kw in PART_GENERIC_KEYWORDS)

CYCLE_TIME_EXACT = set(CYCLE_TIME_COL_KEYWORDS)
CYCLE_TIME_PREFIXES = tuple(f"{kw} " for kw in CYCLE_TIME_COL_KEYWORDS)
CYCLE_TIME_SUFFIXES = tuple(f" {kw}" for kw in CYCLE_TIME_COL_KEYWORDS)

DESCRIPTION_EXACT = set(DESCRIPTION_COL_KEYWORDS)
DESCRIPTION_PREFIXES = tuple(f"{kw} " for kw in DESCRIPTION_COL_KEYWORDS)
DESCRIPTION_SUFFIXES = tuple(f" {kw}" for kw in DESCRIPTION_COL_KEYWORDS)

MATERIAL_GRADE_EXACT = set(MATERIAL_GRADE_KEYWORDS)
MATERIAL_GRADE_PREFIXES = tuple(f"{kw} " for kw in MATERIAL_GRADE_KEYWORDS)
MATERIAL_GRADE_SUFFIXES = tuple(f" {kw}" for kw in MATERIAL_GRADE_KEYWORDS)

MATERIAL_GENERIC_EXACT = set(MATERIAL_GENERIC_KEYWORDS)
MATERIAL_GENERIC_PREFIXES = tuple(f"{kw} " for kw in MATERIAL_GENERIC_KEYWORDS)
MATERIAL_GENERIC_SUFFIXES = tuple(f" {kw}" for kw in MATERIAL_GENERIC_KEYWORDS)

OPS_EXACT = set(OPS_COL_KEYWORDS)
OPS_PREFIXES = tuple(f"{kw} " for kw in OPS_COL_KEYWORDS)
OPS_SUFFIXES = tuple(f" {kw}" for kw in OPS_COL_KEYWORDS)

WEIGHT_EXACT = set(WEIGHT_COL_KEYWORDS)
WEIGHT_PREFIXES = tuple(f"{kw} " for kw in WEIGHT_COL_KEYWORDS)
WEIGHT_SUFFIXES = tuple(f" {kw}" for kw in WEIGHT_COL_KEYWORDS)

VOLUME_EXACT = set(VOLUME_COL_KEYWORDS)
VOLUME_PREFIXES = tuple(f"{kw} " for kw in VOLUME_COL_KEYWORDS)
VOLUME_SUFFIXES = tuple(f" {kw}" for kw in VOLUME_COL_KEYWORDS)


@functools.lru_cache(maxsize=4096)
def is_item_header(norm: str) -> bool:
    """Checks if header represents an item/line number column."""
    if not norm:
        return False
    return norm in ITEM_EXACT or norm.startswith(ITEM_PREFIXES) or norm.endswith(ITEM_SUFFIXES)


@functools.lru_cache(maxsize=4096)
def is_tool_header(norm: str) -> bool:
    """Checks if header represents a tool/mold/die column."""
    if not norm:
        return False
    if norm in TOOL_EXACT or norm.startswith(TOOL_PREFIXES) or norm.endswith(TOOL_SUFFIXES):
        return True
    return any(kw in norm for kw in TOOL_COL_KEYWORDS)


PART_HEADER_EXCLUDE_KEYWORDS = (
    "review",
    "evaluation",
    "evaluate",
    "assessment",
    "checklist",
    "feasibility",
    "complexity",
    "criteria",
    "status",
    "question",
    "questions",
    "fit",
    "awarded",
    "package",
    "complete",
    "priority",
)


@functools.lru_cache(maxsize=4096)
def is_part_specific_header(norm: str) -> bool:
    """Checks if header represents a specific part number column (e.g. 'Part Number', 'Part #', 'Customer Part #')."""
    if not norm:
        return False
    if is_item_header(norm) or is_tool_header(norm) or is_drawing_anchor(norm) or is_description_header(norm):
        return False
    if any(ex in norm for ex in PART_HEADER_EXCLUDE_KEYWORDS):
        return False
    return norm in PART_SPECIFIC_EXACT or norm.startswith(PART_SPECIFIC_PREFIXES) or norm.endswith(PART_SPECIFIC_SUFFIXES)


@functools.lru_cache(maxsize=4096)
def is_part_generic_header(norm: str) -> bool:
    """Checks if header represents a generic part column (e.g. 'PART ID', 'Part')."""
    if not norm:
        return False
    if is_item_header(norm) or is_tool_header(norm) or is_drawing_anchor(norm) or is_description_header(norm):
        return False
    if any(ex in norm for ex in PART_HEADER_EXCLUDE_KEYWORDS):
        return False
    return norm in PART_GENERIC_EXACT or norm.startswith(PART_GENERIC_PREFIXES) or norm.endswith(PART_GENERIC_SUFFIXES)


@functools.lru_cache(maxsize=4096)
def is_description_header(norm: str) -> bool:
    """Checks if header represents part description, strictly rejecting resin/material/generic descriptions."""
    if not norm:
        return False
    if any(ex in norm for ex in DESCRIPTION_EXCLUDE_KEYWORDS):
        return False
    return norm in DESCRIPTION_EXACT or norm.startswith(DESCRIPTION_PREFIXES) or norm.endswith(DESCRIPTION_SUFFIXES)


@functools.lru_cache(maxsize=4096)
def is_material_grade_header(norm: str) -> bool:
    """Checks if header represents specific Resin Grade / Material Grade."""
    if not norm:
        return False
    if any(ex in norm for ex in MATERIAL_EXCLUDE_KEYWORDS):
        return False
    return norm in MATERIAL_GRADE_EXACT or norm.startswith(MATERIAL_GRADE_PREFIXES) or norm.endswith(MATERIAL_GRADE_SUFFIXES)


@functools.lru_cache(maxsize=4096)
def is_material_generic_header(norm: str) -> bool:
    """Checks if header represents generic material / resin, excluding supplier and generic descriptions."""
    if not norm:
        return False
    if any(ex in norm for ex in MATERIAL_EXCLUDE_KEYWORDS):
        return False
    return norm in MATERIAL_GENERIC_EXACT or norm.startswith(MATERIAL_GENERIC_PREFIXES) or norm.endswith(MATERIAL_GENERIC_SUFFIXES)


@functools.lru_cache(maxsize=4096)
def is_cycle_time_header(norm: str) -> bool:
    """Checks if header represents cycle time."""
    if not norm:
        return False
    return norm in CYCLE_TIME_EXACT or norm.startswith(CYCLE_TIME_PREFIXES) or norm.endswith(CYCLE_TIME_SUFFIXES)


@functools.lru_cache(maxsize=4096)
def is_quote_header(norm: str) -> bool:
    """Checks if header represents quote number column."""
    if not norm:
        return False
    if norm in QUOTE_EXACT or norm.startswith(QUOTE_PREFIXES) or norm.endswith(QUOTE_SUFFIXES):
        return True
    return bool(re.search(r"\bQ[0-9]{4,5}\b", norm))


@functools.lru_cache(maxsize=4096)
def is_ops_header(norm: str) -> bool:
    """Checks if header represents #ops/operators column."""
    if not norm:
        return False
    return norm in OPS_EXACT or norm.startswith(OPS_PREFIXES) or norm.endswith(OPS_SUFFIXES)


@functools.lru_cache(maxsize=4096)
def is_weight_header(norm: str) -> bool:
    """Checks if header represents part weight column."""
    if not norm:
        return False
    return norm in WEIGHT_EXACT or norm.startswith(WEIGHT_PREFIXES) or norm.endswith(WEIGHT_SUFFIXES)


@functools.lru_cache(maxsize=4096)
def is_volume_header(norm: str) -> bool:
    """Checks if header represents annual volume column."""
    if not norm:
        return False
    return norm in VOLUME_EXACT or norm.startswith(VOLUME_PREFIXES) or norm.endswith(VOLUME_SUFFIXES)


def is_sheet_row_hidden(ws: Any, row: int) -> bool:
    """Safely check if a row is hidden for both reader.Worksheet and openpyxl Worksheet."""
    if hasattr(ws, "is_row_hidden"):
        return ws.is_row_hidden(row)
    if hasattr(ws, "row_dimensions") and row in ws.row_dimensions:
        return bool(ws.row_dimensions[row].hidden)
    return False


def is_column_index_row(row_cells: List[Any]) -> bool:
    """
    Checks if a row consists purely of column index numbers (e.g. 1, 2, 3, 4, 5 or 2, 3, 4, 5).
    Handles integers, floats (e.g. 1.0, 2.0), string digits, and sequential series.
    """
    non_empty = [c for c in row_cells if c.value is not None and str(c.value).strip() != ""]
    if not non_empty or len(non_empty) < 2:
        return False

    int_pairs: List[Tuple[int, int]] = []  # (col_idx, int_val)
    for c in non_empty:
        v = c.value
        val_int: Optional[int] = None
        if isinstance(v, (int, float)):
            try:
                f = float(v)
                if f.is_integer() and 0 <= f <= 100:
                    val_int = int(f)
            except (ValueError, OverflowError):
                pass
        else:
            s = str(v).strip()
            try:
                f = float(s)
                if f.is_integer() and 0 <= f <= 100:
                    val_int = int(f)
            except (ValueError, OverflowError):
                pass

        if val_int is None:
            # If col 1 has a header label like 'Col #' or 'No.', allow it; otherwise non-integer cell means not index row
            norm = normalize_text(v)
            if c.column == 1 and norm in ("col #", "col no", "col", "column", "no.", "no", "#"):
                continue
            return False

        int_pairs.append((c.column, val_int))

    if len(int_pairs) < 2:
        return False

    # Condition 1: Values match column indices (e.g. col 2 has 2, col 3 has 3...)
    matching_col_indices = sum(1 for col_idx, val in int_pairs if val == col_idx)
    if matching_col_indices >= 2:
        return True

    # Condition 2: Values form a sequential increasing sequence (e.g. 1, 2, 3, 4 across columns)
    int_pairs.sort(key=lambda x: x[0])
    vals = [val for _, val in int_pairs]
    is_sequential = all(vals[i + 1] == vals[i] + 1 for i in range(len(vals) - 1))
    if is_sequential and vals[0] in (0, 1, 2):
        return True

    # Condition 3: At least 3 values and all consecutive step differences are 1
    diffs = [vals[i + 1] - vals[i] for i in range(len(vals) - 1)]
    if len(diffs) >= 2 and all(d == 1 for d in diffs):
        return True

    return False


@dataclass
class GridHeaderInfo:
    """Stores grid table header row and identified column indices."""
    row: int
    col_item: int
    col_part: int
    col_ct: int
    col_desc: Optional[int] = None
    col_mat: Optional[int] = None
    col_ops: Optional[int] = None
    col_weight: Optional[int] = None
    col_vol: Optional[int] = None
    col_quote: Optional[int] = None
    col_drawing: Optional[int] = None

    def __iter__(self):
        # Backward compatibility for tuple unpacking: (header_row, col_item, col_part, col_ct)
        return iter((self.row, self.col_item, self.col_part, self.col_ct))


@functools.lru_cache(maxsize=4096)
def quote_number_has_item_suffix(quote_str: str) -> bool:
    """
    Checks if a quote string already contains an explicit item suffix.
    Examples with item suffix (returns True):
      'Q3428-1', 'Q3428-01', 'Q3428-2', 'Q4378-4 (2140)', 'Q4231-R15-7',
      'Q1001.1', 'Q1001_1', '3428-1'
    Examples missing item suffix (returns False):
      'Q3428', 'Q3428-R3', 'Q3428-Rev3', 'Q4355', '3428', 'Q1001', 'None', '', '0'
    """
    if not quote_str:
        return False
    s_clean = re.sub(r"\s*[\(\[].*?[\)\]]", "", quote_str).strip()
    # Matches quote ID with item suffix, allowing optional revision e.g. -R15-7, -1, -01, -A
    # But strictly NOT ending with just revision like -R3 or -Rev3
    if re.search(r"\bQ?[0-9]{4,5}(?:[-._]r(?:ev)?[0-9]+)?[-._](?!r(?:ev)?[0-9]+$)[A-Za-z0-9]+\b", s_clean, re.IGNORECASE):
        return True
    return False


@functools.lru_cache(maxsize=4096)
def extract_base_quote_id(quote_str: str, fallback_quote_id: Optional[str] = None) -> str:
    """Extracts bare quote ID e.g. 'Q3428' from a string or fallback."""
    if quote_str:
        m = re.search(r"\b(Q[0-9]{4,5})\b", quote_str, re.IGNORECASE)
        if m:
            return m.group(1).upper()
        m_no_q = re.search(r"\b([0-9]{4,5})\b", quote_str)
        if m_no_q:
            return f"Q{m_no_q.group(1)}"
    if fallback_quote_id:
        m_fb = re.search(r"\b(Q[0-9]{4,5})\b", fallback_quote_id, re.IGNORECASE)
        if m_fb:
            return m_fb.group(1).upper()
        return fallback_quote_id.upper()
    return "Q0000"


@functools.lru_cache(maxsize=4096)
def extract_item_suffix(raw_item_val: Any) -> Optional[str]:
    """Cleans an item number cell value (e.g. 1, 1.0, '2', '01', '#1', 'Item 1') into '1', '2', etc."""
    if raw_item_val is None:
        return None
    if isinstance(raw_item_val, float):
        if raw_item_val.is_integer():
            return str(int(raw_item_val))
        return str(raw_item_val)
    if isinstance(raw_item_val, int):
        return str(raw_item_val)

    it_s = str(raw_item_val).strip()
    if not it_s or it_s in ("0", "0.0", "-", "--", "None", "null", "N/A"):
        return None

    # Check if raw_item_val itself is a full quote item with item suffix like Q3428-1
    if quote_number_has_item_suffix(it_s):
        m_full = re.search(r"\b(Q[0-9]{4,5}(?:[-._]r(?:ev)?[0-9]+)?(?:[-._][A-Za-z0-9]+)+)\b", it_s, re.IGNORECASE)
        if m_full:
            return m_full.group(1).upper().replace(".", "-").replace("_", "-")

    # Strip parenthetical notes like (2140) if present
    it_s_no_parens = re.sub(r"\s*[\(\[].*?[\)\]]", "", it_s).strip()

    # Strip prefixes like 'Item', 'Line', '#', 'Pos'
    clean = re.sub(r"(?i)^(?:item|line|pos)[\s#:]*", "", it_s_no_parens).strip(" \t\r\n:;-#")
    if clean and clean not in ("0", "0.0", "-", "--"):
        return clean
    return None


def resolve_quote_item_number(
    raw_quote_val: Any,
    raw_item_val: Any,
    quote_id: Optional[str] = None,
    item_seq: int = 1,
) -> Tuple[str, List[str]]:
    """
    Resolves canonical quote item number (e.g. 'Q3428-1', 'Q3428-2')
    supporting cases where Quote Number and Item Number are in separate columns,
    or where Quote Number already contains the item suffix.

    1. First, checks if raw_quote_val ALREADY contains an item suffix (e.g. 'Q3428-1', 'Q3428-01', 'Q4378-4 (2140)').
       If so, raw_quote_val is NOT missing an item number -> use it directly (never append raw_item_val again).
    2. If raw_quote_val is missing an item number (e.g. bare 'Q3428', 'Q3428-R3', '3428', or None):
       - If raw_item_val is present (e.g. 1, 2, '01', 'A'):
         Combine: f"{base_quote}-{clean_item_val}".
       - If raw_item_val is missing:
         Fallback to sequential item suffix: f"{base_quote}-{item_seq}".
    """
    warnings: List[str] = []
    quote_str = str(raw_quote_val).strip() if raw_quote_val is not None else ""

    # 1. Does raw_quote_val ALREADY contain an explicit item suffix?
    if quote_number_has_item_suffix(quote_str):
        m = re.search(r"\b(Q[0-9]{4,5}(?:[-._]r(?:ev)?[0-9]+)?(?:[-._][A-Za-z0-9]+)+)\b", quote_str, re.IGNORECASE)
        if m:
            return m.group(1).upper().replace(".", "-").replace("_", "-"), warnings
        m_no_q = re.search(r"\b([0-9]{4,5}(?:[-._]r(?:ev)?[0-9]+)?(?:[-._][A-Za-z0-9]+)+)\b", quote_str)
        if m_no_q:
            return f"Q{m_no_q.group(1).upper().replace('.', '-').replace('_', '-')}", warnings
        clean_q = re.sub(r"\s*[\(\[].*?[\)\]]", "", quote_str).strip()
        return clean_q.upper(), warnings

    # 2. Quote number is missing an item number! (e.g. bare "Q3428", "Q3428-R3", "3428", or None)
    base_quote = extract_base_quote_id(quote_str, fallback_quote_id=quote_id)

    # Check if raw_item_val provides the missing item number
    clean_item = extract_item_suffix(raw_item_val)

    # If clean_item is just the base quote number itself (e.g. 'Q3428' or '3428' or 'Q3428-R3'), it is NOT an item suffix
    if clean_item and (
        clean_item.upper() == base_quote
        or clean_item.upper() == base_quote.lstrip("Q")
        or re.match(r"^Q?[0-9]{4,5}(?:[-._]r(?:ev)?[0-9]+)?$", clean_item, re.IGNORECASE)
    ):
        clean_item = None

    if clean_item:
        if quote_number_has_item_suffix(clean_item):
            m = re.search(r"\b(Q[0-9]{4,5}(?:[-._]r(?:ev)?[0-9]+)?(?:[-._][A-Za-z0-9]+)+)\b", clean_item, re.IGNORECASE)
            if m:
                return m.group(1).upper().replace(".", "-").replace("_", "-"), warnings
            return clean_item.upper(), warnings
        clean_item_no_dash = clean_item.lstrip("-")
        result = f"{base_quote}-{clean_item_no_dash}"
        return result, warnings

    # 3. No separate item number cell available -> infer sequential item suffix
    item_num = f"{base_quote}-{item_seq}"
    warnings.append(f"Inferred quote item number suffix '-{item_seq}' from base quote {base_quote}")

    # Guaranteed invariant: A quote item number MUST contain an item suffix
    if not quote_number_has_item_suffix(item_num):
        item_num = f"{base_quote}-{item_seq}"

    return item_num, warnings



class GridMatrixStrategy(BaseStrategy):
    """Strategy for 2D tabular row-column grid matrix quote layouts."""

    def light_scan_table_candidates(self, ws: Worksheet, max_scan_row: int = 250) -> List[int]:
        """
        Tier 1 fast light sweep down the sheet (from row 1 down to max_scan_row).
        Rapidly samples non-empty row text to locate candidate table header bands
        before deep evaluation.
        """
        limit_row = min(ws.max_row + 1, max_scan_row)
        max_scan_col = min(ws.max_column + 1, 40)
        candidate_rows: List[int] = []
        cells_dict = getattr(ws, "_cells", None)

        if cells_dict is not None:
            row_cells_map: Dict[int, List[Cell]] = {}
            for (r, c), cell in cells_dict.items():
                if 1 <= r < limit_row and 1 <= c < max_scan_col and cell.value is not None:
                    row_cells_map.setdefault(r, []).append(cell)

            for r in sorted(row_cells_map.keys()):
                if is_sheet_row_hidden(ws, r):
                    continue
                row_cells = row_cells_map[r]
                if len(row_cells) < 2:
                    continue
                if is_column_index_row(row_cells):
                    continue

                has_part = has_quote = has_item = has_ct = has_desc = has_mat = has_vol = has_ops = has_weight = False
                non_empty_count = 0

                for cell in row_cells:
                    s = str(cell.value).strip()
                    if not s:
                        continue
                    non_empty_count += 1
                    norm = normalize_text(s)

                    if is_tool_header(norm):
                        continue

                    if not has_ct and is_cycle_time_header(norm):
                        has_ct = True
                    elif not has_desc and is_description_header(norm):
                        has_desc = True
                    elif not has_mat and (is_material_grade_header(norm) or is_material_generic_header(norm)):
                        has_mat = True
                    elif not has_ops and is_ops_header(norm):
                        has_ops = True
                    elif not has_weight and is_weight_header(norm):
                        has_weight = True
                    elif not has_vol and is_volume_header(norm):
                        has_vol = True
                    elif not has_part and (is_part_specific_header(norm) or is_part_generic_header(norm)):
                        has_part = True
                    elif not has_quote and is_quote_header(norm):
                        has_quote = True
                    elif not has_item and is_item_header(norm):
                        has_item = True

                if non_empty_count < 2:
                    continue

                cat_count = sum([has_part, has_quote, has_item, has_ct, has_desc, has_mat, has_vol, has_ops, has_weight])
                if (has_part and (has_quote or has_item or has_ct or has_desc or has_mat or has_vol or has_ops or has_weight)) \
                   or (has_quote and (has_ct or has_desc or has_mat or has_vol or has_item)) \
                   or (cat_count >= 2 and (has_part or has_quote or has_item or has_ct)):
                    candidate_rows.append(r)

            if not candidate_rows:
                candidate_rows = list(range(1, min(ws.max_row + 1, 50)))
            return candidate_rows

        return candidate_rows

    def find_grid_header(self, ws: Worksheet) -> Optional[GridHeaderInfo]:
        """
        Two-tier search engine:
        1. Fast light scan sweeps down the sheet to find candidate table header rows.
        2. Targeted deep search evaluates candidate rows based on recognized column types
           and lookahead verification of valid data rows below.

        Prioritizes 'Quote Number' / 'MPC RFQ #' over generic 'Item #' or 'Tool #'.
        Strictly prioritizes 'Resin Grade' over 'Resin Supplier' or 'Generic Resin Description'.
        Strictly selects Part Description (including 'Desc') while excluding resin descriptions.

        Returns highest scoring GridHeaderInfo or None.
        """
        if hasattr(ws, "_cached_grid_header"):
            return ws._cached_grid_header

        scan_limit = max(150, min(ws.max_row + 1, max(45, self.config.grid_matrix.header_max_row + 1)))
        candidate_rows = self.light_scan_table_candidates(ws, max_scan_row=scan_limit)
        max_scan_col = min(ws.max_column + 1, 60)
        stop_kws_tuple = tuple(s.lower() for s in self.config.grid_matrix.stop_keywords)
        cells_dict = getattr(ws, "_cells", None)

        candidates: List[Tuple[float, GridHeaderInfo]] = []

        for r in candidate_rows:
            col_quote: Optional[int] = None
            col_item: Optional[int] = None
            col_part_specific: Optional[int] = None
            col_part_generic: Optional[int] = None
            col_ct: Optional[int] = None
            col_desc: Optional[int] = None
            col_mat_grade: Optional[int] = None
            col_mat_generic: Optional[int] = None
            col_ops: Optional[int] = None
            col_weight: Optional[int] = None
            col_vol: Optional[int] = None
            col_drawing: Optional[int] = None

            for c in range(1, max_scan_col):
                if cells_dict is not None:
                    cell = cells_dict.get((r, c))
                    if cell is None or cell.value is None:
                        continue
                else:
                    cell = ws.cell(row=r, column=c)
                    if cell.value is None:
                        continue

                norm = normalize_text(cell.value)
                if not norm:
                    continue

                # Strictly reject tool columns from being matched as quote item or part
                if is_tool_header(norm):
                    continue

                # 1. Check Cycle Time
                if col_ct is None and is_cycle_time_header(norm):
                    col_ct = c
                    continue

                # 2. Check Description (strictly excluding resin/material/generic descriptions)
                if col_desc is None and is_description_header(norm):
                    col_desc = c
                    continue

                # 3. Check Material Grade (Priority 1 for material)
                if col_mat_grade is None and is_material_grade_header(norm):
                    col_mat_grade = c
                    continue

                # 4. Check Generic Material (Priority 2 for material)
                if col_mat_generic is None and is_material_generic_header(norm):
                    col_mat_generic = c
                    continue

                # 5. Check #Ops / Labour
                if col_ops is None and is_ops_header(norm):
                    col_ops = c
                    continue

                # 6. Check Weight
                if col_weight is None and is_weight_header(norm):
                    col_weight = c
                    continue

                # 7. Check Volume
                if col_vol is None and is_volume_header(norm):
                    col_vol = c
                    continue

                # 8. Check Part Number (specific vs generic)
                if col_part_specific is None and is_part_specific_header(norm):
                    col_part_specific = c
                    continue

                if col_part_generic is None and is_part_generic_header(norm):
                    col_part_generic = c
                    continue

                # 9. Check Quote Number / MPC RFQ # (Priority 1 for Item identifier)
                if col_quote is None and is_quote_header(norm):
                    col_quote = c
                    continue

                # 10. Check Item Number / Line # (Priority 2 for Item identifier)
                if col_item is None and is_item_header(norm):
                    col_item = c
                    continue

                # 11. Check Drawing Number column
                if col_drawing is None and is_drawing_anchor(norm):
                    col_drawing = c
                    continue

            # Prioritize specific part column over generic
            col_part = col_part_specific if col_part_specific is not None else col_part_generic

            # A valid grid header must contain Part # and at least one of (Cycle Time, Quote Number, Item #)
            if col_part is not None and (col_ct is not None or col_quote is not None or col_item is not None):
                # Select material column (prefer grade over generic)
                col_mat = col_mat_grade if col_mat_grade is not None else col_mat_generic

                # Select item and quote columns
                # If Item# column is present, col_item is Item#; if Quote Number column is present, col_quote is Quote Number
                if col_item is not None:
                    effective_col_item = col_item
                elif col_quote is not None:
                    effective_col_item = col_quote
                else:
                    # Only fallback to col 1 if col 1 is NOT tool and NOT drawing
                    c1_val = normalize_text(ws.cell(row=r, column=1).value)
                    if c1_val and not is_tool_header(c1_val) and not is_drawing_anchor(c1_val) and col_part != 1 and (col_ct is None or col_ct != 1):
                        effective_col_item = 1
                    else:
                        effective_col_item = 0

                header_info = GridHeaderInfo(
                    row=r,
                    col_item=effective_col_item,
                    col_part=col_part,
                    col_ct=col_ct if col_ct is not None else 0,
                    col_desc=col_desc,
                    col_mat=col_mat,
                    col_ops=col_ops,
                    col_weight=col_weight,
                    col_vol=col_vol,
                    col_quote=col_quote,
                    col_drawing=col_drawing,
                )

                # Score the candidate header
                score = 0.0
                if col_part_specific is not None:
                    score += 35.0
                elif col_part_generic is not None:
                    score += 10.0

                if col_quote is not None:
                    score += 25.0
                if col_item is not None:
                    score += 15.0
                if col_ct is not None:
                    score += 25.0
                if col_desc is not None:
                    score += 15.0
                if col_mat is not None:
                    score += 15.0
                if col_ops is not None:
                    score += 10.0
                if col_weight is not None:
                    score += 10.0
                if col_vol is not None:
                    score += 15.0
                if col_drawing is not None:
                    score += 5.0

                # Data rows lookahead verification (check up to 25 rows below r)
                data_rows_found = 0
                consecutive_blank = 0
                max_check_row = min(ws.max_row + 1, r + 26)
                for dr in range(r + 1, max_check_row):
                    if is_sheet_row_hidden(ws, dr):
                        continue
                    dr_cells = [ws.cell(row=dr, column=c) for c in range(1, min(ws.max_column + 1, 60))]
                    dr_vals = [c.value for c in dr_cells]
                    if not any(v is not None and str(v).strip() != "" for v in dr_vals):
                        consecutive_blank += 1
                        if (data_rows_found > 0 and consecutive_blank >= 2) or consecutive_blank >= 4:
                            break
                        continue
                    if is_column_index_row(dr_cells):
                        consecutive_blank = 0
                        continue
                    if self._row_matches_stop_keywords(dr_vals, stop_kws_tuple):
                        break

                    consecutive_blank = 0
                    pn_val = clean_part_number_string(ws.cell(row=dr, column=col_part).value)
                    qt_val = ws.cell(row=dr, column=col_quote).value if col_quote else None
                    ct_val = clean_cycle_time_value(ws.cell(row=dr, column=col_ct).value) if col_ct else None

                    if (pn_val and len(pn_val) >= 2) or (qt_val and re.search(r"Q[0-9]{4,5}", str(qt_val))) or ct_val is not None:
                        data_rows_found += 1
                        if data_rows_found >= 10:
                            break

                score += data_rows_found * 50.0
                candidates.append((score, header_info))

        if not candidates:
            ws._cached_grid_header = None
            return None

        # Sort candidate headers by score descending (highest score first)
        candidates.sort(key=lambda x: x[0], reverse=True)
        result = candidates[0][1]
        ws._cached_grid_header = result
        return result

    def sheet_has_grid(self, ws: Worksheet) -> bool:
        """Check if worksheet contains a recognizable 2D grid matrix table."""
        return self.find_grid_header(ws) is not None

    def can_handle(self, workbook: Workbook) -> bool:
        """Returns True if any candidate quoting worksheet contains a 2D grid matrix."""
        if not self.config.grid_matrix.enabled:
            return False

        candidate_sheets = self.get_candidate_sheets(workbook)
        for ws in candidate_sheets:
            if self.sheet_has_grid(ws):
                return True
        return False

    @staticmethod
    @functools.lru_cache(maxsize=4096)
    def _matches_stop_keyword_cached(norm: str, stop_kws_tuple: Tuple[str, ...]) -> bool:
        if not norm:
            return False

        norm_clean = norm.strip("*_# :;-")
        norm_no_dash = norm_clean.replace("-", "")

        for kw_clean in stop_kws_tuple:
            if not kw_clean:
                continue

            # Exact match
            if norm_clean == kw_clean or norm_no_dash == kw_clean:
                return True

            # Plural/singular variants
            if kw_clean == "total" and norm_clean in ("totals", "total"):
                return True
            if kw_clean == "notes" and norm_clean in ("note", "notes"):
                return True

            # Prefix matches
            if (
                norm_clean.startswith(f"{kw_clean} ")
                or norm_clean.startswith(f"{kw_clean}:")
                or norm_clean.startswith(f"{kw_clean}-")
                or norm_no_dash.startswith(f"{kw_clean} ")
            ):
                return True
            if kw_clean == "notes" and (norm_clean.startswith("note ") or norm_clean.startswith("note:")):
                return True

            # Suffix matches (e.g. "grand total", "project total")
            if norm_clean.endswith(f" {kw_clean}") or norm_no_dash.endswith(f" {kw_clean}"):
                return True
            if kw_clean == "total" and "grand total" in norm_clean:
                return True

        return False

    @classmethod
    def _matches_stop_keyword(cls, norm: str, stop_kws: Union[List[str], Tuple[str, ...]]) -> bool:
        """Check if normalized text matches any stop keyword."""
        stop_kws_tuple = tuple(stop_kws) if not isinstance(stop_kws, tuple) else stop_kws
        return cls._matches_stop_keyword_cached(norm, stop_kws_tuple)

    @classmethod
    def _row_matches_stop_keywords(cls, row_values: List[Any], stop_kws: Union[List[str], Tuple[str, ...]]) -> bool:
        """Check if any cell across the row matches stop/footer keywords."""
        stop_kws_tuple = tuple(stop_kws) if not isinstance(stop_kws, tuple) else stop_kws
        for val in row_values:
            if val is None:
                continue
            norm = normalize_text(val)
            if norm and cls._matches_stop_keyword_cached(norm, stop_kws_tuple):
                return True
        return False

    def extract(self, workbook: Workbook, quote_id: Optional[str] = None) -> List[QuoteItem]:
        """Extract items from all candidate grid matrix sheets in the workbook."""
        items: List[QuoteItem] = []
        stop_kws = tuple(s.lower() for s in self.config.grid_matrix.stop_keywords)

        # Scan for global base quote ID if not passed
        locator = AnchorLocator(self.config)
        if not quote_id:
            for ws in workbook.worksheets:
                if not self.is_sheet_visible(ws):
                    continue
                qid, _, _ = locator.find_quote_id(ws)
                if qid:
                    m = re.search(r"\b(Q[0-9]{4,5})\b", qid)
                    if m:
                        quote_id = m.group(1).upper()
                        break

        candidate_sheets = self.get_candidate_sheets(workbook)
        # If candidate sheets don't contain any grid header, check best_qinfo or best_mpc
        if not any(self.find_grid_header(ws) is not None for ws in candidate_sheets):
            best_qinfo = self.get_closest_qinfo_sheet(workbook)
            if best_qinfo and self.find_grid_header(best_qinfo) is not None:
                candidate_sheets = [best_qinfo]
            else:
                best_mpc = self.get_closest_mpc_sheet(workbook)
                if best_mpc and self.find_grid_header(best_mpc) is not None:
                    candidate_sheets = [best_mpc]

        for ws in candidate_sheets:
            header_info = self.find_grid_header(ws)
            if not header_info:
                continue

            # Check if table has multi-field columns
            has_multi_fields = any([
                header_info.col_desc is not None,
                header_info.col_mat is not None,
                header_info.col_ops is not None,
                header_info.col_weight is not None,
                header_info.col_vol is not None,
            ])

            # Resolve static labour rate for this sheet (directly below Rate $CAD/Hr)
            static_labour_rate, lr_coord = locator.find_labour_rate(ws)
            if static_labour_rate is None:
                # Fallback to search other visible sheets in the workbook (especially MPC sheets)
                best_mpc_ws = self.get_closest_mpc_sheet(workbook)
                if best_mpc_ws and best_mpc_ws != ws:
                    lr_val, lr_c = locator.find_labour_rate(best_mpc_ws)
                    if lr_val is not None:
                        static_labour_rate = lr_val
                        lr_coord = f"{best_mpc_ws.title}!{lr_c}" if len(workbook.worksheets) > 1 else lr_c

                if static_labour_rate is None:
                    for other_ws in workbook.worksheets:
                        if other_ws != ws and other_ws != best_mpc_ws and self.is_sheet_visible(other_ws) and not self.is_ignored_sheet(other_ws.title):
                            lr_val, lr_c = locator.find_labour_rate(other_ws)
                            if lr_val is not None:
                                static_labour_rate = lr_val
                                lr_coord = f"{other_ws.title}!{lr_c}" if len(workbook.worksheets) > 1 else lr_c
                                break

            # Build QInfo validation lookup from closest QInfo sheet (if different from primary sheet)
            closest_qinfo_ws = self.get_closest_qinfo_sheet(workbook)
            qinfo_pn_by_item: dict[str, str] = {}
            qinfo_pn_by_idx: dict[int, str] = {}
            qinfo_item_by_pn: dict[str, str] = {}
            qinfo_item_by_idx: dict[int, str] = {}
            if closest_qinfo_ws and closest_qinfo_ws != ws and self.sheet_has_grid(closest_qinfo_ws):
                q_hdr = self.find_grid_header(closest_qinfo_ws)
                if q_hdr and q_hdr.col_part:
                    q_idx = 1
                    for qr in range(q_hdr.row + 1, closest_qinfo_ws.max_row + 1):
                        if is_sheet_row_hidden(closest_qinfo_ws, qr):
                            continue
                        q_pn_raw = closest_qinfo_ws.cell(row=qr, column=q_hdr.col_part).value
                        q_pn = clean_part_number_string(q_pn_raw)
                        if not q_pn:
                            continue
                        qinfo_pn_by_idx[q_idx] = q_pn
                        q_it_cell = closest_qinfo_ws.cell(row=qr, column=q_hdr.col_item) if q_hdr.col_item > 0 else None
                        q_qt_cell = closest_qinfo_ws.cell(row=qr, column=q_hdr.col_quote) if q_hdr.col_quote and q_hdr.col_quote > 0 else None
                        if q_it_cell or q_qt_cell:
                            q_resolved_num, _ = resolve_quote_item_number(
                                raw_quote_val=q_qt_cell.value if q_qt_cell else None,
                                raw_item_val=q_it_cell.value if q_it_cell else None,
                                quote_id=quote_id,
                                item_seq=q_idx,
                            )
                            qinfo_pn_by_item[q_resolved_num] = q_pn
                            qinfo_item_by_pn[q_pn] = q_resolved_num
                            qinfo_item_by_idx[q_idx] = q_resolved_num
                        q_idx += 1


            # Row traversal starting below header
            item_seq = 1
            sheet_items_count = 0
            consecutive_empty = 0
            gap_rows_skipped = 0
            cells_dict = getattr(ws, "_cells", None)
            max_c_bound = min(ws.max_column + 1, 60)

            for r in range(header_info.row + 1, ws.max_row + 1):
                # Disregard hidden rows
                if is_sheet_row_hidden(ws, r):
                    continue

                if cells_dict is not None:
                    row_values = [(cells_dict.get((r, c)).value if (r, c) in cells_dict else None) for c in range(1, max_c_bound)]
                else:
                    row_cells_temp = [ws.cell(row=r, column=c) for c in range(1, max_c_bound)]
                    row_values = [c.value for c in row_cells_temp]

                # Stop if any cell across the row matches stop keywords
                if self._row_matches_stop_keywords(row_values, stop_kws):
                    break

                is_empty_row = not any(v is not None and str(v).strip() != "" for v in row_values)

                # 1. Stop if row is completely empty
                if is_empty_row:
                    consecutive_empty += 1
                    # If we have already extracted items from this table, an empty row signals end-of-table
                    if sheet_items_count > 0 or consecutive_empty >= 2:
                        break
                    continue

                # 2. If before data and row is a column index row (e.g. 1, 2, 3, 4, 5), skip it
                if sheet_items_count == 0:
                    if cells_dict is not None:
                        check_cells = [cells_dict.get((r, c)) or Cell(row=r, col=c) for c in range(1, max_c_bound)]
                    else:
                        check_cells = row_cells_temp
                    if is_column_index_row(check_cells):
                        consecutive_empty = 0
                        continue

                # Extract raw cell values
                if cells_dict is not None:
                    item_cell = cells_dict.get((r, header_info.col_item)) if header_info.col_item > 0 else None
                    quote_cell = cells_dict.get((r, header_info.col_quote)) if header_info.col_quote and header_info.col_quote > 0 else None
                    part_cell = cells_dict.get((r, header_info.col_part)) or Cell(row=r, col=header_info.col_part)
                    ct_cell = cells_dict.get((r, header_info.col_ct)) if header_info.col_ct > 0 else None
                    desc_cell = cells_dict.get((r, header_info.col_desc)) if header_info.col_desc else None
                    mat_cell = cells_dict.get((r, header_info.col_mat)) if header_info.col_mat else None
                    ops_cell = cells_dict.get((r, header_info.col_ops)) if header_info.col_ops else None
                    weight_cell = cells_dict.get((r, header_info.col_weight)) if header_info.col_weight else None
                    vol_cell = cells_dict.get((r, header_info.col_vol)) if header_info.col_vol else None
                    drawing_cell = cells_dict.get((r, header_info.col_drawing)) if header_info.col_drawing else None
                else:
                    item_cell = ws.cell(row=r, column=header_info.col_item) if header_info.col_item > 0 else None
                    quote_cell = ws.cell(row=r, column=header_info.col_quote) if header_info.col_quote and header_info.col_quote > 0 else None
                    part_cell = ws.cell(row=r, column=header_info.col_part)
                    ct_cell = ws.cell(row=r, column=header_info.col_ct) if header_info.col_ct > 0 else None
                    desc_cell = ws.cell(row=r, column=header_info.col_desc) if header_info.col_desc else None
                    mat_cell = ws.cell(row=r, column=header_info.col_mat) if header_info.col_mat else None
                    ops_cell = ws.cell(row=r, column=header_info.col_ops) if header_info.col_ops else None
                    weight_cell = ws.cell(row=r, column=header_info.col_weight) if header_info.col_weight else None
                    vol_cell = ws.cell(row=r, column=header_info.col_vol) if header_info.col_vol else None
                    drawing_cell = ws.cell(row=r, column=header_info.col_drawing) if header_info.col_drawing else None

                raw_item_val = item_cell.value if item_cell else None
                raw_quote_val = quote_cell.value if quote_cell else None
                raw_part_val = part_cell.value
                raw_ct_val = ct_cell.value if ct_cell else None
                raw_desc_val = desc_cell.value if desc_cell else None
                raw_mat_val = mat_cell.value if mat_cell else None
                raw_ops_val = ops_cell.value if ops_cell else None
                raw_weight_val = weight_cell.value if weight_cell else None
                raw_vol_val = vol_cell.value if vol_cell else None
                raw_drawing_val = drawing_cell.value if drawing_cell else None

                # Clean part number and cycle time
                cleaned_pn = clean_part_number_string(raw_part_val)
                cleaned_ct = clean_cycle_time_value(raw_ct_val) if ct_cell else None
                cleaned_desc = clean_text_value(raw_desc_val)
                cleaned_mat = clean_text_value(raw_mat_val)
                cleaned_ops = clean_ops_value(raw_ops_val)
                cleaned_weight = clean_weight_value(raw_weight_val)
                cleaned_vol = clean_volume_value(raw_vol_val)
                cleaned_drawing = clean_text_value(raw_drawing_val)
                if not cleaned_drawing:
                    cleaned_drawing = extract_bracketed_drawing(raw_part_val)

                # Disregard rows where the quote item has no information
                if cleaned_pn is None and (cleaned_ct is None or header_info.col_ct == 0):
                    if sheet_items_count == 0:
                        gap_rows_skipped += 1
                        if gap_rows_skipped >= 10:
                            break
                        continue
                    else:
                        consecutive_empty += 1
                        if consecutive_empty >= 2:
                            break
                        continue

                # Determine quote item number (supporting separate Quote Number & Item# columns)
                warnings = []
                item_num, item_warnings = resolve_quote_item_number(
                    raw_quote_val=raw_quote_val,
                    raw_item_val=raw_item_val,
                    quote_id=quote_id,
                    item_seq=item_seq,
                )
                warnings.extend(item_warnings)


                # Validation against closest QInfo sheet if part number missing on primary sheet
                if cleaned_pn is None:
                    if item_num in qinfo_pn_by_item:
                        cleaned_pn = qinfo_pn_by_item[item_num]
                    elif item_seq in qinfo_pn_by_idx:
                        cleaned_pn = qinfo_pn_by_idx[item_seq]
                elif cleaned_pn in qinfo_item_by_pn and (raw_item_val is None or str(raw_item_val).strip() in ("", "-", "None", "0", "0.0") or str(raw_item_val).strip() == str(raw_quote_val).strip()):
                    item_num = qinfo_item_by_pn[cleaned_pn]

                # A valid quote item row must have a valid part number (not 0, not empty, not error)
                if cleaned_pn is None:
                    continue

                consecutive_empty = 0
                gap_rows_skipped = 0

                item_seq += 1

                # If table has multi-field columns or is MPC quote, validate presence of all required fields
                if has_multi_fields or any(kw in ws.title.lower() for kw in ("mpc", "qinfo", "quote")):
                    if cleaned_desc is None:
                        warnings.append("Missing description")
                    if cleaned_mat is None:
                        warnings.append("Missing material")
                    if cleaned_ct is None:
                        warnings.append("Missing cycle_time_sec")
                    if cleaned_ops is None:
                        warnings.append("Missing ops_labour")
                    if static_labour_rate is None:
                        warnings.append("Missing labour_rate")
                    if cleaned_weight is None:
                        warnings.append("Missing weight_g")
                    if cleaned_vol is None:
                        warnings.append("Missing annual_volume")

                # Coordinates (prefix with sheet title only if workbook has multiple sheets)
                target_q_cell = quote_cell if quote_cell else item_cell
                q_c = target_q_cell.coordinate if target_q_cell else make_coord(r, 1)

                p_c = part_cell.coordinate
                dwg_c = drawing_cell.coordinate if drawing_cell else None
                c_c = ct_cell.coordinate if ct_cell else None
                d_c = desc_cell.coordinate if desc_cell else None
                m_c = mat_cell.coordinate if mat_cell else None
                ops_c = ops_cell.coordinate if ops_cell else None
                lr_c = lr_coord
                w_c = weight_cell.coordinate if weight_cell else None
                v_c = vol_cell.coordinate if vol_cell else None

                if len(workbook.worksheets) > 1:
                    q_c = f"{ws.title}!{q_c}"
                    p_c = f"{ws.title}!{p_c}"
                    if dwg_c: dwg_c = f"{ws.title}!{dwg_c}"
                    if c_c: c_c = f"{ws.title}!{c_c}"
                    if d_c: d_c = f"{ws.title}!{d_c}"
                    if m_c: m_c = f"{ws.title}!{m_c}"
                    if ops_c: ops_c = f"{ws.title}!{ops_c}"
                    if lr_c: lr_c = f"{ws.title}!{lr_c}"
                    if w_c: w_c = f"{ws.title}!{w_c}"
                    if v_c: v_c = f"{ws.title}!{v_c}"

                coords = SourceCellCoords(
                    quote_num=q_c,
                    part_num=p_c,
                    drawing_num=dwg_c,
                    description=d_c,
                    material=m_c,
                    cycle_time=c_c,
                    ops_labour=ops_c,
                    labour_rate=lr_c,
                    weight_g=w_c,
                    annual_volume=v_c,
                )

                quote_item = QuoteItem(
                    quote_item_number=item_num,
                    part_number=cleaned_pn,
                    drawing_number=cleaned_drawing,
                    description=cleaned_desc,
                    material=cleaned_mat,
                    cycle_time_sec=cleaned_ct,
                    ops_labour=cleaned_ops,
                    labour_rate=static_labour_rate,
                    weight_g=cleaned_weight,
                    annual_volume=cleaned_vol,
                    source_cell_coords=coords,
                    warnings=warnings,
                )
                items.append(quote_item)
                sheet_items_count += 1

        return items
