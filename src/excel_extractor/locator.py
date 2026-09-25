"""
Anchor locator and spatial resolver for Excel Quote Extractor.

Handles:
- Normalized and fuzzy string matching against anchor keywords.
- Disambiguation (e.g. excluding drawing numbers from part numbers).
- Merged cell boundary resolution (advancing past merged labels and resolving merged values).
- Directional offset scanning and empty spacer cell traversal.
- Exact cell coordinate tracking (e.g. 'C2', 'B8', 'E14').
"""

from __future__ import annotations

import functools
import re
from typing import Any, List, Optional, Tuple, Union

from excel_extractor.config import ExtractorConfig
from excel_extractor.reader import Cell, MergedCellRange, Worksheet, make_coord, parse_coord


@functools.lru_cache(maxsize=8192)
def normalize_text(val: Any) -> str:
    """Normalize text by replacing non-breaking spaces, stripping punctuation, and lowercasing."""
    if val is None:
        return ""
    if not isinstance(val, str):
        val = str(val)
    # Fast collapse of whitespace and replacement of non-breaking spaces
    s = " ".join(val.replace("\u00a0", " ").split()).strip()
    if not s:
        return ""
    # Strip trailing punctuation common in label cells (:, ;, -, etc.)
    s = s.rstrip(" :;-. \t\r\n")
    return s.lower()


DWG_KEYWORDS = ("drawing number", "drawing no", "drawing #", "drawing", "dwg number", "dwg no", "dwg #", "dwg")
DWG_EXACT = set(DWG_KEYWORDS)
DWG_PREFIXES = tuple(f"{kw} " for kw in DWG_KEYWORDS)
DWG_SUFFIXES = tuple(f" {kw}" for kw in DWG_KEYWORDS)


@functools.lru_cache(maxsize=4096)
def is_drawing_anchor(text: str) -> bool:
    """Check if label text represents a drawing number anchor rather than a part number."""
    norm = normalize_text(text)
    if not norm:
        return False
    return norm in DWG_EXACT or norm.startswith(DWG_PREFIXES) or norm.endswith(DWG_SUFFIXES)


def is_drawing_value(val: Any) -> bool:
    """Check if a value represents a drawing identifier (e.g. DWG-8841-B)."""
    if val is None:
        return False
    s = str(val).strip().upper()
    return s.startswith("DWG-") or s.startswith("DWG #") or s.startswith("DRAWING-")


ITEM_LABEL_KEYWORDS = (
    "item #", "item#", "item no", "item no.", "item number",
    "line item", "line #", "line#", "item", "pos #", "pos#", "pos",
)
ITEM_LABEL_EXACT = set(ITEM_LABEL_KEYWORDS)
ITEM_LABEL_PREFIXES = tuple(f"{kw} " for kw in ITEM_LABEL_KEYWORDS)
ITEM_LABEL_SUFFIXES = tuple(f" {kw}" for kw in ITEM_LABEL_KEYWORDS)


@functools.lru_cache(maxsize=4096)
def is_item_label(text: str) -> bool:
    """Check if label text represents an item or line number rather than a part number."""
    norm = normalize_text(text)
    if not norm:
        return False
    return norm in ITEM_LABEL_EXACT or norm.startswith(ITEM_LABEL_PREFIXES) or norm.endswith(ITEM_LABEL_SUFFIXES)


PART_EXCLUDE_PHRASES = (
    "strategic fit",
    "customer awarded",
    "package complete",
    "priority grade",
    "current material",
    "feasibility",
    "assessment",
    "evaluation",
    "checklist",
    "yes/no",
    "approved",
    "question",
)


def clean_part_number_string(val: Any) -> Optional[str]:
    """Clean and normalize part number string. Rejects empty, zero, dashes, error cells, sentences, questions, single chars."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        if val <= 0:
            return None
        if isinstance(val, float) and val.is_integer():
            val = int(val)
        if isinstance(val, int) and 0 <= val <= 9:
            return None
        return str(val)

    s = str(val)
    s = s.replace("\u00a0", " ").replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s).strip(" \t\r\n:;-")

    # Reject placeholders, zeros, dashes, or formula errors
    if not s or s in ("0", "0.0", "-", "--", "---", "#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!", "None", "null", "N/A", "NA"):
        return None

    # Reject standalone single-character strings (e.g. single digit "1"-"9" or single letter "A"-"Z", "T")
    if len(s) <= 1:
        return None

    # Reject questions (ending with or containing '?')
    if "?" in s:
        return None

    # Strip trailing drawing/bracketed tags (e.g. "3042605C (DWG 4022-A)" -> "3042605C", "858007801000 (P0900612)" -> "858007801000")
    if re.search(r"(?i)\b(?:dwg|drawing)\b", s):
        s = re.sub(r"(?i)\s*[\(\[]?\s*(?:dwg|drawing)\b.*$", "", s).strip(" \t\r\n:;-()[]")
    elif re.search(r"\s*[\(\[].*?[\)\]]\s*$", s):
        trimmed = re.sub(r"\s*[\(\[].*?[\)\]]\s*$", "", s).strip(" \t\r\n:;-")
        if trimmed:
            s = trimmed
        else:
            s = s.strip(" \t\r\n:;-()[]")
    else:
        s = s.strip(" \t\r\n:;-")

    if not s or s in ("0", "0.0", "-", "--", "#REF!", "#VALUE!", "#N/A"):
        return None
    if len(s) <= 1:
        return None
    if "?" in s:
        return None

    # Reject review / audit / checklist phrases
    s_lower = s.lower()
    if any(phrase in s_lower for phrase in PART_EXCLUDE_PHRASES):
        return None

    # Reject long sentence-like text (4 or more words)
    if len(s.split()) >= 4:
        return None

    return s if s else None


def extract_bracketed_drawing(val: Any) -> Optional[str]:
    """Extract bracketed drawing identifier (e.g. '(P0900612)' -> 'P0900612')."""
    if val is None:
        return None
    s = str(val).strip()
    m = re.search(r"[\(\[]\s*(?:dwg|drawing\s*#?\s*)?([a-zA-Z0-9\s_-]+?)\s*[\)\]]", s, re.IGNORECASE)
    if m:
        clean = m.group(1).strip()
        if clean and clean.lower() not in ("dwg", "drawing", "none", "n/a", "null"):
            return clean
    return None



def clean_cycle_time_value(val: Any) -> Optional[float]:
    """Extract numeric cycle time float in seconds from raw cell value. Rejects 0 and non-positive numbers."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        if val <= 0:
            return None
        return round(float(val), 2)

    s = str(val).strip().replace("\u00a0", " ")
    if not s or s.startswith("#") or s in ("0", "0.0", "0.00", "-", "--", "N/A", "NA", "None"):
        return None

    # Strip trailing units like 's', 'sec', 'seconds', 'sec/shot'
    s_clean = re.sub(r"(?i)\s*(?:seconds?|secs?|s|sec/shot)\s*$", "", s).strip()
    try:
        val_f = round(float(s_clean), 2)
        return val_f if val_f > 0 else None
    except ValueError:
        # Regex search for first float/int in string
        match = re.search(r"([0-9]+(?:\.[0-9]+)?)", s)
        if match:
            try:
                val_f = round(float(match.group(1)), 2)
                return val_f if val_f > 0 else None
            except ValueError:
                return None
        return None


def clean_text_value(val: Any) -> Optional[str]:
    """Clean text string, rejecting formula errors and placeholders."""
    if val is None:
        return None
    s = str(val).strip(" \t\r\n:;-")
    if not s or s in ("0", "0.0", "-", "--", "---", "#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!", "None", "null", "N/A", "NA"):
        return None
    return s


def clean_ops_value(val: Any) -> Optional[float]:
    """Parse operator count / labour float."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return round(float(val), 2)
    s = str(val).strip().replace("\u00a0", " ")
    if not s or s.startswith("#") or s in ("-", "--", "N/A", "NA", "None"):
        return None
    s_clean = re.sub(r"(?i)\s*(?:ops|labour|labor|operators?)\s*$", "", s).strip()
    try:
        return round(float(s_clean), 2)
    except ValueError:
        match = re.search(r"([0-9]+(?:\.[0-9]+)?)", s_clean)
        if match:
            try:
                return round(float(match.group(1)), 2)
            except ValueError:
                return None
    return None


def clean_currency_value(val: Any) -> Optional[float]:
    """Extract currency float from cell value (e.g. '$25.50' or '25.5')."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return round(float(val), 2)
    s = str(val).strip().replace("\u00a0", " ").replace("$", "").replace(",", "")
    if not s or s.startswith("#") or s in ("-", "--", "N/A", "NA", "None"):
        return None
    match = re.search(r"([0-9]+(?:\.[0-9]+)?)", s)
    if match:
        try:
            return round(float(match.group(1)), 2)
        except ValueError:
            return None
    return None


def clean_weight_value(val: Any) -> Optional[float]:
    """Extract weight in grams from raw cell value. Rejects non-positive."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        if val <= 0:
            return None
        return round(float(val), 2)
    s = str(val).strip().replace("\u00a0", " ")
    if not s or s.startswith("#") or s in ("0", "0.0", "-", "--", "N/A", "NA", "None"):
        return None
    s_clean = re.sub(r"(?i)\s*(?:grams?|g)\s*$", "", s).strip()
    try:
        v = round(float(s_clean), 2)
        return v if v > 0 else None
    except ValueError:
        match = re.search(r"([0-9]+(?:\.[0-9]+)?)", s)
        if match:
            try:
                v = round(float(match.group(1)), 2)
                return v if v > 0 else None
            except ValueError:
                return None
    return None


def clean_volume_value(val: Any) -> Optional[int]:
    """Extract annual volume (EAU) integer from raw cell value."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        if val <= 0:
            return None
        return int(round(float(val)))
    s = str(val).strip().replace("\u00a0", " ").replace(",", "")
    if not s or s.startswith("#") or s in ("0", "0.0", "-", "--", "N/A", "NA", "None"):
        return None
    s_clean = re.sub(r"(?i)\s*(?:eau|units?|pcs?)\s*$", "", s).strip()
    try:
        v = int(round(float(s_clean)))
        return v if v > 0 else None
    except ValueError:
        match = re.search(r"([0-9]+)", s)
        if match:
            try:
                v = int(match.group(1))
                return v if v > 0 else None
            except ValueError:
                return None
    return None


class AnchorLocator:
    """
    Locates quote data fields within a worksheet using configured anchor keywords,
    directional search offsets, and merged cell boundary traversal.
    """

    def __init__(self, config: Optional[ExtractorConfig] = None):
        self.config = config or ExtractorConfig()

    def resolve_target_cell(
        self,
        ws: Worksheet,
        anchor_row: int,
        anchor_col: int,
        offset_row: int,
        offset_col: int,
        allow_spacer_advance: bool = True,
    ) -> Tuple[Optional[Cell], str]:
        """
        Calculates the true target cell from an anchor coordinate and directional offset,
        taking merged cell boundaries and spacer cells into account.

        Returns (resolved_cell, coordinate_string).
        """
        anchor_merge = ws.get_merged_range(anchor_row, anchor_col)

        # 1. Determine effective starting point past anchor's merged boundary
        if anchor_merge:
            if offset_col > 0:
                base_col = anchor_merge.max_col
            elif offset_col < 0:
                base_col = anchor_merge.min_col
            else:
                base_col = anchor_col

            if offset_row > 0:
                base_row = anchor_merge.max_row
            elif offset_row < 0:
                base_row = anchor_merge.min_row
            else:
                base_row = anchor_row
        else:
            base_row = anchor_row
            base_col = anchor_col

        target_row = base_row + offset_row
        target_col = base_col + offset_col

        # 2. Check if initial target is an empty spacer cell
        cell = ws.cell(row=target_row, column=target_col)
        if allow_spacer_advance and offset_col > 0 and (cell.value is None or str(cell.value).strip() == ""):
            # Advance up to 2 spacer columns to the right
            for spacer_step in range(1, 3):
                adv_col = target_col + spacer_step
                adv_cell = ws.cell(row=target_row, column=adv_col)
                if adv_cell.value is not None and str(adv_cell.value).strip() != "":
                    cell = adv_cell
                    target_col = adv_col
                    break

        # 3. If target cell is inside a merged range, resolve to top-left cell
        target_merge = ws.get_merged_range(cell.row, cell.column)
        if target_merge:
            top_cell = ws.cell(row=target_merge.min_row, column=target_merge.min_col)
            return top_cell, top_cell.coordinate

        return cell, cell.coordinate

    def find_quote_id(self, ws: Worksheet) -> Tuple[Optional[str], Optional[str], bool]:
        """
        Searches worksheet for quote identification.

        Returns (quote_item_or_id, coordinate, is_inferred_item).
        """
        item_patterns = [re.compile(p, re.IGNORECASE) for p in self.config.quote_item_patterns]
        base_patterns = [re.compile(p, re.IGNORECASE) for p in self.config.quote_id_patterns]

        # Check sheet title
        for pat in item_patterns:
            m = pat.search(ws.title)
            if m:
                return m.group(1).upper(), None, False
        for pat in base_patterns:
            m = pat.search(ws.title)
            if m:
                return m.group(1).upper(), None, True

        if hasattr(ws, "_ensure_loaded"):
            ws._ensure_loaded()
        cells = getattr(ws, "_cells", None)
        anchor_keywords = [normalize_text(a) for a in self.config.quote_id_anchors]
        offsets = self.config.quote_id.search_offsets or [[0, 1], [0, 2], [1, 0]]

        if cells is not None:
            max_r = min(ws.max_row + 1, 150)
            max_c = min(ws.max_column + 1, 20)
            base_candidate = None
            anchor_cells = []

            for (r, c), cell in cells.items():
                if r >= max_r or c >= max_c or cell.value is None:
                    continue
                val_str = str(cell.value).strip()
                if not val_str:
                    continue

                for pat in item_patterns:
                    m = pat.search(val_str)
                    if m:
                        return m.group(1).upper(), cell.coordinate, False

                norm = normalize_text(val_str)
                if any(norm == ak or norm.startswith(f"{ak} ") or norm.endswith(f" {ak}") for ak in anchor_keywords):
                    anchor_cells.append((r, c))

                if base_candidate is None:
                    for pat in base_patterns:
                        m = pat.search(val_str)
                        if m:
                            base_candidate = (m.group(1).upper(), cell.coordinate, True)
                            break

            for r, c in anchor_cells:
                for ro, co in offsets:
                    t_cell, coord = self.resolve_target_cell(ws, r, c, ro, co)
                    if t_cell and t_cell.value is not None:
                        t_str = str(t_cell.value).strip()
                        for pat in item_patterns:
                            m = pat.search(t_str)
                            if m:
                                return m.group(1).upper(), coord, False
                        for pat in base_patterns:
                            m = pat.search(t_str)
                            if m:
                                return m.group(1).upper(), coord, True

            if base_candidate is not None:
                return base_candidate

            return None, None, False

        # Fallback if cells is None
        for r in range(1, min(ws.max_row + 1, 150)):
            for c in range(1, min(ws.max_column + 1, 20)):
                cell = ws.cell(row=r, column=c)
                if cell.value is not None:
                    val_str = str(cell.value).strip()
                    for pat in item_patterns:
                        m = pat.search(val_str)
                        if m:
                            return m.group(1).upper(), cell.coordinate, False

        for r in range(1, min(ws.max_row + 1, 150)):
            for c in range(1, min(ws.max_column + 1, 15)):
                cell = ws.cell(row=r, column=c)
                if cell.value is not None:
                    norm = normalize_text(cell.value)
                    if any(norm == ak or norm.startswith(f"{ak} ") or norm.endswith(f" {ak}") for ak in anchor_keywords):
                        for ro, co in offsets:
                            t_cell, coord = self.resolve_target_cell(ws, r, c, ro, co)
                            if t_cell and t_cell.value is not None:
                                t_str = str(t_cell.value).strip()
                                for pat in item_patterns:
                                    m = pat.search(t_str)
                                    if m:
                                        return m.group(1).upper(), coord, False
                                for pat in base_patterns:
                                    m = pat.search(t_str)
                                    if m:
                                        return m.group(1).upper(), coord, True

        for r in range(1, min(ws.max_row + 1, 150)):
            for c in range(1, min(ws.max_column + 1, 15)):
                cell = ws.cell(row=r, column=c)
                if cell.value is not None:
                    val_str = str(cell.value).strip()
                    for pat in base_patterns:
                        m = pat.search(val_str)
                        if m:
                            return m.group(1).upper(), cell.coordinate, True

        return None, None, False

    def find_part_number(self, ws: Worksheet) -> Tuple[Optional[str], Optional[str]]:
        """
        Locates and cleans the part number in the worksheet.

        Returns (cleaned_part_number, coordinate).
        """
        anchor_keywords = [normalize_text(k) for k in self.config.part_number_anchors]
        offsets = self.config.part_number_offsets or [[0, 1], [0, 2], [1, 0]]
        candidates: List[Tuple[str, str, int]] = []  # (part_num, coord, priority)

        if hasattr(ws, "_ensure_loaded"):
            ws._ensure_loaded()
        cells = getattr(ws, "_cells", None)
        max_r = min(ws.max_row + 1, 150)
        max_c = min(ws.max_column + 1, 20)

        cell_items = cells.values() if cells is not None else [
            ws.cell(row=r, column=c) for r in range(1, max_r) for c in range(1, max_c)
        ]

        for cell in cell_items:
            if cell.row >= max_r or cell.column >= max_c or cell.value is None:
                continue

            raw_text = str(cell.value)
            norm = normalize_text(raw_text)

            if is_drawing_anchor(norm) or is_item_label(norm):
                continue

            matched_kw = None
            for kw in anchor_keywords:
                if norm == kw or norm.startswith(f"{kw} ") or norm.endswith(f" {kw}"):
                    matched_kw = kw
                    break

            if not matched_kw:
                continue

            if ":" in raw_text:
                after_colon = raw_text.split(":", 1)[1].strip()
                cleaned_inline = clean_part_number_string(after_colon)
                if cleaned_inline and not is_drawing_value(cleaned_inline):
                    candidates.append((cleaned_inline, cell.coordinate, 0))

            for priority, (ro, co) in enumerate(offsets, start=1):
                t_cell, coord = self.resolve_target_cell(ws, cell.row, cell.column, ro, co)
                if t_cell and t_cell.value is not None:
                    t_norm = normalize_text(t_cell.value)
                    if is_drawing_anchor(t_norm) or is_item_label(t_norm):
                        continue
                    cleaned = clean_part_number_string(t_cell.value)
                    if cleaned and not is_drawing_value(cleaned):
                        candidates.append((cleaned, coord, priority))
                        break

        if candidates:
            candidates.sort(key=lambda item: item[2])
            return candidates[0][0], candidates[0][1]

        return None, None

    def find_cycle_time(self, ws: Worksheet) -> Tuple[Optional[float], Optional[str]]:
        """
        Locates and parses molding cycle time in seconds.

        Returns (cycle_time_sec, coordinate).
        """
        anchor_keywords = [normalize_text(k) for k in self.config.cycle_time_anchors]
        offsets = self.config.cycle_time_offsets or [[0, 1], [0, 2], [1, 0]]
        candidates: List[Tuple[Optional[float], str, int]] = []

        if hasattr(ws, "_ensure_loaded"):
            ws._ensure_loaded()
        cells = getattr(ws, "_cells", None)
        max_r = min(ws.max_row + 1, 150)
        max_c = min(ws.max_column + 1, 20)

        cell_items = cells.values() if cells is not None else [
            ws.cell(row=r, column=c) for r in range(1, max_r) for c in range(1, max_c)
        ]

        for cell in cell_items:
            if cell.row >= max_r or cell.column >= max_c or cell.value is None:
                continue

            raw_text = str(cell.value)
            norm = normalize_text(raw_text)

            matched_kw = None
            for kw in anchor_keywords:
                if norm == kw or norm.startswith(f"{kw} ") or norm.endswith(f" {kw}"):
                    matched_kw = kw
                    break

            if not matched_kw:
                continue

            if ":" in raw_text:
                after_colon = raw_text.split(":", 1)[1].strip()
                ct_inline = clean_cycle_time_value(after_colon)
                if ct_inline is not None:
                    candidates.append((ct_inline, cell.coordinate, 0))

            for priority, (ro, co) in enumerate(offsets, start=1):
                t_cell, coord = self.resolve_target_cell(ws, cell.row, cell.column, ro, co)
                if t_cell is not None:
                    raw_val = t_cell.value
                    ct_val = clean_cycle_time_value(raw_val)
                    if ct_val is not None:
                        candidates.append((ct_val, coord, priority))
                        break
                    elif raw_val is not None and str(raw_val).strip() in ("#N/A", "#VALUE!", "#REF!"):
                        candidates.append((None, coord, priority))
                        break

        if candidates:
            candidates.sort(key=lambda item: (0 if item[0] is not None else 1, item[2]))
            return candidates[0][0], candidates[0][1]

        return None, None

    def find_labour_rate(self, ws: Worksheet) -> Tuple[Optional[float], Optional[str]]:
        """
        Locates the static Labour Rate on the worksheet.
        Searches for labels matching Rate $CAD/Hr (or configured labour rate anchors),
        and reads the rate from the cell directly below it (row + 1 or row + 2).
        """
        anchor_keywords = [normalize_text(k) for k in self.config.labour_rate_anchors]
        max_scan_row = min(ws.max_row + 1, 150)
        max_scan_col = min(ws.max_column + 1, 30)

        if hasattr(ws, "_ensure_loaded"):
            ws._ensure_loaded()
        cells = getattr(ws, "_cells", None)

        if cells is not None:
            for (r, c), cell in cells.items():
                if r >= max_scan_row or c >= max_scan_col or cell.value is None:
                    continue
                norm = normalize_text(cell.value)
                if not norm:
                    continue

                matched = any(norm == kw or norm.startswith(f"{kw} ") or norm.endswith(f" {kw}") for kw in anchor_keywords)
                if not matched:
                    matched = any(kw in norm for kw in ("rate $cad/hr", "rate cad/hr", "$cad/hr", "rate $/hr", "rate $cad / hr", "labour rate", "labor rate"))

                if not matched and norm in ("labour", "labor"):
                    cell_below = cells.get((r + 1, c))
                    if cell_below and cell_below.value is not None and any(kw in normalize_text(cell_below.value) for kw in ("rate", "$cad", "cad/hr")):
                        matched = True

                if matched:
                    for next_r in range(r + 1, min(r + 4, ws.max_row + 1)):
                        c_cell = cells.get((next_r, c))
                        if c_cell and c_cell.value is not None:
                            rate_val = clean_currency_value(c_cell.value)
                            if rate_val is not None:
                                return rate_val, c_cell.coordinate
                    for next_c in range(c + 1, min(c + 3, ws.max_column + 1)):
                        c_cell = cells.get((r, next_c))
                        if c_cell and c_cell.value is not None:
                            rate_val = clean_currency_value(c_cell.value)
                            if rate_val is not None:
                                return rate_val, c_cell.coordinate
            return None, None

        for r in range(1, max_scan_row):
            for c in range(1, max_scan_col):
                cell = ws.cell(row=r, column=c)
                if cell.value is None:
                    continue
                norm = normalize_text(cell.value)
                if not norm:
                    continue

                matched = any(norm == kw or norm.startswith(f"{kw} ") or norm.endswith(f" {kw}") for kw in anchor_keywords)
                if not matched:
                    matched = any(kw in norm for kw in ("rate $cad/hr", "rate cad/hr", "$cad/hr", "rate $/hr", "rate $cad / hr", "labour rate", "labor rate"))

                if not matched and norm in ("labour", "labor"):
                    cell_below = ws.cell(row=r + 1, column=c)
                    if cell_below.value is not None and any(kw in normalize_text(cell_below.value) for kw in ("rate", "$cad", "cad/hr")):
                        matched = True

                if matched:
                    for next_r in range(r + 1, min(r + 4, ws.max_row + 1)):
                        c_cell = ws.cell(row=next_r, column=c)
                        rate_val = clean_currency_value(c_cell.value)
                        if rate_val is not None:
                            return rate_val, c_cell.coordinate
                    for next_c in range(c + 1, min(c + 3, ws.max_column + 1)):
                        c_cell = ws.cell(row=r, column=next_c)
                        rate_val = clean_currency_value(c_cell.value)
                        if rate_val is not None:
                            return rate_val, c_cell.coordinate

        return None, None

    def _find_field_by_anchors(
        self,
        ws: Worksheet,
        keywords: List[str],
        cleaner_fn,
        offsets: Optional[List[List[int]]] = None,
        exclude_keywords: Optional[List[str]] = None,
    ) -> Tuple[Optional[Any], Optional[str]]:
        """Generic helper to find a field using anchor keywords and offsets."""
        anchor_keywords = [normalize_text(k) for k in keywords]
        search_offsets = offsets or [[0, 1], [0, 2], [1, 0]]
        candidates: List[Tuple[Any, str, int]] = []
        max_r = min(ws.max_row + 1, 150)
        max_c = min(ws.max_column + 1, 20)

        if hasattr(ws, "_ensure_loaded"):
            ws._ensure_loaded()
        cells = getattr(ws, "_cells", None)

        cell_items = cells.values() if cells is not None else [
            ws.cell(row=r, column=c) for r in range(1, max_r) for c in range(1, max_c)
        ]

        for cell in cell_items:
            if cell.row >= max_r or cell.column >= max_c or cell.value is None:
                continue

            raw_text = str(cell.value)
            norm = normalize_text(raw_text)

            if exclude_keywords and any(ex in norm for ex in exclude_keywords):
                continue

            matched = any(norm == kw or norm.startswith(f"{kw} ") or norm.endswith(f" {kw}") for kw in anchor_keywords)
            if not matched:
                continue

            if ":" in raw_text:
                after_colon = raw_text.split(":", 1)[1].strip()
                val_inline = cleaner_fn(after_colon)
                if val_inline is not None:
                    candidates.append((val_inline, cell.coordinate, 0))

            for priority, (ro, co) in enumerate(search_offsets, start=1):
                t_cell, coord = self.resolve_target_cell(ws, cell.row, cell.column, ro, co)
                if t_cell is not None and t_cell.value is not None:
                    val = cleaner_fn(t_cell.value)
                    if val is not None:
                        candidates.append((val, coord, priority))
                        break

        if candidates:
            candidates.sort(key=lambda item: item[2])
            return candidates[0][0], candidates[0][1]

        return None, None

    def find_description(self, ws: Worksheet) -> Tuple[Optional[str], Optional[str]]:
        """Locates and cleans the part description, strictly rejecting resin/generic descriptions."""
        exclude = ["resin", "material", "generic", "raw", "tool", "machine", "press", "supplier", "vendor", "grade"]
        return self._find_field_by_anchors(
            ws, self.config.description_anchors, clean_text_value, self.config.description_offsets, exclude_keywords=exclude
        )

    def find_material(self, ws: Worksheet) -> Tuple[Optional[str], Optional[str]]:
        """Locates and cleans the raw material / resin description, preferring grade and rejecting suppliers."""
        exclude = ["supplier", "vendor", "distributor", "generic", "mfr", "manufacturer", "tool"]
        # Priority 1: Check specifically for resin grade / material grade anchors first
        grade_anchors = ["resin grade", "material grade", "grade"]
        grade_val, grade_coord = self._find_field_by_anchors(
            ws, grade_anchors, clean_text_value, self.config.material_offsets, exclude_keywords=exclude
        )
        if grade_val is not None:
            return grade_val, grade_coord

        return self._find_field_by_anchors(
            ws, self.config.material_anchors, clean_text_value, self.config.material_offsets, exclude_keywords=exclude
        )

    def find_ops_labour(self, ws: Worksheet) -> Tuple[Optional[float], Optional[str]]:
        """Locates and parses #Ops / Labour float."""
        return self._find_field_by_anchors(
            ws, self.config.ops_labour_anchors, clean_ops_value, self.config.ops_labour_offsets
        )

    def find_weight(self, ws: Worksheet) -> Tuple[Optional[float], Optional[str]]:
        """Locates and parses part weight in grams."""
        return self._find_field_by_anchors(
            ws, self.config.weight_anchors, clean_weight_value, self.config.weight_offsets
        )

    def find_annual_volume(self, ws: Worksheet) -> Tuple[Optional[int], Optional[str]]:
        """Locates and parses annual volume (EAU)."""
        return self._find_field_by_anchors(
            ws, self.config.annual_volume_anchors, clean_volume_value, self.config.annual_volume_offsets
        )
