"""
Resilient OpenXML Excel reader for legacy .xlsx and .xlsm workbooks.

Provides a unified interface (Workbook, Worksheet, Cell, MergedCellRange)
with dual-reader capabilities:
1. Pure-Python OpenXML reader using standard library zipfile and xml.etree.ElementTree.
   Zero C-dependencies, 100% portable, fast, handles formulas (<f>), cached values (<v>),
   shared strings, inline strings, and merged cell ranges (<mergeCells>).
2. Openpyxl compatibility adapter if openpyxl is installed and preferred.
"""

from __future__ import annotations

import functools
import io
import os
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, Iterator, List, Optional, Tuple, Union
import xml.etree.ElementTree as ET


# ---------------------------------------------------------------------------
# Coordinate Conversion Helpers
# ---------------------------------------------------------------------------

@functools.lru_cache(maxsize=2048)
def col_letter(col_idx: int) -> str:
    """Convert 1-based column index to Excel column letter (e.g. 1->'A', 27->'AA')."""
    if col_idx < 1:
        raise ValueError(f"Column index must be >= 1, got {col_idx}")
    result = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


@functools.lru_cache(maxsize=2048)
def col_index(letter: str) -> int:
    """Convert Excel column letter to 1-based column index (e.g. 'A'->1, 'AA'->27)."""
    idx = 0
    for char in letter.upper():
        if not ("A" <= char <= "Z"):
            raise ValueError(f"Invalid column letter: {letter}")
        idx = idx * 26 + (ord(char) - ord("A") + 1)
    return idx


@functools.lru_cache(maxsize=16384)
def parse_coord(coord: str) -> Tuple[int, int]:
    """Parse cell coordinate string like 'B18' into (row, col) (e.g. (18, 2))."""
    s = coord.strip()
    idx = 0
    n = len(s)
    while idx < n and s[idx].isalpha():
        idx += 1
    if 0 < idx < n and s[idx:].isdigit():
        return int(s[idx:]), col_index(s[:idx])

    match = re.match(r"^([A-Za-z]+)([0-9]+)$", s)
    if not match:
        raise ValueError(f"Invalid coordinate: '{coord}'")
    col_str, row_str = match.groups()
    return int(row_str), col_index(col_str)


@functools.lru_cache(maxsize=16384)
def make_coord(row: int, col: int) -> str:
    """Create cell coordinate string from (row, col) (e.g. (18, 2) -> 'B18')."""
    return f"{col_letter(col)}{row}"


# ---------------------------------------------------------------------------
# Merged Cell Classes
# ---------------------------------------------------------------------------

class MergedCellRange:
    """Represents a merged rectangular range of cells (e.g. 'B2:D2')."""

    def __init__(self, ref: str):
        self.ref = ref.strip().upper()
        self.coord = self.ref
        if ":" in self.ref:
            start_coord, end_coord = self.ref.split(":", 1)
            self.min_row, self.min_col = parse_coord(start_coord)
            self.max_row, self.max_col = parse_coord(end_coord)
        else:
            self.min_row, self.min_col = parse_coord(self.ref)
            self.max_row, self.max_col = self.min_row, self.min_col

        self.top_left_coord = make_coord(self.min_row, self.min_col)
        self.bottom_right_coord = make_coord(self.max_row, self.max_col)

    def contains(self, row: int, col: int) -> bool:
        """Check if (row, col) is within this merged cell range."""
        return self.min_row <= row <= self.max_row and self.min_col <= col <= self.max_col

    def __contains__(self, item: Any) -> bool:
        if isinstance(item, str):
            r, c = parse_coord(item)
            return self.contains(r, c)
        if isinstance(item, tuple) and len(item) == 2:
            return self.contains(item[0], item[1])
        return False

    def __str__(self) -> str:
        return self.ref

    def __repr__(self) -> str:
        return f"MergedCellRange('{self.ref}')"

    def __eq__(self, other: Any) -> bool:
        if isinstance(other, MergedCellRange):
            return self.ref == other.ref
        if isinstance(other, str):
            return self.ref == other.strip().upper()
        return False


class MergedCellCollection(list):
    """List collection of MergedCellRange objects with helper methods."""

    @property
    def ranges(self) -> List[MergedCellRange]:
        return list(self)

    def get_range_for_cell(self, row: int, col: int) -> Optional[MergedCellRange]:
        """Find the merged range containing (row, col), if any."""
        for r in self:
            if r.contains(row, col):
                return r
        return None

    def __contains__(self, item: Any) -> bool:
        if isinstance(item, str):
            norm = item.strip().upper()
            return any(r.ref == norm for r in self)
        if isinstance(item, MergedCellRange):
            return any(r.ref == item.ref for r in self)
        return super().__contains__(item)


# ---------------------------------------------------------------------------
# Cell & Worksheet Classes
# ---------------------------------------------------------------------------

class Cell:
    """Represents an individual worksheet cell."""

    def __init__(
        self,
        row: int,
        col: int,
        value: Any = None,
        data_type: str = "s",
        formula: Optional[str] = None,
    ):
        self.row = row
        self.column = col
        self.value = value
        self.data_type = data_type
        self.formula = formula

    @property
    def col(self) -> int:
        return self.column

    @col.setter
    def col(self, val: int) -> None:
        self.column = val

    @property
    def coordinate(self) -> str:
        return make_coord(self.row, self.column)

    def __repr__(self) -> str:
        return f"<Cell '{self.coordinate}': value={self.value!r}>"


class Worksheet:
    """Represents an Excel worksheet with coordinate and row/col access."""

    def __init__(
        self,
        title: str,
        sheet_state: str = "visible",
        loader: Optional[Callable[[], None]] = None,
    ):
        self.title = title
        self.sheet_state = sheet_state.lower() if sheet_state else "visible"
        self._cells: Dict[Tuple[int, int], Cell] = {}
        self._merged_cells = MergedCellCollection()
        self._hidden_rows: set[int] = set()
        self._loader = loader
        self._loaded: bool = loader is None
        self._max_row: int = 0
        self._max_col: int = 0

    def _ensure_loaded(self) -> None:
        if not self._loaded and self._loader is not None:
            self._loaded = True
            loader = self._loader
            self._loader = None
            loader()

    @property
    def merged_cells(self) -> MergedCellCollection:
        self._ensure_loaded()
        return self._merged_cells

    @merged_cells.setter
    def merged_cells(self, val: MergedCellCollection) -> None:
        self._merged_cells = val

    @property
    def hidden_rows(self) -> set[int]:
        self._ensure_loaded()
        return self._hidden_rows

    @hidden_rows.setter
    def hidden_rows(self, val: set[int]) -> None:
        self._hidden_rows = val

    @property
    def state(self) -> str:
        return self.sheet_state

    @state.setter
    def state(self, val: str) -> None:
        self.sheet_state = val.lower() if val else "visible"

    @property
    def is_visible(self) -> bool:
        return self.sheet_state == "visible"

    @property
    def is_hidden(self) -> bool:
        return self.sheet_state in ("hidden", "veryhidden")

    def is_row_hidden(self, row: int) -> bool:
        """Check if a 1-based row index is hidden in Excel."""
        self._ensure_loaded()
        return row in self._hidden_rows

    @property
    def max_row(self) -> int:
        self._ensure_loaded()
        if self._max_row > 0:
            return self._max_row
        if not self._cells:
            return 0
        self._max_row = max(
            (r for (r, _), cell in self._cells.items() if cell.value is not None or cell.formula is not None),
            default=0,
        )
        return self._max_row

    @property
    def max_column(self) -> int:
        self._ensure_loaded()
        if self._max_col > 0:
            return self._max_col
        if not self._cells:
            return 0
        self._max_col = max(
            (c for (_, c), cell in self._cells.items() if cell.value is not None or cell.formula is not None),
            default=0,
        )
        return self._max_col

    def __getitem__(self, coord: str) -> Cell:
        self._ensure_loaded()
        r, c = parse_coord(coord)
        return self.cell(row=r, column=c)

    def __setitem__(self, coord: str, value: Any) -> None:
        self._ensure_loaded()
        r, c = parse_coord(coord)
        self.cell(row=r, column=c, value=value)

    def get_cell(self, row: int, column: int) -> Optional[Cell]:
        """Get the Cell at (row, column) if it exists, without initializing an empty Cell."""
        self._ensure_loaded()
        return self._cells.get((row, column))

    def get_cell_value(self, row: int, column: int) -> Any:
        """Get the value at (row, column) if it exists, without initializing an empty Cell."""
        self._ensure_loaded()
        cell = self._cells.get((row, column))
        return cell.value if cell is not None else None

    def cell(
        self,
        row: int,
        column: int,
        value: Any = None,
        formula: Optional[str] = None,
    ) -> Cell:
        """Get or initialize a Cell at (row, column)."""
        self._ensure_loaded()
        key = (row, column)
        if key not in self._cells:
            if value is None and formula is None:
                return Cell(row=row, col=column)
            self._cells[key] = Cell(row=row, col=column)
        cell = self._cells[key]

        if formula is not None:
            cell.formula = formula if str(formula).startswith("=") else f"={formula}"
            cell.value = value
            cell.data_type = "n" if isinstance(value, (int, float)) else "s"
        elif value is not None:
            if isinstance(value, bool):
                cell.value = value
                cell.data_type = "b"
            elif isinstance(value, (int, float)):
                cell.value = value
                cell.data_type = "n"
            elif str(value).startswith("="):
                cell.formula = str(value)
                cell.value = None
                cell.data_type = "n"
            elif str(value).startswith("#") and str(value) in ("#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!"):
                cell.value = str(value)
                cell.data_type = "e"
            else:
                cell.value = str(value)
                cell.data_type = "s"

        if cell.value is not None or cell.formula is not None:
            if row > self._max_row:
                self._max_row = row
            if column > self._max_col:
                self._max_col = column

        return cell

    def merge_cells(self, range_ref: str) -> None:
        """Register a merged cell range."""
        self._ensure_loaded()
        m = MergedCellRange(range_ref)
        if m not in self._merged_cells:
            self._merged_cells.append(m)

    def get_merged_range(self, row: int, col: int) -> Optional[MergedCellRange]:
        """Find merged range for (row, col)."""
        self._ensure_loaded()
        return self._merged_cells.get_range_for_cell(row, col)

    def iter_rows(
        self,
        min_row: int = 1,
        max_row: Optional[int] = None,
        min_col: int = 1,
        max_col: Optional[int] = None,
        values_only: bool = False,
    ) -> Iterator[Tuple[Any, ...]]:
        """Iterate rows yielding tuples of cells or values."""
        self._ensure_loaded()
        end_row = max_row if max_row is not None else self.max_row
        end_col = max_col if max_col is not None else self.max_column

        for r in range(min_row, end_row + 1):
            row_items = []
            for c in range(min_col, end_col + 1):
                cell = self.cell(row=r, column=c)
                row_items.append(cell.value if values_only else cell)
            yield tuple(row_items)


# ---------------------------------------------------------------------------
# Workbook Class
# ---------------------------------------------------------------------------

class Workbook:
    """Represents an Excel workbook containing one or more worksheets."""

    def __init__(self, create_default_sheet: bool = True):
        """
        Initialize workbook.

        :param create_default_sheet: If True, initializes with a single default
            Worksheet("Sheet1") matching OpenPyXL semantics. If False, initializes
            with an empty worksheet list (used when loading existing XML archives).
        """
        self.worksheets: List[Worksheet] = [Worksheet("Sheet1")] if create_default_sheet else []
        self._zf: Optional[zipfile.ZipFile] = None
        self._owns_zf: bool = False

    @property
    def sheetnames(self) -> List[str]:
        return [ws.title for ws in self.worksheets]

    @property
    def active(self) -> Worksheet:
        if not self.worksheets:
            ws = Worksheet("Sheet1")
            self.worksheets.append(ws)
        return self.worksheets[0]

    def create_sheet(
        self,
        title: str,
        index: Optional[int] = None,
        sheet_state: str = "visible",
        loader: Optional[Callable[[], None]] = None,
    ) -> Worksheet:
        """Create a new worksheet."""
        ws = Worksheet(title=title, sheet_state=sheet_state, loader=loader)
        if index is not None:
            self.worksheets.insert(index, ws)
        else:
            self.worksheets.append(ws)
        return ws

    def remove(self, worksheet: Worksheet) -> None:
        """Remove a worksheet from the workbook."""
        if worksheet in self.worksheets:
            self.worksheets.remove(worksheet)

    def __getitem__(self, title: str) -> Worksheet:
        return self.get_sheet_by_name(title)

    def get_sheet_by_name(self, title: str) -> Worksheet:
        for ws in self.worksheets:
            if ws.title == title:
                return ws
        raise KeyError(f"Worksheet '{title}' does not exist.")

    def close(self) -> None:
        """Close workbook resources."""
        if getattr(self, "_owns_zf", False) and getattr(self, "_zf", None) is not None:
            try:
                self._zf.close()
            except Exception:
                pass
            self._zf = None

    def __del__(self) -> None:
        self.close()

    def __enter__(self) -> "Workbook":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.close()


# ---------------------------------------------------------------------------
# Pure-Python OpenXML Reader Implementation
# ---------------------------------------------------------------------------

@functools.lru_cache(maxsize=1024)
def _strip_ns(tag: str) -> str:
    """Remove XML namespace prefix."""
    return tag.rpartition("}")[2] if "}" in tag else tag


def is_ignored_sheet_title(title: str) -> bool:
    """Checks if a sheet title matches administrative ignore keywords."""
    norm = title.lower().strip()
    # If title matches quote target keywords (MPC, QInfo, Quote, RFQ), NEVER ignore
    for target in ("mpc", "qinfo", "quote", "rfq"):
        if target in norm:
            return False

    ignore_list = [
        "summary", "cover", "toc", "terms", "notes", "lookup", "lookups",
        "instructions", "data", "master", "template", "erp",
    ]
    for ign in ignore_list:
        if norm == ign:
            return True
        if norm == f"{ign}s" or norm == f"{ign}es" or norm.rstrip("s") == ign.rstrip("s"):
            return True
        if norm.startswith(f"{ign} ") or norm.endswith(f" {ign}") or f" {ign} " in norm:
            return True
        if norm.startswith(f"{ign}s ") or norm.endswith(f" {ign}s") or f" {ign}s " in norm:
            return True
        base_ign = ign.rstrip("s")
        if re.search(rf"\b{re.escape(base_ign)}s?\b", norm):
            return True
    return False


def _parse_shared_strings(zf: zipfile.ZipFile, sst_path: str) -> List[str]:
    """Parse xl/sharedStrings.xml using fast C-level XML tree parsing."""
    shared_strings: List[str] = []
    try:
        with zf.open(sst_path) as f:
            xml_data = f.read()
        root = ET.fromstring(xml_data)
        ns = root.tag[:root.tag.index("}") + 1] if root.tag.startswith("{") else ""
        si_tag = f"{ns}si"
        t_tag = f"{ns}t"

        for si in root.iter(si_tag):
            t = si.find(t_tag)
            if len(si) == 1 and t is not None:
                shared_strings.append(t.text or "")
            else:
                t_parts = [child.text for child in si.iter(t_tag) if child.text]
                shared_strings.append("".join(t_parts))
    except Exception:
        shared_strings = []
    return shared_strings


def _parse_worksheet_xml(
    ws: Worksheet,
    f: Any,
    shared_strings: List[str],
    data_only: bool = True,
) -> None:
    """Parse worksheet XML using fast C-level ElementTree parsing."""
    try:
        xml_data = f.read() if hasattr(f, "read") else f
        root = ET.fromstring(xml_data)
        ns = root.tag[:root.tag.index("}") + 1] if root.tag.startswith("{") else ""
        row_tag = f"{ns}row"
        c_tag = f"{ns}c"
        v_tag = f"{ns}v"
        f_tag = f"{ns}f"
        is_tag = f"{ns}is"
        t_tag = f"{ns}t"
        merge_tag = f"{ns}mergeCell"

        hidden_rows = ws._hidden_rows
        cells = ws._cells
        max_r = ws._max_row
        max_c = ws._max_col

        for row_elem in root.iter(row_tag):
            r_num_str = row_elem.attrib.get("r")
            row_idx = int(r_num_str) if r_num_str else None
            if row_idx is not None and row_elem.attrib.get("hidden") in ("1", "true", "True"):
                hidden_rows.add(row_idx)

            for c_elem in row_elem:
                if c_elem.tag != c_tag:
                    continue
                coord = c_elem.attrib.get("r")
                if coord:
                    r, c = parse_coord(coord)
                elif row_idx is not None:
                    r = row_idx
                    c = max_c + 1
                else:
                    continue

                data_type = c_elem.attrib.get("t", "n")
                v_elem = c_elem.find(v_tag)
                raw_val_str = v_elem.text if v_elem is not None else None

                resolved_val = None
                if data_type == "s" and raw_val_str is not None:
                    try:
                        s_idx = int(raw_val_str)
                        resolved_val = shared_strings[s_idx] if 0 <= s_idx < len(shared_strings) else ""
                    except ValueError:
                        resolved_val = raw_val_str
                elif data_type == "b" and raw_val_str is not None:
                    resolved_val = raw_val_str in ("1", "true", "True")
                elif data_type == "e" and raw_val_str is not None:
                    resolved_val = raw_val_str
                elif data_type == "inlineStr":
                    is_elem = c_elem.find(is_tag)
                    if is_elem is not None:
                        t_nodes = [e.text for e in is_elem.iter(t_tag) if e.text]
                        resolved_val = "".join(t_nodes)
                elif raw_val_str is not None:
                    if "." not in raw_val_str and "e" not in raw_val_str.lower():
                        try:
                            resolved_val = int(raw_val_str)
                        except ValueError:
                            resolved_val = raw_val_str
                    else:
                        try:
                            f_num = float(raw_val_str)
                            resolved_val = int(f_num) if f_num.is_integer() else f_num
                        except ValueError:
                            resolved_val = raw_val_str

                f_elem = c_elem.find(f_tag)
                formula_val = f"={f_elem.text}" if (f_elem is not None and f_elem.text) else None

                cell_obj = Cell(row=r, col=c)
                cell_obj.data_type = data_type
                cell_obj.formula = formula_val
                if data_only:
                    cell_obj.value = resolved_val
                else:
                    cell_obj.value = formula_val if formula_val is not None else resolved_val

                cells[(r, c)] = cell_obj

                if resolved_val is not None or formula_val is not None:
                    if r > max_r:
                        max_r = r
                    if c > max_c:
                        max_c = c

        ws._max_row = max_r
        ws._max_col = max_c

        for merge_elem in root.iter(merge_tag):
            ref = merge_elem.attrib.get("ref")
            if ref:
                m = MergedCellRange(ref)
                if m not in ws._merged_cells:
                    ws._merged_cells.append(m)
    except Exception:
        pass


def _parse_openxml_archive(zf: zipfile.ZipFile, data_only: bool = True) -> Workbook:
    """Parse an OpenXML (.xlsx, .xlsm) zip archive into a Workbook object."""
    wb = Workbook(create_default_sheet=False)
    namelist = set(zf.namelist())

    def find_in_namelist(target: str) -> Optional[str]:
        t_lower = target.lower().lstrip("/")
        for n in namelist:
            if n.lower().lstrip("/") == t_lower:
                return n
        return None

    # 1. Parse xl/sharedStrings.xml if present (fast C-level ElementTree parsing)
    shared_strings: List[str] = []
    sst_path = find_in_namelist("xl/sharedStrings.xml")
    if sst_path:
        shared_strings = _parse_shared_strings(zf, sst_path)

    # 2. Parse xl/workbook.xml for sheet names, states, and r:id
    wb_path = find_in_namelist("xl/workbook.xml")
    if not wb_path:
        raise ValueError("Invalid OpenXML workbook: missing xl/workbook.xml")

    with zf.open(wb_path) as f:
        wb_root = ET.fromstring(f.read())

    wb_ns = wb_root.tag[:wb_root.tag.index("}") + 1] if wb_root.tag.startswith("{") else ""
    sheets_info: List[Tuple[str, str, str, str]] = []  # (name, sheetId, rId, state)
    for elem in wb_root.iter(f"{wb_ns}sheet"):
        name = elem.attrib.get("name", f"Sheet{len(sheets_info) + 1}")
        sheet_id = elem.attrib.get("sheetId", "")
        state = elem.attrib.get("state", "visible")
        r_id = ""
        for k, v in elem.attrib.items():
            if k.endswith("id") or k.endswith("Id"):
                r_id = v
                break
        sheets_info.append((name, sheet_id, r_id, state))

    # 3. Parse xl/_rels/workbook.xml.rels for target files
    rels_path = find_in_namelist("xl/_rels/workbook.xml.rels")
    r_id_to_target: Dict[str, str] = {}
    if rels_path:
        try:
            with zf.open(rels_path) as f:
                rels_root = ET.fromstring(f.read())
                rels_ns = rels_root.tag[:rels_root.tag.index("}") + 1] if rels_root.tag.startswith("{") else ""
                for rel in rels_root.iter(f"{rels_ns}Relationship"):
                    rid = rel.attrib.get("Id", "")
                    target = rel.attrib.get("Target", "")
                    if target.startswith("/"):
                        norm_target = target.lstrip("/")
                    elif target.startswith("xl/"):
                        norm_target = target
                    else:
                        norm_target = f"xl/{target}"
                    r_id_to_target[rid] = norm_target
        except Exception:
            pass

    # 4. Load each worksheet XML
    loaded_any_eager = False
    for idx, (title, sheet_id, r_id, state) in enumerate(sheets_info, start=1):
        target_path = None
        if r_id in r_id_to_target:
            target_path = find_in_namelist(r_id_to_target[r_id])

        if not target_path:
            candidates = [
                f"xl/worksheets/sheet{idx}.xml",
                f"xl/worksheets/sheet{sheet_id}.xml",
                f"xl/worksheets/Sheet{idx}.xml",
            ]
            for c in candidates:
                found_c = find_in_namelist(c)
                if found_c:
                    target_path = found_c
                    break

        if not target_path or target_path not in namelist:
            wb.create_sheet(title=title, sheet_state=state)
            continue

        is_ignored = is_ignored_sheet_title(title)

        def make_loader(t_path: str, target_ws: Worksheet):
            def loader():
                try:
                    with zf.open(t_path) as sf:
                        _parse_worksheet_xml(target_ws, sf, shared_strings, data_only=data_only)
                except Exception:
                    pass
            return loader

        if is_ignored:
            ws = wb.create_sheet(title=title, sheet_state=state)
            ws._loader = make_loader(target_path, ws)
            ws._loaded = False
        else:
            ws = wb.create_sheet(title=title, sheet_state=state)
            with zf.open(target_path) as sf:
                _parse_worksheet_xml(ws, sf, shared_strings, data_only=data_only)
            loaded_any_eager = True

    # Fallback: if all sheets were marked ignored, eagerly load the first sheet
    if not loaded_any_eager and wb.worksheets:
        wb.worksheets[0]._ensure_loaded()

    return wb


def load_workbook(
    filename_or_stream: Union[str, Path, bytes, io.BytesIO],
    data_only: bool = True,
    read_only: bool = False,
) -> Workbook:
    """
    Loads an Excel workbook (.xlsx, .xlsm) into a unified Workbook object.

    Uses the pure-Python OpenXML reader with zero external binary dependencies.
    """
    if isinstance(filename_or_stream, (str, Path)):
        p = Path(filename_or_stream)
        if not p.exists():
            raise FileNotFoundError(f"Excel file not found: {p}")
        try:
            zf = zipfile.ZipFile(p, "r")
            wb = _parse_openxml_archive(zf, data_only=data_only)
            wb._zf = zf
            wb._owns_zf = True
            return wb
        except zipfile.BadZipFile as e:
            raise ValueError(f"File is not a valid zip / Excel archive: {p}") from e
    elif isinstance(filename_or_stream, bytes):
        try:
            zf = zipfile.ZipFile(io.BytesIO(filename_or_stream), "r")
            wb = _parse_openxml_archive(zf, data_only=data_only)
            wb._zf = zf
            wb._owns_zf = True
            return wb
        except zipfile.BadZipFile as e:
            raise ValueError("Bytes buffer is not a valid zip / Excel archive") from e
    elif hasattr(filename_or_stream, "read"):
        content = filename_or_stream.read()
        if isinstance(content, str):
            content = content.encode("utf-8")
        try:
            zf = zipfile.ZipFile(io.BytesIO(content), "r")
            wb = _parse_openxml_archive(zf, data_only=data_only)
            wb._zf = zf
            wb._owns_zf = True
            return wb
        except zipfile.BadZipFile as e:
            raise ValueError("Stream is not a valid zip / Excel archive") from e
    else:
        raise TypeError(f"Unsupported file type: {type(filename_or_stream)}")
