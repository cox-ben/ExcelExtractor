"""Synthetic Reference Workbook Generator for Excel Quote Extractor.

Generates pure-Python valid OpenXML (.xlsx and .xlsm) test fixtures adhering
strictly to the architectural specifications in workbooks_report.md and PROJECT.md.
Does NOT require external heavy tools or Microsoft Excel installation.
"""

from __future__ import annotations

import os
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union


# ---------------------------------------------------------------------------
# Coordinate and XML Helpers
# ---------------------------------------------------------------------------

def col_letter(col_idx: int) -> str:
    """Convert 1-based column index to Excel column letter (e.g. 1->'A', 27->'AA')."""
    if col_idx < 1:
        raise ValueError(f"Column index must be >= 1, got {col_idx}")
    result = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def col_index(letter: str) -> int:
    """Convert Excel column letter to 1-based column index (e.g. 'A'->1, 'AA'->27)."""
    idx = 0
    for char in letter.upper():
        if not ("A" <= char <= "Z"):
            raise ValueError(f"Invalid column letter: {letter}")
        idx = idx * 26 + (ord(char) - ord("A") + 1)
    return idx


def parse_coord(coord: str) -> Tuple[int, int]:
    """Parse cell coordinate string like 'B18' into (row, col) (e.g. (18, 2))."""
    match = re.match(r"^([A-Za-z]+)([0-9]+)$", coord.strip())
    if not match:
        raise ValueError(f"Invalid coordinate: '{coord}'")
    col_str, row_str = match.groups()
    return int(row_str), col_index(col_str)


def make_coord(row: int, col: int) -> str:
    """Create cell coordinate string from (row, col) (e.g. (18, 2) -> 'B18')."""
    return f"{col_letter(col)}{row}"


def escape_xml(s: Any) -> str:
    """Escape string for inclusion in XML text nodes."""
    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    # Strip illegal XML control characters except \t, \n, \r
    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", s)
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


# ---------------------------------------------------------------------------
# Pure-Python OpenXML Builder Classes
# ---------------------------------------------------------------------------

@dataclass
class Cell:
    """Represents a single worksheet cell."""
    row: int
    col: int
    value: Any = None
    data_type: str = "s"  # "s" (string), "n" (number), "b" (bool), "e" (error)
    formula: Optional[str] = None


class Worksheet:
    """Worksheet supporting cell access and merged ranges."""

    def __init__(self, title: str, state: str = "visible"):
        self.title = title
        self.state = state
        self._cells: Dict[Tuple[int, int], Cell] = {}
        self.merged_cells: List[str] = []
        self.hidden_rows: set[int] = set()

    def __setitem__(self, coord: str, value: Any) -> None:
        row, col = parse_coord(coord)
        self.cell(row=row, column=col, value=value)

    def __getitem__(self, coord: str) -> Cell:
        row, col = parse_coord(coord)
        if (row, col) not in self._cells:
            self._cells[(row, col)] = Cell(row=row, col=col, value=None, data_type="s")
        return self._cells[(row, col)]

    def cell(
        self,
        row: int,
        column: int,
        value: Any = None,
        formula: Optional[str] = None,
    ) -> Cell:
        """Get or set a cell value with automatic type inference."""
        key = (row, column)
        if key not in self._cells:
            self._cells[key] = Cell(row=row, col=column)
        cell = self._cells[key]

        if formula is not None:
            cell.formula = formula
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
                cell.value = 0
                cell.data_type = "n"
            elif str(value).startswith("#") and str(value).endswith("!") or str(value) in ("#N/A", "#VALUE!", "#REF!"):
                cell.value = str(value)
                cell.data_type = "e"
            else:
                cell.value = str(value)
                cell.data_type = "s"
        return cell

    def merge_cells(self, range_ref: str) -> None:
        """Merge a cell range e.g. 'B2:D2'."""
        if range_ref not in self.merged_cells:
            self.merged_cells.append(range_ref)


class PureXmlWorkbook:
    """Pure-Python OpenXML Workbook creator with zero external dependencies."""

    def __init__(self):
        self.worksheets: List[Worksheet] = [Worksheet(title="Sheet1")]

    @property
    def active(self) -> Worksheet:
        return self.worksheets[0]

    @property
    def sheetnames(self) -> List[str]:
        return [ws.title for ws in self.worksheets]

    def create_sheet(self, title: str, index: Optional[int] = None, state: str = "visible") -> Worksheet:
        """Create a new sheet with given title."""
        ws = Worksheet(title=title, state=state)
        if index is not None:
            self.worksheets.insert(index, ws)
        else:
            self.worksheets.append(ws)
        return ws

    def get_sheet_by_name(self, title: str) -> Worksheet:
        """Retrieve sheet by title."""
        for ws in self.worksheets:
            if ws.title == title:
                return ws
        raise KeyError(f"Worksheet '{title}' does not exist.")

    def save(self, filename: Union[str, Path]) -> None:
        """Write valid OpenXML zip archive to file (.xlsx or .xlsm)."""
        filepath = Path(filename)
        filepath.parent.mkdir(parents=True, exist_ok=True)
        is_xlsm = filepath.suffix.lower() == ".xlsm"

        # 1. Collect all shared strings
        shared_strings: List[str] = []
        string_map: Dict[str, int] = {}
        total_string_refs = 0

        for ws in self.worksheets:
            for cell in ws._cells.values():
                if cell.data_type == "s" and cell.value is not None and cell.formula is None:
                    total_string_refs += 1
                    s = str(cell.value)
                    if s not in string_map:
                        string_map[s] = len(shared_strings)
                        shared_strings.append(s)

        # 2. Build xl/sharedStrings.xml
        sst_parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
            f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'count="{total_string_refs}" uniqueCount="{len(shared_strings)}">\n',
        ]
        for s in shared_strings:
            sst_parts.append(f'  <si><t>{escape_xml(s)}</t></si>\n')
        sst_parts.append("</sst>")
        shared_strings_xml = "".join(sst_parts)

        # 3. Build each xl/worksheets/sheetN.xml
        sheet_xmls: List[str] = []
        for ws in self.worksheets:
            sh_parts = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n',
                "  <sheetData>\n",
            ]
            # Group cells by row in ascending order
            all_rows = sorted(set(r for (r, c) in ws._cells.keys()))
            for r in all_rows:
                cells_in_row = sorted(
                    [cell for (row_idx, _), cell in ws._cells.items() if row_idx == r],
                    key=lambda c: c.col,
                )
                hidden_attr = ' hidden="1"' if r in getattr(ws, "hidden_rows", set()) else ""
                sh_parts.append(f'    <row r="{r}"{hidden_attr}>\n')
                for cell in cells_in_row:
                    coord = make_coord(cell.row, cell.col)
                    if cell.formula:
                        f_esc = escape_xml(cell.formula.lstrip("="))
                        v_str = str(cell.value) if cell.value is not None else ""
                        sh_parts.append(f'      <c r="{coord}"><f>{f_esc}</f><v>{escape_xml(v_str)}</v></c>\n')
                    elif cell.data_type == "e":
                        sh_parts.append(f'      <c r="{coord}" t="e"><v>{escape_xml(cell.value)}</v></c>\n')
                    elif cell.data_type == "b":
                        b_val = "1" if cell.value else "0"
                        sh_parts.append(f'      <c r="{coord}" t="b"><v>{b_val}</v></c>\n')
                    elif cell.data_type == "n":
                        sh_parts.append(f'      <c r="{coord}"><v>{cell.value}</v></c>\n')
                    elif cell.data_type == "s" and cell.value is not None:
                        idx = string_map[str(cell.value)]
                        sh_parts.append(f'      <c r="{coord}" t="s"><v>{idx}</v></c>\n')
                    else:
                        # Empty cell
                        sh_parts.append(f'      <c r="{coord}"/>\n')
                sh_parts.append("    </row>\n")
            sh_parts.append("  </sheetData>\n")

            if ws.merged_cells:
                sh_parts.append(f'  <mergeCells count="{len(ws.merged_cells)}">\n')
                for m in ws.merged_cells:
                    sh_parts.append(f'    <mergeCell ref="{m}"/>\n')
                sh_parts.append("  </mergeCells>\n")

            sh_parts.append("</worksheet>")
            sheet_xmls.append("".join(sh_parts))

        # 4. Build xl/workbook.xml
        wb_parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n',
            "  <sheets>\n",
        ]
        for i, ws in enumerate(self.worksheets, start=1):
            state_attr = f' state="{ws.state}"' if getattr(ws, "state", "visible") != "visible" else ""
            wb_parts.append(
                f'    <sheet name="{escape_xml(ws.title)}" sheetId="{i}"{state_attr} r:id="rId{i}"/>\n'
            )
        wb_parts.append("  </sheets>\n</workbook>")
        workbook_xml = "".join(wb_parts)

        # 5. Build xl/_rels/workbook.xml.rels
        rels_parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n',
        ]
        for i in range(1, len(self.worksheets) + 1):
            rels_parts.append(
                f'  <Relationship Id="rId{i}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                f'Target="worksheets/sheet{i}.xml"/>\n'
            )
        sst_rid = len(self.worksheets) + 1
        styles_rid = len(self.worksheets) + 2
        rels_parts.append(
            f'  <Relationship Id="rId{sst_rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
            f'Target="sharedStrings.xml"/>\n'
        )
        rels_parts.append(
            f'  <Relationship Id="rId{styles_rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            f'Target="styles.xml"/>\n'
        )
        rels_parts.append("</Relationships>")
        workbook_rels_xml = "".join(rels_parts)

        # 6. Build [Content_Types].xml
        wb_mime = (
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
            if is_xlsm
            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        )
        ct_parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n',
            '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n',
            '  <Default Extension="xml" ContentType="application/xml"/>\n',
            f'  <Override PartName="/xl/workbook.xml" ContentType="{wb_mime}"/>\n',
            '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>\n',
            '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedString+xml"/>\n',
        ]
        for i in range(1, len(self.worksheets) + 1):
            ct_parts.append(
                f'  <Override PartName="/xl/worksheets/sheet{i}.xml" '
                f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n'
            )
        ct_parts.append("</Types>")
        content_types_xml = "".join(ct_parts)

        # 7. Static _rels/.rels
        root_rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            '  <Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/>\n'
            "</Relationships>"
        )

        # 8. Static xl/styles.xml
        styles_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
            '  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>\n'
            '  <fills count="2"><fill><patternFill fillType="none"/></fill><fill><patternFill fillType="gray125"/></fill></fills>\n'
            '  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\n'
            '  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\n'
            '  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>\n'
            '  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>\n'
            "</styleSheet>"
        )

        # 9. Pack into ZipFile
        with zipfile.ZipFile(filepath, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", content_types_xml.encode("utf-8"))
            zf.writestr("_rels/.rels", root_rels_xml.encode("utf-8"))
            zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml.encode("utf-8"))
            zf.writestr("xl/workbook.xml", workbook_xml.encode("utf-8"))
            zf.writestr("xl/styles.xml", styles_xml.encode("utf-8"))
            zf.writestr("xl/sharedStrings.xml", shared_strings_xml.encode("utf-8"))
            for i, sh_xml in enumerate(sheet_xmls, start=1):
                zf.writestr(f"xl/worksheets/sheet{i}.xml", sh_xml.encode("utf-8"))


# ---------------------------------------------------------------------------
# Synthetic Reference Workbook Recipes
# ---------------------------------------------------------------------------

def create_synthetic_modern_single(path: Union[str, Path]) -> Path:
    """Generate standard modern single-item quote workbook (.xlsx).

    Specifications from workbooks_report.md §4.1:
    - Sheet: 'Quote Summary'
    - Quote Number: Q1001-1 at C2 (anchor 'Quote Number:' at B2)
    - Part Number: 2202046213000 at C10 (anchor 'Part Number:' at B10)
    - Cycle Time: 24.5s at C18 (anchor 'Cycle Time (sec):' at B18)
    - Non-target DWG # at F10 (must not be confused with Part #)
    """
    wb = PureXmlWorkbook()
    ws = wb.active
    ws.title = "Quote Summary"

    # Metadata & Quote ID
    ws["B2"] = "Quote Number:"
    ws["C2"] = "Q1001-1"
    ws["B4"] = "Customer:"
    ws["C4"] = "Acme Automotive Corp"
    ws["B6"] = "Quote Date:"
    ws["C6"] = "2026-09-18"

    # Part Details & DWG Disambiguation
    ws["B10"] = "Part Number:"
    ws["C10"] = "2202046213000"
    ws["B11"] = "Part Description:"
    ws["C11"] = "Housing Sensor Bracket"
    ws["E10"] = "Drawing Number:"
    ws["F10"] = "DWG-8841-B"

    # Process Parameters
    ws["B18"] = "Cycle Time (sec):"
    ws["C18"] = 24.5
    ws["B19"] = "Cavities:"
    ws["C19"] = 4

    target = Path(path)
    wb.save(target)
    return target


def create_synthetic_multitab(path: Union[str, Path]) -> Path:
    """Generate multi-tab quote workbook (.xlsx).

    Specifications from workbooks_report.md §4.2:
    - Sheet 'Summary': Global overview (excluded from item extraction)
    - Sheet 'Q2005-1': Item 1 (P/N 'MPC-7011-A', C/T 18.2s)
    - Sheet 'Q2005-2': Item 2 (P/N 'MPC-7012-B', C/T 32.0s)
    - Sheet 'Q2005-3': Item 3 (P/N 'MPC-7013-C', C/T 45.5s)
    - Sheet 'Lookups': Auxiliary press database (excluded from extraction)
    """
    wb = PureXmlWorkbook()

    # 1. Summary Sheet
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_sum["A1"] = "MASTER QUOTE SUMMARY"
    ws_sum["B2"] = "Quote Reference:"
    ws_sum["C2"] = "Q2005"
    ws_sum["A5"] = "Total Items: 3"

    # 2. Item Sheets
    items = [
        ("Q2005-1", "MPC-7011-A", 18.2, "Customer P/N:", "C/T (s):"),
        ("Q2005-2", "MPC-7012-B", 32.0, "Part #:", "Cycle Time:"),
        ("Q2005-3", "MPC-7013-C", 45.5, "P/N:", "Molding Cycle (sec):"),
    ]
    for item_num, part_num, ct, part_anchor, ct_anchor in items:
        ws = wb.create_sheet(title=item_num)
        ws["B2"] = "Quote Item #:"
        ws["C2"] = item_num
        ws["B5"] = part_anchor
        ws["C5"] = part_num
        ws["E12"] = ct_anchor
        ws["F12"] = ct

    # 3. Lookups Sheet
    ws_look = wb.create_sheet(title="Lookups")
    ws_look["A1"] = "Machine Tonnage Rates"
    ws_look["A2"] = "100T"
    ws_look["B2"] = 45.0

    target = Path(path)
    wb.save(target)
    return target


def create_synthetic_grid_matrix(path: Union[str, Path]) -> Path:
    """Generate 2D Tabular Grid Matrix quote workbook (.xlsx).

    Specifications from workbooks_report.md §4.3:
    - Sheet: 'QuoteMatrix'
    - Header in Row 7: Item #, Customer Part #, Description, Cavities, Cycle Time (sec), Tool Cost, Piece Price
    - Row 8: Q4000-1 | MED-10492 | Luer Lock Fitting | 8 | 14.5 | 45000 | 0.185
    - Row 9: Q4000-2 | MED-10493 | Filter Housing Cap | 4 | 22.0 | 38000 | 0.320
    - Row 10: Q4000-3 | MED-10494 | Syringe Plunger | 16 | 9.8 | 62000 | 0.095
    - Row 11: Total row (must be ignored)
    """
    wb = PureXmlWorkbook()
    ws = wb.active
    ws.title = "QuoteMatrix"

    ws["B2"] = "Quote Reference Number:"
    ws["C2"] = "Q4000"
    ws["B3"] = "Customer:"
    ws["C3"] = "Apex Medical Systems"

    headers = [
        "Item #",
        "Customer Part #",
        "Description",
        "Cavities",
        "Cycle Time (sec)",
        "Tool Cost",
        "Piece Price",
    ]
    for c_idx, h in enumerate(headers, start=1):
        ws.cell(row=7, column=c_idx, value=h)

    rows_data = [
        ("Q4000-1", "MED-10492", "Luer Lock Fitting", 8, 14.5, 45000, 0.185),
        ("Q4000-2", "MED-10493", "Filter Housing Cap", 4, 22.0, 38000, 0.320),
        ("Q4000-3", "MED-10494", "Syringe Plunger", 16, 9.8, 62000, 0.095),
    ]
    for r_idx, r_data in enumerate(rows_data, start=8):
        for c_idx, val in enumerate(r_data, start=1):
            ws.cell(row=r_idx, column=c_idx, value=val)

    # Footer row
    ws.cell(row=11, column=1, value="Total")
    ws.cell(row=11, column=6, value=145000)

    target = Path(path)
    wb.save(target)
    return target


def create_synthetic_merged_cells(path: Union[str, Path]) -> Path:
    """Generate quote workbook with merged cells and spatial offsets (.xlsx).

    Specifications from workbooks_report.md §4.4:
    - Merged B2:D2: 'MPC MANUFACTURING QUOTE Q5000-1'
    - Merged B6:C6: Label 'Customer Part Number:'
    - Merged D6:F6: Value '3042605C' (value stored at D6, E6 and F6 are empty/merged)
    - Merged B14:C14: Label 'Cycle Time (Seconds):'
    - Cell D14: spacer column (empty)
    - Cell E14: value 28.0 (offset +3 cols from B14, or +1 from merge edge C14 past spacer)
    """
    wb = PureXmlWorkbook()
    ws = wb.active
    ws.title = "Quote Detail"

    ws.merge_cells("B2:D2")
    ws["B2"] = "MPC MANUFACTURING QUOTE Q5000-1"

    ws.merge_cells("B6:C6")
    ws["B6"] = "Customer Part Number:"
    ws.merge_cells("D6:F6")
    ws["D6"] = "3042605C"

    ws.merge_cells("B14:C14")
    ws["B14"] = "Cycle Time (Seconds):"
    ws["D14"] = ""  # spacer
    ws["E14"] = 28.0

    target = Path(path)
    wb.save(target)
    return target


def create_synthetic_legacy_raw(path: Union[str, Path]) -> Path:
    """Generate legacy macro quote workbook (.xlsm) with base quote and formula.

    Specifications from workbooks_report.md §4.5:
    - Extension: .xlsm
    - Sheet: 'Sheet1'
    - C2: 'Quote No:', D2: 'Q3050' (no explicit -1 suffix, requires suffix inference)
    - C8: 'Part No', D8: '994021A-REV2'
    - D14: 12.5 (injection & hold), D15: 15.0 (cooling & ejection)
    - C16: 'Cycle Time:'
    - D16: Formula '=D14+D15' with cached evaluated value 27.5
    """
    wb = PureXmlWorkbook()
    ws = wb.active
    ws.title = "Sheet1"

    ws["C2"] = "Quote No:"
    ws["D2"] = "Q3050"

    ws["C8"] = "Part No"
    ws["D8"] = "994021A-REV2"

    ws["C14"] = "Inject Time:"
    ws["D14"] = 12.5
    ws["C15"] = "Cool Time:"
    ws["D15"] = 15.0

    ws["C16"] = "Cycle Time:"
    ws.cell(row=16, column=4, value=27.5, formula="=D14+D15")

    target = Path(path)
    wb.save(target)
    return target


def create_synthetic_edge_cases(path: Union[str, Path]) -> Path:
    """Generate edge-case quote workbook (.xlsx) testing boundaries and warnings.

    Specifications from workbooks_report.md §4.6:
    - Sheet 'Out_Of_Bounds_Low': Q6000-1, PART-FAST-01, Cycle Time 2.4s (< 5.0s warning)
    - Sheet 'Out_Of_Bounds_High': Q6000-2, PART-SLOW-02, Cycle Time 420.0s (> 300.0s warning)
    - Sheet 'Text_Formatted_Units': Q6000-3, PART-TXT-03, Cycle Time ' 26.5 sec '
    - Sheet 'Missing_Fields': Q6000-4, missing Part #, Cycle Time '#N/A'
    """
    wb = PureXmlWorkbook()

    # Sheet 1: Out_Of_Bounds_Low
    ws1 = wb.active
    ws1.title = "Out_Of_Bounds_Low"
    ws1["B2"] = "Quote Item #:"
    ws1["C2"] = "Q6000-1"
    ws1["B5"] = "Part Number:"
    ws1["C5"] = "PART-FAST-01"
    ws1["B8"] = "Cycle Time:"
    ws1["C8"] = 2.4

    # Sheet 2: Out_Of_Bounds_High
    ws2 = wb.create_sheet(title="Out_Of_Bounds_High")
    ws2["B2"] = "Quote Item #:"
    ws2["C2"] = "Q6000-2"
    ws2["B5"] = "Part Number:"
    ws2["C5"] = "PART-SLOW-02"
    ws2["B8"] = "Cycle Time:"
    ws2["C8"] = 420.0

    # Sheet 3: Text_Formatted_Units
    ws3 = wb.create_sheet(title="Text_Formatted_Units")
    ws3["B2"] = "Quote Item #:"
    ws3["C2"] = "Q6000-3"
    ws3["B5"] = "Part Number:"
    ws3["C5"] = "PART-TXT-03"
    ws3["B8"] = "Cycle Time:"
    ws3["C8"] = " 26.5 sec "

    # Sheet 4: Missing_Fields
    ws4 = wb.create_sheet(title="Missing_Fields")
    ws4["B2"] = "Quote Item #:"
    ws4["C2"] = "Q6000-4"
    ws4["B5"] = "Part Number:"
    ws4["C5"] = ""
    ws4["B8"] = "Cycle Time:"
    ws4.cell(row=8, column=3, value="#N/A")

    target = Path(path)
    wb.save(target)
    return target


def create_synthetic_empty_workbook(path: Union[str, Path]) -> Path:
    """Generate an empty workbook with no items (.xlsx)."""
    wb = PureXmlWorkbook()
    ws = wb.active
    ws.title = "Sheet1"
    target = Path(path)
    wb.save(target)
    return target


def create_synthetic_non_quote_workbook(path: Union[str, Path]) -> Path:
    """Generate a valid Excel sheet containing general data without quote structure."""
    wb = PureXmlWorkbook()
    ws = wb.active
    ws.title = "Warehouse Inventory"
    ws["A1"] = "SKU"
    ws["B1"] = "Description"
    ws["C1"] = "Quantity"
    ws["A2"] = "SKU-9901"
    ws["B2"] = "Hex Bolt M8"
    ws["C2"] = 1500
    target = Path(path)
    wb.save(target)
    return target


def create_synthetic_scientific_notation(path: Union[str, Path]) -> Path:
    """Generate workbook with 13-digit part number formatted as numeric integer/float."""
    wb = PureXmlWorkbook()
    ws = wb.active
    ws.title = "SciNotationQuote"
    ws["B2"] = "Quote Number:"
    ws["C2"] = "Q7000-1"
    ws["B5"] = "Part Number:"
    ws["C5"] = 2202046213000  # Stored as numeric, often displayed as 2.20205E+12
    ws["B8"] = "Cycle Time (sec):"
    ws["C8"] = 19.5
    target = Path(path)
    wb.save(target)
    return target


def create_synthetic_whitespace_variation(path: Union[str, Path]) -> Path:
    """Generate workbook with irregular casing, spaces, and non-breaking spaces."""
    wb = PureXmlWorkbook()
    ws = wb.active
    ws.title = "MessyQuote"
    ws["B2"] = "  qUoTe  # :  "
    ws["C2"] = "  Q8000-1  "
    ws["B5"] = "  pArT  nUmBeR  : \u00a0"
    ws["C5"] = "  3042605C-REV1; "
    ws["B8"] = "  c/t  : "
    ws["C8"] = 25.0
    target = Path(path)
    wb.save(target)
    return target


def create_synthetic_dwg_disambiguation(path: Union[str, Path]) -> Path:
    """Generate workbook with both Part Number and Drawing Number anchors."""
    wb = PureXmlWorkbook()
    ws = wb.active
    ws.title = "DwgQuote"
    ws["B2"] = "Quote #:"
    ws["C2"] = "Q9000-1"
    ws["B4"] = "Drawing Number:"
    ws["C4"] = "DWG-9999-X"
    ws["B6"] = "Part Number:"
    ws["C6"] = "REAL-PART-1234"
    ws["B8"] = "Cycle Time:"
    ws["C8"] = 22.0
    target = Path(path)
    wb.save(target)
    return target


def create_synthetic_hybrid_quote(path: Union[str, Path]) -> Path:
    """Generate hybrid workbook with 2D summary matrix and child detail sheets."""
    wb = PureXmlWorkbook()

    # Sheet 1: Master Matrix
    ws_mat = wb.active
    ws_mat.title = "Master Summary"
    ws_mat["B2"] = "Quote Number:"
    ws_mat["C2"] = "Q10050"
    headers = ["Item #", "Part #", "Cycle Time (sec)"]
    for c_idx, h in enumerate(headers, start=1):
        ws_mat.cell(row=5, column=c_idx, value=h)
    ws_mat.cell(row=6, column=1, value="Q10050-1")
    ws_mat.cell(row=6, column=2, value="HYBRID-PART-A")
    ws_mat.cell(row=6, column=3, value=15.0)
    ws_mat.cell(row=7, column=1, value="Q10050-2")
    ws_mat.cell(row=7, column=2, value="HYBRID-PART-B")
    ws_mat.cell(row=7, column=3, value=25.0)

    # Sheet 2 & 3: Child Detail Tabs
    ws1 = wb.create_sheet(title="Q10050-1")
    ws1["B2"] = "Quote Item #:"
    ws1["C2"] = "Q10050-1"
    ws1["B5"] = "Part Number:"
    ws1["C5"] = "HYBRID-PART-A"
    ws1["B8"] = "Cycle Time:"
    ws1["C8"] = 15.0

    ws2 = wb.create_sheet(title="Q10050-2")
    ws2["B2"] = "Quote Item #:"
    ws2["C2"] = "Q10050-2"
    ws2["B5"] = "Part Number:"
    ws2["C5"] = "HYBRID-PART-B"
    ws2["B8"] = "Cycle Time:"
    ws2["C8"] = 25.0

    target = Path(path)
    wb.save(target)
    return target


def create_all_fixtures(output_dir: Union[str, Path]) -> Dict[str, Path]:
    """Generate all reference test fixtures into output_dir and return file mapping."""
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    fixtures = {
        "single_modern": create_synthetic_modern_single(out / "quote_single_modern.xlsx"),
        "multitab": create_synthetic_multitab(out / "quote_multitab_q2005.xlsx"),
        "grid_matrix": create_synthetic_grid_matrix(out / "quote_grid_matrix_q4000.xlsx"),
        "merged_cells": create_synthetic_merged_cells(out / "quote_merged_cells_q5000.xlsx"),
        "legacy_raw": create_synthetic_legacy_raw(out / "quote_legacy_raw_q3050.xlsm"),
        "edge_cases": create_synthetic_edge_cases(out / "quote_edge_cases_q6000.xlsx"),
        "empty": create_synthetic_empty_workbook(out / "quote_empty.xlsx"),
        "non_quote": create_synthetic_non_quote_workbook(out / "non_quote_spreadsheet.xlsx"),
        "scientific_notation": create_synthetic_scientific_notation(out / "quote_scientific_notation.xlsx"),
        "whitespace_variation": create_synthetic_whitespace_variation(out / "quote_whitespace_variation.xlsx"),
        "dwg_disambiguation": create_synthetic_dwg_disambiguation(out / "quote_dwg_disambiguation.xlsx"),
        "hybrid": create_synthetic_hybrid_quote(out / "quote_hybrid.xlsx"),
    }
    return fixtures


# ---------------------------------------------------------------------------
# Lightweight Standalone Workbook Inspector
# ---------------------------------------------------------------------------

def inspect_openxml_zip(filepath: Union[str, Path]) -> Dict[str, Any]:
    """Inspect contents of an OpenXML .xlsx/.xlsm file without external libraries.

    Returns dict with sheet names, shared strings, and zip entries.
    """
    path = Path(filepath)
    if not path.is_file():
        raise FileNotFoundError(f"File not found: {path}")

    with zipfile.ZipFile(path, "r") as zf:
        namelist = zf.namelist()

        # Read shared strings
        shared_strings = []
        if "xl/sharedStrings.xml" in namelist:
            xml_data = zf.read("xl/sharedStrings.xml").decode("utf-8")
            shared_strings = re.findall(r"<t(?:[^>]*)>([^<]*)</t>", xml_data)

        # Read sheet names from xl/workbook.xml
        sheets = []
        if "xl/workbook.xml" in namelist:
            wb_xml = zf.read("xl/workbook.xml").decode("utf-8")
            sheets = re.findall(r'<sheet[^>]*name="([^"]+)"', wb_xml)

        return {
            "namelist": namelist,
            "sheets": sheets,
            "shared_strings": shared_strings,
            "is_valid_openxml": "[Content_Types].xml" in namelist and "xl/workbook.xml" in namelist,
        }
