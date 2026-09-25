"""
Exporters for QuoteDocument data.

Provides serialization and export functions for QuoteDocument and QuoteItem records
to JSON (structured contract), CSV (flat tabular format), and Excel (.xlsx).
Supports single documents and batch collections of documents.
"""

from __future__ import annotations

import csv
import io
import json
import math
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Sequence, Union
import zipfile

from excel_extractor.models import QuoteDocument

# Canonical column order for CSV and tabular Excel exports
CSV_COLUMNS: List[str] = [
    "source_file",
    "quote_id",
    "quote_item_number",
    "part_number",
    "description",
    "material",
    "cycle_time_sec",
    "ops_labour",
    "labour_rate",
    "weight_g",
    "annual_volume",
    "confidence",
    "quote_num_coord",
    "part_num_coord",
    "description_coord",
    "material_coord",
    "cycle_time_coord",
    "ops_labour_coord",
    "labour_rate_coord",
    "weight_g_coord",
    "annual_volume_coord",
    "warnings",
]


def export_to_json(
    documents: Union[QuoteDocument, Sequence[QuoteDocument]],
    file_path: Optional[Union[str, Path]] = None,
    indent: int = 2,
    as_list: Optional[bool] = None,
) -> str:
    """
    Exports QuoteDocument(s) to a JSON formatted string adhering to the structured data contract.
    Optionally writes output to file_path.

    Parameters:
        documents: A single QuoteDocument or a sequence of QuoteDocuments.
        file_path: Optional destination file path on disk.
        indent: JSON indentation spacing (default: 2).
        as_list: If True, always serializes as a list of document dictionaries.
                 If False, serializes a single document as a dictionary.
                 If None, serializes single doc as dict and sequence as list.

    Returns:
        JSON string representation.
    """
    if isinstance(documents, QuoteDocument):
        if as_list is True:
            payload: Any = [documents.to_dict()]
        else:
            payload = documents.to_dict()
    else:
        doc_list = list(documents)
        if as_list is False and len(doc_list) == 1:
            payload = doc_list[0].to_dict()
        else:
            payload = [d.to_dict() for d in doc_list]

    json_str = json.dumps(payload, indent=indent, ensure_ascii=False)

    if file_path is not None:
        target = Path(file_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(json_str, encoding="utf-8")

    return json_str


def export_to_csv(
    documents: Union[QuoteDocument, Sequence[QuoteDocument]],
    file_path: Optional[Union[str, Path]] = None,
) -> str:
    """
    Exports QuoteDocument(s) into flat tabular CSV format.
    Columns: source_file, quote_id, quote_item_number, part_number,
             cycle_time_sec, confidence, quote_num_coord, part_num_coord,
             cycle_time_coord, warnings.

    Parameters:
        documents: A single QuoteDocument or a sequence of QuoteDocuments.
        file_path: Optional destination file path on disk.

    Returns:
        CSV formatted string.
    """
    if isinstance(documents, QuoteDocument):
        docs = [documents]
    else:
        docs = list(documents)

    rows: List[Dict[str, Any]] = []
    for doc in docs:
        rows.extend(doc.to_flat_rows())

    output = io.StringIO()
    writer = csv.DictWriter(
        output,
        fieldnames=CSV_COLUMNS,
        extrasaction="ignore",
        lineterminator="\n",
    )
    writer.writeheader()
    for row in rows:
        writer.writerow(row)

    csv_str = output.getvalue()

    if file_path is not None:
        target = Path(file_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(csv_str, encoding="utf-8", newline="")

    return csv_str


# ---------------------------------------------------------------------------
# Excel (.xlsx) Exporter Helpers
# ---------------------------------------------------------------------------

def _col_letter(col_idx: int) -> str:
    """Convert 1-based column index to Excel column letter."""
    result = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def _escape_xml(s: Any) -> str:
    """Escape string for XML text content."""
    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", s)
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def _build_pure_xml_xlsx(rows: List[Dict[str, Any]], columns: List[str], sheet_name: str = "ExtractedQuotes") -> bytes:
    """
    Builds a valid OpenXML (.xlsx) file in pure Python with zero external dependencies.
    """
    shared_strings: List[str] = []
    string_map: Dict[str, int] = {}
    total_string_refs = 0

    def add_string(s: str) -> int:
        nonlocal total_string_refs
        total_string_refs += 1
        if s not in string_map:
            string_map[s] = len(shared_strings)
            shared_strings.append(s)
        return string_map[s]

    # Prepare rows XML
    row_xmls: List[str] = []

    # Row 1: Header row
    header_cells: List[str] = []
    for c_idx, col in enumerate(columns, start=1):
        coord = f"{_col_letter(c_idx)}1"
        s_idx = add_string(col)
        header_cells.append(f'<c r="{coord}" t="s"><v>{s_idx}</v></c>')
    row_xmls.append(f'    <row r="1">\n      {"".join(header_cells)}\n    </row>\n')

    # Data rows
    for r_idx, row in enumerate(rows, start=2):
        cell_xmls: List[str] = []
        for c_idx, col in enumerate(columns, start=1):
            coord = f"{_col_letter(c_idx)}{r_idx}"
            val = row.get(col)
            if val is None or val == "" or (isinstance(val, float) and (math.isnan(val) or math.isinf(val))):
                cell_xmls.append(f'<c r="{coord}"/>')
            elif isinstance(val, bool):
                b_val = "1" if val else "0"
                cell_xmls.append(f'<c r="{coord}" t="b"><v>{b_val}</v></c>')
            elif isinstance(val, (int, float)):
                cell_xmls.append(f'<c r="{coord}"><v>{val}</v></c>')
            else:
                s_idx = add_string(str(val))
                cell_xmls.append(f'<c r="{coord}" t="s"><v>{s_idx}</v></c>')
        row_xmls.append(f'    <row r="{r_idx}">\n      {"".join(cell_xmls)}\n    </row>\n')

    # Worksheet XML
    worksheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
        '  <sheetData>\n'
        f'{"".join(row_xmls)}'
        '  </sheetData>\n'
        '</worksheet>'
    )

    # Shared Strings XML
    sst_parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{total_string_refs}" uniqueCount="{len(shared_strings)}">\n',
    ]
    for s in shared_strings:
        sst_parts.append(f'  <si><t>{_escape_xml(s)}</t></si>\n')
    sst_parts.append('</sst>')
    shared_strings_xml = "".join(sst_parts)

    # Workbook XML
    clean_sheet_name = _escape_xml(sheet_name)
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
        '  <sheets>\n'
        f'    <sheet name="{clean_sheet_name}" sheetId="1" r:id="rId1"/>\n'
        '  </sheets>\n'
        '</workbook>'
    )

    # Relationships
    wb_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>\n'
        '  <Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
        'Target="sharedStrings.xml"/>\n'
        '  <Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>\n'
        '</Relationships>'
    )

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        '  <Default Extension="xml" ContentType="application/xml"/>\n'
        '  <Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\n'
        '  <Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>\n'
        '  <Override PartName="/xl/sharedStrings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedString+xml"/>\n'
        '  <Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n'
        '</Types>'
    )

    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>\n'
        '</Relationships>'
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
        '  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>\n'
        '  <fills count="2"><fill><patternFill fillType="none"/></fill><fill><patternFill fillType="gray125"/></fill></fills>\n'
        '  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>\n'
        '  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\n'
        '  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>\n'
        '  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>\n'
        '</styleSheet>'
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml.encode("utf-8"))
        zf.writestr("_rels/.rels", root_rels_xml.encode("utf-8"))
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels_xml.encode("utf-8"))
        zf.writestr("xl/workbook.xml", workbook_xml.encode("utf-8"))
        zf.writestr("xl/styles.xml", styles_xml.encode("utf-8"))
        zf.writestr("xl/sharedStrings.xml", shared_strings_xml.encode("utf-8"))
        zf.writestr("xl/worksheets/sheet1.xml", worksheet_xml.encode("utf-8"))

    return buf.getvalue()


def export_to_excel(
    documents: Union[QuoteDocument, Sequence[QuoteDocument]],
    file_path: Optional[Union[str, Path]] = None,
    sheet_name: str = "ExtractedQuotes",
) -> bytes:
    """
    Exports QuoteDocument(s) into an Excel (.xlsx) workbook.
    If openpyxl is installed, it is utilized; otherwise, the resilient pure-Python
    OpenXML generator produces a fully valid standard workbook.

    Parameters:
        documents: A single QuoteDocument or a sequence of QuoteDocuments.
        file_path: Optional destination file path on disk.
        sheet_name: Worksheet title (default: 'ExtractedQuotes').

    Returns:
        Bytes of the generated .xlsx workbook.
    """
    if isinstance(documents, QuoteDocument):
        docs = [documents]
    else:
        docs = list(documents)

    rows: List[Dict[str, Any]] = []
    for doc in docs:
        rows.extend(doc.to_flat_rows())

    xlsx_bytes: bytes
    use_openpyxl = False

    try:
        import openpyxl  # type: ignore
        use_openpyxl = True
    except ImportError:
        use_openpyxl = False

    if use_openpyxl:
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = sheet_name
            ws.append(CSV_COLUMNS)
            for row in rows:
                ws.append([row.get(c, "") if row.get(c) is not None else "" for c in CSV_COLUMNS])
            bio = io.BytesIO()
            wb.save(bio)
            xlsx_bytes = bio.getvalue()
        except Exception:
            # Fallback to pure OpenXML builder
            xlsx_bytes = _build_pure_xml_xlsx(rows, CSV_COLUMNS, sheet_name=sheet_name)
    else:
        xlsx_bytes = _build_pure_xml_xlsx(rows, CSV_COLUMNS, sheet_name=sheet_name)

    if file_path is not None:
        target = Path(file_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(xlsx_bytes)

    return xlsx_bytes


def export_dataframe_to_excel(
    df: Any,
    sheet_name: str = "Extracted Quotes",
    file_path: Optional[Union[str, Path]] = None,
) -> bytes:
    """
    Exports a pandas DataFrame into an Excel (.xlsx) workbook bytes representation
    using the pure-Python OpenXML generator (_build_pure_xml_xlsx) with zero
    external dependencies.
    """
    try:
        import pandas as pd
        clean_df = df.where(pd.notnull(df), None)
        rows = clean_df.to_dict(orient="records")
        columns = [str(c) for c in df.columns]
    except Exception:
        rows = df if isinstance(df, list) else []
        columns = list(rows[0].keys()) if rows else []

    xlsx_bytes = _build_pure_xml_xlsx(rows, columns, sheet_name=sheet_name)

    if file_path is not None:
        target = Path(file_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(xlsx_bytes)

    return xlsx_bytes


# Convenience aliases
export_json = export_to_json
export_csv = export_to_csv
export_excel = export_to_excel
export_dataframe = export_dataframe_to_excel
