"""
MPC Excel Quote Extractor & Validator - Streamlit Interactive Testing Dashboard.

Provides:
- Automated extraction from legacy quote workbooks (.xlsx, .xlsm).
- Dark Mode client-side interactive table with instant cell hover origin tooltips & click modal.
- Sanity Check engine with yellow (range warnings) and red (wrong column) highlighting.
- Job# - Quote# Linker workspace:
  * Upload or paste Job numbers and Part numbers
  * Automatic matching via smart normalized part numbers
  * Preserves exact input order of jobs (blanks for missing quotes)
  * Dedicated view for quotes missing jobs
  * Multi-sheet Excel workbook export ('Linked Jobs' + 'Quotes Missing Jobs')
"""

from __future__ import annotations

import functools
import html
import io
import json
from pathlib import Path
import re
import sys
from typing import Any, Dict, List, Optional

# Ensure src/ is on sys.path regardless of execution directory or runner
SRC_DIR = Path(__file__).resolve().parent / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

import pandas as pd
_ = pd.DataFrame({"_": [0.0]})  # Pre-warm pandas DataFrame type engines
import streamlit as st
import streamlit.components.v1 as components

from excel_extractor.engine import QuoteExtractorEngine
from excel_extractor.exporters import export_dataframe_to_excel
from excel_extractor.linker import (
    JobItem,
    LinkedRecord,
    LinkResult,
    build_linked_dataframe,
    build_unlinked_quotes_dataframe,
    export_linked_data_to_multi_sheet_excel,
    link_jobs_to_quotes,
    parse_job_input_file,
    parse_job_input_text,
)
from excel_extractor.models import QuoteDocument

APP_VERSION = "v3.0.0"
BUILD_DATE = "2026-09-23"

FIELD_TO_COORD_MAP: Dict[str, tuple[str, str]] = {
    "Quote Item": ("quote_num", "Quote Item Number"),
    "Part Number": ("part_num", "Part Number"),
    "Description": ("description", "Part Description"),
    "Material": ("material", "Material / Resin Grade"),
    "Cycle Sec": ("cycle_time", "Cycle Time (s)"),
    "#Ops": ("ops_labour", "Number of Operators (#Ops)"),
    "Labour Rate ($CAD/Hr)": ("labour_rate", "Labour Rate ($CAD/Hr)"),
    "Weight (g)": ("weight_g", "Part Weight (g)"),
    "Ann. Volume (EAU)": ("annual_volume", "Annual Volume (EAU)"),
}


def render_footer() -> None:
    """Render version counter and build info at the bottom of the page."""
    st.markdown("---")
    col_info, col_ver = st.columns([7, 3])
    with col_info:
        st.caption("ExcelExtractor • Quote Variable Extractor & Linker")
    with col_ver:
        st.caption(f"**Version**: `{APP_VERSION}` ({BUILD_DATE})")


def initialize_page() -> None:
    """Set Streamlit page configuration and title banner."""
    st.set_page_config(
        page_title=f"ExcelExtractor ({APP_VERSION})",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.title("ExcelExtractor - Legacy Quote Variable Extractor")



def setup_sidebar() -> str:
    """
    Render sidebar navigation controls.

    Returns:
        workspace_mode (str)
    """
    with st.sidebar:
        st.header("Workspace Navigation")
        workspace_mode = st.radio(
            "Select Workspace Mode:",
            options=["Quote Extractor", "Job# - Quote# Linker"],
            index=0,
            key="nav_workspace_mode",
            help="Switch between extracting quote spreadsheets and linking quote items to manufacturing Job numbers.",
        )

        st.markdown("---")
        st.caption(f"App Version: `{APP_VERSION}`")

    return workspace_mode


def format_coordinates_string(item: Any) -> str:
    """Format cell coordinates into a concise readable string."""
    coords = []
    sc = getattr(item, "source_cell_coords", None)
    if sc:
        if getattr(sc, "quote_num", None): coords.append(f"Q: {sc.quote_num}")
        if getattr(sc, "part_num", None): coords.append(f"PN: {sc.part_num}")
        if getattr(sc, "description", None): coords.append(f"Desc: {sc.description}")
        if getattr(sc, "material", None): coords.append(f"Mat: {sc.material}")
        if getattr(sc, "cycle_time", None): coords.append(f"CT: {sc.cycle_time}")
        if getattr(sc, "ops_labour", None): coords.append(f"Ops: {sc.ops_labour}")
        if getattr(sc, "labour_rate", None): coords.append(f"LR: {sc.labour_rate}")
        if getattr(sc, "weight_g", None): coords.append(f"Wt: {sc.weight_g}")
        if getattr(sc, "annual_volume", None): coords.append(f"Vol: {sc.annual_volume}")
    return ", ".join(coords) if coords else "-"


@functools.lru_cache(maxsize=4096)
def parse_cell_coord_sheet_and_cell(raw_coord: Optional[str]) -> tuple[str, str]:
    """Parse 'Sheet!A1' or 'A1' into (sheet_name, cell_coord)."""
    if not raw_coord or raw_coord.strip() in ("-", "None", ""):
        return "N/A", "N/A"
    raw_coord = raw_coord.strip()
    if "!" in raw_coord:
        parts = raw_coord.split("!", 1)
        return parts[0], parts[1]
    return "Current Sheet", raw_coord


RE_NUM_PN = re.compile(r"^\d{5,}$")
RE_DWG_PN = re.compile(r"^[A-Z0-9]{2,6}[-_][A-Z0-9]{3,}$")
RE_NUM_6 = re.compile(r"^\d{6,}$")
RESIN_KEYWORDS = (
    "pa", "gf", "pp", "bk", "nylon", "resin", "poly", "schul",
    "ziam", "griv", "abs", "pom", "pbt", "pc", "tpe", "compound",
    "delrin", "celcon", "dupont", "basf", "sabic", "covestro"
)
DESC_WORDS = (
    "bracket", "spring", "cover", "socket", "ball", "motor", "side",
    "housing", "bushing", "assembly", "bobbin", "schulamid", "ziamid",
    "grivory", "pa66", "gf30"
)
BLANK_STRINGS = frozenset(("", "-", "--", "---", "None", "nan", "null", "N/A", "na", "#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!"))


@functools.lru_cache(maxsize=4096)
def evaluate_cell_sanity(
    col: str,
    val: Any,
    min_ct: Optional[float] = None,
    max_ct: Optional[float] = None,
    **kwargs: Any,
) -> tuple[str, str]:
    """
    Evaluate cell value for domain sanity.
    Returns (severity, message):
      - 'red': wrong thing in the wrong column / critical error / blank data cell
      - 'yellow': number outside of range or slightly different format
      - 'clean': passes sanity check
    """
    # 0. Check if value is blank / missing across extracted data fields (counts as RED error)
    is_blank = (
        val is None
        or (isinstance(val, float) and pd.isna(val))
        or str(val).strip() in BLANK_STRINGS
    )
    if is_blank:
        if col in FIELD_TO_COORD_MAP:
            field_name = FIELD_TO_COORD_MAP[col][1]
            return "red", f"Blank cell: missing {field_name}"
        return "clean", ""

    # 1. Material / Resin Column
    if col == "Material":
        v = str(val).strip()
        if v.upper() == "NO MAT IDENTIFIED":
            return "yellow", "Material unassigned (NO MAT IDENTIFIED)"
        is_num_pn = bool(RE_NUM_PN.match(v))
        is_dwg_pn = bool(RE_DWG_PN.match(v))
        if is_num_pn or is_dwg_pn:
            v_lower = v.lower()
            if not any(k in v_lower for k in RESIN_KEYWORDS):
                return "red", f"Part number '{v}' placed under Resin / Material column"
        return "clean", ""

    # 2. #Ops (Operators / Labour)
    if col == "#Ops":
        try:
            num = float(val)
            if num < 0:
                return "red", f"Negative operator count ({num})"
            if num > 2.0:
                return "red", f"Extreme operator count ({num} ops). Typical range is 0.0 to 1.0 (e.g. 0.3)"
            if num > 1.0:
                return "yellow", f"Operator count ({num} ops) exceeds standard 1.0 threshold"
        except (ValueError, TypeError):
            return "red", f"Non-numeric value in #Ops column: {val}"
        return "clean", ""

    # 3. Labour Rate ($CAD/Hr)
    if col == "Labour Rate ($CAD/Hr)":
        try:
            num = float(val)
            if num <= 0:
                return "red", f"Invalid labour rate (${num:.2f}/hr)"
            if num > 33.0:
                return "yellow", f"Labour rate (${num:.2f}/hr) exceeds standard threshold ($33.00/hr)"
            if num < 15.0:
                return "yellow", f"Labour rate (${num:.2f}/hr) is unusually low (< $15.00/hr)"
        except (ValueError, TypeError):
            return "red", f"Non-numeric value in Labour Rate column: {val}"
        return "clean", ""

    # 4. Part Number
    if col == "Part Number":
        v = str(val).strip()
        words = v.split()
        v_lower = v.lower()
        if len(words) >= 3 or any(w in v_lower for w in DESC_WORDS) or len(v) > 40:
            return "red", f"Description or Resin found in Part Number column: '{v}'"
        return "clean", ""

    # 5. Description
    if col == "Description":
        v = str(val).strip()
        if RE_NUM_6.match(v):
            return "red", f"Part number '{v}' placed under Description column"
        return "clean", ""

    # 6. Cycle Sec
    if col == "Cycle Sec":
        try:
            num = float(val)
            if num <= 0:
                return "red", f"Invalid cycle time ({num}s <= 0)"
        except (ValueError, TypeError):
            return "red", f"Non-numeric value in Cycle Sec column: {val}"
        return "clean", ""

    # 7. Weight (g)
    if col == "Weight (g)":
        try:
            num = float(val)
            if num <= 0:
                return "yellow", "Part weight is zero or negative"
            if num > 5000.0:
                return "yellow", f"Unusually high part weight ({num:.1f}g)"
        except (ValueError, TypeError):
            return "red", f"Non-numeric value in Weight column: {val}"
        return "clean", ""

    # 8. Ann. Volume (EAU)
    if col == "Ann. Volume (EAU)":
        try:
            num = float(val)
            if num <= 0:
                return "yellow", "Annual volume is zero or negative"
            if num > 10_000_000:
                return "yellow", f"Annual volume ({int(num):,}) exceeds standard 10M threshold"
        except (ValueError, TypeError):
            return "red", f"Non-numeric value in Ann. Volume column: {val}"
        return "clean", ""

    # 9. Quote Item
    if col == "Quote Item":
        return "clean", ""

    return "clean", ""


def render_summary_cards(
    doc: QuoteDocument,
    min_ct: Optional[float] = None,
    max_ct: Optional[float] = None,
    **kwargs: Any,
) -> None:
    """Render 3 summary KPI metric cards for extracted quotes."""
    quote_id_display = doc.quote_id or "N/A"
    total_items = len(doc.items)

    confidences = [getattr(item, "confidence", "medium") for item in doc.items]
    if not confidences:
        overall_conf = "N/A"
    elif any(c == "low" for c in confidences):
        overall_conf = "LOW"
    elif any(c == "medium" for c in confidences):
        overall_conf = "MEDIUM"
    else:
        overall_conf = "HIGH"

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Quote ID", quote_id_display)
    with col2:
        st.metric("Total Items Extracted", total_items)
    with col3:
        st.metric("Overall Extraction Confidence", overall_conf)



def build_dataframe_from_document(
    doc: QuoteDocument,
    min_ct: Optional[float] = None,
    max_ct: Optional[float] = None,
    include_coords_col: bool = False,
    **kwargs: Any,
) -> pd.DataFrame:
    """Transform extracted QuoteDocument items into tabular DataFrame with the manufacturing variables."""
    items = doc.items or []
    if not items:
        cols = [
            "Quote Item", "Part Number", "Description", "Material",
            "Cycle Sec", "#Ops", "Labour Rate ($CAD/Hr)", "Weight (g)",
            "Ann. Volume (EAU)", "Confidence"
        ]
        if include_coords_col:
            cols.append("Cell Coordinates")
        cols.append("Warnings")
        return pd.DataFrame(columns=cols)

    data: Dict[str, list] = {
        "Quote Item": [getattr(it, "quote_item_number", "") or "" for it in items],
        "Part Number": [getattr(it, "part_number", "") or "" for it in items],
        "Description": [getattr(it, "description", "") or "" for it in items],
        "Material": [getattr(it, "material", "") or "" for it in items],
        "Cycle Sec": [float(it.cycle_time_sec) if getattr(it, "cycle_time_sec", None) is not None else None for it in items],
        "#Ops": [float(it.ops_labour) if getattr(it, "ops_labour", None) is not None else None for it in items],
        "Labour Rate ($CAD/Hr)": [float(it.labour_rate) if getattr(it, "labour_rate", None) is not None else None for it in items],
        "Weight (g)": [float(it.weight_g) if getattr(it, "weight_g", None) is not None else None for it in items],
        "Ann. Volume (EAU)": [int(it.annual_volume) if getattr(it, "annual_volume", None) is not None else None for it in items],
        "Confidence": [getattr(it, "confidence", "medium") for it in items],
    }
    if include_coords_col:
        data["Cell Coordinates"] = [format_coordinates_string(it) for it in items]
    data["Warnings"] = ["; ".join(getattr(it, "warnings", [])) for it in items]

    return pd.DataFrame(data)


def parse_coordinates_string(coord_str: Any) -> Dict[str, Optional[str]]:
    """Parse cell coordinates string into structured dict."""
    coords: Dict[str, Optional[str]] = {
        "quote_num": None,
        "part_num": None,
        "description": None,
        "material": None,
        "cycle_time": None,
        "ops_labour": None,
        "labour_rate": None,
        "weight_g": None,
        "annual_volume": None,
    }
    if pd.isna(coord_str) or not coord_str or str(coord_str).strip() in ("-", "None", ""):
        return coords

    text = str(coord_str).strip()
    q_match = re.search(r"\bQ:\s*([^\s,;]+)", text, re.IGNORECASE)
    pn_match = re.search(r"\bPN:\s*([^\s,;]+)", text, re.IGNORECASE)
    desc_match = re.search(r"\bDesc:\s*([^\s,;]+)", text, re.IGNORECASE)
    mat_match = re.search(r"\bMat:\s*([^\s,;]+)", text, re.IGNORECASE)
    ct_match = re.search(r"\bCT:\s*([^\s,;]+)", text, re.IGNORECASE)
    ops_match = re.search(r"\bOps:\s*([^\s,;]+)", text, re.IGNORECASE)
    lr_match = re.search(r"\bLR:\s*([^\s,;]+)", text, re.IGNORECASE)
    wt_match = re.search(r"\bWt:\s*([^\s,;]+)", text, re.IGNORECASE)
    vol_match = re.search(r"\bVol:\s*([^\s,;]+)", text, re.IGNORECASE)

    if q_match: coords["quote_num"] = q_match.group(1)
    if pn_match: coords["part_num"] = pn_match.group(1)
    if desc_match: coords["description"] = desc_match.group(1)
    if mat_match: coords["material"] = mat_match.group(1)
    if ct_match: coords["cycle_time"] = ct_match.group(1)
    if ops_match: coords["ops_labour"] = ops_match.group(1)
    if lr_match: coords["labour_rate"] = lr_match.group(1)
    if wt_match: coords["weight_g"] = wt_match.group(1)
    if vol_match: coords["annual_volume"] = vol_match.group(1)

    return coords


def build_canonical_export_dict(
    edited_df: pd.DataFrame,
    source_filename: str,
    quote_id: Optional[str],
    doc_items: Optional[List[Any]] = None,
    doc_warnings: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Construct canonical QuoteDocument schema dictionary from edited DataFrame.
    """
    base_id = quote_id if quote_id else "Q1001"
    items: List[Dict[str, Any]] = []

    for idx, (_, row) in enumerate(edited_df.iterrows()):
        raw_item = row.get("Quote Item")
        item_num = str(raw_item).strip() if pd.notna(raw_item) and str(raw_item).strip() else f"{base_id}-{idx + 1}"

        raw_pn = row.get("Part Number")
        part_num = str(raw_pn).strip() if not pd.isna(raw_pn) and raw_pn is not None and str(raw_pn).strip() else None

        raw_desc = row.get("Description")
        desc = str(raw_desc).strip() if not pd.isna(raw_desc) and raw_desc is not None and str(raw_desc).strip() else None

        raw_mat = row.get("Material")
        mat = str(raw_mat).strip() if not pd.isna(raw_mat) and raw_mat is not None and str(raw_mat).strip() else None

        raw_ct = row.get("Cycle Sec") if "Cycle Sec" in row else row.get("Cycle Time (s)")
        cycle_time: Optional[float] = None
        if not pd.isna(raw_ct) and raw_ct is not None and str(raw_ct).strip() != "":
            try:
                cycle_time = round(float(raw_ct), 2)
            except (ValueError, TypeError):
                cycle_time = None

        raw_ops = row.get("#Ops")
        ops_val: Optional[float] = None
        if not pd.isna(raw_ops) and raw_ops is not None and str(raw_ops).strip() != "":
            try:
                ops_val = round(float(raw_ops), 2)
            except (ValueError, TypeError):
                ops_val = None

        raw_lr = row.get("Labour Rate ($CAD/Hr)")
        lr_val: Optional[float] = None
        if not pd.isna(raw_lr) and raw_lr is not None and str(raw_lr).strip() != "":
            try:
                lr_val = round(float(raw_lr), 2)
            except (ValueError, TypeError):
                lr_val = None

        raw_wt = row.get("Weight (g)")
        wt_val: Optional[float] = None
        if not pd.isna(raw_wt) and raw_wt is not None and str(raw_wt).strip() != "":
            try:
                wt_val = round(float(raw_wt), 2)
            except (ValueError, TypeError):
                wt_val = None

        raw_vol = row.get("Ann. Volume (EAU)")
        vol_val: Optional[int] = None
        if not pd.isna(raw_vol) and raw_vol is not None and str(raw_vol).strip() != "":
            try:
                vol_val = int(round(float(raw_vol)))
            except (ValueError, TypeError):
                vol_val = None

        raw_conf = row.get("Confidence")
        confidence = str(raw_conf).strip().lower() if pd.notna(raw_conf) and str(raw_conf).strip().lower() in ("high", "medium", "low") else "medium"

        if doc_items and idx < len(doc_items) and hasattr(doc_items[idx], "source_cell_coords"):
            sc = doc_items[idx].source_cell_coords
            if hasattr(sc, "model_dump"):
                source_coords = sc.model_dump()
            elif hasattr(sc, "to_dict"):
                source_coords = sc.to_dict()
            elif isinstance(sc, dict):
                source_coords = sc
            else:
                source_coords = parse_coordinates_string(row.get("Cell Coordinates"))
        else:
            source_coords = parse_coordinates_string(row.get("Cell Coordinates"))

        raw_warn = row.get("Warnings")
        warnings_list: List[str] = [w.strip() for w in str(raw_warn).split(";") if w.strip()] if pd.notna(raw_warn) and str(raw_warn).strip() else []

        items.append({
            "quote_item_number": item_num,
            "part_number": part_num,
            "description": desc,
            "material": mat,
            "cycle_time_sec": cycle_time,
            "ops_labour": ops_val,
            "labour_rate": lr_val,
            "weight_g": wt_val,
            "annual_volume": vol_val,
            "source_cell_coords": source_coords,
            "confidence": confidence,
            "warnings": warnings_list,
        })

    return {
        "source_file": source_filename,
        "quote_id": quote_id,
        "total_items": len(items),
        "items": items,
        "warnings": list(doc_warnings) if doc_warnings else [],
    }


def render_export_buttons(
    edited_df: pd.DataFrame,
    source_filename: str,
    quote_id: Optional[str],
    doc_items: Optional[List[Any]] = None,
    doc_warnings: Optional[List[str]] = None,
) -> None:
    """Render export download buttons for Quote Document JSON, CSV, and Excel."""
    st.subheader("Export Extracted Data")

    base_id = quote_id if quote_id else "quote"
    json_filename = f"{base_id}_extracted.json"
    csv_filename = f"{base_id}_extracted.csv"
    excel_filename = f"{base_id}_extracted.xlsx"

    col_json, col_csv, col_excel = st.columns(3)

    export_dict = build_canonical_export_dict(
        edited_df, source_filename, quote_id, doc_items=doc_items, doc_warnings=doc_warnings
    )
    json_str = json.dumps(export_dict, indent=2, ensure_ascii=False)

    with col_json:
        st.download_button(
            label="Download JSON",
            data=json_str.encode("utf-8"),
            file_name=json_filename,
            mime="application/json",
            key="btn_download_json",
            use_container_width=True,
        )

    csv_str = edited_df.to_csv(index=False)
    with col_csv:
        st.download_button(
            label="Download CSV",
            data=csv_str.encode("utf-8"),
            file_name=csv_filename,
            mime="text/csv",
            key="btn_download_csv",
            use_container_width=True,
        )

    excel_bytes = export_dataframe_to_excel(edited_df, sheet_name="Extracted Quotes")
    with col_excel:
        st.download_button(
            label="Download Excel",
            data=excel_bytes,
            file_name=excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_download_excel",
            use_container_width=True,
        )


def generate_interactive_table_html(
    doc: QuoteDocument,
    df: pd.DataFrame,
    min_ct: Optional[float] = None,
    max_ct: Optional[float] = None,
    initial_sanity_active: bool = False,
    **kwargs: Any,
) -> str:
    """
    Generate client-side Dark Mode interactive table HTML for Quote Items.
    """
    items_data: List[Dict[str, Any]] = []
    table_rows_html: List[str] = []

    columns = [
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
        "Warnings",
    ]

    total_red = 0
    total_yellow = 0
    df_records = df.to_dict(orient="records")

    for idx, item in enumerate(doc.items):
        item_coords: List[Dict[str, str]] = []
        row_tds: List[str] = []
        row_record = df_records[idx] if idx < len(df_records) else {}

        item_num = str(item.quote_item_number or f"Item {idx + 1}")
        part_num = str(item.part_number or "-")

        for col_name, (attr, label) in FIELD_TO_COORD_MAP.items():
            raw_c = getattr(item.source_cell_coords, attr, None) or ""
            s_name, c_name = parse_cell_coord_sheet_and_cell(raw_c)
            v = row_record.get(col_name, "-")
            val_str = str(v) if pd.notna(v) and v is not None else "-"
            item_coords.append({
                "col": col_name,
                "label": label,
                "val": val_str,
                "sheet": s_name,
                "cell": c_name,
                "coord": raw_c or "-",
            })

        items_data.append({
            "index": idx,
            "item_num": item_num,
            "part_num": part_num,
            "confidence": str(item.confidence or "medium").upper(),
            "coords": item_coords,
        })

        row_has_red = False
        row_has_yellow = False

        for col in columns:
            raw_val = row_record.get(col, "")
            val_display = ""
            is_cell_blank = (
                raw_val is None
                or (isinstance(raw_val, float) and pd.isna(raw_val))
                or str(raw_val).strip() in ("", "-", "--", "---", "None", "nan", "null", "N/A", "na", "#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!")
            )
            if not is_cell_blank:
                if col in ("Cycle Sec", "#Ops", "Weight (g)"):
                    try:
                        val_display = f"{float(raw_val):.2f}"
                    except (ValueError, TypeError):
                        val_display = str(raw_val)
                elif col == "Labour Rate ($CAD/Hr)":
                    try:
                        val_display = f"${float(raw_val):.2f}"
                    except (ValueError, TypeError):
                        val_display = str(raw_val)
                elif col == "Ann. Volume (EAU)":
                    try:
                        val_display = f"{int(raw_val):,}"
                    except (ValueError, TypeError):
                        val_display = str(raw_val)
                else:
                    val_display = str(raw_val)
            else:
                if col == "Warnings":
                    val_display = ""
                else:
                    val_display = "-"

            escaped_display = html.escape(val_display)
            sanity_status, sanity_msg = evaluate_cell_sanity(col, raw_val)
            if sanity_status == "red":
                total_red += 1
                row_has_red = True
            elif sanity_status == "yellow":
                total_yellow += 1
                row_has_yellow = True

            init_cls = ""
            if initial_sanity_active and sanity_status in ("red", "yellow"):
                init_cls = f" sanity-{sanity_status}"

            if col in FIELD_TO_COORD_MAP:
                attr, field_label = FIELD_TO_COORD_MAP[col]
                raw_coord = getattr(item.source_cell_coords, attr, None) or ""
                s_name, c_name = parse_cell_coord_sheet_and_cell(raw_coord)

                row_tds.append(
                    f'<td class="origin-cell{init_cls}" '
                    f'data-rowidx="{idx}" '
                    f'data-col="{html.escape(col)}" '
                    f'data-field="{html.escape(field_label)}" '
                    f'data-val="{escaped_display}" '
                    f'data-sheet="{html.escape(s_name)}" '
                    f'data-cell="{html.escape(c_name)}" '
                    f'data-coord="{html.escape(raw_coord)}" '
                    f'data-sanity-status="{sanity_status}" '
                    f'data-sanity-msg="{html.escape(sanity_msg)}">'
                    f'{escaped_display}'
                    f'</td>'
                )
            elif col == "Confidence":
                conf_val = str(raw_val).lower() if pd.notna(raw_val) else "medium"
                badge_class = f"badge badge-{conf_val}"
                row_tds.append(f'<td class="regular-cell"><span class="{badge_class}">{html.escape(str(raw_val).upper())}</span></td>')
            else:
                row_tds.append(
                    f'<td class="regular-cell{init_cls}" '
                    f'data-sanity-status="{sanity_status}" '
                    f'data-sanity-msg="{html.escape(sanity_msg)}">'
                    f'{escaped_display}</td>'
                )

        row_flag_attr = "red" if row_has_red else ("yellow" if row_has_yellow else "clean")
        table_rows_html.append(f'<tr class="data-row" data-row-flag="{row_flag_attr}">{"".join(row_tds)}</tr>')

    items_json_str = json.dumps(items_data, ensure_ascii=False)
    tbody_content = "\n".join(table_rows_html)

    btn_active_class = " active" if initial_sanity_active else ""
    btn_text = "Hide Errors & Warnings" if initial_sanity_active else "Show Errors & Warnings"

    html_template = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; }}
  body {{ background-color: transparent; color: #e2e8f0; padding: 2px; font-size: 13px; }}

  .table-container {{
    border: 1px solid #1f2937;
    border-radius: 8px;
    background: #111827;
    box-shadow: 0 4px 20px rgba(0,0,0,0.5);
    overflow: hidden;
  }}

  .toolbar {{
    padding: 10px 14px;
    background: #0b0f19;
    border-bottom: 1px solid #1f2937;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 12px;
    flex-wrap: wrap;
  }}

  .search-box {{
    flex: 1;
    min-width: 220px;
    max-width: 340px;
    padding: 7px 12px;
    background: #030712;
    border: 1px solid #374151;
    color: #f8fafc;
    border-radius: 6px;
    font-size: 12.5px;
    outline: none;
    transition: border-color 0.15s ease, box-shadow 0.15s ease;
  }}
  .search-box:focus {{
    border-color: #38bdf8;
    box-shadow: 0 0 0 2px rgba(56,189,248,0.25);
  }}

  .toolbar-actions {{
    display: flex;
    align-items: center;
    gap: 12px;
  }}

  .btn-sanity {{
    background: linear-gradient(135deg, #1e3a8a, #2563eb);
    color: #ffffff;
    border: 1px solid #3b82f6;
    padding: 6px 14px;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 600;
    cursor: pointer;
    display: flex;
    align-items: center;
    gap: 6px;
    transition: all 0.15s ease;
  }}
  .btn-sanity:hover {{
    background: linear-gradient(135deg, #2563eb, #3b82f6);
    box-shadow: 0 0 10px rgba(59,130,246,0.5);
  }}
  .btn-sanity.active {{
    background: linear-gradient(135deg, #b91c1c, #dc2626);
    border-color: #ef4444;
    box-shadow: 0 0 10px rgba(239,68,68,0.5);
  }}

  .filter-toggle {{
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 11.5px;
    color: #cbd5e1;
    cursor: pointer;
    user-select: none;
  }}
  .filter-toggle input {{
    accent-color: #38bdf8;
    cursor: pointer;
  }}

  .sanity-badges {{
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 11.5px;
  }}
  .stat-badge-red {{
    background: rgba(220, 38, 38, 0.25);
    color: #fca5a5;
    border: 1px solid #ef4444;
    padding: 2px 8px;
    border-radius: 4px;
    font-weight: 600;
  }}
  .stat-badge-yellow {{
    background: rgba(217, 119, 6, 0.25);
    color: #fde047;
    border: 1px solid #f59e0b;
    padding: 2px 8px;
    border-radius: 4px;
    font-weight: 600;
  }}

  .table-wrapper {{
    overflow-x: auto;
    overflow-y: visible;
  }}

  .table-wrapper::-webkit-scrollbar {{
    width: 8px;
    height: 8px;
  }}
  .table-wrapper::-webkit-scrollbar-track {{
    background: #0b0f19;
  }}
  .table-wrapper::-webkit-scrollbar-thumb {{
    background: #334155;
    border-radius: 4px;
  }}
  .table-wrapper::-webkit-scrollbar-thumb:hover {{
    background: #475569;
  }}

  table {{
    width: 100%;
    border-collapse: collapse;
    text-align: left;
    white-space: nowrap;
  }}

  thead th {{
    position: sticky;
    top: 0;
    background: #030712;
    color: #f8fafc;
    font-weight: 600;
    font-size: 12px;
    padding: 10px 12px;
    border-bottom: 2px solid #374151;
    z-index: 10;
    letter-spacing: 0.02em;
  }}

  tbody tr {{
    border-bottom: 1px solid #1f2937;
    background-color: #111827;
    transition: background-color 0.1s ease;
  }}
  tbody tr:nth-child(even) {{
    background-color: #0d131f;
  }}
  tbody tr:hover {{
    background-color: #1a2234;
  }}

  td {{
    padding: 8px 12px;
    font-size: 12.5px;
    color: #e2e8f0;
  }}

  td.origin-cell {{
    cursor: pointer;
    position: relative;
    transition: background-color 0.12s ease, color 0.12s ease;
  }}
  td.origin-cell:hover {{
    background-color: #1e3a5f !important;
    color: #93c5fd !important;
    font-weight: 500;
  }}

  td.sanity-red {{
    background-color: rgba(220, 38, 38, 0.32) !important;
    color: #fca5a5 !important;
    box-shadow: inset 0 0 0 1.5px #ef4444 !important;
    font-weight: 600 !important;
  }}
  td.sanity-red:hover {{
    background-color: rgba(220, 38, 38, 0.48) !important;
  }}

  td.sanity-yellow {{
    background-color: rgba(217, 119, 6, 0.28) !important;
    color: #fde047 !important;
    box-shadow: inset 0 0 0 1.5px #f59e0b !important;
    font-weight: 500 !important;
  }}
  td.sanity-yellow:hover {{
    background-color: rgba(217, 119, 6, 0.42) !important;
  }}

  .badge {{
    display: inline-block;
    padding: 2px 7px;
    border-radius: 4px;
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
  }}
  .badge-high {{ background-color: #064e3b; color: #6ee7b7; border: 1px solid #059669; }}
  .badge-medium {{ background-color: #713f12; color: #fde047; border: 1px solid #d97706; }}
  .badge-low {{ background-color: #7f1d1d; color: #fca5a5; border: 1px solid #dc2626; }}

  #hover-tooltip {{
    position: fixed;
    display: none;
    background: #030712;
    color: #f8fafc;
    padding: 8px 12px;
    border-radius: 6px;
    font-size: 11.5px;
    line-height: 1.4;
    pointer-events: none;
    z-index: 99999;
    box-shadow: 0 10px 25px rgba(0,0,0,0.75);
    border: 1px solid #38bdf8;
    max-width: 340px;
    transition: opacity 0.08s ease;
  }}
  .tt-issue-red {{
    background: rgba(220, 38, 38, 0.25);
    border-left: 3px solid #ef4444;
    padding: 4px 7px;
    margin-bottom: 6px;
    color: #fca5a5;
    font-weight: 600;
    font-size: 11px;
    border-radius: 2px;
  }}
  .tt-issue-yellow {{
    background: rgba(217, 119, 6, 0.25);
    border-left: 3px solid #f59e0b;
    padding: 4px 7px;
    margin-bottom: 6px;
    color: #fde047;
    font-weight: 600;
    font-size: 11px;
    border-radius: 2px;
  }}
  #hover-tooltip .tt-title {{
    font-weight: 600;
    color: #38bdf8;
    margin-bottom: 2px;
  }}
  #hover-tooltip .tt-loc {{
    color: #f8fafc;
    font-family: monospace;
    font-weight: 600;
  }}
  #hover-tooltip .tt-sub {{
    color: #94a3b8;
    font-size: 10.5px;
  }}

  .modal-backdrop {{
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(3, 7, 18, 0.75);
    backdrop-filter: blur(4px);
    display: none;
    align-items: center;
    justify-content: center;
    z-index: 100000;
    padding: 16px;
  }}
  .modal-card {{
    background: #111827;
    border-radius: 10px;
    box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.6);
    width: 100%;
    max-width: 680px;
    min-height: 480px;
    max-height: 90vh;
    overflow-y: auto;
    border: 1px solid #374151;
    color: #e2e8f0;
    animation: modalAnim 0.15s ease-out;
  }}
  @keyframes modalAnim {{
    from {{ opacity: 0; transform: scale(0.96); }}
    to {{ opacity: 1; transform: scale(1); }}
  }}
  .modal-header {{
    padding: 13px 18px;
    background: #030712;
    color: #ffffff;
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-top-left-radius: 9px;
    border-top-right-radius: 9px;
    border-bottom: 1px solid #1f2937;
  }}
  .modal-header h3 {{
    font-size: 14.5px;
    font-weight: 600;
    display: flex;
    align-items: center;
    gap: 8px;
  }}
  .close-btn {{
    background: transparent;
    border: none;
    color: #94a3b8;
    font-size: 20px;
    line-height: 1;
    cursor: pointer;
    padding: 4px 8px;
    border-radius: 4px;
    transition: color 0.15s, background-color 0.15s;
  }}
  .close-btn:hover {{
    color: #ffffff;
    background-color: rgba(255,255,255,0.15);
  }}
  .modal-body {{
    padding: 16px 20px;
  }}
  .modal-alert {{
    padding: 10px 13px;
    border-radius: 6px;
    font-size: 12px;
    margin-bottom: 14px;
    font-weight: 500;
    line-height: 1.4;
  }}
  .modal-alert-red {{
    background: rgba(220, 38, 38, 0.25);
    border: 1px solid #ef4444;
    color: #fca5a5;
  }}
  .modal-alert-yellow {{
    background: rgba(217, 119, 6, 0.25);
    border: 1px solid #f59e0b;
    color: #fde047;
  }}
  .origin-grid {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 10px;
    margin-bottom: 16px;
  }}
  .origin-stat {{
    background: #1f2937;
    border: 1px solid #374151;
    border-radius: 6px;
    padding: 8px 12px;
  }}
  .origin-stat .stat-label {{
    font-size: 10.5px;
    color: #9ca3af;
    text-transform: uppercase;
    font-weight: 600;
  }}
  .origin-stat .stat-value {{
    font-size: 13.5px;
    font-weight: 600;
    color: #f3f4f6;
    margin-top: 2px;
    word-break: break-all;
  }}
  .full-map-title {{
    font-size: 12.5px;
    font-weight: 600;
    color: #e2e8f0;
    margin-bottom: 8px;
  }}
  .map-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 12px;
    margin-bottom: 14px;
  }}
  .map-table th {{
    background: #1f2937;
    color: #cbd5e1;
    font-weight: 600;
    padding: 6px 10px;
    text-align: left;
    border-bottom: 1px solid #374151;
  }}
  .map-table td {{
    padding: 6px 10px;
    border-bottom: 1px solid #1f2937;
    color: #d1d5db;
  }}
  .map-table tr:hover td {{
    background-color: #1f2937;
  }}
  .coord-pill {{
    display: inline-block;
    background: #0369a1;
    color: #e0f2fe;
    font-family: monospace;
    font-weight: 600;
    padding: 2px 6px;
    border-radius: 4px;
  }}
  .modal-footer {{
    display: flex;
    justify-content: flex-end;
    padding-top: 10px;
    border-top: 1px solid #374151;
  }}
  .btn-close {{
    background: #374151;
    color: #ffffff;
    border: none;
    padding: 6px 16px;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 500;
    cursor: pointer;
    transition: background-color 0.15s;
  }}
  .btn-close:hover {{
    background: #4b5563;
  }}
</style>
</head>
<body>

<div class="table-container">
  <div class="toolbar">
    <input type="text" class="search-box" id="tableSearch" placeholder="Search part #, description, material, quote item...">
    <div class="toolbar-actions">
      <div class="sanity-badges" id="sanityBadges">
        <span class="stat-badge-red" id="badgeRed">{total_red} Errors</span>
        <span class="stat-badge-yellow" id="badgeYellow">{total_yellow} Warnings</span>
      </div>
      <button class="btn-sanity{btn_active_class}" id="btnSanityCheck" title="Highlight all cells with errors and warnings">
        {btn_text}
      </button>
      <label class="filter-toggle" title="Filter table to only show rows with flagged errors or warnings">
        <input type="checkbox" id="toggleOnlyFlagged"> Flagged Only
      </label>
    </div>

  </div>
  <div class="table-wrapper">
    <table id="quoteTable">
      <thead>
        <tr>
          <th>Quote Item</th>
          <th>Part Number</th>
          <th>Description</th>
          <th>Material</th>
          <th>Cycle Sec</th>
          <th>#Ops</th>
          <th>Labour Rate ($CAD/Hr)</th>
          <th>Weight (g)</th>
          <th>Ann. Volume (EAU)</th>
          <th>Confidence</th>
          <th>Warnings</th>
        </tr>
      </thead>
      <tbody>
        {tbody_content}
      </tbody>
    </table>
  </div>
</div>

<!-- Hover Tooltip -->
<div id="hover-tooltip"></div>

<!-- Click Modal -->
<div class="modal-backdrop" id="modalBackdrop">
  <div class="modal-card">
    <div class="modal-header">
      <h3 id="modalTitle">Excel Data Origin Inspector</h3>
      <button class="close-btn" id="modalCloseBtn" title="Close (Esc)">✕</button>
    </div>
    <div class="modal-body">
      <div id="modalAlert" class="modal-alert" style="display:none;"></div>

      <div class="origin-grid">
        <div class="origin-stat">
          <div class="stat-label">Field Name</div>
          <div class="stat-value" id="modalFieldName">-</div>
        </div>
        <div class="origin-stat">
          <div class="stat-label">Extracted Value</div>
          <div class="stat-value" id="modalFieldValue">-</div>
        </div>
        <div class="origin-stat">
          <div class="stat-label">Excel Sheet</div>
          <div class="stat-value" id="modalSheetName">-</div>
        </div>
        <div class="origin-stat">
          <div class="stat-label">Cell Coordinate</div>
          <div class="stat-value" id="modalCellCoord">-</div>
        </div>
      </div>

      <div class="full-map-title" id="modalMapTitle">All Excel Source Coordinates for Item</div>
      <table class="map-table">
        <thead>
          <tr>
            <th>Field</th>
            <th>Value</th>
            <th>Excel Sheet</th>
            <th>Coordinate</th>
          </tr>
        </thead>
        <tbody id="modalMapTbody">
        </tbody>
      </table>

      <div class="modal-footer">
        <button class="btn-close" id="modalFooterCloseBtn">Close</button>
      </div>
    </div>
  </div>
</div>

<script>
  const itemsData = {items_json_str};
  let sanityActive = {'true' if initial_sanity_active else 'false'};

  const tooltip = document.getElementById('hover-tooltip');
  const modalBackdrop = document.getElementById('modalBackdrop');
  const modalCloseBtn = document.getElementById('modalCloseBtn');
  const modalFooterCloseBtn = document.getElementById('modalFooterCloseBtn');
  const btnSanityCheck = document.getElementById('btnSanityCheck');
  const toggleOnlyFlagged = document.getElementById('toggleOnlyFlagged');

  function positionTooltip(e) {{
    let left = e.clientX + 14;
    let top = e.clientY + 14;
    if (left + 280 > window.innerWidth) left = e.clientX - 280;
    if (top + 90 > window.innerHeight) top = e.clientY - 85;
    tooltip.style.left = Math.max(8, left) + 'px';
    tooltip.style.top = Math.max(8, top) + 'px';
  }}

  function applySanityClasses() {{
    document.querySelectorAll('td[data-sanity-status]').forEach(td => {{
      const st = td.getAttribute('data-sanity-status');
      if (sanityActive && st && st !== 'clean') {{
        td.classList.add('sanity-' + st);
      }} else {{
        td.classList.remove('sanity-red', 'sanity-yellow');
      }}
    }});

    if (btnSanityCheck) {{
      if (sanityActive) {{
        btnSanityCheck.classList.add('active');
        btnSanityCheck.innerHTML = 'Hide Errors & Warnings';
      }} else {{

        btnSanityCheck.classList.remove('active');
        btnSanityCheck.innerHTML = 'Show Errors & Warnings';
      }}
    }}

    filterTable();
  }}

  const sanityBadges = document.getElementById('sanityBadges');
  if (sanityBadges) {{
    sanityBadges.style.cursor = 'pointer';
    sanityBadges.setAttribute('title', 'Click to toggle error & warning highlighting');
    sanityBadges.addEventListener('click', () => {{
      sanityActive = !sanityActive;
      applySanityClasses();
    }});
  }}

  if (btnSanityCheck) {{
    btnSanityCheck.addEventListener('click', () => {{
      sanityActive = !sanityActive;
      applySanityClasses();
    }});
  }}

  if (toggleOnlyFlagged) {{
    toggleOnlyFlagged.addEventListener('change', () => {{
      filterTable();
    }});
  }}

  document.querySelectorAll('td.origin-cell').forEach(cell => {{
    cell.addEventListener('mouseenter', (e) => {{
      const field = cell.getAttribute('data-field') || 'Field';
      const coord = cell.getAttribute('data-coord') || 'Not Recorded';
      const sheet = cell.getAttribute('data-sheet') || 'N/A';
      const cellAddr = cell.getAttribute('data-cell') || 'N/A';
      const sanityStatus = cell.getAttribute('data-sanity-status');
      const sanityMsg = cell.getAttribute('data-sanity-msg');

      let sanityBanner = '';
      if (sanityActive && sanityStatus && sanityStatus !== 'clean') {{
        const cls = sanityStatus === 'red' ? 'tt-issue-red' : 'tt-issue-yellow';
        const icon = sanityStatus === 'red' ? 'ERROR:' : 'WARNING:';
        sanityBanner = `<div class="${{cls}}"><b>${{icon}}</b> ${{sanityMsg}}</div>`;
      }}

      tooltip.innerHTML = `
        ${{sanityBanner}}
        <div class="tt-title">${{field}}</div>
        <div class="tt-loc">${{coord}}</div>
        <div class="tt-sub">Sheet: <b>${{sheet}}</b> • Cell: <b>${{cellAddr}}</b></div>
      `;
      tooltip.style.display = 'block';
      positionTooltip(e);
    }});

    cell.addEventListener('mousemove', (e) => {{
      positionTooltip(e);
    }});

    cell.addEventListener('mouseleave', () => {{
      tooltip.style.display = 'none';
    }});

    cell.addEventListener('click', () => {{
      tooltip.style.display = 'none';
      const rowIdx = parseInt(cell.getAttribute('data-rowidx'), 10);
      const field = cell.getAttribute('data-field');
      const val = cell.getAttribute('data-val');
      const sheet = cell.getAttribute('data-sheet');
      const cellAddr = cell.getAttribute('data-cell');
      const coord = cell.getAttribute('data-coord');
      const sanityStatus = cell.getAttribute('data-sanity-status');
      const sanityMsg = cell.getAttribute('data-sanity-msg');
      openModal(rowIdx, field, val, sheet, cellAddr, coord, sanityStatus, sanityMsg);
    }});
  }});

  let isModalOpen = false;

  function sendHeight(forceHeight) {{
    try {{
      const container = document.querySelector('.table-container');
      const tableHeight = container ? Math.ceil(container.getBoundingClientRect().height + 16) : Math.ceil(document.body.scrollHeight + 16);
      let h;
      if (forceHeight) {{
        h = forceHeight;
      }} else if (isModalOpen) {{
        h = Math.max(tableHeight, 650);
      }} else if (container) {{
        h = tableHeight;
      }} else {{
        h = Math.ceil(document.body.scrollHeight + 16);
      }}
      if (window.frameElement) {{
        window.frameElement.style.height = h + 'px';
        window.frameElement.height = h;
        let p = window.frameElement.parentElement;
        if (p) {{
          p.style.height = h + 'px';
          p.style.minHeight = h + 'px';
        }}
      }}
      window.parent.postMessage({{
        isStreamlitMessage: true,
        type: 'streamlit:setFrameHeight',
        height: h
      }}, '*');
    }} catch (e) {{
      try {{
        const h = forceHeight || (isModalOpen ? 650 : Math.ceil(document.body.scrollHeight + 16));
        window.parent.postMessage({{
          isStreamlitMessage: true,
          type: 'streamlit:setFrameHeight',
          height: h
        }}, '*');
      }} catch (err) {{}}
    }}
  }}

  function openModal(rowIdx, field, val, sheet, cellAddr, coord, sanityStatus, sanityMsg) {{
    const item = itemsData[rowIdx];
    if (!item) return;

    isModalOpen = true;

    document.getElementById('modalTitle').textContent = `Excel Origin: ${{field}}`;
    document.getElementById('modalFieldName').textContent = field;
    document.getElementById('modalFieldValue').textContent = val || '-';
    document.getElementById('modalSheetName').textContent = sheet || 'N/A';
    document.getElementById('modalCellCoord').textContent = cellAddr !== 'N/A' ? cellAddr : (coord || 'Not Recorded');
    document.getElementById('modalMapTitle').textContent = `All Excel Source Coordinates for Item ${{item.item_num}}`;

    const modalAlert = document.getElementById('modalAlert');
    if (sanityActive && sanityStatus && sanityStatus !== 'clean') {{
      modalAlert.style.display = 'block';
      modalAlert.className = 'modal-alert modal-alert-' + sanityStatus;
      const icon = sanityStatus === 'red' ? 'Error:' : 'Warning:';
      modalAlert.innerHTML = `<b>${{icon}}</b> ${{sanityMsg}}`;
    }} else {{
      modalAlert.style.display = 'none';
    }}

    const mapTbody = document.getElementById('modalMapTbody');
    mapTbody.innerHTML = '';
    item.coords.forEach(c => {{
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><b>${{c.label}}</b></td>
        <td>${{c.val}}</td>
        <td>${{c.sheet}}</td>
        <td><span class="coord-pill">${{c.cell !== 'N/A' ? c.cell : c.coord}}</span></td>
      `;
      mapTbody.appendChild(tr);
    }});

    modalBackdrop.style.display = 'flex';
    const container = document.querySelector('.table-container');
    const tableHeight = container ? Math.ceil(container.getBoundingClientRect().height + 16) : 210;
    sendHeight(Math.max(tableHeight, 650));
  }}

  function closeModal() {{
    isModalOpen = false;
    modalBackdrop.style.display = 'none';
    sendHeight();
  }}

  modalCloseBtn.addEventListener('click', closeModal);
  modalFooterCloseBtn.addEventListener('click', closeModal);
  modalBackdrop.addEventListener('click', (e) => {{
    if (e.target === modalBackdrop) closeModal();
  }});

  document.addEventListener('keydown', (e) => {{
    if (e.key === 'Escape') closeModal();
  }});

  function filterTable() {{
    const query = document.getElementById('tableSearch').value.toLowerCase().trim();
    const onlyFlagged = toggleOnlyFlagged ? toggleOnlyFlagged.checked : false;

    const rows = document.querySelectorAll('tbody tr.data-row');
    rows.forEach(row => {{
      const text = row.textContent.toLowerCase();
      const flag = row.getAttribute('data-row-flag');
      const matchesSearch = !query || text.includes(query);
      const matchesFlag = !onlyFlagged || (flag === 'red' || flag === 'yellow');

      if (matchesSearch && matchesFlag) {{
        row.style.display = '';
      }} else {{
        row.style.display = 'none';
      }}
    }});

    sendHeight();
  }}

  document.getElementById('tableSearch').addEventListener('input', filterTable);

  window.addEventListener('load', () => sendHeight());
  window.addEventListener('resize', () => sendHeight());
  if (window.ResizeObserver) {{
    const ro = new ResizeObserver(() => sendHeight());
    const c = document.querySelector('.table-container');
    if (c) ro.observe(c);
  }}
</script>

</body>
</html>"""
    return html_template


def render_quote_extractor_workspace(engine: QuoteExtractorEngine) -> None:
    """Render the primary Quote Extractor workspace."""
    # 1. Main File Uploader
    uploaded_file = st.file_uploader(
        "Upload Legacy Quote Workbook",
        type=["xlsx", "xlsm"],
        help="Select a legacy quoting spreadsheet (.xlsx or .xlsm) to parse.",
        key="main_workbook_uploader",
    )

    if uploaded_file is None:
        st.info("Please upload a legacy quote workbook (.xlsx or .xlsm) to begin extraction.")
        render_footer()
        return

    # Cache extracted workbook in session state to eliminate re-extraction latency
    file_bytes = uploaded_file.getvalue()
    cache_key = f"{APP_VERSION}_{uploaded_file.name}_{len(file_bytes)}_{hash((file_bytes[:1024], file_bytes[-1024:]))}"
    if st.session_state.get("last_cache_key") == cache_key and "cached_doc" in st.session_state:
        doc = st.session_state["cached_doc"]
        df = st.session_state.get("cached_df")
        table_html = st.session_state.get("cached_table_html")
        if df is None:
            df = build_dataframe_from_document(doc)
            st.session_state["cached_df"] = df
        if table_html is None:
            table_html = generate_interactive_table_html(doc, df, initial_sanity_active=False)
            st.session_state["cached_table_html"] = table_html
    else:
        try:
            doc = engine.extract_file(io.BytesIO(file_bytes), filename=uploaded_file.name)
            st.session_state["cached_doc"] = doc
            st.session_state["cached_filename"] = uploaded_file.name
            st.session_state["last_cache_key"] = cache_key
            df = build_dataframe_from_document(doc)
            st.session_state["cached_df"] = df
            table_html = generate_interactive_table_html(doc, df, initial_sanity_active=False)
            st.session_state["cached_table_html"] = table_html
        except Exception as exc:
            st.error(f"Failed to parse workbook '{uploaded_file.name}': {exc}")
            render_footer()
            return

    if not doc.items or doc.total_items == 0:
        st.warning("No quote line items were found in the uploaded workbook.")
        if doc.warnings:
            for w in doc.warnings:
                st.caption(f"• {w}")
        render_footer()
        return

    # Summary Metric Cards
    render_summary_cards(doc)

    # Extracted Items Subheader (single sanity toggle inside table toolbar)
    st.subheader("Extracted Quote Items")

    # Render Dark Mode interactive table with dynamic height
    # Calculate initial iframe height dynamically based on row count:
    # Hugs the table tightly without blank space; moves up/down as rows change
    calc_height = max(135 + len(df) * 40, 210)
    components.html(table_html, height=calc_height, scrolling=False)

    # Standard data table view for headless test compatibility
    column_config_dict = {
        "Quote Item": st.column_config.TextColumn("Quote Item", required=True),
        "Part Number": st.column_config.TextColumn("Part Number"),
        "Description": st.column_config.TextColumn("Description"),
        "Material": st.column_config.TextColumn("Material"),
        "Cycle Sec": st.column_config.NumberColumn("Cycle Sec", format="%.2f"),
        "#Ops": st.column_config.NumberColumn("#Ops", format="%.2f"),
        "Labour Rate ($CAD/Hr)": st.column_config.NumberColumn("Labour Rate ($CAD/Hr)", format="$%.2f"),
        "Weight (g)": st.column_config.NumberColumn("Weight (g)", format="%.2f"),
        "Ann. Volume (EAU)": st.column_config.NumberColumn("Ann. Volume (EAU)", format="%d"),
        "Confidence": st.column_config.SelectboxColumn("Confidence", options=["high", "medium", "low"]),
        "Warnings": st.column_config.TextColumn("Warnings"),
    }

    with st.expander("View Standard Data Table", expanded=False):
        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
            key=f"data_df_{uploaded_file.name}",
            column_config=column_config_dict,
        )

    render_export_buttons(df, uploaded_file.name, doc.quote_id, doc_items=doc.items, doc_warnings=doc.warnings)
    render_footer()


def render_job_linker_workspace() -> None:
    """Render the Job# - Quote# Linker workspace."""
    st.subheader("Job# - Quote# Linker Workspace")
    st.caption("Paste or upload manufacturing Job numbers with Part numbers to automatically align quote variables in exact Job order.")

    doc: Optional[QuoteDocument] = st.session_state.get("cached_doc", None)
    quote_filename = st.session_state.get("cached_filename", "Current Quote")

    if doc is None or not doc.items:
        st.warning("No active quote workbook found. Please switch to 'Quote Extractor' to upload a quote workbook first, or upload below.")
        uploaded_fallback = st.file_uploader("Upload Quote Workbook (.xlsx, .xlsm)", type=["xlsx", "xlsm"], key="linker_fallback_quote")
        if uploaded_fallback:
            try:
                engine = QuoteExtractorEngine()
                doc = engine.extract_file(io.BytesIO(uploaded_fallback.getvalue()), filename=uploaded_fallback.name)
                st.session_state["cached_doc"] = doc
                st.session_state["cached_filename"] = uploaded_fallback.name
                st.success(f"Loaded quote: {uploaded_fallback.name} ({len(doc.items)} items)")
                st.rerun()
            except Exception as e:
                st.error(f"Error parsing quote: {e}")
                return
        render_footer()
        return

    st.success(f"Active Quote Workbook: `{quote_filename}` ({len(doc.items)} line items available for matching)")

    # Input Container: Tabs for Paste or File Upload
    with st.container(border=True):
        st.markdown("#### Provide Job Numbers & Part Numbers")
        tab_paste, tab_upload = st.tabs(["Paste from Excel / ERP", "Upload Spreadsheet (.xlsx, .csv)"])

        jobs: List[JobItem] = []

        with tab_paste:
            st.caption("Copy two columns from Excel (`Job #` and `Part #`) and paste them directly below:")
            paste_text = st.text_area(
                "Paste Job# & Part# data:",
                height=140,
                placeholder="Job #\tPart Number\nJ1001\t2202046213000\nJ1002\t3001950000\nJ1003\t858007801000",
                key="paste_jobs_area",
            )
            if paste_text.strip():
                jobs = parse_job_input_text(paste_text)

        with tab_upload:
            uploaded_job_file = st.file_uploader(
                "Select Job file (.xlsx, .csv)",
                type=["xlsx", "xlsm", "csv"],
                key="upload_jobs_file",
            )
            if uploaded_job_file:
                try:
                    file_jobs = parse_job_input_file(uploaded_job_file.getvalue(), uploaded_job_file.name)
                    if file_jobs:
                        jobs = file_jobs
                        st.info(f"Loaded {len(jobs)} jobs from `{uploaded_job_file.name}`")
                except Exception as ex:
                    st.error(f"Failed to read file: {ex}")

    if not jobs:
        st.info("Paste or upload Job numbers above to view the aligned manufacturing table.")
        render_footer()
        return

    # Execute Linking
    link_res: LinkResult = link_jobs_to_quotes(jobs, doc.items)

    # 4 Summary KPI Cards
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Jobs Provided", link_res.total_jobs)
    with col2:
        st.metric("Successfully Linked", link_res.matched_count)
    with col3:
        st.metric("Jobs Missing Quotes", link_res.unmatched_jobs_count)
    with col4:
        st.metric("Quotes Missing Jobs", link_res.unmatched_quotes_count)

    # Build DataFrames
    linked_df = build_linked_dataframe(link_res.linked_records)
    unlinked_quotes_df = build_unlinked_quotes_dataframe(link_res.unlinked_quotes)

    # Display Tabs
    tab_linked, tab_unlinked = st.tabs([
        f"Linked Jobs ({len(linked_df)} in Exact Order)",
        f"Quotes Missing Jobs ({len(unlinked_quotes_df)})",
    ])

    with tab_linked:
        st.caption("Ordered strictly by Job Number in your exact input sequence. Jobs without matching quotes display with blank quote fields.")
        st.dataframe(
            linked_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Job Number": st.column_config.TextColumn("Job Number", required=True),
                "Part Number": st.column_config.TextColumn("Part Number"),
                "Quote Item": st.column_config.TextColumn("Quote Item"),
                "Description": st.column_config.TextColumn("Description"),
                "Material": st.column_config.TextColumn("Material"),
                "Cycle Sec": st.column_config.NumberColumn("Cycle Sec", format="%.2f"),
                "#Ops": st.column_config.NumberColumn("#Ops", format="%.2f"),
                "Labour Rate ($CAD/Hr)": st.column_config.NumberColumn("Labour Rate ($CAD/Hr)", format="$%.2f"),
                "Weight (g)": st.column_config.NumberColumn("Weight (g)", format="%.2f"),
                "Ann. Volume (EAU)": st.column_config.NumberColumn("Ann. Volume (EAU)", format="%d"),
                "Link Status": st.column_config.TextColumn("Link Status"),
                "Match Confidence": st.column_config.TextColumn("Match Confidence"),
            },
        )

    with tab_unlinked:
        if len(unlinked_quotes_df) == 0:
            st.success("All quote line items were successfully matched to Job numbers!")
        else:
            st.warning(f"Found {len(unlinked_quotes_df)} quote item(s) from `{quote_filename}` that did not match any Job Number.")
            st.dataframe(
                unlinked_quotes_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Quote Item": st.column_config.TextColumn("Quote Item"),
                    "Part Number": st.column_config.TextColumn("Part Number"),
                    "Description": st.column_config.TextColumn("Description"),
                    "Material": st.column_config.TextColumn("Material"),
                    "Cycle Sec": st.column_config.NumberColumn("Cycle Sec", format="%.2f"),
                    "#Ops": st.column_config.NumberColumn("#Ops", format="%.2f"),
                    "Labour Rate ($CAD/Hr)": st.column_config.NumberColumn("Labour Rate ($CAD/Hr)", format="$%.2f"),
                    "Weight (g)": st.column_config.NumberColumn("Weight (g)", format="%.2f"),
                    "Ann. Volume (EAU)": st.column_config.NumberColumn("Ann. Volume (EAU)", format="%d"),
                    "Link Status": st.column_config.TextColumn("Link Status"),
                },
            )

    # Export Actions for Linked Jobs
    st.subheader("Export Linked Data")
    col_x1, col_x2, col_x3 = st.columns(3)

    # 1. Multi-sheet Excel
    multi_excel_bytes = export_linked_data_to_multi_sheet_excel(linked_df, unlinked_quotes_df)
    with col_x1:
        st.download_button(
            label="Download Excel (Multi-Sheet: Linked Jobs + Missing)",
            data=multi_excel_bytes,
            file_name=f"{doc.quote_id or 'quote'}_jobs_linked.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_download_linked_excel",
            use_container_width=True,
        )

    # 2. CSV
    csv_linked_str = linked_df.to_csv(index=False)
    with col_x2:
        st.download_button(
            label="Download CSV (Job Order)",
            data=csv_linked_str.encode("utf-8"),
            file_name=f"{doc.quote_id or 'quote'}_jobs_linked.csv",
            mime="text/csv",
            key="btn_download_linked_csv",
            use_container_width=True,
        )

    # 3. JSON
    linked_json = json.dumps(link_res.model_dump(), indent=2, ensure_ascii=False)
    with col_x3:
        st.download_button(
            label="Download JSON",
            data=linked_json.encode("utf-8"),
            file_name=f"{doc.quote_id or 'quote'}_jobs_linked.json",
            mime="application/json",
            key="btn_download_linked_json",
            use_container_width=True,
        )

    render_footer()


def main() -> None:
    """Main application loop."""
    initialize_page()

    workspace_mode = setup_sidebar()
    engine = QuoteExtractorEngine()

    if workspace_mode == "Job# - Quote# Linker":
        render_job_linker_workspace()
    else:
        render_quote_extractor_workspace(engine)


if __name__ == "__main__":
    main()
