"""
Single-item dedicated form extraction strategy.

Handles single-item quoting worksheets with key-value label offsets,
inferring item suffix '-1' if only base quote 'QXXXX' is found.
"""

from __future__ import annotations

import re
from typing import List, Optional

from excel_extractor.config import ExtractorConfig
from excel_extractor.locator import AnchorLocator
from excel_extractor.models import QuoteItem, SourceCellCoords
from excel_extractor.reader import Workbook, Worksheet
from excel_extractor.strategies.base import BaseStrategy


class SingleItemStrategy(BaseStrategy):
    """Strategy for single-item form layout workbooks."""

    def _is_ignored_sheet(self, title: str) -> bool:
        """Checks if a sheet title matches configured administrative ignore keywords."""
        norm = title.lower().strip()
        ignore_list = [s.lower().strip() for s in self.config.ignore_sheets]
        for ign in ignore_list:
            if norm == ign:
                return True
            # Handle plurals / singulars (e.g. lookup vs lookups, notes vs note)
            if norm == f"{ign}s" or norm == f"{ign}es" or norm.rstrip("s") == ign.rstrip("s"):
                return True
            if norm.startswith(f"{ign} ") or norm.endswith(f" {ign}") or f" {ign} " in norm:
                return True
            if norm.startswith(f"{ign}s ") or norm.endswith(f" {ign}s") or f" {ign}s " in norm:
                return True
            # Word boundary regex check for compound tab names (e.g. "Project Cover Sheet")
            base_ign = ign.rstrip("s")
            if re.search(rf"\b{re.escape(base_ign)}s?\b", norm):
                return True
        return False

    def can_handle(self, workbook: Workbook) -> bool:
        """Single-item strategy applies when workbook has 1 sheet or exactly 1 valid candidate sheet."""
        if len(workbook.worksheets) == 1:
            return True
        valid_sheets = self.get_candidate_sheets(workbook)
        return len(valid_sheets) == 1

    def extract(self, workbook: Workbook, quote_id: Optional[str] = None) -> List[QuoteItem]:
        """Extracts single quote item from the primary sheet."""
        if not workbook.worksheets:
            return []

        locator = AnchorLocator(self.config)

        # Resolve base quote ID from visible sheets if not provided
        if not quote_id:
            for w in workbook.worksheets:
                if not self.is_sheet_visible(w):
                    continue
                qid, _, _ = locator.find_quote_id(w)
                if qid:
                    m = re.search(r"\b(Q[0-9]{4,5})\b", qid)
                    if m:
                        quote_id = m.group(1).upper()
                        break

        # Identify candidate worksheets
        candidate_sheets = self.get_candidate_sheets(workbook)
        if not candidate_sheets:
            if workbook.worksheets:
                # Fallback to first visible sheet
                visible = [w for w in workbook.worksheets if self.is_sheet_visible(w)]
                candidate_sheets = visible if visible else [workbook.worksheets[0]]
            else:
                return []

        # Select candidate worksheet: inspect candidate sheets for quote items / anchors
        selected_ws = candidate_sheets[0]
        selected_qid_info = locator.find_quote_id(selected_ws)
        selected_pn_info = locator.find_part_number(selected_ws)
        selected_ct_info = locator.find_cycle_time(selected_ws)

        def _has_item_anchors(pn_val: Optional[str], ct_val: Optional[float]) -> bool:
            return pn_val is not None or ct_val is not None

        def _has_any_anchors(qid_val: Optional[str], pn_val: Optional[str], ct_val: Optional[float]) -> bool:
            return qid_val is not None or pn_val is not None or ct_val is not None

        # If the selected sheet has no quote item anchors (pn or ct), inspect other candidate sheets
        if not _has_item_anchors(selected_pn_info[0], selected_ct_info[0]):
            for other_ws in candidate_sheets[1:]:
                other_qid_info = locator.find_quote_id(other_ws)
                other_pn_info = locator.find_part_number(other_ws)
                other_ct_info = locator.find_cycle_time(other_ws)

                if _has_item_anchors(other_pn_info[0], other_ct_info[0]):
                    selected_ws = other_ws
                    selected_qid_info = other_qid_info
                    selected_pn_info = other_pn_info
                    selected_ct_info = other_ct_info
                    break
                elif not _has_any_anchors(selected_qid_info[0], selected_pn_info[0], selected_ct_info[0]) and _has_any_anchors(other_qid_info[0], other_pn_info[0], other_ct_info[0]):
                    selected_ws = other_ws
                    selected_qid_info = other_qid_info
                    selected_pn_info = other_pn_info
                    selected_ct_info = other_ct_info

        ws = selected_ws
        if ws.max_row == 0:
            return []

        qid_found, qid_coord, is_inferred = selected_qid_info
        pn, pn_coord = selected_pn_info
        ct, ct_coord = selected_ct_info
        desc, desc_coord = locator.find_description(ws)
        mat, mat_coord = locator.find_material(ws)
        ops, ops_coord = locator.find_ops_labour(ws)
        l_rate, l_rate_coord = locator.find_labour_rate(ws)
        weight, weight_coord = locator.find_weight(ws)
        vol, vol_coord = locator.find_annual_volume(ws)

        target_qid = qid_found or quote_id

        # Negative control check: If no quote ID, no part number, and no cycle time, return []
        if not target_qid and pn is None and ct is None:
            return []

        # 3. Determine canonical quote item number
        warnings = []
        if target_qid:
            m_item = re.search(r"\b(Q[0-9]{4,5}(?:-[A-Za-z0-9]+)+)\b", target_qid)
            if m_item:
                item_number = m_item.group(1).upper()
            else:
                m_base = re.search(r"\b(Q[0-9]{4,5})\b", target_qid)
                base_val = m_base.group(1).upper() if m_base else target_qid
                item_number = f"{base_val}-1"
                warnings.append(f"Inferred quote item number suffix '-1' from base quote {base_val}")
        else:
            # Fallback if quote ID was not detected
            item_number = "Q0000-1"
            warnings.append("Inferred fallback quote item number 'Q0000-1'")

        # 4. Source cell coordinates
        q_coord_str = qid_coord or "B2"
        pn_coord_str = pn_coord
        ct_coord_str = ct_coord
        desc_coord_str = desc_coord
        mat_coord_str = mat_coord
        ops_coord_str = ops_coord
        lr_coord_str = l_rate_coord
        w_coord_str = weight_coord
        vol_coord_str = vol_coord

        if len(workbook.worksheets) > 1:
            q_coord_str = f"{ws.title}!{q_coord_str}"
            if pn_coord_str: pn_coord_str = f"{ws.title}!{pn_coord_str}"
            if ct_coord_str: ct_coord_str = f"{ws.title}!{ct_coord_str}"
            if desc_coord_str: desc_coord_str = f"{ws.title}!{desc_coord_str}"
            if mat_coord_str: mat_coord_str = f"{ws.title}!{mat_coord_str}"
            if ops_coord_str: ops_coord_str = f"{ws.title}!{ops_coord_str}"
            if lr_coord_str: lr_coord_str = f"{ws.title}!{lr_coord_str}"
            if w_coord_str: w_coord_str = f"{ws.title}!{w_coord_str}"
            if vol_coord_str: vol_coord_str = f"{ws.title}!{vol_coord_str}"

        coords = SourceCellCoords(
            quote_num=q_coord_str,
            part_num=pn_coord_str,
            description=desc_coord_str,
            material=mat_coord_str,
            cycle_time=ct_coord_str,
            ops_labour=ops_coord_str,
            labour_rate=lr_coord_str,
            weight_g=w_coord_str,
            annual_volume=vol_coord_str,
        )

        item = QuoteItem(
            quote_item_number=item_number,
            part_number=pn,
            description=desc,
            material=mat,
            cycle_time_sec=ct,
            ops_labour=ops,
            labour_rate=l_rate,
            weight_g=weight,
            annual_volume=vol,
            source_cell_coords=coords,
            warnings=warnings,
        )

        return [item]
