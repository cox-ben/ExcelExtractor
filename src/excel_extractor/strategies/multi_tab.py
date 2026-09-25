"""
Multi-tab multi-item extraction strategy.

Isolates quote items modeled across dedicated worksheets (sheet-per-item),
excluding administrative, summary, and lookup sheets.
"""

from __future__ import annotations

import re
from typing import List, Optional, Set

from excel_extractor.config import ExtractorConfig
from excel_extractor.locator import AnchorLocator
from excel_extractor.models import QuoteItem, SourceCellCoords
from excel_extractor.reader import Workbook, Worksheet
from excel_extractor.strategies.base import BaseStrategy
from excel_extractor.strategies.grid_matrix import GridMatrixStrategy


class MultiTabStrategy(BaseStrategy):
    """Strategy for multi-tab quoting workbooks where each tab models an item."""

    def can_handle(self, workbook: Workbook) -> bool:
        """Applies if workbook has multiple candidate quoting worksheets."""
        valid_sheets = self.get_candidate_sheets(workbook)
        return len(valid_sheets) > 1

    def extract(self, workbook: Workbook, quote_id: Optional[str] = None) -> List[QuoteItem]:
        """Extracts quote items across all candidate item sheets."""
        items: List[QuoteItem] = []
        locator = AnchorLocator(self.config)
        grid_strategy = GridMatrixStrategy(self.config)

        # 1. Resolve base quote ID from visible sheets if not provided
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

        # 2. Retrieve filtered candidate worksheets
        candidate_sheets = self.get_candidate_sheets(workbook)

        sheet_item_idx = 1
        for ws in candidate_sheets:
            # If this sheet contains a 2D grid matrix table, extract via grid strategy
            if grid_strategy.sheet_has_grid(ws):
                grid_wb = Workbook()
                grid_wb.worksheets = [ws]
                grid_items = grid_strategy.extract(grid_wb, quote_id=quote_id)
                items.extend(grid_items)
                continue

            # Otherwise, extract as single-item form sheet
            qid_found, qid_coord, is_inferred = locator.find_quote_id(ws)
            pn, pn_coord = locator.find_part_number(ws)
            ct, ct_coord = locator.find_cycle_time(ws)
            desc, desc_coord = locator.find_description(ws)
            mat, mat_coord = locator.find_material(ws)
            ops, ops_coord = locator.find_ops_labour(ws)
            l_rate, l_rate_coord = locator.find_labour_rate(ws)
            weight, weight_coord = locator.find_weight(ws)
            vol, vol_coord = locator.find_annual_volume(ws)

            # If sheet has no quote ID, no part number, and no cycle time, skip
            if not qid_found and pn is None and ct is None:
                continue

            # Determine quote item number
            item_num = None
            warnings = []

            # Priority A: Canonical pattern in sheet cells (e.g. Q2005-1, Q4476-R1-1)
            if qid_found and re.search(r"\b(Q[0-9]{4,5}(?:-[A-Za-z0-9]+)+)\b", qid_found):
                item_num = re.search(r"\b(Q[0-9]{4,5}(?:-[A-Za-z0-9]+)+)\b", qid_found).group(1).upper()
            # Priority B: Canonical pattern in sheet title (e.g. tab named "Q2005-1")
            elif re.search(r"\b(Q[0-9]{4,5}(?:-[A-Za-z0-9]+)+)\b", ws.title):
                item_num = re.search(r"\b(Q[0-9]{4,5}(?:-[A-Za-z0-9]+)+)\b", ws.title).group(1).upper()
            # Priority C: Sheet title like "Item 1", "Item 2", "MPC 1", "MPC-2", "QInfo 1", "Quote Info 2"
            elif re.search(r"(?i)\b(?:item|mpc|qinfo|quote\s*info|part)[-_ ]*([0-9]+)\b", ws.title):
                item_idx = re.search(r"(?i)\b(?:item|mpc|qinfo|quote\s*info|part)[-_ ]*([0-9]+)\b", ws.title).group(1)
                base = quote_id or (qid_found if qid_found and qid_found.startswith("Q") else "Q0000")
                item_num = f"{base}-{item_idx}"
            # Priority D: Base quote ID + sheet sequence index
            elif quote_id or (qid_found and qid_found.startswith("Q")):
                base = quote_id or qid_found
                item_num = f"{base}-{sheet_item_idx}"
            else:
                item_num = f"Q0000-{sheet_item_idx}"

            sheet_item_idx += 1

            # Scoped coordinates: prefix with sheet title
            q_coord_raw = qid_coord or "B2"
            q_coord_str = f"{ws.title}!{q_coord_raw}"
            pn_coord_str = f"{ws.title}!{pn_coord}" if pn_coord else None
            ct_coord_str = f"{ws.title}!{ct_coord}" if ct_coord else None
            desc_coord_str = f"{ws.title}!{desc_coord}" if desc_coord else None
            mat_coord_str = f"{ws.title}!{mat_coord}" if mat_coord else None
            ops_coord_str = f"{ws.title}!{ops_coord}" if ops_coord else None
            lr_coord_str = f"{ws.title}!{l_rate_coord}" if l_rate_coord else None
            w_coord_str = f"{ws.title}!{weight_coord}" if weight_coord else None
            vol_coord_str = f"{ws.title}!{vol_coord}" if vol_coord else None

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
                quote_item_number=item_num,
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
            items.append(item)

        # 3. Deduplicate items by quote_item_number (retaining first occurrence)
        seen_items: Set[str] = set()
        deduped_items: List[QuoteItem] = []
        for it in items:
            if it.quote_item_number not in seen_items:
                seen_items.add(it.quote_item_number)
                deduped_items.append(it)

        return deduped_items
