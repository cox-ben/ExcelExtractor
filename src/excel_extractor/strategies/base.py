"""
Base extraction strategy interface for Excel Quote Extractor.
"""

from __future__ import annotations

import functools
import re
from abc import ABC, abstractmethod
from typing import List, Optional, Tuple

from excel_extractor.config import ExtractorConfig
from excel_extractor.models import QuoteItem
from excel_extractor.reader import Workbook, Worksheet


class BaseStrategy(ABC):
    """Abstract base strategy for extracting quote items from a workbook."""

    def __init__(self, config: Optional[ExtractorConfig] = None):
        self.config = config or ExtractorConfig()
        self._cached_ignore_tuple: Optional[Tuple[str, ...]] = None
        self._cached_target_tuple: Optional[Tuple[str, ...]] = None

    def is_sheet_visible(self, ws: Worksheet) -> bool:
        """Checks if a worksheet is visible (not hidden or veryHidden)."""
        if hasattr(ws, "is_visible"):
            return bool(ws.is_visible)
        state = getattr(ws, "sheet_state", getattr(ws, "state", "visible"))
        if isinstance(state, str) and state.lower() in ("hidden", "veryhidden"):
            return False
        if getattr(ws, "is_hidden", False) or getattr(ws, "hidden", False):
            return False
        return True

    @staticmethod
    @functools.lru_cache(maxsize=1024)
    def _is_ignored_sheet_cached(norm: str, ignore_tuple: Tuple[str, ...]) -> bool:
        for ign in ignore_tuple:
            if not ign:
                continue
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

    def is_ignored_sheet(self, title: str) -> bool:
        """Checks if a sheet title matches configured administrative ignore keywords."""
        norm = title.lower().strip()
        if self._cached_ignore_tuple is None:
            self._cached_ignore_tuple = tuple(s.lower().strip() for s in self.config.ignore_sheets)
        return self._is_ignored_sheet_cached(norm, self._cached_ignore_tuple)

    @staticmethod
    @functools.lru_cache(maxsize=1024)
    def _is_target_sheet_cached(norm: str, target_tuple: Tuple[str, ...]) -> bool:
        for kw_norm in target_tuple:
            if not kw_norm:
                continue
            if norm == kw_norm:
                return True
            if norm.startswith(f"{kw_norm} ") or norm.startswith(f"{kw_norm}-") or norm.startswith(f"{kw_norm}_"):
                return True
            if norm.endswith(f" {kw_norm}") or norm.endswith(f"-{kw_norm}") or norm.endswith(f"_{kw_norm}"):
                return True
            if f" {kw_norm} " in norm or f"-{kw_norm}-" in norm or f"_{kw_norm}_" in norm:
                return True
            if re.search(rf"\b{re.escape(kw_norm)}\b", norm):
                return True

        # Fallback regex for MPC and QInfo / Quote Info variations
        if re.search(r"(?i)\bmpc\b|^mpc[-_ ]?[0-9]*", norm):
            return True
        if re.search(r"(?i)\bq[-_ ]?info\b|\bquote[-_ ]?info\b|\bquoteinfo\b", norm):
            return True

        return False

    def is_target_sheet(self, title: str) -> bool:
        """Checks if a sheet title matches MPC quote target keywords ('MPC', 'mpc', 'QInfo', 'Quote Info', etc.)."""
        norm = title.lower().strip()
        if self._cached_target_tuple is None:
            target_keywords = getattr(self.config, "target_sheet_keywords", [
                "mpc", "qinfo", "quote info", "quoteinfo", "q-info", "quote_info"
            ])
            self._cached_target_tuple = tuple(kw.lower().strip() for kw in target_keywords)
        return self._is_target_sheet_cached(norm, self._cached_target_tuple)

    @staticmethod
    @functools.lru_cache(maxsize=1024)
    def score_sheet_title_mpc(title: str) -> int:
        """Scores how closely a sheet title matches 'MPC'."""
        norm = title.lower().strip()
        if not norm:
            return 0
        if norm == "mpc":
            return 1000
        # Specific quoting suffixes e.g. "mpc quote", "mpc quotes", "mpc_quote", "mpc-quote", "mpc rfq"
        if re.match(r"^mpc[-_ ]?(?:quote|rfq|pricing|master)", norm):
            return 800 - len(norm)
        # Tab with item index like "mpc 1", "mpc-1", "mpc_1"
        if re.match(r"^mpc[-_ ]*[0-9]+$", norm):
            return 400 - len(norm)
        # Tab starting with mpc
        if re.match(r"^mpc\b|^mpc[-_ ]", norm):
            return 300 - len(norm)
        if re.search(r"\bmpc\b", norm):
            return 200 - len(norm)
        if "mpc" in norm:
            return 100 - len(norm)
        return 0

    @staticmethod
    @functools.lru_cache(maxsize=1024)
    def score_sheet_title_qinfo(title: str) -> int:
        """Scores how closely a sheet title matches 'QInfo' / 'Quote Info'."""
        norm = title.lower().strip()
        if not norm:
            return 0
        target_exact = ("qinfo", "quote info", "quoteinfo", "q-info", "quote_info")
        if norm in target_exact:
            return 1000
        if re.match(r"^(?:q[-_ ]?info|quote[-_ ]?info)\b", norm):
            return 500 - len(norm)
        if re.search(r"\b(?:qinfo|quote info|quoteinfo|q-info)\b", norm):
            return 300 - len(norm)
        if "qinfo" in norm or "quote info" in norm or "quoteinfo" in norm:
            return 100 - len(norm)
        return 0

    def get_closest_mpc_sheet(self, workbook: Workbook) -> Optional[Worksheet]:
        """Finds the visible, non-ignored sheet that matches 'MPC' the closest."""
        scored = []
        for ws in workbook.worksheets:
            if not self.is_sheet_visible(ws) or self.is_ignored_sheet(ws.title):
                continue
            score = self.score_sheet_title_mpc(ws.title)
            if score > 0 and ws.max_row > 0:
                scored.append((score, ws))
        if not scored:
            return None
        scored.sort(key=lambda x: x[0], reverse=True)
        return scored[0][1]

    def get_closest_qinfo_sheet(self, workbook: Workbook) -> Optional[Worksheet]:
        """Finds the visible, non-ignored sheet that matches 'QInfo' the closest."""
        scored = []
        for ws in workbook.worksheets:
            if not self.is_sheet_visible(ws) or self.is_ignored_sheet(ws.title):
                continue
            score = self.score_sheet_title_qinfo(ws.title)
            if score > 0 and ws.max_row > 0:
                scored.append((score, ws))
        if not scored:
            return None
        scored.sort(key=lambda x: x[0], reverse=True)
        return scored[0][1]

    def get_candidate_sheets(self, workbook: Workbook) -> List[Worksheet]:
        """
        Filters and returns valid candidate quoting worksheets:
        1. If any sheet matches 'QInfo' / 'Quote Info', grabs information from the sheet matching 'QInfo' closest.
           (If all QInfo sheets are item-tab numbered sheets, returns all item tabs).
        2. If no 'QInfo' sheet, grabs information from the sheet matching 'MPC' closest.
        3. Otherwise, falls back to all visible, non-ignored sheets (e.g. multi-tab workbooks with Item 1, Item 2).
        """
        valid_visible = [
            ws for ws in workbook.worksheets
            if self.is_sheet_visible(ws)
            and not self.is_ignored_sheet(ws.title)
        ]

        if not valid_visible:
            return []

        # Filter out empty sheets
        valid_visible = [ws for ws in valid_visible if ws.max_row > 0]
        if not valid_visible:
            return []

        def _sheet_has_cycle_time(ws: Worksheet) -> bool:
            if hasattr(ws, "_cached_has_cycle_time"):
                return ws._cached_has_cycle_time
            ws._ensure_loaded()
            ct_kws = ("cycle time", "cycle sec", "cycle (sec)", "c/t", "sec/shot", "molding cycle")
            res = False
            for cell in ws._cells.values():
                if cell.row <= 60 and cell.column <= 50 and cell.value is not None:
                    s = str(cell.value).lower()
                    if any(kw in s for kw in ct_kws):
                        res = True
                        break
            ws._cached_has_cycle_time = res
            return res

        # 1. Grab information from the sheet that matches QInfo the closest
        best_qinfo = self.get_closest_qinfo_sheet(workbook)
        best_mpc = self.get_closest_mpc_sheet(workbook)

        # If QInfo exists but has NO cycle time, and MPC exists AND has cycle time, prefer MPC
        if best_qinfo is not None and best_mpc is not None:
            if not _sheet_has_cycle_time(best_qinfo) and _sheet_has_cycle_time(best_mpc):
                mpc_sheets = [ws for ws in valid_visible if self.score_sheet_title_mpc(ws.title) > 0]
                exact_mpc = [
                    ws for ws in mpc_sheets
                    if ws.title.lower().strip() in ("mpc", "mpc quote", "mpc quotes", "mpc_quote", "mpc-quote", "mpc rfq")
                ]
                if exact_mpc:
                    return [exact_mpc[0]]
                return [best_mpc]

        if best_qinfo is not None:
            qinfo_sheets = [ws for ws in valid_visible if self.score_sheet_title_qinfo(ws.title) > 0]
            exact_qinfo = [
                ws for ws in qinfo_sheets
                if ws.title.lower().strip() in ("qinfo", "quote info", "quoteinfo", "q-info", "quote_info")
            ]
            if exact_qinfo:
                return [exact_qinfo[0]]

            all_numbered = all(
                re.search(r"(?i)\b(?:item|qinfo|quote\s*info|part)[-_ ]*[0-9]+\b", ws.title)
                for ws in qinfo_sheets
            )
            if len(qinfo_sheets) > 1 and all_numbered:
                return qinfo_sheets

            return [best_qinfo]

        # 2. If no QInfo sheet, grab information from the sheet that matches MPC the closest
        if best_mpc is not None:
            mpc_sheets = [ws for ws in valid_visible if self.score_sheet_title_mpc(ws.title) > 0]
            # Exact match or primary quote title
            exact_mpc = [
                ws for ws in mpc_sheets
                if ws.title.lower().strip() in ("mpc", "mpc quote", "mpc quotes", "mpc_quote", "mpc-quote", "mpc rfq")
            ]
            if exact_mpc:
                return [exact_mpc[0]]

            # If all MPC sheets are item-tab numbered sheets (e.g. 'MPC 1', 'mpc 2'), return all item tabs
            all_numbered = all(
                re.search(r"(?i)\b(?:item|mpc|part)[-_ ]*[0-9]+\b", ws.title)
                for ws in mpc_sheets
            )
            if len(mpc_sheets) > 1 and all_numbered:
                return mpc_sheets

            return [best_mpc]

        # 3. Fallback to all valid visible non-ignored sheets
        return valid_visible

    @abstractmethod
    def can_handle(self, workbook: Workbook) -> bool:
        """Determines if this strategy can extract items from the provided workbook."""
        pass

    @abstractmethod
    def extract(self, workbook: Workbook, quote_id: Optional[str] = None) -> List[QuoteItem]:
        """Extracts and returns a list of QuoteItem objects from the workbook."""
        pass
