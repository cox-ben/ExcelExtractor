"""
Quote Extractor Engine coordinator.

Orchestrates workbook loading, layout strategy selection, quote item parsing,
confidence computation, validation, and batch directory processing.
"""

from __future__ import annotations

import io
from pathlib import Path
import re
from typing import BinaryIO, List, Optional, Set, Union

from excel_extractor.config import ConfigManager, ExtractorConfig
from excel_extractor.locator import AnchorLocator
from excel_extractor.models import QuoteDocument, QuoteItem
from excel_extractor.reader import Workbook, load_workbook
from excel_extractor.strategies.grid_matrix import GridMatrixStrategy
from excel_extractor.strategies.multi_tab import MultiTabStrategy
from excel_extractor.strategies.single_item import SingleItemStrategy


class QuoteExtractorEngine:
    """
    Main extraction engine coordinating workbook ingestion and multi-item extraction.
    """

    def __init__(
        self,
        config: Optional[Union[str, Path, ConfigManager, ExtractorConfig]] = None,
    ):
        if isinstance(config, ConfigManager):
            self.config_manager = config
            self.config = config.config
        elif isinstance(config, ExtractorConfig):
            self.config_manager = ConfigManager()
            self.config = config
        elif isinstance(config, (str, Path)):
            self.config_manager = ConfigManager(config_path=config)
            self.config = self.config_manager.config
        else:
            self.config_manager = ConfigManager()
            self.config = self.config_manager.config

        self.grid_strategy = GridMatrixStrategy(self.config)
        self.multi_tab_strategy = MultiTabStrategy(self.config)
        self.single_item_strategy = SingleItemStrategy(self.config)

    def extract_file(
        self,
        file_path_or_bytes: Union[str, Path, bytes, BinaryIO],
        filename: Optional[str] = None,
    ) -> QuoteDocument:
        """
        Extracts quote metadata and line items from an Excel workbook.

        Parameters:
            file_path_or_bytes: Path to workbook file on disk or bytes stream.
            filename: Optional source file name override.

        Returns:
            Validated QuoteDocument conforming to data contracts.
        """
        # Determine source file name
        source_name = filename
        if isinstance(file_path_or_bytes, (str, Path)):
            p = Path(file_path_or_bytes)
            if not p.exists():
                raise FileNotFoundError(f"Quote file not found: {p}")
            if not source_name:
                source_name = p.name

        if not source_name:
            source_name = "in_memory_quote.xlsx"

        # Load workbook via resilient reader
        wb = load_workbook(file_path_or_bytes, data_only=True)

        # Check for completely empty workbook
        if not wb.worksheets or not any(ws.max_row > 0 for ws in wb.worksheets):
            return QuoteDocument(
                source_file=source_name,
                quote_id=None,
                total_items=0,
                items=[],
                warnings=["Workbook contains no populated worksheets."],
            )

        # 1. Discover base quote ID across visible sheets and filenames
        locator = AnchorLocator(self.config)
        global_quote_id: Optional[str] = None

        # Prioritize candidate/target sheets, then all visible sheets (skipping hidden)
        candidate_sheets = self.single_item_strategy.get_candidate_sheets(wb)
        sheets_to_search = candidate_sheets if candidate_sheets else [
            ws for ws in wb.worksheets if self.single_item_strategy.is_sheet_visible(ws)
        ]

        # Search cells for quote ID
        for ws in sheets_to_search:
            qid_found, _, _ = locator.find_quote_id(ws)
            if qid_found:
                m = re.search(r"\b(Q[0-9]{4,5})\b", qid_found)
                if m:
                    global_quote_id = m.group(1).upper()
                    break

        # Fallback to file name
        if not global_quote_id and source_name:
            m = re.search(r"\b(Q[0-9]{4,5})\b", source_name, re.IGNORECASE)
            if m:
                global_quote_id = m.group(1).upper()

        # 2. Select and execute extraction strategies
        extracted_items: List[QuoteItem] = []

        # Check for hybrid layout: both a 2D grid matrix AND multiple tabs
        has_grid = self.grid_strategy.can_handle(wb)
        has_multitab = self.multi_tab_strategy.can_handle(wb)

        if has_grid and has_multitab:
            # Hybrid quote: run grid on grid sheets and multi_tab on child tabs, then deduplicate
            grid_items = self.grid_strategy.extract(wb, quote_id=global_quote_id)
            tab_items = self.multi_tab_strategy.extract(wb, quote_id=global_quote_id)

            combined = grid_items + tab_items
            seen_items: Set[str] = set()
            for it in combined:
                if it.quote_item_number not in seen_items:
                    seen_items.add(it.quote_item_number)
                    extracted_items.append(it)
        elif has_multitab:
            extracted_items = self.multi_tab_strategy.extract(wb, quote_id=global_quote_id)
        elif has_grid:
            extracted_items = self.grid_strategy.extract(wb, quote_id=global_quote_id)
        else:
            extracted_items = self.single_item_strategy.extract(wb, quote_id=global_quote_id)

        # 3. Synchronize quote_id if items were extracted
        if extracted_items and not global_quote_id:
            m = re.search(r"\b(Q[0-9]{4,5})\b", extracted_items[0].quote_item_number)
            if m:
                global_quote_id = m.group(1).upper()

        doc_warnings: List[str] = []
        if len(extracted_items) == 0:
            doc_warnings.append("No quote items could be extracted from workbook.")

        return QuoteDocument(
            source_file=source_name,
            quote_id=global_quote_id,
            total_items=len(extracted_items),
            items=extracted_items,
            warnings=doc_warnings,
        )

    def extract_directory(
        self,
        dir_path: Union[str, Path],
        recursive: bool = True,
    ) -> List[QuoteDocument]:
        """
        Scans a directory for Excel files (.xlsx, .xlsm), strictly ignoring
        temporary lock files (~$*), extracts quotes, and returns aggregated documents.

        Parameters:
            dir_path: Directory path to scan.
            recursive: If True, scans all subdirectories recursively.

        Returns:
            List of QuoteDocument objects.
        """
        root_dir = Path(dir_path)
        if not root_dir.exists():
            raise FileNotFoundError(f"Directory not found: {root_dir}")

        pattern = "**/*" if recursive else "*"
        all_candidates = sorted(root_dir.glob(pattern))
        results: List[QuoteDocument] = []

        for p in all_candidates:
            if not p.is_file():
                continue

            # Strictly exclude Excel temporary lock files
            if p.name.startswith("~$"):
                continue

            # Only accept .xlsx and .xlsm
            if p.suffix.lower() not in (".xlsx", ".xlsm"):
                continue

            try:
                doc = self.extract_file(p)
                results.append(doc)
            except Exception as e:
                # Record error document rather than aborting batch run
                results.append(
                    QuoteDocument(
                        source_file=p.name,
                        quote_id=None,
                        total_items=0,
                        items=[],
                        warnings=[f"Extraction failed: {e}"],
                    )
                )

        return results
