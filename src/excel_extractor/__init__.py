"""
Excel Quote Extractor package.

Automated extraction utility for legacy Excel (.xlsx/.xlsm) quote workbooks,
supporting single-item, multi-tab, and 2D grid matrix layouts.
"""

from excel_extractor.config import (
    AnchorConfig,
    BoundsConfig,
    ConfigManager,
    ExtractorConfig,
    GridMatrixConfig,
    QuoteIdConfig,
)
from excel_extractor.engine import QuoteExtractorEngine
from excel_extractor.locator import AnchorLocator
from excel_extractor.models import (
    QuoteDocument,
    QuoteItem,
    SourceCellCoords,
)
from excel_extractor.reader import (
    Cell,
    MergedCellRange,
    Workbook,
    Worksheet,
    load_workbook,
)

__version__ = "3.0.0"

__all__ = [
    "SourceCellCoords",
    "QuoteItem",
    "QuoteDocument",
    "ConfigManager",
    "ExtractorConfig",
    "AnchorConfig",
    "BoundsConfig",
    "QuoteIdConfig",
    "GridMatrixConfig",
    "QuoteExtractorEngine",
    "AnchorLocator",
    "load_workbook",
    "Workbook",
    "Worksheet",
    "Cell",
    "MergedCellRange",
    "__version__",
]
