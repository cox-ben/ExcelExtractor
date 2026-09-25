"""Shared Pytest Fixtures and Environment Configuration for Excel Quote Extractor.

Provides self-contained, isolated temporary workbooks for all four layout archetypes,
edge cases, and batch directory structures without modifying global state.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path
from typing import Any, Dict, Generator, Optional
import pytest

# Ensure workspace root and src/ are in sys.path
WORKSPACE_ROOT = Path(__file__).resolve().parent.parent
SRC_DIR = WORKSPACE_ROOT / "src"
if str(WORKSPACE_ROOT) not in sys.path:
    sys.path.insert(0, str(WORKSPACE_ROOT))
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from tests.fixtures import generator


# ---------------------------------------------------------------------------
# Individual Workbook Fixtures (Scoped per-test in tmp_path)
# ---------------------------------------------------------------------------

@pytest.fixture
def single_modern_quote_file(tmp_path: Path) -> Path:
    """Generate quote_single_modern.xlsx in a temporary directory."""
    path = tmp_path / "quote_single_modern.xlsx"
    return generator.create_synthetic_modern_single(path)


@pytest.fixture
def multitab_quote_file(tmp_path: Path) -> Path:
    """Generate quote_multitab_q2005.xlsx in a temporary directory."""
    path = tmp_path / "quote_multitab_q2005.xlsx"
    return generator.create_synthetic_multitab(path)


@pytest.fixture
def grid_matrix_quote_file(tmp_path: Path) -> Path:
    """Generate quote_grid_matrix_q4000.xlsx in a temporary directory."""
    path = tmp_path / "quote_grid_matrix_q4000.xlsx"
    return generator.create_synthetic_grid_matrix(path)


@pytest.fixture
def merged_cells_quote_file(tmp_path: Path) -> Path:
    """Generate quote_merged_cells_q5000.xlsx in a temporary directory."""
    path = tmp_path / "quote_merged_cells_q5000.xlsx"
    return generator.create_synthetic_merged_cells(path)


@pytest.fixture
def legacy_raw_quote_file(tmp_path: Path) -> Path:
    """Generate quote_legacy_raw_q3050.xlsm in a temporary directory."""
    path = tmp_path / "quote_legacy_raw_q3050.xlsm"
    return generator.create_synthetic_legacy_raw(path)


@pytest.fixture
def edge_cases_quote_file(tmp_path: Path) -> Path:
    """Generate quote_edge_cases_q6000.xlsx in a temporary directory."""
    path = tmp_path / "quote_edge_cases_q6000.xlsx"
    return generator.create_synthetic_edge_cases(path)


@pytest.fixture
def empty_quote_file(tmp_path: Path) -> Path:
    """Generate quote_empty.xlsx in a temporary directory."""
    path = tmp_path / "quote_empty.xlsx"
    return generator.create_synthetic_empty_workbook(path)


@pytest.fixture
def non_quote_file(tmp_path: Path) -> Path:
    """Generate non_quote_spreadsheet.xlsx in a temporary directory."""
    path = tmp_path / "non_quote_spreadsheet.xlsx"
    return generator.create_synthetic_non_quote_workbook(path)


@pytest.fixture
def scientific_notation_quote_file(tmp_path: Path) -> Path:
    """Generate quote_scientific_notation.xlsx in a temporary directory."""
    path = tmp_path / "quote_scientific_notation.xlsx"
    return generator.create_synthetic_scientific_notation(path)


@pytest.fixture
def whitespace_variation_quote_file(tmp_path: Path) -> Path:
    """Generate quote_whitespace_variation.xlsx in a temporary directory."""
    path = tmp_path / "quote_whitespace_variation.xlsx"
    return generator.create_synthetic_whitespace_variation(path)


@pytest.fixture
def dwg_disambiguation_quote_file(tmp_path: Path) -> Path:
    """Generate quote_dwg_disambiguation.xlsx in a temporary directory."""
    path = tmp_path / "quote_dwg_disambiguation.xlsx"
    return generator.create_synthetic_dwg_disambiguation(path)


@pytest.fixture
def hybrid_quote_file(tmp_path: Path) -> Path:
    """Generate quote_hybrid.xlsx in a temporary directory."""
    path = tmp_path / "quote_hybrid.xlsx"
    return generator.create_synthetic_hybrid_quote(path)


@pytest.fixture
def all_synthetic_fixtures(tmp_path: Path) -> Dict[str, Path]:
    """Generate all 12 synthetic fixtures into a dedicated temp folder."""
    fixtures_dir = tmp_path / "all_fixtures"
    return generator.create_all_fixtures(fixtures_dir)


@pytest.fixture
def batch_directory_fixture(tmp_path: Path) -> Path:
    """Create a realistic directory tree with valid quotes and temp lock files (~$*).

    Structure:
    tmp_path/batch_run/
    ├── Q1001_single.xlsx
    ├── Q2005_multitab.xlsx
    ├── Q4000_grid.xlsx
    ├── ~$Q1001_single.xlsx           <-- Lock file, MUST BE FILTERED
    ├── subfolder_legacy/
    │   ├── Q3050_macro.xlsm
    │   └── ~$Q3050_macro.xlsm        <-- Lock file, MUST BE FILTERED
    └── subfolder_auxiliary/
        └── inventory.xlsx            <-- Non-quote file
    """
    batch_dir = tmp_path / "batch_run"
    batch_dir.mkdir(parents=True, exist_ok=True)
    sub_legacy = batch_dir / "subfolder_legacy"
    sub_legacy.mkdir(parents=True, exist_ok=True)
    sub_aux = batch_dir / "subfolder_auxiliary"
    sub_aux.mkdir(parents=True, exist_ok=True)

    # Valid files
    generator.create_synthetic_modern_single(batch_dir / "Q1001_single.xlsx")
    generator.create_synthetic_multitab(batch_dir / "Q2005_multitab.xlsx")
    generator.create_synthetic_grid_matrix(batch_dir / "Q4000_grid.xlsx")
    generator.create_synthetic_legacy_raw(sub_legacy / "Q3050_macro.xlsm")
    generator.create_synthetic_non_quote_workbook(sub_aux / "inventory.xlsx")

    # Excel temporary lock files (should be ignored by CLI and batch scanner)
    (batch_dir / "~$Q1001_single.xlsx").write_bytes(b"temp lock bytes")
    (sub_legacy / "~$Q3050_macro.xlsm").write_bytes(b"temp lock bytes")

    return batch_dir


# ---------------------------------------------------------------------------
# Engine / Model Import Helper
# ---------------------------------------------------------------------------

def get_engine():
    """Import and return QuoteExtractorEngine class if implemented, else None."""
    try:
        from excel_extractor.engine import QuoteExtractorEngine
        return QuoteExtractorEngine
    except ImportError:
        try:
            from src.excel_extractor.engine import QuoteExtractorEngine
            return QuoteExtractorEngine
        except ImportError:
            return None


def get_models():
    """Import and return (QuoteDocument, QuoteItem) classes if implemented, else None."""
    try:
        from excel_extractor.models import QuoteDocument, QuoteItem, SourceCellCoords
        return QuoteDocument, QuoteItem, SourceCellCoords
    except ImportError:
        try:
            from src.excel_extractor.models import QuoteDocument, QuoteItem, SourceCellCoords
            return QuoteDocument, QuoteItem, SourceCellCoords
        except ImportError:
            return None, None, None


@pytest.fixture
def extractor_engine():
    """Provide QuoteExtractorEngine instance or skip test if not yet implemented."""
    engine_cls = get_engine()
    if engine_cls is None:
        pytest.skip("QuoteExtractorEngine is not implemented yet (Milestone 2 dependency).")
    return engine_cls()
