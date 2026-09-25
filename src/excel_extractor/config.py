"""
Configuration management for the Excel Quote Extractor.

Loads and validates configuration from config.yaml using PyYAML and Pydantic V2,
with resilient built-in defaults if the configuration file is missing, empty, or corrupted.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import yaml
from pydantic import BaseModel, ConfigDict, Field

logger = logging.getLogger(__name__)

# Canonical defaults matching MPC domain specifications
DEFAULT_QUOTE_ID_PATTERNS = [r"\b(Q[0-9]{4,5})\b"]
DEFAULT_QUOTE_ITEM_PATTERNS = [r"\b(Q[0-9]{4,5}-[0-9]+)\b"]
DEFAULT_QUOTE_ID_ANCHORS = [
    "quote #",
    "quote number",
    "quote id",
    "quote no",
    "quote no.",
    "rfq #",
    "quote ref",
]

DEFAULT_PART_NUMBER_ANCHORS = [
    "part number",
    "part no",
    "part #",
    "p/n",
    "customer part",
    "customer p/n",
]

DEFAULT_CYCLE_TIME_ANCHORS = [
    "cycle time",
    "cycle time (sec)",
    "cycle time (s)",
    "cycle sec",
    "c/t",
    "sec/shot",
    "molding cycle",
    "cycle (sec)",
    "est cycle time",
]

DEFAULT_DESCRIPTION_ANCHORS = [
    "part description",
    "part desc",
    "part desc.",
    "part discription",
    "description",
    "discription",
    "desc",
    "desc.",
]

DEFAULT_MATERIAL_ANCHORS = [
    "resin grade",
    "material grade",
    "grade",
    "material",
    "resin",
    "material description",
    "raw material",
]

DEFAULT_OPS_LABOUR_ANCHORS = [
    "#ops",
    "# ops",
    "ops",
    "labour",
    "labor",
    "operators",
    "operator",
    "# operators",
]

DEFAULT_LABOUR_RATE_ANCHORS = [
    "rate $cad/hr",
    "rate cad/hr",
    "rate $/hr",
    "rate $cad / hr",
    "$cad/hr",
    "cad/hr",
    "labour rate",
    "labor rate",
]

DEFAULT_WEIGHT_ANCHORS = [
    "weight (g)",
    "weight(g)",
    "part weight (g)",
    "part wt (g)",
    "part weight",
    "weight",
    "wt (g)",
]

DEFAULT_ANNUAL_VOLUME_ANCHORS = [
    "ann. volume (eau)",
    "ann. volume",
    "annual volume (eau)",
    "annual volume",
    "volume (eau)",
    "volume",
    "eau",
    "ann volume",
]

DEFAULT_OFFSETS = [[0, 1], [0, 2], [1, 0]]
DEFAULT_MIN_CYCLE_TIME = 5.0
DEFAULT_MAX_CYCLE_TIME = 300.0
DEFAULT_MIN_OPS_LABOUR = 0.0
DEFAULT_MAX_OPS_LABOUR = 1.0
DEFAULT_IGNORE_SHEETS = [
    "summary",
    "cover",
    "toc",
    "terms",
    "notes",
    "lookup",
    "instructions",
    "data",
    "master",
    "template",
]

DEFAULT_TARGET_SHEET_KEYWORDS = [
    "mpc",
    "qinfo",
    "quote info",
    "quoteinfo",
    "q-info",
    "quote_info",
]

DEFAULT_TARGET_SHEET_PATTERNS = [
    r"(?i)\bmpc\b",
    r"(?i)^mpc.*",
    r"(?i)\bq[-_ ]?info\b",
    r"(?i)\bquote[-_ ]?info\b",
    r"(?i)\bquoteinfo\b",
    r"(?i)\bQ[0-9]{4,5}(?:-[0-9]+)?\b",
    r"(?i)\b(?:item|part)[-_ ]?[0-9]+\b",
]


class AnchorConfig(BaseModel):
    """Configuration for an anchor search label and directional search offsets."""
    keywords: List[str] = Field(default_factory=list)
    search_offsets: List[List[int]] = Field(default_factory=lambda: list(DEFAULT_OFFSETS))
    clean_patterns: List[str] = Field(default_factory=list)

    model_config = ConfigDict(extra="ignore")


class BoundsConfig(BaseModel):
    """Numeric bounds configuration for manufacturing metrics (e.g. cycle time)."""
    min: float = DEFAULT_MIN_CYCLE_TIME
    max: float = DEFAULT_MAX_CYCLE_TIME

    model_config = ConfigDict(extra="ignore")


class QuoteIdConfig(BaseModel):
    """Regex patterns and anchors for discovering Quote IDs and Item tags."""
    patterns: List[str] = Field(default_factory=lambda: list(DEFAULT_QUOTE_ID_PATTERNS))
    item_patterns: List[str] = Field(default_factory=lambda: list(DEFAULT_QUOTE_ITEM_PATTERNS))
    anchors: List[str] = Field(default_factory=lambda: list(DEFAULT_QUOTE_ID_ANCHORS))
    search_offsets: List[List[int]] = Field(default_factory=lambda: list(DEFAULT_OFFSETS))

    model_config = ConfigDict(extra="ignore")


class GridMatrixConfig(BaseModel):
    """Heuristic settings for 2D tabular row-column grid extraction."""
    enabled: bool = True
    header_max_row: int = 100
    stop_keywords: List[str] = Field(
        default_factory=lambda: [
            "total",
            "subtotal",
            "notes",
            "approved",
            "revision",
            "revisions",
            "quote revision",
        ]
    )

    model_config = ConfigDict(extra="ignore")


class SheetsConfig(BaseModel):
    """Configuration for sheet filtering, ignoring, and targeting."""
    ignore: List[str] = Field(default_factory=lambda: list(DEFAULT_IGNORE_SHEETS))
    target_keywords: List[str] = Field(default_factory=lambda: list(DEFAULT_TARGET_SHEET_KEYWORDS))
    target_patterns: List[str] = Field(default_factory=lambda: list(DEFAULT_TARGET_SHEET_PATTERNS))

    model_config = ConfigDict(extra="ignore")


class ExtractorConfig(BaseModel):
    """Complete extraction engine configuration."""
    version: str = "1.0"
    quote_id: QuoteIdConfig = Field(default_factory=QuoteIdConfig)
    anchors: Dict[str, AnchorConfig] = Field(
        default_factory=lambda: {
            "part_number": AnchorConfig(
                keywords=list(DEFAULT_PART_NUMBER_ANCHORS),
                search_offsets=list(DEFAULT_OFFSETS),
                clean_patterns=[r"(?i)dwg\s*#?:?.*", r"(?i)drawing.*"],
            ),
            "cycle_time": AnchorConfig(
                keywords=list(DEFAULT_CYCLE_TIME_ANCHORS),
                search_offsets=list(DEFAULT_OFFSETS),
            ),
            "description": AnchorConfig(
                keywords=list(DEFAULT_DESCRIPTION_ANCHORS),
                search_offsets=list(DEFAULT_OFFSETS),
            ),
            "material": AnchorConfig(
                keywords=list(DEFAULT_MATERIAL_ANCHORS),
                search_offsets=list(DEFAULT_OFFSETS),
            ),
            "ops_labour": AnchorConfig(
                keywords=list(DEFAULT_OPS_LABOUR_ANCHORS),
                search_offsets=list(DEFAULT_OFFSETS),
            ),
            "labour_rate": AnchorConfig(
                keywords=list(DEFAULT_LABOUR_RATE_ANCHORS),
                search_offsets=[[1, 0], [2, 0]],
            ),
            "weight_g": AnchorConfig(
                keywords=list(DEFAULT_WEIGHT_ANCHORS),
                search_offsets=list(DEFAULT_OFFSETS),
            ),
            "annual_volume": AnchorConfig(
                keywords=list(DEFAULT_ANNUAL_VOLUME_ANCHORS),
                search_offsets=list(DEFAULT_OFFSETS),
            ),
        }
    )
    bounds: Dict[str, BoundsConfig] = Field(
        default_factory=lambda: {
            "cycle_time_sec": BoundsConfig(min=DEFAULT_MIN_CYCLE_TIME, max=DEFAULT_MAX_CYCLE_TIME),
            "ops_labour": BoundsConfig(min=DEFAULT_MIN_OPS_LABOUR, max=DEFAULT_MAX_OPS_LABOUR),
        }
    )
    sheets_ignore: List[str] = Field(default_factory=lambda: list(DEFAULT_IGNORE_SHEETS))
    sheets_target_keywords: List[str] = Field(
        default_factory=lambda: list(DEFAULT_TARGET_SHEET_KEYWORDS)
    )
    sheets_target_patterns: List[str] = Field(
        default_factory=lambda: list(DEFAULT_TARGET_SHEET_PATTERNS)
    )
    grid_matrix: GridMatrixConfig = Field(default_factory=GridMatrixConfig)
    config_source: str = "defaults"

    model_config = ConfigDict(extra="ignore")

    @property
    def quote_id_patterns(self) -> List[str]:
        return self.quote_id.patterns

    @property
    def quote_item_patterns(self) -> List[str]:
        return self.quote_id.item_patterns

    @property
    def quote_id_anchors(self) -> List[str]:
        return self.quote_id.anchors

    @property
    def part_number_anchors(self) -> List[str]:
        if "part_number" in self.anchors:
            return self.anchors["part_number"].keywords
        return list(DEFAULT_PART_NUMBER_ANCHORS)

    @property
    def cycle_time_anchors(self) -> List[str]:
        if "cycle_time" in self.anchors:
            return self.anchors["cycle_time"].keywords
        return list(DEFAULT_CYCLE_TIME_ANCHORS)

    @property
    def part_number_offsets(self) -> List[List[int]]:
        if "part_number" in self.anchors:
            return self.anchors["part_number"].search_offsets
        return list(DEFAULT_OFFSETS)

    @property
    def cycle_time_offsets(self) -> List[List[int]]:
        if "cycle_time" in self.anchors:
            return self.anchors["cycle_time"].search_offsets
        return list(DEFAULT_OFFSETS)

    @property
    def min_cycle_time(self) -> float:
        if "cycle_time_sec" in self.bounds:
            return self.bounds["cycle_time_sec"].min
        return DEFAULT_MIN_CYCLE_TIME

    @property
    def max_cycle_time(self) -> float:
        if "cycle_time_sec" in self.bounds:
            return self.bounds["cycle_time_sec"].max
        return DEFAULT_MAX_CYCLE_TIME

    @property
    def ignore_sheets(self) -> List[str]:
        return self.sheets_ignore

    @property
    def target_sheet_keywords(self) -> List[str]:
        return self.sheets_target_keywords

    @property
    def target_sheet_patterns(self) -> List[str]:
        return self.sheets_target_patterns

    @property
    def description_anchors(self) -> List[str]:
        if "description" in self.anchors:
            return self.anchors["description"].keywords
        return list(DEFAULT_DESCRIPTION_ANCHORS)

    @property
    def description_offsets(self) -> List[List[int]]:
        if "description" in self.anchors:
            return self.anchors["description"].search_offsets
        return list(DEFAULT_OFFSETS)

    @property
    def material_anchors(self) -> List[str]:
        if "material" in self.anchors:
            return self.anchors["material"].keywords
        return list(DEFAULT_MATERIAL_ANCHORS)

    @property
    def material_offsets(self) -> List[List[int]]:
        if "material" in self.anchors:
            return self.anchors["material"].search_offsets
        return list(DEFAULT_OFFSETS)

    @property
    def ops_labour_anchors(self) -> List[str]:
        if "ops_labour" in self.anchors:
            return self.anchors["ops_labour"].keywords
        return list(DEFAULT_OPS_LABOUR_ANCHORS)

    @property
    def ops_labour_offsets(self) -> List[List[int]]:
        if "ops_labour" in self.anchors:
            return self.anchors["ops_labour"].search_offsets
        return list(DEFAULT_OFFSETS)

    @property
    def labour_rate_anchors(self) -> List[str]:
        if "labour_rate" in self.anchors:
            return self.anchors["labour_rate"].keywords
        return list(DEFAULT_LABOUR_RATE_ANCHORS)

    @property
    def labour_rate_offsets(self) -> List[List[int]]:
        if "labour_rate" in self.anchors:
            return self.anchors["labour_rate"].search_offsets
        return [[1, 0], [2, 0]]

    @property
    def weight_anchors(self) -> List[str]:
        if "weight_g" in self.anchors:
            return self.anchors["weight_g"].keywords
        return list(DEFAULT_WEIGHT_ANCHORS)

    @property
    def weight_offsets(self) -> List[List[int]]:
        if "weight_g" in self.anchors:
            return self.anchors["weight_g"].search_offsets
        return list(DEFAULT_OFFSETS)

    @property
    def annual_volume_anchors(self) -> List[str]:
        if "annual_volume" in self.anchors:
            return self.anchors["annual_volume"].keywords
        return list(DEFAULT_ANNUAL_VOLUME_ANCHORS)

    @property
    def annual_volume_offsets(self) -> List[List[int]]:
        if "annual_volume" in self.anchors:
            return self.anchors["annual_volume"].search_offsets
        return list(DEFAULT_OFFSETS)


class ConfigManager:
    """
    Loads, parses, and provides access to extractor configuration.
    Falls back gracefully to robust domain defaults if the configuration file is
    missing, inaccessible, or corrupted.
    """

    def __init__(self, config_path: Optional[Union[str, Path]] = None):
        self._config_path: Optional[Path] = Path(config_path) if config_path else None
        self._config: ExtractorConfig = self.load(self._config_path)

    @property
    def config(self) -> ExtractorConfig:
        return self._config

    @classmethod
    def get_default_config(cls) -> ExtractorConfig:
        """Returns a freshly instantiated ExtractorConfig with default domain parameters."""
        return ExtractorConfig(config_source="defaults")

    def reload(self) -> ExtractorConfig:
        """Reloads configuration from the stored path."""
        self._config = self.load(self._config_path)
        return self._config

    def load(self, config_path: Optional[Union[str, Path]] = None) -> ExtractorConfig:
        """
        Loads configuration from YAML file or returns default configuration.
        """
        target_path = self._resolve_config_path(config_path)
        if target_path is None or not target_path.exists():
            logger.info("Configuration file not found. Utilizing built-in MPC defaults.")
            return self.get_default_config()

        try:
            with open(target_path, "r", encoding="utf-8") as f:
                raw_data = yaml.safe_load(f)

            if not isinstance(raw_data, dict):
                logger.warning(
                    f"Config at {target_path} is not a valid YAML mapping. Using defaults."
                )
                return self.get_default_config()

            return self._parse_raw_dict(raw_data, source=str(target_path))

        except Exception as e:
            logger.warning(
                f"Failed to load or parse config from {target_path}: {e}. Falling back to defaults."
            )
            return self.get_default_config()

    def _resolve_config_path(self, config_path: Optional[Union[str, Path]]) -> Optional[Path]:
        """Resolves target config path checking provided path, CWD, and project root."""
        if config_path is not None:
            p = Path(config_path)
            return p

        # Check standard default locations
        cwd_candidate = Path.cwd() / "config.yaml"
        if cwd_candidate.exists():
            return cwd_candidate

        # Check relative to module (project root)
        project_candidate = Path(__file__).resolve().parent.parent.parent / "config.yaml"
        if project_candidate.exists():
            return project_candidate

        return None

    def _parse_raw_dict(self, data: Dict[str, Any], source: str) -> ExtractorConfig:
        """
        Normalizes and parses raw dictionary data into ExtractorConfig,
        handling both nested 'extraction' structures and flat formats.
        """
        config = self.get_default_config()
        config.config_source = source

        if "version" in data and isinstance(data["version"], str):
            config.version = data["version"]

        # If data is nested under 'extraction'
        extraction_data = data.get("extraction", data)

        # 1. Parse Quote ID patterns & anchors
        if "quote_id" in extraction_data and isinstance(extraction_data["quote_id"], dict):
            qid_data = extraction_data["quote_id"]
            if "patterns" in qid_data and isinstance(qid_data["patterns"], list):
                config.quote_id.patterns = [str(p) for p in qid_data["patterns"]]
            elif "pattern" in qid_data and isinstance(qid_data["pattern"], str):
                config.quote_id.patterns = [qid_data["pattern"]]

            if "item_patterns" in qid_data and isinstance(qid_data["item_patterns"], list):
                config.quote_id.item_patterns = [str(p) for p in qid_data["item_patterns"]]

            if "anchors" in qid_data and isinstance(qid_data["anchors"], list):
                config.quote_id.anchors = [str(a) for a in qid_data["anchors"]]

            if "search_offsets" in qid_data and isinstance(qid_data["search_offsets"], list):
                config.quote_id.search_offsets = [
                    list(offset) for offset in qid_data["search_offsets"]
                ]
            elif "offsets" in qid_data and isinstance(qid_data["offsets"], list):
                config.quote_id.search_offsets = [
                    list(offset) for offset in qid_data["offsets"]
                ]

        # 2. Parse Anchors (part_number, cycle_time)
        anchors_dict = extraction_data.get("anchors", {})
        # Support flat top-level part_number/cycle_time keys as well
        if "part_number" in extraction_data and isinstance(extraction_data["part_number"], dict):
            anchors_dict["part_number"] = extraction_data["part_number"]
        if "cycle_time" in extraction_data and isinstance(extraction_data["cycle_time"], dict):
            anchors_dict["cycle_time"] = extraction_data["cycle_time"]

        all_anchor_names = [
            "part_number",
            "cycle_time",
            "description",
            "material",
            "ops_labour",
            "labour_rate",
            "weight_g",
            "annual_volume",
        ]
        for anchor_name in all_anchor_names:
            if anchor_name in anchors_dict and isinstance(anchors_dict[anchor_name], dict):
                a_data = anchors_dict[anchor_name]
                keywords = a_data.get("keywords", a_data.get("anchors"))
                offsets = a_data.get("search_offsets", a_data.get("offsets"))
                clean_patterns = a_data.get("clean_patterns")

                if anchor_name not in config.anchors:
                    config.anchors[anchor_name] = AnchorConfig()

                if keywords and isinstance(keywords, list):
                    config.anchors[anchor_name].keywords = [str(k) for k in keywords]
                if offsets and isinstance(offsets, list):
                    config.anchors[anchor_name].search_offsets = [list(o) for o in offsets]
                if clean_patterns and isinstance(clean_patterns, list):
                    config.anchors[anchor_name].clean_patterns = [str(cp) for cp in clean_patterns]

        # 3. Parse Bounds
        bounds_data = extraction_data.get("bounds", {})
        if "cycle_time_sec" in bounds_data and isinstance(bounds_data["cycle_time_sec"], dict):
            ct_bounds = bounds_data["cycle_time_sec"]
            min_val = float(ct_bounds.get("min", config.bounds["cycle_time_sec"].min))
            max_val = float(ct_bounds.get("max", config.bounds["cycle_time_sec"].max))
            config.bounds["cycle_time_sec"] = BoundsConfig(min=min_val, max=max_val)
        elif "cycle_time" in extraction_data and isinstance(extraction_data["cycle_time"], dict):
            ct_data = extraction_data["cycle_time"]
            if "min_seconds" in ct_data or "max_seconds" in ct_data:
                min_val = float(ct_data.get("min_seconds", DEFAULT_MIN_CYCLE_TIME))
                max_val = float(ct_data.get("max_seconds", DEFAULT_MAX_CYCLE_TIME))
                config.bounds["cycle_time_sec"] = BoundsConfig(min=min_val, max=max_val)

        if "ops_labour" in bounds_data and isinstance(bounds_data["ops_labour"], dict):
            ops_bounds = bounds_data["ops_labour"]
            min_val = float(ops_bounds.get("min", config.bounds["ops_labour"].min))
            max_val = float(ops_bounds.get("max", config.bounds["ops_labour"].max))
            config.bounds["ops_labour"] = BoundsConfig(min=min_val, max=max_val)

        # 4. Parse Sheet Ignore and Target Lists
        sheets_data = extraction_data.get("sheets", {})
        if isinstance(sheets_data, dict):
            if "ignore" in sheets_data and isinstance(sheets_data["ignore"], list):
                config.sheets_ignore = [str(s).lower() for s in sheets_data["ignore"]]
            if "target_keywords" in sheets_data and isinstance(sheets_data["target_keywords"], list):
                config.sheets_target_keywords = [str(s).lower() for s in sheets_data["target_keywords"]]
            if "target_patterns" in sheets_data and isinstance(sheets_data["target_patterns"], list):
                config.sheets_target_patterns = [str(s) for s in sheets_data["target_patterns"]]
        elif "layout" in data and isinstance(data["layout"], dict):
            mt_data = data["layout"].get("multi_tab", {})
            if isinstance(mt_data, dict) and "ignore_sheets" in mt_data:
                config.sheets_ignore = [str(s).lower() for s in mt_data["ignore_sheets"]]

        # 5. Parse Grid Matrix Config
        gm_data = extraction_data.get("grid_matrix", {})
        if isinstance(gm_data, dict) and gm_data:
            if "enabled" in gm_data:
                config.grid_matrix.enabled = bool(gm_data["enabled"])
            if "header_max_row" in gm_data:
                config.grid_matrix.header_max_row = int(gm_data["header_max_row"])
            if "stop_keywords" in gm_data and isinstance(gm_data["stop_keywords"], list):
                config.grid_matrix.stop_keywords = [str(k) for k in gm_data["stop_keywords"]]

        return config
