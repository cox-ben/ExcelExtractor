"""
Data contracts and models for the Excel Quote Extractor.

This module defines Pydantic V2 data models for extracted quote data:
- SourceCellCoords: Exact cell coordinates of extracted fields.
- QuoteItem: Individual quote line item with validation and confidence scoring.
- QuoteDocument: Top-level document container with serialization and DataFrame export helpers.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Literal, Optional
from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


class SourceCellCoords(BaseModel):
    """
    Tracks the Excel cell coordinates (e.g., 'B2', 'Sheet1!C10') for extracted fields.
    """
    quote_num: Optional[str] = Field(default=None, description="Cell coordinate for quote number")
    part_num: Optional[str] = Field(default=None, description="Cell coordinate for part number")
    drawing_num: Optional[str] = Field(default=None, description="Cell coordinate for drawing number")
    description: Optional[str] = Field(default=None, description="Cell coordinate for description")
    material: Optional[str] = Field(default=None, description="Cell coordinate for material")
    cycle_time: Optional[str] = Field(default=None, description="Cell coordinate for cycle time")
    ops_labour: Optional[str] = Field(default=None, description="Cell coordinate for ops/labour")
    labour_rate: Optional[str] = Field(default=None, description="Cell coordinate for labour rate")
    weight_g: Optional[str] = Field(default=None, description="Cell coordinate for weight (g)")
    annual_volume: Optional[str] = Field(default=None, description="Cell coordinate for annual volume")

    model_config = ConfigDict(extra="ignore")

    def to_dict(self) -> Dict[str, Optional[str]]:
        """Return coordinates as a dictionary."""
        return self.model_dump()


class QuoteItem(BaseModel):
    """
    Represents an individual quoted part item within a quoting workbook.
    """
    quote_item_number: str = Field(
        ...,
        description="Canonical quote item identifier (e.g., 'Q1001-1', 'Q4476-R1-1')",
    )
    part_number: Optional[str] = Field(
        default=None,
        description="Customer or manufacturing part number",
    )
    drawing_number: Optional[str] = Field(
        default=None,
        description="Drawing number or bracketed drawing tag",
    )
    description: Optional[str] = Field(
        default=None,
        description="Part description",
    )
    material: Optional[str] = Field(
        default=None,
        description="Raw material / resin description",
    )
    cycle_time_sec: Optional[float] = Field(
        default=None,
        description="Cycle time strictly in seconds (5.0s to 300.0s expected)",
    )
    ops_labour: Optional[float] = Field(
        default=None,
        description="Number of operators / labour (0.0 to 1.0 expected)",
    )
    labour_rate: Optional[float] = Field(
        default=None,
        description="Static labour rate ($CAD/Hr) for all parts on sheet",
    )
    weight_g: Optional[float] = Field(
        default=None,
        description="Part weight in grams (> 0g expected)",
    )
    annual_volume: Optional[int] = Field(
        default=None,
        description="Annual estimated volume (EAU > 0 expected)",
    )
    source_cell_coords: SourceCellCoords = Field(
        default_factory=SourceCellCoords,
        description="Coordinates of the source cells in the Excel workbook",
    )
    confidence: Literal["high", "medium", "low"] = Field(
        default="high",
        description="Extraction confidence rating ('high', 'medium', or 'low')",
    )
    warnings: List[str] = Field(
        default_factory=list,
        description="Validation or extraction warning messages",
    )

    model_config = ConfigDict(validate_assignment=True, extra="ignore")

    @field_validator("part_number", mode="before")
    @classmethod
    def clean_part_number(cls, v: Any) -> Optional[str]:
        """Normalize part number string, strip trailing drawing labels, and reject empty/zero/error values."""
        if v is None:
            return None
        if isinstance(v, (int, float)):
            if v == 0:
                return None
            if isinstance(v, float) and v.is_integer():
                return str(int(v))
            return str(v)

        s = str(v).strip(" :\t\r\n")
        if not s or s in ("0", "0.0", "-", "--", "---", "#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!", "None", "null", "N/A", "NA"):
            return None

        # Strip trailing drawing/bracketed tags (e.g. "3042605C (DWG 4022-A)" -> "3042605C", "858007801000 (P0900612)" -> "858007801000")
        if re.search(r"(?i)\b(?:dwg|drawing)\b", s):
            cleaned = re.sub(
                r"(?i)(?<=[a-zA-Z0-9])\s*[\(\[]?\s*(?:dwg|drawing)\b.*$", "", s
            ).strip(" :\t\r\n()[]")
        elif re.search(r"\s*[\(\[].*?[\)\]]\s*$", s):
            trimmed = re.sub(r"\s*[\(\[].*?[\)\]]\s*$", "", s).strip(" :\t\r\n")
            if trimmed:
                cleaned = trimmed
            else:
                cleaned = s.strip(" :\t\r\n()[]")
        else:
            cleaned = s.strip(" :\t\r\n")
        if not cleaned or cleaned in ("0", "0.0", "-", "--", "#REF!", "#VALUE!", "#N/A"):
            return None
        return cleaned

    @field_validator("description", mode="before")
    @classmethod
    def clean_description(cls, v: Any) -> Optional[str]:
        """Normalize description text, rejecting formulas and placeholders."""
        if v is None:
            return None
        s = str(v).strip(" :\t\r\n")
        if not s or s in ("0", "0.0", "-", "--", "---", "#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!", "None", "null", "N/A", "NA"):
            return None
        return s

    @field_validator("material", mode="before")
    @classmethod
    def clean_material(cls, v: Any) -> Optional[str]:
        """Normalize material text, rejecting formulas and placeholders."""
        if v is None:
            return None
        s = str(v).strip(" :\t\r\n")
        if not s or s in ("0", "0.0", "-", "--", "---", "#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!", "None", "null", "N/A", "NA"):
            return None
        return s

    @field_validator("cycle_time_sec", mode="before")
    @classmethod
    def clean_cycle_time(cls, v: Any) -> Optional[float]:
        """Parse cycle time float, stripping common unit text."""
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return round(float(v), 2)
        if isinstance(v, str):
            cleaned = v.strip().lower()
            if not cleaned or cleaned in ("-", "--", "n/a", "na", "none") or cleaned.startswith("#"):
                return None
            # Strip trailing unit words like 's', 'sec', 'seconds' (e.g. '24.5s', '15 sec')
            cleaned = re.sub(r"(?i)\s*(?:seconds?|secs?|s)\s*$", "", cleaned).strip()
            try:
                return round(float(cleaned), 2)
            except ValueError:
                return None
        return None

    @field_validator("ops_labour", mode="before")
    @classmethod
    def clean_ops_labour(cls, v: Any) -> Optional[float]:
        """Parse operator count / labour float."""
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return round(float(v), 2)
        if isinstance(v, str):
            cleaned = v.strip().lower()
            if not cleaned or cleaned in ("-", "--", "n/a", "na", "none") or cleaned.startswith("#"):
                return None
            cleaned = re.sub(r"(?i)\s*(?:ops|labour|labor|operators?)\s*$", "", cleaned).strip()
            try:
                return round(float(cleaned), 2)
            except ValueError:
                match = re.search(r"([0-9]+(?:\.[0-9]+)?)", cleaned)
                if match:
                    try:
                        return round(float(match.group(1)), 2)
                    except ValueError:
                        return None
                return None
        return None

    @field_validator("labour_rate", mode="before")
    @classmethod
    def clean_labour_rate(cls, v: Any) -> Optional[float]:
        """Parse labour rate currency float."""
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return round(float(v), 2)
        if isinstance(v, str):
            cleaned = v.strip().replace("$", "").replace(",", "")
            if not cleaned or cleaned.lower() in ("-", "--", "n/a", "na", "none") or cleaned.startswith("#"):
                return None
            match = re.search(r"([0-9]+(?:\.[0-9]+)?)", cleaned)
            if match:
                try:
                    return round(float(match.group(1)), 2)
                except ValueError:
                    return None
            return None
        return None

    @field_validator("weight_g", mode="before")
    @classmethod
    def clean_weight_g(cls, v: Any) -> Optional[float]:
        """Parse part weight float in grams."""
        if v is None:
            return None
        if isinstance(v, (int, float)):
            if v == 0:
                return None
            return round(float(v), 2)
        if isinstance(v, str):
            cleaned = v.strip().lower()
            if not cleaned or cleaned in ("0", "0.0", "-", "--", "n/a", "na", "none") or cleaned.startswith("#"):
                return None
            cleaned = re.sub(r"(?i)\s*(?:grams?|g)\s*$", "", cleaned).strip()
            try:
                val = round(float(cleaned), 2)
                return val if val > 0 else None
            except ValueError:
                match = re.search(r"([0-9]+(?:\.[0-9]+)?)", cleaned)
                if match:
                    try:
                        val = round(float(match.group(1)), 2)
                        return val if val > 0 else None
                    except ValueError:
                        return None
                return None
        return None

    @field_validator("annual_volume", mode="before")
    @classmethod
    def clean_annual_volume(cls, v: Any) -> Optional[int]:
        """Parse annual volume integer (EAU)."""
        if v is None:
            return None
        if isinstance(v, (int, float)):
            if v <= 0:
                return None
            return int(round(float(v)))
        if isinstance(v, str):
            cleaned = v.strip().replace(",", "").replace(" ", "")
            if not cleaned or cleaned in ("0", "0.0", "-", "--", "n/a", "na", "none") or cleaned.startswith("#"):
                return None
            cleaned = re.sub(r"(?i)\s*(?:eau|units?|pcs?)\s*$", "", cleaned).strip()
            try:
                val = int(round(float(cleaned)))
                return val if val > 0 else None
            except ValueError:
                match = re.search(r"([0-9]+)", cleaned)
                if match:
                    try:
                        val = int(match.group(1))
                        return val if val > 0 else None
                    except ValueError:
                        return None
                return None
        return None

    @model_validator(mode="after")
    def validate_business_rules(self) -> "QuoteItem":
        """
        Validates business rules:
        - Strict cycle time bounds (5.0s <= cycle_time <= 300.0s).
          If violated, appends a warning WITHOUT discarding the numeric value and downgrades confidence to 'low'.
        - Checks for missing required fields (part_number, cycle_time_sec).
        - Validates #Ops bounds (0.0 <= ops_labour <= 1.0).
        - Validates weight_g (> 0.0) and annual_volume (> 0).
        - Validates quote item number format and updates confidence score deterministically.
        """
        target_confidence: Literal["high", "medium", "low"] = self.confidence

        # 1. Validate Cycle Time Bounds
        if self.cycle_time_sec is not None:
            if self.cycle_time_sec < 5.0 or self.cycle_time_sec > 300.0:
                warning_msg = (
                    f"Cycle time {self.cycle_time_sec}s outside expected range (5.0s–300.0s)"
                )
                if warning_msg not in self.warnings:
                    self.warnings.append(warning_msg)
                target_confidence = "low"
        else:
            missing_ct_msg = "Missing cycle_time_sec"
            if missing_ct_msg not in self.warnings:
                self.warnings.append(missing_ct_msg)
            target_confidence = "low"

        # 2. Validate Part Number presence
        if not self.part_number or not self.part_number.strip():
            missing_pn_msg = "Missing part_number"
            if missing_pn_msg not in self.warnings:
                self.warnings.append(missing_pn_msg)
            target_confidence = "low"

        # 3. Validate #Ops / Labour bounds (strictly 0.0 to 1.0)
        if self.ops_labour is not None:
            if self.ops_labour < 0.0 or self.ops_labour > 1.0:
                ops_msg = f"#Ops / Labour {self.ops_labour} outside expected range (0.0–1.0)"
                if ops_msg not in self.warnings:
                    self.warnings.append(ops_msg)
                target_confidence = "low"

        # 4. Validate Weight bounds (> 0g)
        if self.weight_g is not None and self.weight_g <= 0.0:
            wt_msg = f"Weight {self.weight_g}g outside expected range (>0g)"
            if wt_msg not in self.warnings:
                self.warnings.append(wt_msg)
            target_confidence = "low"

        # 5. Validate Annual Volume bounds (> 0)
        if self.annual_volume is not None and self.annual_volume <= 0:
            vol_msg = f"Annual volume {self.annual_volume} outside expected range (>0)"
            if vol_msg not in self.warnings:
                self.warnings.append(vol_msg)
            target_confidence = "low"

        # 6. Validate Quote Item Number canonical structure
        if not re.match(r"^Q[0-9]{4,5}(?:-[A-Za-z0-9]+)+$", self.quote_item_number):
            non_standard_msg = (
                f"Non-standard quote item number format: {self.quote_item_number}"
            )
            # If not already flagged by an inference warning
            if not any("inferred" in w.lower() for w in self.warnings):
                if non_standard_msg not in self.warnings:
                    self.warnings.append(non_standard_msg)
            if target_confidence == "high":
                target_confidence = "medium"

        # 7. Overall Confidence Determination
        # If any missing field or bounds warning was recorded in warnings:
        if any(w.startswith("Missing ") for w in self.warnings):
            target_confidence = "low"

        if target_confidence != "low":
            if self.warnings:
                target_confidence = "medium"

        # Apply target confidence bypassing Pydantic assignment validator
        # to prevent RecursionError when validate_assignment=True
        if self.confidence != target_confidence:
            object.__setattr__(self, "confidence", target_confidence)
            self.__pydantic_fields_set__.add("confidence")

        return self

    def to_flat_dict(self, parent_doc: Optional["QuoteDocument"] = None) -> Dict[str, Any]:
        """
        Serializes item into a flat tabular dictionary suitable for CSV/Excel export.
        Order matches MPC manufacturing specification:
        Quote Item, Part Number, Description, Material, Cycle Sec, #Ops, Labour Rate, Weight (g), Ann. Volume.
        """
        warnings_str = "; ".join(self.warnings) if self.warnings else ""
        return {
            "source_file": parent_doc.source_file if parent_doc else "",
            "quote_id": (parent_doc.quote_id or "") if parent_doc else "",
            "quote_item_number": self.quote_item_number,
            "part_number": self.part_number or "",
            "description": self.description or "",
            "material": self.material or "",
            "cycle_time_sec": self.cycle_time_sec,
            "ops_labour": self.ops_labour,
            "labour_rate": self.labour_rate,
            "weight_g": self.weight_g,
            "annual_volume": self.annual_volume,
            "confidence": self.confidence,
            "quote_num_coord": self.source_cell_coords.quote_num or "",
            "part_num_coord": self.source_cell_coords.part_num or "",
            "description_coord": self.source_cell_coords.description or "",
            "material_coord": self.source_cell_coords.material or "",
            "cycle_time_coord": self.source_cell_coords.cycle_time or "",
            "ops_labour_coord": self.source_cell_coords.ops_labour or "",
            "labour_rate_coord": self.source_cell_coords.labour_rate or "",
            "weight_g_coord": self.source_cell_coords.weight_g or "",
            "annual_volume_coord": self.source_cell_coords.annual_volume or "",
            "warnings": warnings_str,
        }


class QuoteDocument(BaseModel):
    """
    Top-level data model representing an analyzed quoting workbook.
    """
    source_file: str = Field(..., description="Base filename of the analyzed quote workbook")
    quote_id: Optional[str] = Field(default=None, description="Global quote identifier (e.g. 'Q1001')")
    total_items: int = Field(default=0, ge=0, description="Total count of extracted quote items")
    items: List[QuoteItem] = Field(default_factory=list, description="List of extracted quote items")
    warnings: List[str] = Field(default_factory=list, description="Document-level parsing warnings")

    model_config = ConfigDict(validate_assignment=True, extra="ignore")

    @model_validator(mode="after")
    def sync_total_items_and_quote_id(self) -> "QuoteDocument":
        """Synchronizes total_items count and infers quote_id from items if not explicitly set."""
        if len(self.items) > 0:
            if self.total_items == 0 or self.total_items != len(self.items):
                object.__setattr__(self, "total_items", len(self.items))
                self.__pydantic_fields_set__.add("total_items")

            if self.quote_id is None:
                match = re.match(r"^(Q[0-9]{4,5})", self.items[0].quote_item_number)
                if match:
                    object.__setattr__(self, "quote_id", match.group(1))
                    self.__pydantic_fields_set__.add("quote_id")

        return self

    def add_item(self, item: QuoteItem) -> None:
        """Appends a QuoteItem to items and updates total_items and quote_id."""
        self.items.append(item)
        object.__setattr__(self, "total_items", len(self.items))
        self.__pydantic_fields_set__.add("total_items")
        if self.quote_id is None:
            match = re.match(r"^(Q[0-9]{4,5})", item.quote_item_number)
            if match:
                object.__setattr__(self, "quote_id", match.group(1))
                self.__pydantic_fields_set__.add("quote_id")

    def to_dict(self) -> Dict[str, Any]:
        """Serializes the document to a standard dictionary."""
        return self.model_dump()

    def to_json(self, indent: int = 2) -> str:
        """Serializes the document to a JSON formatted string."""
        return self.model_dump_json(indent=indent)

    def to_flat_rows(self) -> List[Dict[str, Any]]:
        """
        Serializes all quote items into a list of flat dictionaries for CSV / Excel / DataFrame export.
        """
        return [item.to_flat_dict(parent_doc=self) for item in self.items]

    def to_dataframe(self) -> Any:
        """
        Converts the extracted items into a Pandas DataFrame.
        """
        try:
            import pandas as pd
            rows = self.to_flat_rows()
            columns = [
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
                "warnings",
                "source_file",
                "quote_id",
                "quote_num_coord",
                "part_num_coord",
                "description_coord",
                "material_coord",
                "cycle_time_coord",
                "ops_labour_coord",
                "labour_rate_coord",
                "weight_g_coord",
                "annual_volume_coord",
            ]
            if not rows:
                return pd.DataFrame(columns=columns)
            return pd.DataFrame(rows, columns=columns)
        except ImportError:
            raise ImportError("pandas is required for QuoteDocument.to_dataframe()")

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "QuoteDocument":
        """Instantiates a QuoteDocument from a dictionary."""
        return cls.model_validate(data)

    @classmethod
    def from_json(cls, json_str: str) -> "QuoteDocument":
        """Instantiates a QuoteDocument from a JSON string."""
        return cls.model_validate_json(json_str)
