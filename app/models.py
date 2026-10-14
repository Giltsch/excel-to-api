"""Internal data structures for Excel analysis results."""

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Optional


class DetectedType(Enum):
    STRING = "str"
    INTEGER = "int"
    FLOAT = "float"
    BOOLEAN = "bool"
    DATETIME = "datetime"
    DATE = "date"
    EMAIL = "EmailStr"
    URL = "HttpUrl"
    ENUM = "enum"
    TEXT = "str"


@dataclass
class ColumnConstraint:
    nullable: bool = True
    unique: bool = False
    gt: Optional[float] = None
    ge: Optional[float] = None
    lt: Optional[float] = None
    le: Optional[float] = None
    min_length: Optional[int] = None
    max_length: Optional[int] = None
    regex: Optional[str] = None
    default: Optional[str] = None


@dataclass
class ColumnAnalysis:
    original_name: str
    python_name: str
    detected_type: DetectedType
    constraint: ColumnConstraint
    is_primary_key: bool = False
    is_foreign_key: bool = False
    foreign_table: Optional[str] = None
    foreign_column: Optional[str] = None
    enum_values: list[str] = field(default_factory=list)
    sample_values: list[Any] = field(default_factory=list)
    values_unique: bool = False
    confidence: float = 1.0


@dataclass
class SheetAnalysis:
    original_name: str
    table_name: str
    model_name: str
    columns: list[ColumnAnalysis]
    row_count: int
    sample_data: list[dict] = field(default_factory=list)


@dataclass
class RelationshipInfo:
    source_table: str
    source_column: str
    target_table: str
    target_column: str


@dataclass
class AnalysisResult:
    sheets: list[SheetAnalysis]
    relationships: list[RelationshipInfo]
    warnings: list[str] = field(default_factory=list)
