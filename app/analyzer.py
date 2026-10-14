"""Rule-based Excel analyzer for type detection, relationships, and constraints."""

import re
from collections import Counter
from datetime import datetime, date
from typing import Any, Optional

import openpyxl

from .models import (
    AnalysisResult,
    ColumnAnalysis,
    ColumnConstraint,
    DetectedType,
    RelationshipInfo,
    SheetAnalysis,
)

# Name-based pattern rules: (regex, DetectedType, constraint_factory)
# Checked in order; first match wins.
NAME_PATTERNS: list[tuple[re.Pattern, DetectedType, callable]] = [
    (
        re.compile(r"e[-_]?mail", re.IGNORECASE),
        DetectedType.EMAIL,
        lambda: ColumnConstraint(nullable=False),
    ),
    (
        re.compile(r"(url|website|homepage|webseite|link)", re.IGNORECASE),
        DetectedType.URL,
        lambda: ColumnConstraint(nullable=True),
    ),
    (
        re.compile(
            r"(datum|date|created|updated|modified|timestamp|_at$|_on$)", re.IGNORECASE
        ),
        DetectedType.DATETIME,
        lambda: ColumnConstraint(nullable=True),
    ),
    (
        re.compile(
            r"(preis|price|betrag|amount|cost|kosten|total|summe|gebühr|gebuehr|fee)",
            re.IGNORECASE,
        ),
        DetectedType.FLOAT,
        lambda: ColumnConstraint(nullable=False, ge=0.0),
    ),
    (
        re.compile(
            r"(menge|quantity|anzahl|count|bestand|stock|stueck|stück|lagerbestand)",
            re.IGNORECASE,
        ),
        DetectedType.INTEGER,
        lambda: ColumnConstraint(nullable=False, ge=0),
    ),
    (
        re.compile(
            r"(aktiv|active|enabled|disabled|is_|has_|kann_|verified|published)",
            re.IGNORECASE,
        ),
        DetectedType.BOOLEAN,
        lambda: ColumnConstraint(nullable=False),
    ),
    (
        re.compile(r"(name|bezeichnung|titel|title)", re.IGNORECASE),
        DetectedType.STRING,
        lambda: ColumnConstraint(nullable=False, min_length=1, max_length=255),
    ),
    (
        re.compile(
            r"(beschreibung|description|kommentar|comment|notes|notiz|\btext\b|bemerkung)",
            re.IGNORECASE,
        ),
        DetectedType.TEXT,
        lambda: ColumnConstraint(nullable=True, max_length=2000),
    ),
    (
        re.compile(r"(telefon|phone|\btel\b|\btfn\b|fax|mobil|mobile|handy)", re.IGNORECASE),
        DetectedType.STRING,
        lambda: ColumnConstraint(nullable=True, max_length=30),
    ),
    (
        re.compile(r"(plz|zip|postal)", re.IGNORECASE),
        DetectedType.STRING,
        lambda: ColumnConstraint(nullable=True, max_length=10),
    ),
]


# Error values Excel stores as text in cached formula results.
EXCEL_ERRORS = {
    "#N/A", "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!",
    "#GETTING_DATA", "#SPILL!", "#CALC!",
}


def clean_cell(value: Any) -> Any:
    """Treat Excel error values (e.g. "#N/A") as empty cells."""
    if isinstance(value, str) and value.strip() in EXCEL_ERRORS:
        return None
    return value


def analyze_excel(file_path: str) -> AnalysisResult:
    """Main entry point. Analyzes all sheets and returns complete result."""
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheets: list[SheetAnalysis] = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = [
            tuple(clean_cell(v) for v in row)
            for row in ws.iter_rows(values_only=True)
        ]
        if len(rows) < 2:
            continue

        headers = rows[0]
        data_rows = rows[1:]

        table_name = _sanitize_to_snake(sheet_name)
        model_name = _sanitize_to_pascal(sheet_name)

        columns: list[ColumnAnalysis] = []
        for col_idx, header in enumerate(headers):
            if header is None:
                continue
            col_values = [
                row[col_idx] for row in data_rows if col_idx < len(row)
            ]
            analysis = _analyze_column(str(header), col_values, table_name)
            columns.append(analysis)

        _ensure_primary_key(columns, table_name)

        # Skip sheets with no usable columns (e.g. ReadMe sheets)
        if not columns:
            continue

        sample_data = _build_sample_data(headers, data_rows[:5], columns)

        sheets.append(
            SheetAnalysis(
                original_name=sheet_name,
                table_name=table_name,
                model_name=model_name,
                columns=columns,
                row_count=len(data_rows),
                sample_data=sample_data,
            )
        )

    wb.close()

    relationships = _detect_relationships(sheets)
    warnings = _collect_warnings(sheets)

    return AnalysisResult(
        sheets=sheets,
        relationships=relationships,
        warnings=warnings,
    )


def _analyze_column(
    header: str, values: list[Any], table_name: str
) -> ColumnAnalysis:
    """Analyze a single column by name patterns + data inspection."""
    python_name = _sanitize_to_snake(header)
    non_null_values = [v for v in values if v is not None]
    sample_values = non_null_values[:5]

    # Phase 1: Name-based detection
    detected_type, constraint = _detect_by_name(header)

    # Phase 1b: Validate name-based detection against actual data
    # If name says numeric/bool but data is mostly strings, fall back
    if detected_type is not None and non_null_values:
        detected_type, constraint = _validate_name_detection(
            detected_type, constraint, non_null_values
        )

    # Phase 2: If name didn't match (or was invalidated), detect by data
    if detected_type is None:
        detected_type, constraint = _detect_by_data(non_null_values)

    # Phase 3: Check for enum pattern
    enum_values: list[str] = []
    if detected_type in (DetectedType.STRING, DetectedType.TEXT):
        enum_values = _check_enum_pattern(non_null_values)
        if enum_values:
            detected_type = DetectedType.ENUM

    # Phase 4: Refine nullability from data (data always wins)
    null_count = sum(1 for v in values if v is None)
    if null_count > 0:
        constraint.nullable = True
    elif len(values) > 0:
        constraint.nullable = False

    # Phase 5: Check primary key
    is_pk = _is_primary_key_column(python_name, table_name, non_null_values)

    # Phase 6: Check foreign key pattern
    is_fk, fk_table, fk_col = _check_foreign_key_pattern(python_name)
    if is_pk:
        is_fk = False
        fk_table = None
        fk_col = None

    return ColumnAnalysis(
        original_name=header,
        python_name=python_name,
        detected_type=detected_type,
        constraint=constraint,
        is_primary_key=is_pk,
        is_foreign_key=is_fk,
        foreign_table=fk_table,
        foreign_column=fk_col,
        enum_values=enum_values,
        sample_values=sample_values,
        values_unique=bool(non_null_values) and null_count == 0 and _all_unique(non_null_values),
    )


def _detect_by_name(
    header: str,
) -> tuple[Optional[DetectedType], ColumnConstraint]:
    """Match column name against NAME_PATTERNS."""
    for pattern, dtype, constraint_factory in NAME_PATTERNS:
        if pattern.search(header):
            return dtype, constraint_factory()
    return None, ColumnConstraint()


def _validate_name_detection(
    detected_type: DetectedType,
    constraint: ColumnConstraint,
    values: list[Any],
) -> tuple[Optional[DetectedType], ColumnConstraint]:
    """Validate name-based detection against actual data. Returns (None, default)
    if the data contradicts the name-based type, causing fallback to data detection."""
    # Sample up to 100 values for validation
    sample = values[:100]
    if not sample:
        return detected_type, constraint

    NUMERIC_TYPES = {DetectedType.FLOAT, DetectedType.INTEGER}
    if detected_type in NUMERIC_TYPES:
        numeric_count = sum(1 for v in sample if isinstance(v, (int, float)))
        if numeric_count / len(sample) < 0.5:
            return None, ColumnConstraint()

    if detected_type == DetectedType.BOOLEAN:
        bool_count = sum(
            1 for v in sample
            if isinstance(v, bool)
            or (isinstance(v, str) and _looks_like_bool(v))
        )
        if bool_count / len(sample) < 0.5:
            return None, ColumnConstraint()

    if detected_type == DetectedType.DATETIME:
        dt_count = sum(
            1 for v in sample
            if isinstance(v, (datetime, date))
            or (isinstance(v, str) and _looks_like_date(v))
        )
        if dt_count / len(sample) < 0.5:
            return None, ColumnConstraint()

    if detected_type == DetectedType.EMAIL:
        email_count = sum(
            1 for v in sample
            if isinstance(v, str) and _looks_like_email(v)
        )
        if email_count / len(sample) < 0.5:
            return None, ColumnConstraint()

    return detected_type, constraint


def _detect_by_data(
    values: list[Any],
) -> tuple[DetectedType, ColumnConstraint]:
    """Infer type purely from data values."""
    if not values:
        return DetectedType.STRING, ColumnConstraint()

    type_counts: Counter = Counter()
    for v in values:
        if isinstance(v, bool):
            type_counts["bool"] += 1
        elif isinstance(v, int):
            type_counts["int"] += 1
        elif isinstance(v, float):
            type_counts["float"] += 1
        elif isinstance(v, datetime):
            type_counts["datetime"] += 1
        elif isinstance(v, date):
            type_counts["date"] += 1
        elif isinstance(v, str):
            if _looks_like_email(v):
                type_counts["email"] += 1
            elif _looks_like_url(v):
                type_counts["url"] += 1
            elif _looks_like_date(v):
                type_counts["datetime"] += 1
            elif _looks_like_int(v):
                type_counts["int"] += 1
            elif _looks_like_float(v):
                type_counts["float"] += 1
            elif _looks_like_bool(v):
                type_counts["bool"] += 1
            else:
                type_counts["str"] += 1
        else:
            type_counts["str"] += 1

    if not type_counts:
        return DetectedType.STRING, ColumnConstraint()

    dominant_type, dominant_count = type_counts.most_common(1)[0]
    threshold = 0.7 * len(values)

    TYPE_MAP = {
        "bool": (DetectedType.BOOLEAN, ColumnConstraint()),
        "int": (DetectedType.INTEGER, ColumnConstraint()),
        "float": (DetectedType.FLOAT, ColumnConstraint()),
        "datetime": (DetectedType.DATETIME, ColumnConstraint()),
        "date": (DetectedType.DATE, ColumnConstraint()),
        "email": (DetectedType.EMAIL, ColumnConstraint()),
        "url": (DetectedType.URL, ColumnConstraint()),
    }

    if dominant_count >= threshold and dominant_type in TYPE_MAP:
        return TYPE_MAP[dominant_type]

    # int + float mix -> float
    if type_counts["int"] + type_counts["float"] >= threshold:
        return DetectedType.FLOAT, ColumnConstraint()

    return DetectedType.STRING, ColumnConstraint()


# ---- Helper functions for data type detection ----


def _looks_like_email(v: str) -> bool:
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", v))


def _looks_like_url(v: str) -> bool:
    return bool(re.match(r"^https?://", v, re.IGNORECASE))


def _looks_like_date(v: str) -> bool:
    patterns = [
        r"^\d{4}-\d{2}-\d{2}",
        r"^\d{2}\.\d{2}\.\d{4}",
        r"^\d{2}/\d{2}/\d{4}",
    ]
    return any(re.match(p, v) for p in patterns)


def _looks_like_int(v: str) -> bool:
    try:
        int(v)
        return "." not in v and "," not in v
    except (ValueError, TypeError):
        return False


def _looks_like_float(v: str) -> bool:
    try:
        float(v.replace(",", "."))
        return "." in v or "," in v
    except (ValueError, TypeError, AttributeError):
        return False


def _looks_like_bool(v: str) -> bool:
    return v.lower() in {"true", "false", "ja", "nein", "yes", "no", "1", "0"}


def _check_enum_pattern(values: list[Any]) -> list[str]:
    """Detect enum: <= 10 distinct values with repetition."""
    if len(values) < 3:
        return []
    str_values = [str(v) for v in values]
    distinct = set(str_values)
    if 2 <= len(distinct) <= 10 and len(str_values) / len(distinct) >= 2:
        return sorted(distinct)
    return []


# ---- Primary key / Foreign key detection ----


def _is_primary_key_column(
    python_name: str, table_name: str, values: list[Any]
) -> bool:
    """Detect if this column is a primary key."""
    # _? makes the underscore optional to handle both "category_id" and "categoryid"
    pk_patterns = [r"^id$", f"^{re.escape(table_name)}_?id$", r"^pk$"]
    # Also match singular form: "produkte" table -> "produkt_id" / "produktid"
    singular = _guess_singular(table_name)
    if singular != table_name:
        pk_patterns.append(f"^{re.escape(singular)}_?id$")

    # Also match if the column stem (without _id / id suffix) is related to table name.
    # Handles cases like table "usstates" with column "state_id" or "stateid".
    col_match = re.match(r"^(.+?)_?id$", python_name, re.IGNORECASE)
    if col_match:
        stem = col_match.group(1).lower()
        if (
            table_name.startswith(stem)
            or table_name.endswith(stem)
            or stem in table_name
        ) and len(stem) >= 3:
            pk_patterns.append(f"^{re.escape(python_name)}$")

    name_match = any(
        re.match(p, python_name, re.IGNORECASE) for p in pk_patterns
    )
    if not name_match:
        return False

    # A name match is not enough: duplicate values (e.g. "employeeid" in a
    # junction table) would violate the PK constraint when seeding.
    return not values or _all_unique(values)


def _all_unique(values: list[Any]) -> bool:
    return len(set(values)) == len(values)


def _check_foreign_key_pattern(
    python_name: str,
) -> tuple[bool, Optional[str], Optional[str]]:
    """Detect FK pattern like 'kunde_id' or 'categoryid' -> foreign table."""
    # Don't match bare "id"
    if python_name.lower() == "id":
        return False, None, None
    # Match with underscore: "category_id" -> base "category"
    fk_match = re.match(r"^(.+?)_id$", python_name, re.IGNORECASE)
    if fk_match:
        base = fk_match.group(1).lower()
        return True, base, "id"
    # Match without underscore: "categoryid" -> base "category"
    # Require base >= 3 chars to avoid false positives ("grid", "acid", etc.)
    fk_match = re.match(r"^(.{3,}?)id$", python_name, re.IGNORECASE)
    if fk_match:
        base = fk_match.group(1).lower()
        return True, base, "id"
    return False, None, None


def _detect_relationships(sheets: list[SheetAnalysis]) -> list[RelationshipInfo]:
    """Resolve FK columns against actual sheet/table names."""
    actual_tables = {s.table_name: s for s in sheets}
    relationships: list[RelationshipInfo] = []

    for sheet in sheets:
        for col in sheet.columns:
            if not col.is_foreign_key or col.is_primary_key:
                continue

            target_base = col.foreign_table
            if target_base is None:
                continue

            resolved = _resolve_table_name(target_base, actual_tables)
            if resolved and resolved != sheet.table_name:
                # Find the actual PK column name in the target table
                target_pk_col = _get_pk_column_name(actual_tables[resolved])
                col.foreign_table = resolved
                col.foreign_column = target_pk_col
                relationships.append(
                    RelationshipInfo(
                        source_table=sheet.table_name,
                        source_column=col.python_name,
                        target_table=resolved,
                        target_column=target_pk_col,
                    )
                )
            else:
                col.is_foreign_key = False
                col.foreign_table = None
                col.foreign_column = None

    return relationships


def _get_pk_column_name(sheet: SheetAnalysis) -> str:
    """Return the primary key column name for a sheet, defaulting to 'id'."""
    for col in sheet.columns:
        if col.is_primary_key:
            return col.python_name
    return "id"


def _resolve_table_name(
    base: str, actual_tables: dict[str, SheetAnalysis]
) -> Optional[str]:
    """Try to match 'kunde' to 'kunden', 'customer' to 'customers', etc."""
    candidates = [
        base,
        base + "s",
        base + "n",
        base + "en",
        base + "e",
        base + "er",
    ]
    # English: -y → -ies (category → categories, territory → territories)
    if base.endswith("y"):
        candidates.append(base[:-1] + "ies")
    # English: -s, -x, -z, -ch, -sh → +es (box → boxes)
    if base.endswith(("s", "x", "z", "ch", "sh")):
        candidates.append(base + "es")
    # English: -f → -ves (leaf → leaves)
    if base.endswith("f"):
        candidates.append(base[:-1] + "ves")

    for candidate in candidates:
        if candidate in actual_tables:
            return candidate

    # Fuzzy: check if base is a substring of any table name or vice versa
    for table_name in actual_tables:
        if table_name.startswith(base) or base.startswith(table_name):
            return table_name

    return None


def _guess_singular(name: str) -> str:
    """Guess the singular form of a German/English plural."""
    # English: -ies → -y  (categories → category, territories → territory)
    if name.endswith("ies") and len(name) > 4:
        return name[:-3] + "y"
    # English: -ves → -f  (lives → life — rare but valid)
    if name.endswith("ves") and len(name) > 4:
        return name[:-3] + "f"
    # English: -ses, -xes, -zes, -ches, -shes → remove -es
    if name.endswith(("ses", "xes", "zes")) and len(name) > 4:
        return name[:-2]
    if name.endswith(("ches", "shes")) and len(name) > 5:
        return name[:-2]
    # English/German: -s (but not -ss, -us, -is)
    if name.endswith("s") and not name.endswith(("ss", "us", "is")) and len(name) > 3:
        return name[:-1]
    # German: -en, -er, -n, -e
    for suffix in ["en", "er", "n", "e"]:
        if name.endswith(suffix) and len(name) > len(suffix) + 2:
            return name[: -len(suffix)]
    return name


def _ensure_primary_key(
    columns: list[ColumnAnalysis], table_name: str
) -> None:
    """Ensure at least one column is marked as primary key."""
    pk_cols = [c for c in columns if c.is_primary_key]
    if pk_cols:
        return

    # Look for 'id' column
    for col in columns:
        if col.python_name == "id" and col.values_unique:
            col.is_primary_key = True
            col.detected_type = DetectedType.INTEGER
            col.constraint.nullable = False
            col.constraint.unique = True
            return

    # Look for first column ending with _id that has unique values
    for col in columns:
        if col.python_name.endswith("_id") and col.values_unique:
            col.is_primary_key = True
            col.is_foreign_key = False
            col.foreign_table = None
            col.foreign_column = None
            col.constraint.nullable = False
            col.constraint.unique = True
            return

    # Look for first column ending with 'id' (no underscore, e.g. "categoryid")
    for col in columns:
        if (
            col.python_name.endswith("id")
            and col.python_name != "id"
            and len(col.python_name) > 4
            and col.values_unique
        ):
            col.is_primary_key = True
            col.is_foreign_key = False
            col.foreign_table = None
            col.foreign_column = None
            col.constraint.nullable = False
            col.constraint.unique = True
            return

    # No suitable PK found - generator synthesizes an autoincrement "id"


def _sanitize_to_snake(name: str) -> str:
    """Convert any string to valid Python snake_case identifier."""
    replacements = {
        "ä": "ae", "ö": "oe", "ü": "ue", "ß": "ss",
        "Ä": "Ae", "Ö": "Oe", "Ü": "Ue",
    }
    for k, v in replacements.items():
        name = name.replace(k, v)
    name = re.sub(r"[^a-zA-Z0-9]", "_", name)
    name = re.sub(r"([a-z])([A-Z])", r"\1_\2", name)
    name = name.lower().strip("_")
    name = re.sub(r"_+", "_", name)
    if name and name[0].isdigit():
        name = "col_" + name
    return name or "unnamed"


def _sanitize_to_pascal(name: str) -> str:
    """Convert to PascalCase for class names."""
    snake = _sanitize_to_snake(name)
    return "".join(word.capitalize() for word in snake.split("_"))


def _build_sample_data(
    headers: tuple, rows: list[tuple], columns: list[ColumnAnalysis]
) -> list[dict]:
    """Build list of dicts from first N rows for preview."""
    result = []
    for row in rows:
        d = {}
        for i, col in enumerate(columns):
            val = row[i] if i < len(row) else None
            if isinstance(val, (datetime, date)):
                val = val.isoformat()
            d[col.python_name] = val
        result.append(d)
    return result


def _collect_warnings(sheets: list[SheetAnalysis]) -> list[str]:
    """Collect warnings produced during analysis."""
    warnings = []
    for sheet in sheets:
        pk_cols = [c for c in sheet.columns if c.is_primary_key]
        if not pk_cols:
            warnings.append(
                f"Sheet '{sheet.original_name}': No primary key detected. "
                f"A synthetic 'id' column will be added."
            )
    return warnings
