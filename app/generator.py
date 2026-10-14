"""Code generator: transforms AnalysisResult into a complete FastAPI project."""

import json
import os
import re
import tempfile
import zipfile
from pathlib import Path

from jinja2 import Environment, FileSystemLoader

from .models import AnalysisResult, ColumnAnalysis, DetectedType, SheetAnalysis

TEMPLATE_DIR = Path(__file__).parent.parent / "templates"


class ProjectGenerator:
    """Generates a complete FastAPI project from an AnalysisResult."""

    def __init__(self, analysis: AnalysisResult, excel_data: dict[str, list[dict]]):
        self.analysis = analysis
        self.excel_data = excel_data
        self.env = Environment(
            loader=FileSystemLoader(str(TEMPLATE_DIR)),
            trim_blocks=True,
            lstrip_blocks=True,
            keep_trailing_newline=True,
        )
        self.env.filters["capitalize_first"] = self._capitalize_first
        self.env.filters["sanitize_enum"] = self._sanitize_enum_member

    def generate_to_directory(self, output_dir: str) -> str:
        """Generate all project files into output_dir."""
        os.makedirs(output_dir, exist_ok=True)

        context = self._build_context()

        templates_map = {
            "main.py": "fastapi_main.py.j2",
            "models.py": "models.py.j2",
            "schemas.py": "schemas.py.j2",
            "database.py": "database.py.j2",
            "crud.py": "crud.py.j2",
            "seed_data.py": "seed_data.py.j2",
            "requirements.txt": "requirements.txt.j2",
            "Dockerfile": "dockerfile.j2",
            "README.md": "readme.md.j2",
        }

        for filename, template_name in templates_map.items():
            template = self.env.get_template(template_name)
            content = template.render(**context)
            filepath = os.path.join(output_dir, filename)
            with open(filepath, "w") as f:
                f.write(content)

        # Write seed data as JSON
        seed_path = os.path.join(output_dir, "seed_data.json")
        with open(seed_path, "w") as f:
            json.dump(self.excel_data, f, indent=2, default=str)

        return output_dir

    def generate_zip(self) -> str:
        """Generate complete project and return path to ZIP file."""
        tmpdir = tempfile.mkdtemp(prefix="excel_api_")
        project_dir = os.path.join(tmpdir, "generated_api")
        self.generate_to_directory(project_dir)

        zip_path = os.path.join(tmpdir, "generated_api.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _dirs, files in os.walk(project_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, project_dir)
                    zf.write(file_path, f"generated_api/{arcname}")
        return zip_path

    def _build_context(self) -> dict:
        """Build the full template rendering context."""
        enriched_sheets = []
        for sheet in self.analysis.sheets:
            enriched = _EnrichedSheet(sheet)
            enriched_sheets.append(enriched)

        enriched_rels = []
        for rel in self.analysis.relationships:
            enriched_rels.append(
                _EnrichedRelationship(rel, self.analysis.sheets)
            )

        return {
            "sheets": enriched_sheets,
            "relationships": enriched_rels,
            "has_email_type": self._uses_type(DetectedType.EMAIL),
            "has_url_type": self._uses_type(DetectedType.URL),
            "has_datetime_type": self._uses_type(DetectedType.DATETIME),
            "has_date_type": self._uses_type(DetectedType.DATE),
            "has_enum_type": self._uses_type(DetectedType.ENUM),
            "api_title": self._derive_api_title(),
        }

    def _uses_type(self, dtype: DetectedType) -> bool:
        return any(
            col.detected_type == dtype
            for sheet in self.analysis.sheets
            for col in sheet.columns
        )

    def _derive_api_title(self) -> str:
        names = [s.model_name for s in self.analysis.sheets]
        if len(names) == 1:
            return f"{names[0]} API"
        return " & ".join(names) + " API"

    @staticmethod
    def _capitalize_first(value: str) -> str:
        if not value:
            return value
        return value[0].upper() + value[1:]

    @staticmethod
    def _sanitize_enum_member(value: str) -> str:
        """Convert a string to a valid Python enum member name."""
        sanitized = re.sub(r"[^a-zA-Z0-9]", "_", str(value))
        sanitized = sanitized.strip("_")
        if not sanitized or sanitized[0].isdigit():
            sanitized = "val_" + sanitized
        return sanitized.upper()


class _EnrichedColumn:
    """Wrapper around ColumnAnalysis that adds template-friendly properties."""

    def __init__(self, col: ColumnAnalysis, sheet_model_name: str):
        self._col = col
        self._sheet_model_name = sheet_model_name
        self._enum_entries: list[tuple[str, str]] | None = None

    def __getattr__(self, name):
        return getattr(self._col, name)

    @property
    def enum_entries(self) -> list[tuple[str, str]]:
        """Returns list of (unique_member_name, original_value) for enum columns."""
        if self._enum_entries is not None:
            return self._enum_entries
        entries = []
        seen: set[str] = set()
        for val in self._col.enum_values:
            if val is None:
                continue
            str_val = str(val).strip()
            if not str_val:
                continue
            # Sanitize member name
            sanitized = re.sub(r"[^a-zA-Z0-9]", "_", str_val)
            sanitized = sanitized.strip("_")
            if not sanitized or sanitized[0].isdigit():
                sanitized = "VAL_" + sanitized
            sanitized = sanitized.upper()
            # Deduplicate
            unique = sanitized
            counter = 2
            while unique in seen:
                unique = f"{sanitized}_{counter}"
                counter += 1
            seen.add(unique)
            # Sanitize value: collapse newlines/control chars into spaces
            clean_val = re.sub(r"[\r\n]+", " ", str_val).strip()
            # Escape backslashes and double quotes for safe Python string literal
            clean_val = clean_val.replace("\\", "\\\\").replace('"', '\\"')
            entries.append((unique, clean_val))
        self._enum_entries = entries
        return entries

    @property
    def sa_type(self) -> str:
        """SQLAlchemy column type string."""
        mapping = {
            DetectedType.STRING: "String(255)",
            DetectedType.TEXT: "Text",
            DetectedType.INTEGER: "Integer",
            DetectedType.FLOAT: "Float",
            DetectedType.BOOLEAN: "Boolean",
            DetectedType.DATETIME: "DateTime",
            DetectedType.DATE: "Date",
            DetectedType.EMAIL: "String(255)",
            DetectedType.URL: "String(2048)",
            DetectedType.ENUM: "String(100)",
        }
        return mapping.get(self._col.detected_type, "String(255)")

    @property
    def pydantic_base_type(self) -> str:
        """Base Pydantic type annotation (no Optional, no Field)."""
        if self._col.detected_type == DetectedType.ENUM and self.enum_entries:
            enum_name = (
                f"{self._sheet_model_name}"
                f"{self._col.python_name[0].upper()}{self._col.python_name[1:]}Enum"
            )
            return enum_name

        mapping = {
            DetectedType.STRING: "str",
            DetectedType.TEXT: "str",
            DetectedType.INTEGER: "int",
            DetectedType.FLOAT: "float",
            DetectedType.BOOLEAN: "bool",
            DetectedType.DATETIME: "datetime",
            DetectedType.DATE: "date",
            DetectedType.EMAIL: "EmailStr",
            DetectedType.URL: "str",
        }
        return mapping.get(self._col.detected_type, "str")

    @property
    def pydantic_annotation(self) -> str:
        """Full Pydantic field annotation with Optional and Field()."""
        base = self.pydantic_base_type
        c = self._col.constraint

        field_args = []
        if c.gt is not None:
            field_args.append(f"gt={c.gt}")
        if c.ge is not None:
            field_args.append(f"ge={c.ge}")
        if c.lt is not None:
            field_args.append(f"lt={c.lt}")
        if c.le is not None:
            field_args.append(f"le={c.le}")
        if c.min_length is not None:
            field_args.append(f"min_length={c.min_length}")
        if c.max_length is not None:
            field_args.append(f"max_length={c.max_length}")

        if c.nullable:
            if field_args:
                return f"Optional[{base}] = Field(None, {', '.join(field_args)})"
            return f"Optional[{base}] = None"
        else:
            if field_args:
                return f"{base} = Field(..., {', '.join(field_args)})"
            return base


class _EnrichedSheet:
    """Wrapper around SheetAnalysis with enriched columns."""

    def __init__(self, sheet: SheetAnalysis):
        self._sheet = sheet
        self.columns = [
            _EnrichedColumn(col, sheet.model_name) for col in sheet.columns
        ]

    def __getattr__(self, name):
        return getattr(self._sheet, name)

    @property
    def has_pk(self) -> bool:
        return any(col.is_primary_key for col in self._sheet.columns)

    @property
    def pk_column(self) -> str:
        for col in self._sheet.columns:
            if col.is_primary_key:
                return col.python_name
        return "id"

    @property
    def pk_python_type(self) -> str:
        """Python type for the primary key (int or str)."""
        for col in self._sheet.columns:
            if col.is_primary_key:
                if col.detected_type in (DetectedType.INTEGER, DetectedType.FLOAT):
                    return "int"
                return "str"
        return "int"

    @property
    def pk_is_auto(self) -> bool:
        """True if PK is auto-generated (int autoincrement or synthetic)."""
        if not self.has_pk:
            return True  # synthetic id is autoincrement
        return self.pk_python_type == "int"


class _EnrichedRelationship:
    """Wrapper around RelationshipInfo with model name lookups."""

    def __init__(self, rel, sheets: list[SheetAnalysis]):
        self._rel = rel
        self._sheets = {s.table_name: s for s in sheets}

    def __getattr__(self, name):
        return getattr(self._rel, name)

    @property
    def source_model(self) -> str:
        sheet = self._sheets.get(self._rel.source_table)
        return sheet.model_name if sheet else self._rel.source_table.capitalize()

    @property
    def target_model(self) -> str:
        sheet = self._sheets.get(self._rel.target_table)
        return sheet.model_name if sheet else self._rel.target_table.capitalize()
