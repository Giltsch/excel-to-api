"""Optional LLM enhancement using OpenRouter to improve analysis results."""

import json
import os
from typing import Optional

from .models import (
    AnalysisResult,
    ColumnConstraint,
    DetectedType,
    RelationshipInfo,
)


def is_llm_available() -> bool:
    """Check if OpenRouter API key is configured."""
    return bool(os.environ.get("OPENROUTER_API_KEY"))


def enhance_analysis(analysis: AnalysisResult) -> tuple[AnalysisResult, Optional[str]]:
    """Send rule-based analysis to LLM via OpenRouter for enhancement.
    
    Returns:
        Tuple of (enhanced_analysis, error_message_or_none).
    """
    if not is_llm_available():
        return analysis, "No OpenRouter API key configured."

    try:
        from openai import OpenAI
    except ImportError:
        return analysis, "Package 'openai' not installed. Run: pip install openai"

    client = OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=os.environ["OPENROUTER_API_KEY"],
    )

    model = os.environ.get("OPENROUTER_MODEL", "anthropic/claude-sonnet-4")
    summary = _build_analysis_summary(analysis)

    prompt = f"""You are an API design expert. I have an Excel file that I'm converting to a REST API.
The rule-based analyzer produced this result:

{json.dumps(summary, indent=2)}

Review each column and suggest improvements. For each column, you may:
1. Change the detected_type if you think a better type fits
2. Add or modify constraints (nullable, unique, gt, ge, lt, le, min_length, max_length)
3. Identify relationships the rule-based system might have missed

Valid detected_type values: STRING, INTEGER, FLOAT, BOOLEAN, DATETIME, DATE, EMAIL, URL, ENUM, TEXT

Respond with ONLY a JSON object:
{{
  "changes": [
    {{
      "table": "table_name",
      "column": "column_name",
      "detected_type": "NEW_TYPE or null if no change",
      "constraint_updates": {{}},
      "reason": "brief explanation"
    }}
  ],
  "new_relationships": [
    {{
      "source_table": "...",
      "source_column": "...",
      "target_table": "...",
      "target_column": "..."
    }}
  ]
}}
Only include columns where you have a meaningful improvement. If everything looks good, return empty arrays."""

    try:
        response = client.chat.completions.create(
            model=model,
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}],
        )
        response_text = response.choices[0].message.content

        # Extract JSON from response (handle markdown code blocks)
        json_text = response_text
        if "```" in json_text:
            start = json_text.find("{")
            end = json_text.rfind("}") + 1
            if start >= 0 and end > start:
                json_text = json_text[start:end]

        changes = json.loads(json_text)
        return _apply_llm_changes(analysis, changes), None
    except Exception as e:
        return analysis, f"LLM call failed ({type(e).__name__}): {e}"


def _build_analysis_summary(analysis: AnalysisResult) -> dict:
    """Build compact representation for the LLM prompt."""
    result: dict = {"tables": []}
    for sheet in analysis.sheets:
        table = {
            "name": sheet.table_name,
            "row_count": sheet.row_count,
            "columns": [],
        }
        for col in sheet.columns:
            table["columns"].append(
                {
                    "name": col.python_name,
                    "original_name": col.original_name,
                    "detected_type": col.detected_type.name,
                    "is_pk": col.is_primary_key,
                    "is_fk": col.is_foreign_key,
                    "nullable": col.constraint.nullable,
                    "sample_values": [str(v) for v in col.sample_values[:3]],
                }
            )
        result["tables"].append(table)

    result["relationships"] = [
        {
            "source": f"{r.source_table}.{r.source_column}",
            "target": f"{r.target_table}.{r.target_column}",
        }
        for r in analysis.relationships
    ]
    return result


def _apply_llm_changes(analysis: AnalysisResult, changes: dict) -> AnalysisResult:
    """Apply LLM-suggested changes to the analysis result."""
    type_map = {t.name: t for t in DetectedType}

    for change in changes.get("changes", []):
        table_name = change.get("table")
        col_name = change.get("column")

        sheet = next(
            (s for s in analysis.sheets if s.table_name == table_name), None
        )
        if not sheet:
            continue
        col = next(
            (c for c in sheet.columns if c.python_name == col_name), None
        )
        if not col:
            continue

        new_type = change.get("detected_type")
        if new_type and new_type in type_map:
            col.detected_type = type_map[new_type]
            col.confidence = 0.8

        for key, value in change.get("constraint_updates", {}).items():
            if hasattr(col.constraint, key):
                setattr(col.constraint, key, value)

    for rel in changes.get("new_relationships", []):
        try:
            analysis.relationships.append(
                RelationshipInfo(
                    source_table=rel["source_table"],
                    source_column=rel["source_column"],
                    target_table=rel["target_table"],
                    target_column=rel["target_column"],
                )
            )
        except KeyError:
            continue

    return analysis
