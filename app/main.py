"""Streamlit UI for Excel-to-API Converter."""

import os
import sys
import tempfile

from dotenv import load_dotenv
import streamlit as st

load_dotenv()

# Add parent dir to path so imports work
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
# Also add parent of app/ for when running from project root
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.analyzer import analyze_excel, clean_cell
from app.api_runner import APIRunner
from app.generator import ProjectGenerator
from app.models import AnalysisResult
from app.security import (
    check_rate_limit,
    is_public_mode,
    rate_limit_message,
    record_upload,
    validate_excel,
)


def _extract_excel_data(
    filepath: str, analysis: AnalysisResult
) -> dict[str, list[dict]]:
    """Extract all data from Excel as dict of table_name -> list[dict]."""
    import openpyxl

    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    result: dict[str, list[dict]] = {}

    for sheet_analysis in analysis.sheets:
        ws = wb[sheet_analysis.original_name]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            result[sheet_analysis.table_name] = []
            continue

        data = []
        for row in rows[1:]:
            record = {}
            for i, col in enumerate(sheet_analysis.columns):
                val = clean_cell(row[i]) if i < len(row) else None
                record[col.python_name] = val
            # Skip completely empty rows
            if all(v is None for v in record.values()):
                continue
            data.append(record)
        result[sheet_analysis.table_name] = data

    wb.close()
    return result


st.set_page_config(
    page_title="Excel to API Converter",
    page_icon="⚡",
    layout="wide",
)

st.title("Excel → API Converter")
st.markdown("Upload an Excel file and get a fully functional REST API in seconds.")

# ---- Session State ----
if "analysis" not in st.session_state:
    st.session_state.analysis = None
if "runner" not in st.session_state:
    st.session_state.runner = None
if "project_dir" not in st.session_state:
    st.session_state.project_dir = None
if "excel_data" not in st.session_state:
    st.session_state.excel_data = None
if "zip_bytes" not in st.session_state:
    st.session_state.zip_bytes = None

# ---- Mode ----
_PUBLIC = is_public_mode()

# ---- Sidebar: Configuration ----
with st.sidebar:
    st.header("Configuration")

    if _PUBLIC:
        st.info("Public mode — Live Preview is disabled.")
        api_port = 8100  # unused but keeps variable defined
    else:
        api_port = st.number_input(
            "API Preview Port", value=8100, min_value=1024, max_value=65535
        )

    st.divider()
    st.subheader("LLM Enhancement (optional)")

    if _PUBLIC:
        # In public mode the server never injects its own key.
        or_key = st.text_input(
            "Your OpenRouter API Key",
            value="",
            type="password",
            help="Bring your own key — get one at https://openrouter.ai/keys",
        )
    else:
        env_key = os.environ.get("OPENROUTER_API_KEY", "")
        or_key = st.text_input(
            "OpenRouter API Key",
            value=env_key,
            type="password",
            help="Loaded from .env if set. Get a key at https://openrouter.ai/keys",
        )

    if or_key:
        os.environ["OPENROUTER_API_KEY"] = or_key
    elif _PUBLIC:
        # Make sure no leaked server key is used
        os.environ.pop("OPENROUTER_API_KEY", None)

    or_model = st.text_input(
        "Model",
        value=os.environ.get("OPENROUTER_MODEL", "anthropic/claude-sonnet-4"),
        help="e.g. openai/gpt-4o, mistralai/mistral-large, google/gemini-2.0-flash-001",
    )
    if or_model:
        os.environ["OPENROUTER_MODEL"] = or_model
    use_llm = st.checkbox(
        "Enhance analysis with LLM",
        value=False,
        disabled=not bool(or_key),
        help="Uses OpenRouter to improve type detection and suggest better constraints.",
        key="use_llm_checkbox",
    )

    st.divider()
    if not _PUBLIC and st.session_state.runner and st.session_state.runner.is_running():
        st.success(f"API running on port {api_port}")
        if st.button("Stop API Server", type="secondary"):
            st.session_state.runner.stop()
            st.session_state.runner = None
            st.rerun()

# ---- Step 1: Upload ----
st.header("1. Upload Excel File")
uploaded = st.file_uploader(
    "Choose an Excel file (.xlsx)",
    type=["xlsx", "xls"],
    help="Upload a single Excel file. Each sheet becomes an API resource.",
)

if uploaded:
    file_bytes = uploaded.read()

    # ── Security: validate uploaded file ──
    validation = validate_excel(file_bytes, uploaded.name)
    if not validation.ok:
        for err in validation.errors:
            st.error(err)
        st.stop()

    # ── Security: rate limiting ──
    if not check_rate_limit():
        st.error(rate_limit_message())
        st.stop()

    # Save to temp file
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(file_bytes)
    tmp.close()

    if st.button("Analyze", type="primary"):
        record_upload()
        with st.spinner("Analyzing Excel file..."):
            analysis = analyze_excel(tmp.name)

            if use_llm:
                from app.llm_analyzer import enhance_analysis

                with st.spinner("Enhancing with LLM..."):
                    analysis, llm_error = enhance_analysis(analysis)
                    if llm_error:
                        st.warning(f"LLM enhancement failed: {llm_error}")
                    else:
                        st.success("LLM enhancement applied successfully.")

            st.session_state.analysis = analysis
            st.session_state.excel_data = _extract_excel_data(tmp.name, analysis)
            # Reset downstream state
            st.session_state.project_dir = None
            st.session_state.runner = None
            st.session_state.zip_bytes = None
            st.rerun()

# ---- Step 2: Analysis Preview ----
if st.session_state.analysis:
    analysis: AnalysisResult = st.session_state.analysis

    st.header("2. Analysis Result")

    # Warnings
    for w in analysis.warnings:
        st.warning(w)

    # Summary metrics
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Tables", len(analysis.sheets))
    with col2:
        total_rows = sum(s.row_count for s in analysis.sheets)
        st.metric("Total Rows", total_rows)
    with col3:
        st.metric("Relationships", len(analysis.relationships))

    # Tabs per sheet
    if analysis.sheets:
        tabs = st.tabs([s.original_name for s in analysis.sheets])
        for tab, sheet in zip(tabs, analysis.sheets):
            with tab:
                st.subheader(
                    f"Table: `{sheet.table_name}` → Model: `{sheet.model_name}` ({sheet.row_count} rows)"
                )

                # Column analysis table
                col_data = []
                for col in sheet.columns:
                    flags = []
                    if col.is_primary_key:
                        flags.append("PK")
                    if col.is_foreign_key:
                        flags.append(f"FK → {col.foreign_table}")
                    if col.enum_values:
                        flags.append(f"Enum({len(col.enum_values)})")
                    if col.confidence < 1.0:
                        flags.append("AI")

                    constraints = []
                    c = col.constraint
                    if c.ge is not None:
                        constraints.append(f">= {c.ge}")
                    if c.gt is not None:
                        constraints.append(f"> {c.gt}")
                    if c.le is not None:
                        constraints.append(f"<= {c.le}")
                    if c.min_length is not None:
                        constraints.append(f"min_len={c.min_length}")
                    if c.max_length is not None:
                        constraints.append(f"max_len={c.max_length}")
                    if not c.nullable:
                        constraints.append("required")

                    col_data.append(
                        {
                            "Column": col.original_name,
                            "Python Name": col.python_name,
                            "Type": col.detected_type.value,
                            "Constraints": ", ".join(constraints) if constraints else "-",
                            "Flags": " ".join(flags) if flags else "-",
                            "Samples": str(col.sample_values[:3]),
                        }
                    )
                st.dataframe(col_data, use_container_width=True, hide_index=True)

                # Sample data
                if sheet.sample_data:
                    with st.expander("Sample Data (first 5 rows)"):
                        st.dataframe(
                            sheet.sample_data,
                            use_container_width=True,
                            hide_index=True,
                        )

        # Relationships
        if analysis.relationships:
            st.subheader("Detected Relationships")
            for rel in analysis.relationships:
                st.markdown(
                    f"- `{rel.source_table}.{rel.source_column}` → `{rel.target_table}.{rel.target_column}`"
                )

    # ---- Step 3: Generate ----
    st.header("3. Generate API")

    gcol1, gcol2, gcol3 = st.columns(3)

    with gcol1:
        if st.button("Generate Code", type="primary"):
            with st.spinner("Generating FastAPI project..."):
                generator = ProjectGenerator(
                    analysis, st.session_state.excel_data
                )
                project_dir = tempfile.mkdtemp(prefix="excel_api_gen_")
                generator.generate_to_directory(project_dir)
                st.session_state.project_dir = project_dir

                # Pre-generate ZIP
                zip_path = generator.generate_zip()
                with open(zip_path, "rb") as f:
                    st.session_state.zip_bytes = f.read()

                st.rerun()

    with gcol2:
        if not _PUBLIC and st.session_state.project_dir:
            if st.button("Start Live Preview"):
                with st.spinner("Starting API server..."):
                    runner = APIRunner(
                        st.session_state.project_dir, port=api_port
                    )
                    if runner.start():
                        st.session_state.runner = runner
                        st.rerun()
                    else:
                        stderr = runner.get_stderr()
                        st.error(
                            f"Failed to start API server.\n\n```\n{stderr}\n```"
                        )

    with gcol3:
        if st.session_state.zip_bytes:
            st.download_button(
                label="Download ZIP",
                data=st.session_state.zip_bytes,
                file_name="generated_api.zip",
                mime="application/zip",
            )

    # ---- Step 4: Live Preview (local mode only) ----
    if not _PUBLIC and st.session_state.runner and st.session_state.runner.is_running():
        st.header("4. Live API Preview")
        runner = st.session_state.runner
        st.success(f"API is running at **{runner.get_url()}**")
        st.markdown(
            f"[Open Swagger UI in new tab]({runner.get_docs_url()})"
        )

        st.components.v1.iframe(runner.get_docs_url(), height=600, scrolling=True)

    # ---- Generated Code Preview ----
    if st.session_state.project_dir:
        st.header("Generated Code")
        files = [
            "main.py",
            "models.py",
            "schemas.py",
            "crud.py",
            "database.py",
            "seed_data.py",
            "requirements.txt",
            "README.md",
        ]
        code_tabs = st.tabs(files)
        for code_tab, filename in zip(code_tabs, files):
            filepath = os.path.join(st.session_state.project_dir, filename)
            if os.path.exists(filepath):
                with code_tab:
                    with open(filepath) as f:
                        lang = "python" if filename.endswith(".py") else "markdown" if filename.endswith(".md") else "text"
                        st.code(f.read(), language=lang)
