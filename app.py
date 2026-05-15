"""
app.py
T.A.N.G.S - Transformative AI-Powered NLP & Generation Suite
Complete Streamlit frontend.

Run:  streamlit run app.py
Docs: README.md

Author: SAKEC Mumbai - Final Year B.Tech Project 2025-26
Python: 3.9+
"""
import os
import tempfile

import pandas as pd
import streamlit as st
from groq import Groq

from pdf_extractor import HybridExtractor
from academic_module import (
    process_result,
    get_kt_students,
    failure_rate_per_subject,
)
from timetable_module import process_timetable
from financial_module import process_financial
from query_router import get_suggested_queries
from output_generator import (
    ChartGenerator,
    ExcelGenerator,
    PowerPointGenerator,
)


# ── Page config (must be the very first Streamlit call) ───────────────
st.set_page_config(
    page_title="T.A.N.G.S",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ── Chart type label → PPT generator key mapping ──────────────────────
_CHART_MAP = {
    "bar":          "bar_chart",
    "pie":          "pie_chart",
    "line":         "line_chart",
    "failure_rate": "failure_rate_chart",
}

# ── Document-type override → internal key ─────────────────────────────
_DOCTYPE_MAP = {
    "Academic Result":  "result",
    "Timetable":        "timetable",
    "Financial Report": "financial",
}


# ── Session-state defaults (initialized once per browser session) ──────
_SS_DEFAULTS: dict = {
    "query_text":       "",
    "analysis_output":  None,
    "analysis_result":  None,
    "last_file_id":     None,
    "excel_bytes":      None,
    "ppt_bytes":        None,
}
for _k, _v in _SS_DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ======================================================================
# SIDEBAR
# ======================================================================
st.sidebar.title("⚙️ Settings")

# API key — sidebar input takes priority; fallback to Streamlit secrets
api_key = st.sidebar.text_input(
    "Groq API Key",
    type="password",
    help="Free key from console.groq.com",
    placeholder="gsk_...",
)
if not api_key:
    try:
        api_key = st.secrets.get("GROQ_API_KEY", "")
    except Exception:
        api_key = ""

if not api_key:
    st.sidebar.warning("⚠️ Enter a Groq API key to enable analysis.")
else:
    st.sidebar.success("✅ API key loaded.")

st.sidebar.markdown("---")

doc_type_override = st.sidebar.selectbox(
    "Document Type Override",
    options=["Auto Detect", "Academic Result", "Timetable", "Financial Report"],
    help=(
        "Override auto-detection when the extractor guesses wrong. "
        "'Auto Detect' uses keyword heuristics from the PDF text."
    ),
)

st.sidebar.markdown("---")
st.sidebar.subheader("Output Options")

export_excel = st.sidebar.checkbox("Export Excel Report", value=True)
export_ppt   = st.sidebar.checkbox("Export PowerPoint Report", value=False)

chart_types = st.sidebar.multiselect(
    "Chart Types (PowerPoint)",
    options=["bar", "pie", "line", "failure_rate"],
    default=["bar", "failure_rate"],
    help="Charts to embed in each PowerPoint slide.",
)

ppt_theme = st.sidebar.selectbox(
    "PPT Theme",
    options=["corporate_blue", "academic_gold", "clean_light", "dark_mode"],
)

pass_mark = st.sidebar.slider(
    "Pass Mark Threshold",
    min_value=30,
    max_value=50,
    value=40,
    step=1,
    help="Minimum ESE marks required to pass a subject (SAKEC default: 40).",
)

st.sidebar.markdown("---")
st.sidebar.caption(
    "**T.A.N.G.S v1.0**  \n"
    "SAKEC ECS Dept | 2025-26  \n"
    "[Get free Groq API key →](https://console.groq.com)"
)


# ======================================================================
# MAIN HEADER
# ======================================================================
hdr_col, ttl_col = st.columns([1, 4])
with hdr_col:
    st.markdown("# 📊 T.A.N.G.S")
with ttl_col:
    st.markdown("### Transformative AI-Powered NLP & Generation Suite")
    st.caption(
        "SAKEC Final Year Project  |  ECS Department  |  2025-26  |  "
        "Powered by Groq LLaMA-3 · pdfplumber · pytesseract"
    )

st.markdown("---")


# ======================================================================
# FILE UPLOADER
# ======================================================================
uploaded_file = st.file_uploader(
    "📂 Upload Document",
    type=["pdf", "csv", "xlsx", "png", "jpg"],
    help=(
        "Upload a SAKEC academic result PDF, class timetable PDF, "
        "or any financial report PDF. Max 50 MB."
    ),
)


# ======================================================================
# BRANCH A — FILE IS PRESENT
# ======================================================================
if uploaded_file is not None:

    # Clear cached analysis when the user swaps to a different file ----
    _fid = f"{uploaded_file.name}::{uploaded_file.size}"
    if st.session_state.last_file_id != _fid:
        st.session_state.analysis_output = None
        st.session_state.analysis_result = None
        st.session_state.excel_bytes     = None
        st.session_state.ppt_bytes       = None
        st.session_state.query_text      = ""
        st.session_state.last_file_id    = _fid

    # File info pill
    _kb = uploaded_file.size / 1024
    _size_str = f"{_kb / 1024:.1f} MB" if _kb > 1024 else f"{_kb:.1f} KB"
    st.caption(f"📄 **{uploaded_file.name}**  ·  {_size_str}")

    # -- Query text input -----------------------------------------------
    query = st.text_input(
        "🔎 Enter your query",
        key="query_text",
        placeholder=(
            "e.g.  Who topped overall?   ·   "
            "Workload of Manjusha Kulkarni?   ·   "
            "List all KT students"
        ),
    )

    # -- Suggested quick-queries (type-appropriate) --------------------
    if doc_type_override == "Timetable":
        _sug_type = "timetable"
    elif doc_type_override == "Financial Report":
        _sug_type = "financial"
    else:
        _sug_type = "result"

    _suggestions = get_suggested_queries(_sug_type)
    if _suggestions:
        st.markdown("**Quick queries:**")
        _sug_cols = st.columns(3)
        for _i, _sug in enumerate(_suggestions):
            if _sug_cols[_i % 3].button(
                _sug, key=f"sug_{_i}", use_container_width=True
            ):
                st.session_state.query_text = _sug
                st.rerun()

    # -- Analyse button ------------------------------------------------
    _btn_disabled = not api_key or not query.strip()
    analyse_btn = st.button(
        "🔍 Analyse Document",
        type="primary",
        disabled=_btn_disabled,
        help=(
            "Enter your Groq API key and a query to enable this button."
            if _btn_disabled
            else "Click to extract, analyse, and generate reports."
        ),
    )

    # ==================================================================
    # ANALYSIS PIPELINE (runs when button is clicked)
    # ==================================================================
    if analyse_btn:

        if not api_key:
            st.error(
                "⚠️ Please enter your Groq API key in the sidebar.  "
                "Get a free key from [console.groq.com](https://console.groq.com)."
            )
            st.stop()

        _groq_client = Groq(api_key=api_key)
        _extractor   = HybridExtractor()

        # ---- Step 1: Extract ----------------------------------------
        with st.spinner("📂 Extracting document content…"):
            _suffix  = f".{uploaded_file.name.rsplit('.', 1)[-1].lower()}"
            _tmp_path = None
            try:
                with tempfile.NamedTemporaryFile(
                    delete=False, suffix=_suffix
                ) as _tmp:
                    _tmp.write(uploaded_file.read())
                    _tmp_path = _tmp.name
                _extract_result = _extractor.extract(_tmp_path)
            except Exception as _exc:
                st.error(f"❌ Extraction failed: {_exc}")
                st.stop()
            finally:
                if _tmp_path and os.path.exists(_tmp_path):
                    try:
                        os.unlink(_tmp_path)
                    except Exception:
                        pass

            # Apply override
            if doc_type_override != "Auto Detect":
                _extract_result["doc_type"] = _DOCTYPE_MAP[doc_type_override]

            st.success(
                f"✅ Extracted using **{_extract_result['method_used']}**  ·  "
                f"Document type: **{_extract_result['doc_type']}**"
            )

        # ---- Step 2: AI Analysis ------------------------------------
        with st.spinner("🤖 Analysing with Groq LLaMA-3…"):
            try:
                _doc_type = _extract_result["doc_type"]
                if _doc_type == "result":
                    _analysis_out = process_result(
                        _extract_result, query, _groq_client
                    )
                    # Re-compute KT / failure rates with custom pass_mark
                    if pass_mark != 40:
                        _df = _analysis_out.get("dataframe", pd.DataFrame())
                        if _df is not None and not _df.empty:
                            _analysis_out["kt_df"] = get_kt_students(
                                _df, pass_mark=pass_mark
                            )
                            _analysis_out["failure_rates"] = failure_rate_per_subject(
                                _df, pass_mark=pass_mark
                            )
                elif _doc_type == "timetable":
                    _analysis_out = process_timetable(
                        _extract_result, query, _groq_client
                    )
                else:
                    _analysis_out = process_financial(
                        _extract_result, query, _groq_client
                    )
            except Exception as _exc:
                st.error(f"❌ Analysis error: {_exc}")
                st.stop()

            # Enrich with metadata so downstream generators can use them
            _analysis_out["query"]    = query
            _analysis_out["doc_type"] = _doc_type

        # ---- Step 3: Pre-generate Excel / PPT -----------------------
        with st.spinner("📊 Generating downloadable reports…"):
            _excel_bytes = None
            _ppt_bytes   = None

            try:
                _excel_gen   = ExcelGenerator()
                _excel_bytes = _excel_gen.generate(_analysis_out, _doc_type)
            except Exception as _exc:
                st.warning(f"Excel generation issue: {_exc}")

            try:
                _ct_ppt = [
                    _CHART_MAP.get(ct, ct + "_chart") for ct in chart_types
                ]
                _ppt_gen   = PowerPointGenerator()
                _ppt_bytes = _ppt_gen.generate(_analysis_out, _ct_ppt, ppt_theme)
            except Exception as _exc:
                st.warning(f"PowerPoint generation issue: {_exc}")

        # ---- Persist to session state --------------------------------
        st.session_state.analysis_output = _analysis_out
        st.session_state.analysis_result = _extract_result
        st.session_state.excel_bytes     = _excel_bytes
        st.session_state.ppt_bytes       = _ppt_bytes

    # ==================================================================
    # DISPLAY RESULTS (shown on every rerun while output is stored)
    # ==================================================================
    if (
        st.session_state.analysis_output is not None
        and st.session_state.analysis_result is not None
    ):
        _output   = st.session_state.analysis_output
        _result_m = st.session_state.analysis_result
        _dtype    = _result_m["doc_type"]

        st.markdown("---")
        st.subheader("📋 Analysis Results")

        # LLM answer banner
        _qa = _output.get("query_answer", "")
        if _qa:
            st.info(f"**🤖 Answer:**  {_qa}")

        # Raw extracted DataFrame
        _main_df = _output.get("dataframe")
        if _main_df is not None and not _main_df.empty:
            with st.expander("🗂️ View Extracted Data", expanded=True):
                st.dataframe(_main_df, use_container_width=True)

        # ----------------------------------------------------------
        # RESULT-specific sections
        # ----------------------------------------------------------
        if _dtype == "result":
            _stats = _output.get("summary_stats", {}) or {}
            _m1, _m2, _m3, _m4 = st.columns(4)
            _m1.metric(
                "Total Students",
                _stats.get("total_students", _stats.get("total", 0)),
            )
            _m2.metric("Passed",  _stats.get("passed", 0))
            _m3.metric("ATKT",    _stats.get("atkt",   0))
            _m4.metric(
                "Pass %",
                f"{_stats.get('overall_pass_percentage', _stats.get('pass_pct', 0.0)):.1f}%",
            )

            _tab_top, _tab_kt, _tab_sub = st.tabs(
                ["🏆 Toppers", "⚠️ KT Students", "📊 Subject Stats"]
            )

            with _tab_top:
                _topper_df = _output.get("topper_df")
                if _topper_df is not None and not _topper_df.empty:
                    st.markdown(
                        f"**Top {len(_topper_df)} students by total ESE marks**"
                    )
                    st.dataframe(_topper_df, use_container_width=True)
                    _cg = ChartGenerator()
                    _x  = "name" if "name" in _topper_df.columns else _topper_df.columns[0]
                    _y  = "total" if "total" in _topper_df.columns else (
                        next(
                            (
                                c for c in _topper_df.columns
                                if pd.api.types.is_numeric_dtype(_topper_df[c])
                            ),
                            None,
                        )
                    )
                    if _y:
                        _chart = _cg.bar_chart(
                            _topper_df.head(10), _x, _y,
                            "Top 10 Students — Total ESE Marks",
                        )
                        st.image(_chart, use_container_width=True)
                else:
                    st.caption("No topper data available.")

            with _tab_kt:
                _kt_df = _output.get("kt_df")
                if _kt_df is not None and not _kt_df.empty:
                    st.markdown(
                        f"**{len(_kt_df)} student(s) with KT / backlog  "
                        f"(pass mark used: {pass_mark})**"
                    )
                    st.dataframe(_kt_df, use_container_width=True)
                else:
                    st.success(
                        f"🎉 No KT students found at pass mark {pass_mark}."
                    )

            with _tab_sub:
                _fr = _output.get("failure_rates")
                if _fr:
                    _cg2 = ChartGenerator()
                    _fr_chart = _cg2.failure_rate_chart(_fr)
                    st.image(_fr_chart, use_container_width=True)

                    _fr_df = pd.DataFrame(
                        [
                            {
                                "Subject":         s,
                                "Passed":          v.get("passed", 0),
                                "Failed":          v.get("failed", 0),
                                "Failure Rate %":  v.get("rate",   0.0),
                            }
                            for s, v in sorted(
                                _fr.items(),
                                key=lambda x: x[1].get("rate", 0),
                                reverse=True,
                            )
                        ]
                    )
                    st.dataframe(_fr_df, use_container_width=True)
                else:
                    st.caption("No subject failure-rate data available.")

        # ----------------------------------------------------------
        # TIMETABLE-specific sections
        # ----------------------------------------------------------
        elif _dtype == "timetable":
            _tab_wl, _tab_sched = st.tabs(
                ["👨‍🏫 Teacher Workload", "📅 Full Schedule"]
            )
            with _tab_wl:
                _wl_df = _output.get("workload_df")
                if _wl_df is not None and not _wl_df.empty:
                    st.dataframe(_wl_df, use_container_width=True)
                    if (
                        "teacher_name" in _wl_df.columns
                        and "total_weekly_lectures" in _wl_df.columns
                    ):
                        _cg3  = ChartGenerator()
                        _wchart = _cg3.bar_chart(
                            _wl_df,
                            "teacher_name",
                            "total_weekly_lectures",
                            "Weekly Lecture Load per Teacher",
                        )
                        st.image(_wchart, use_container_width=True)
                else:
                    st.caption("No workload data available.")

            with _tab_sched:
                _sdf = _output.get("dataframe")
                if _sdf is not None and not _sdf.empty:
                    st.dataframe(_sdf, use_container_width=True)
                else:
                    st.caption("No schedule data available.")

            # Subject frequency
            _freq = _output.get("subject_freq", {}) or {}
            if _freq:
                with st.expander("📖 Subject Frequency", expanded=False):
                    _freq_df = pd.DataFrame(
                        sorted(_freq.items(), key=lambda x: -x[1]),
                        columns=["Subject", "Weekly Slots"],
                    )
                    st.dataframe(_freq_df, use_container_width=True)

        # ----------------------------------------------------------
        # FINANCIAL-specific sections
        # ----------------------------------------------------------
        elif _dtype == "financial":
            _kpis = _output.get("kpis", {}) or {}
            if _kpis:
                _kpi_names = list(_kpis.items())
                _n_cols    = min(len(_kpi_names), 4)
                _kcols     = st.columns(_n_cols)
                for _ki, (_kname, _kval) in enumerate(_kpi_names):
                    _kcols[_ki % _n_cols].metric(
                        _kname.replace("_", " ").title(),
                        f"{_kval:,.2f}",
                    )

            _yoy = _output.get("yoy_growth")
            if _yoy is not None:
                st.metric("📈 YoY Growth", f"{_yoy:.2f}%")

            if _kpis:
                _cg4 = ChartGenerator()
                _kdf = pd.DataFrame(
                    list(_kpis.items()), columns=["Metric", "Value"]
                )
                _kchart = _cg4.bar_chart(
                    _kdf, "Metric", "Value", "Financial KPIs Overview"
                )
                st.image(_kchart, use_container_width=True)

        # ----------------------------------------------------------
        # DOWNLOAD SECTION
        # ----------------------------------------------------------
        st.markdown("---")
        st.subheader("📥 Download Reports")
        _dl1, _dl2 = st.columns(2)

        with _dl1:
            if export_excel and st.session_state.excel_bytes:
                st.download_button(
                    label="📥 Download Excel Report",
                    data=st.session_state.excel_bytes,
                    file_name=f"TANGS_{_dtype}_report.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument"
                        ".spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                )
            elif export_excel:
                st.caption("Excel generation failed — re-run analysis.")

        with _dl2:
            if export_ppt and st.session_state.ppt_bytes:
                st.download_button(
                    label="📥 Download PowerPoint Report",
                    data=st.session_state.ppt_bytes,
                    file_name=f"TANGS_{_dtype}_report.pptx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument"
                        ".presentationml.presentation"
                    ),
                    use_container_width=True,
                )
            elif export_ppt:
                st.caption("PowerPoint generation failed — re-run analysis.")

        st.success("✅ Analysis complete!")


# ======================================================================
# BRANCH B — NO FILE UPLOADED — LANDING PAGE
# ======================================================================
else:
    st.markdown(
        """
        ### 🚀 How to use T.A.N.G.S

        1. Enter your **Groq API key** in the sidebar  
           → Free key from [console.groq.com](https://console.groq.com) (no credit card)
        2. **Upload** a PDF document — result sheet, timetable, or financial report
        3. **Type your query** or click a quick-query chip below
        4. Click **🔍 Analyse Document**
        5. View the AI answer, explore result tabs, and **download** Excel / PowerPoint

        ---

        #### 📂 Supported Documents

        | Type | Example | Key Queries |
        |------|---------|-------------|
        | 📝 **Academic Results** | SAKEC B.Tech result PDFs | Who topped? · List KT students · Pass % in Physics |
        | 📅 **Class Timetables** | SAKEC semester timetables | Workload of /MK? · Free slots Friday? · Rooms used? |
        | 💰 **Financial Reports** | Annual reports, P&L sheets | Revenue? · Net profit? · YoY growth? · EPS? |

        ---

        #### 🔧 Extraction Engine

        | Tier | Method | Best For |
        |------|--------|----------|
        | **1** | `pdfplumber` table extraction | Clean digital PDFs |
        | **2** | Regex on raw text | Semi-structured PDFs |
        | **3** | `pytesseract` OCR + auto-rotation | Scanned / image PDFs |

        ---

        #### 🛠 Tech Stack
        `Streamlit` · `Groq LLaMA-3 8B` · `pdfplumber` · `pytesseract` ·
        `openpyxl` · `python-pptx` · `matplotlib` · `fuzzywuzzy` · `scikit-learn`

        ---
        *Built with ❤️ by the ECS Department, SAKEC Mumbai — Final Year B.Tech Project 2025-26*
        """
    )
