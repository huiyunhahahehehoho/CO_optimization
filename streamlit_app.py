from __future__ import annotations

import base64
from io import BytesIO
from pathlib import Path

import streamlit as st
from PIL import Image

from optimizer_backend import (
    FULL_ENUMERATION_LABEL,
    get_algorithm_options,
    get_available_display_products,
    run_optimizer,
)

st.set_page_config(
    page_title="Product Changeover Optimization Tool",
    page_icon="🔄",
    layout="wide",
)

# -------------------------------------------------
# PATHS
# -------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
ABBOTT_LOGO_PATH = BASE_DIR / "abbott-logo.png"
CO_MATRIX_PATH = BASE_DIR / "COMatrix.xlsx"

# -------------------------------------------------
# PRODUCT LIST FROM CO MATRIX
# -------------------------------------------------
PRODUCT_LOAD_ERROR = None
try:
    ALL_PRODUCTS = get_available_display_products(CO_MATRIX_PATH)
except Exception as exc:
    ALL_PRODUCTS = []
    PRODUCT_LOAD_ERROR = str(exc)

DEFAULT_PRODUCTS = []

# -------------------------------------------------
# LOGO
# -------------------------------------------------
def load_logo_base64(path: Path):
    if not path.exists():
        return None

    img = Image.open(path).convert("RGBA")
    buffer = BytesIO()
    img.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


ABBOTT_LOGO_B64 = load_logo_base64(ABBOTT_LOGO_PATH)

# -------------------------------------------------
# SESSION STATE DEFAULTS
# -------------------------------------------------
def init_state():
    defaults = {
        "select_all": False,
        "chosen_products": DEFAULT_PRODUCTS.copy(),
        "problem_type": "Open",
        "time_limit_hours": 1.0,
        "selected_algorithms": [],
        "use_gurobi_wls": False,
        "gurobi_wls_accessid": "",
        "gurobi_wls_secret": "",
        "gurobi_license_id": "",
        "result": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_state()

# -------------------------------------------------
# RESET FUNCTION
# -------------------------------------------------
def reset_all():
    for key in [
        "select_all",
        "chosen_products",
        "problem_type",
        "time_limit_hours",
        "selected_algorithms",
        "use_gurobi_wls",
        "gurobi_wls_accessid",
        "gurobi_wls_secret",
        "gurobi_license_id",
        "result",
        "full_enum_locked_display",
        "algorithm_placeholder_display",
    ]:
        if key in st.session_state:
            del st.session_state[key]
    init_state()


# -------------------------------------------------
# SELECT ALL CALLBACK
# -------------------------------------------------
def select_all_callback():
    if st.session_state.select_all:
        st.session_state.chosen_products = ALL_PRODUCTS.copy()
    else:
        st.session_state.chosen_products = []


# -------------------------------------------------
# MULTISELECT CALLBACK
# -------------------------------------------------
def multiselect_callback():
    st.session_state.select_all = len(st.session_state.chosen_products) == len(ALL_PRODUCTS)


# -------------------------------------------------
# STYLES
# -------------------------------------------------
st.markdown(
    """
    <style>
        :root {
            --bg-1: #dff3ff;
            --bg-2: #b9dcff;
            --bg-3: #79b7ff;
            --bg-4: #3b82f6;
            --bg-5: #1d4ed8;

            --card-white: rgba(255,255,255,0.92);
            --card-soft: rgba(247,250,255,0.82);
            --text-dark: #0f172a;
            --text-mid: #475569;
            --text-soft: #64748b;
            --border: #d9e6f7;
            --accent: #2563eb;
            --accent-2: #1d4ed8;
        }

        .stApp {
            background:
                radial-gradient(circle at top left, rgba(255,255,255,0.35), transparent 28%),
                linear-gradient(135deg, var(--bg-1) 0%, var(--bg-2) 25%, var(--bg-3) 55%, var(--bg-4) 82%, var(--bg-5) 100%);
        }

        header[data-testid="stHeader"] { display: none; }
        div[data-testid="stToolbar"] { display: none; }
        #MainMenu { visibility: hidden; }
        footer { visibility: hidden; }

        .block-container {
            max-width: 1360px;
            padding-top: 1rem;
            padding-bottom: 2rem;
        }

        .hero {
            background: linear-gradient(135deg, rgba(255,255,255,0.88), rgba(255,255,255,0.72));
            border: 1px solid rgba(255,255,255,0.45);
            border-radius: 28px;
            padding: 26px 28px 22px 28px;
            box-shadow: 0 18px 40px rgba(30, 64, 175, 0.14);
            margin-bottom: 16px;
        }

        .hero-inner {
            display: flex;
            align-items: flex-start;
            gap: 24px;
        }

        .hero-logo {
            flex: 0 0 150px;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 92px;
        }

        .hero-logo img {
            max-width: 180px;
            height: auto;
            display: block;
        }

        .hero-copy {
            flex: 1;
            min-width: 0;
        }

        .hero-title {
            font-size: 2.15rem;
            font-weight: 900;
            color: #0f172a;
            line-height: 1.15;
            margin-bottom: 0.45rem;
        }

        .hero-subtitle {
            font-size: 1rem;
            color: #475569;
            line-height: 1.7;
            max-width: 980px;
        }

        div[data-testid="stHorizontalBlock"] > div:nth-child(1),
        div[data-testid="stHorizontalBlock"] > div:nth-child(2) {
            border-radius: 26px;
            border: 1px solid rgba(255,255,255,0.45);
            padding: 26px 24px;
            min-height: 740px;
            box-shadow: 0 16px 38px rgba(30, 64, 175, 0.14);
        }

        div[data-testid="stHorizontalBlock"] > div:nth-child(1) {
            background: linear-gradient(180deg, rgba(241,247,255,0.96) 0%, rgba(232,241,255,0.92) 100%);
        }

        div[data-testid="stHorizontalBlock"] > div:nth-child(2) {
            background: linear-gradient(180deg, rgba(255,255,255,0.97) 0%, rgba(248,251,255,0.96) 100%);
        }

        .panel-title {
            font-size: 1.55rem;
            font-weight: 850;
            color: var(--text-dark);
            margin-bottom: 0.25rem;
            line-height: 1.2;
        }

        .panel-subtitle {
            font-size: 0.98rem;
            color: var(--text-mid);
            margin-bottom: 1.25rem;
            line-height: 1.65;
        }

        .section-label {
            font-size: 0.78rem;
            font-weight: 800;
            color: var(--text-soft);
            text-transform: uppercase;
            letter-spacing: 0.09em;
            margin-top: 0.6rem;
            margin-bottom: 0.42rem;
        }

        .section-header-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            flex-wrap: wrap;
            margin-top: 0.6rem;
            margin-bottom: 0.42rem;
        }

        .algo-pill {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 7px 12px;
            border-radius: 999px;
            background: linear-gradient(180deg, #eef4ff 0%, #e3edff 100%);
            border: 1px solid #d5e4fb;
            color: #294a78;
            font-size: 0.86rem;
            font-weight: 700;
            line-height: 1.2;
            box-shadow: inset 0 1px 0 rgba(255,255,255,0.7);
        }

        .soft-alert {
            border: 1px solid #dce9fb;
            border-radius: 18px;
            padding: 14px 16px;
            background: linear-gradient(180deg, #f8fbff 0%, #eef5ff 100%);
            color: #35507b;
            line-height: 1.6;
            margin-top: 10px;
        }

        .empty-state {
            border: 1.5px dashed #c8daf8;
            border-radius: 20px;
            padding: 28px 22px;
            background: linear-gradient(180deg, #fbfdff 0%, #f3f8ff 100%);
            color: #52627c;
            line-height: 1.8;
        }

        .sequence-box {
            background: linear-gradient(180deg, #f9fbff 0%, #eef4ff 100%);
            border: 1px solid #dce7f8;
            border-radius: 20px;
            padding: 18px;
            margin-top: 10px;
            margin-bottom: 16px;
            font-size: 1rem;
            line-height: 1.9;
            color: #16325b;
            word-break: break-word;
            box-shadow: inset 0 1px 0 rgba(255,255,255,0.75);
        }

        .note-box {
            border: 1px solid #dce7f8;
            border-radius: 18px;
            background: linear-gradient(180deg, #f9fbff 0%, #f1f6ff 100%);
            padding: 15px 16px;
            color: #4b5f7b;
            line-height: 1.65;
            margin-top: 12px;
        }

        .metric-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 12px;
            margin-top: 8px;
            margin-bottom: 14px;
        }

        .metric-card {
            background: linear-gradient(180deg, #ffffff 0%, #f7faff 100%);
            border: 1px solid #dbe7fa;
            border-radius: 18px;
            padding: 16px;
            box-shadow: 0 8px 18px rgba(37, 99, 235, 0.06);
        }

        .metric-label {
            font-size: 0.77rem;
            font-weight: 800;
            color: #64748b;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            margin-bottom: 8px;
        }

        .metric-value {
            font-size: 1.35rem;
            font-weight: 900;
            color: #0f172a;
            line-height: 1.2;
        }

        .metric-sub {
            font-size: 0.9rem;
            color: #64748b;
            margin-top: 4px;
        }

        .muted-note {
            color: #5b6b84;
            font-size: 0.92rem;
            line-height: 1.6;
            margin-top: 16px;
        }

        div[data-baseweb="select"] > div {
            background-color: rgba(255,255,255,0.96) !important;
            border-radius: 16px !important;
            border: 1px solid #d5e3f8 !important;
            min-height: 54px !important;
        }

        div[data-baseweb="select"] > div:focus-within,
        div[data-baseweb="input"] > div:focus-within {
            border-color: var(--accent) !important;
            box-shadow: 0 0 0 1px var(--accent) !important;
            outline: none !important;
        }

        div[data-baseweb="input"] > div {
            background-color: rgba(255,255,255,0.96) !important;
            border-radius: 16px !important;
            border: 1px solid #d5e3f8 !important;
        }

        span[data-baseweb="tag"] {
            background: #ebf2ff !important;
            border: 1px solid #d2e0fb !important;
            color: #314d7d !important;
            border-radius: 10px !important;
        }

        span[data-baseweb="tag"] * {
            color: #314d7d !important;
            fill: #314d7d !important;
        }

        div[role="radiogroup"] label {
            background: #edf3ff !important;
            border: 1px solid #d7e3f8 !important;
            padding: 9px 14px;
            border-radius: 14px;
            margin-right: 10px;
        }

        div[data-testid="stNumberInput"] > div {
            border-radius: 16px !important;
        }

        div[data-testid="stNumberInput"] button {
            background: #eef2f7 !important;
            border: 1px solid #d3dce8 !important;
            color: #475569 !important;
            box-shadow: none !important;
        }

        div[data-testid="stNumberInput"] button:hover {
            background: #e2e8f0 !important;
            border-color: #c7d3e3 !important;
        }

        input[type="checkbox"], input[type="radio"] {
            accent-color: var(--accent) !important;
        }

        div.stButton > button {
            width: 100%;
            border-radius: 16px;
            padding: 0.92rem 1.1rem !important;
            font-weight: 800 !important;
            font-size: 0.98rem !important;
            box-shadow: none !important;
        }

        div.stButton > button[kind="primary"] {
            background: #6482AD !important;
            color: white !important;
            border: none !important;
        }

        div.stButton > button[kind="primary"]:hover {
            background: #99b4dd !important;
            color: white !important;
            filter: none !important;
        }

        div.stButton > button[kind="secondary"] {
            background: #f9fbff !important;
            color: #4a6a92 !important;
            border: 1px solid #dbe7fa !important;
        }

        div.stButton > button[kind="secondary"]:hover {
            background: #f2fbff !important;
            color: #35507b !important;
            border: 1px solid #cfe0f8 !important;
        }

        div[data-testid="stDataFrame"] {
            border-radius: 18px;
            overflow: hidden;
            border: 1px solid #dbe7fa;
        }

        div[data-testid="stDataFrame"] td,
        div[data-testid="stDataFrame"] th {
            white-space: normal !important;
            overflow-wrap: break-word !important;
            padding: 12px !important;
        }

        div.stDownloadButton > button {
            width: 100%;
            border-radius: 14px;
            background: white;
            border: 1px solid #d4e2f7;
            color: #17315c;
            font-weight: 800;
        }

        @media (max-width: 900px) {
            .hero-inner {
                flex-direction: column;
                gap: 12px;
            }

            .hero-logo {
                justify-content: flex-start;
                min-height: auto;
            }

            .hero-logo img {
                max-width: 200px;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# -------------------------------------------------
# HERO
# -------------------------------------------------
logo_html = ""
if ABBOTT_LOGO_B64:
    logo_html = f'<div class="hero-logo"><img src="data:image/png;base64,{ABBOTT_LOGO_B64}" alt="Abbott Logo"></div>'

hero_html = f"""
<div class="hero">
    <div class="hero-inner">
        {logo_html}
        <div class="hero-copy">
            <div class="hero-title">Product Changeover Optimization Tool</div>
            <div class="hero-subtitle">
                A decision-support interface for optimizing the production sequence of a selected set of products across open- and closed-loop formulations to reduce total changeover time, improve production efficiency, and support more effective schedule planning.
            </div>
        </div>
    </div>
</div>
"""

st.markdown(hero_html, unsafe_allow_html=True)

if PRODUCT_LOAD_ERROR:
    st.markdown(
        f"""
        <div class="soft-alert">
            <b>CO Matrix could not be loaded.</b><br>
            {PRODUCT_LOAD_ERROR}
        </div>
        """,
        unsafe_allow_html=True,
    )

left_col, right_col = st.columns([0.78, 1.32], gap="small")

# -------------------------------------------------
# LEFT PANEL
# -------------------------------------------------
with left_col:
    st.markdown('<div class="panel-title">Configuration</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="panel-subtitle">Select the products, choose the loop type, select the algorithm, and set the runtime parameters for the optimization.</div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-label">Product selection</div>', unsafe_allow_html=True)

    st.multiselect(
        "Choose which products to optimize",
        ALL_PRODUCTS,
        key="chosen_products",
        placeholder="Search and select products...",
        on_change=multiselect_callback,
        label_visibility="collapsed",
        disabled=bool(PRODUCT_LOAD_ERROR),
    )

    st.checkbox(
        f"Select all {len(ALL_PRODUCTS)} products" if ALL_PRODUCTS else "Select all products",
        key="select_all",
        on_change=select_all_callback,
        disabled=bool(PRODUCT_LOAD_ERROR),
    )

    st.markdown('<div class="section-label">Loop Type</div>', unsafe_allow_html=True)
    problem_type = st.radio(
        "Choose the loop type",
        ["Open", "Close"],
        horizontal=True,
        key="problem_type",
        label_visibility="collapsed",
        disabled=bool(PRODUCT_LOAD_ERROR),
    )

    st.markdown('<div class="section-label">Time Limit (Hours)</div>', unsafe_allow_html=True)
    time_limit_hours = st.number_input(
        "Time limit in hours",
        min_value=0.1,
        step=0.5,
        key="time_limit_hours",
        label_visibility="collapsed",
        disabled=bool(PRODUCT_LOAD_ERROR),
    )

    st.markdown('<div class="section-label">Algorithm Selection</div>', unsafe_allow_html=True)

    selected_count = len(st.session_state.chosen_products)
    available_algorithms = get_algorithm_options(problem_type, selected_count)
    force_full_enumeration = 1 <= selected_count <= 15

    if selected_count == 0:
        st.session_state.selected_algorithms = []

        st.multiselect(
            "Choose Algorithm(s)",
            options=[],
            key="algorithm_placeholder_display",
            disabled=True,
            label_visibility="collapsed",
            placeholder="Select product(s) first",
        )

    elif force_full_enumeration:
        st.session_state.selected_algorithms = [FULL_ENUMERATION_LABEL]

        st.multiselect(
            "Choose Algorithm(s)",
            options=available_algorithms,
            default=[FULL_ENUMERATION_LABEL],
            key="full_enum_locked_display",
            disabled=True,
            label_visibility="collapsed",
            placeholder="Choose Algorithm(s)",
        )

    else:
        st.session_state.selected_algorithms = [
            algo
            for algo in st.session_state.get("selected_algorithms", [])
            if algo in available_algorithms
        ]

        st.multiselect(
            "Choose Algorithm(s)",
            options=available_algorithms,
            key="selected_algorithms",
            placeholder="Choose Algorithm(s)",
            label_visibility="collapsed",
            disabled=bool(PRODUCT_LOAD_ERROR),
        )

    st.markdown('<div class="section-label">Optional Gurobi WLS License</div>', unsafe_allow_html=True)
    st.checkbox(
        "Use Gurobi WLS credentials entered in this interface",
        key="use_gurobi_wls",
        disabled=bool(PRODUCT_LOAD_ERROR),
    )

    if st.session_state.use_gurobi_wls:
        st.text_input(
            "WLS Access ID",
            key="gurobi_wls_accessid",
            placeholder="Enter WLSACCESSID",
            disabled=bool(PRODUCT_LOAD_ERROR),
        )
        st.text_input(
            "WLS Secret",
            key="gurobi_wls_secret",
            placeholder="Enter WLSSECRET",
            type="password",
            disabled=bool(PRODUCT_LOAD_ERROR),
        )
        st.text_input(
            "License ID",
            key="gurobi_license_id",
            placeholder="Enter LICENSEID",
            type="password",
            disabled=bool(PRODUCT_LOAD_ERROR),
        )

        st.markdown(
            """
            <div class="soft-alert">
                <b>Before using Gurobi from the interface:</b><br>
                1. Enter <b>WLSACCESSID</b>, <b>WLSSECRET</b>, and <b>LICENSEID</b>.<br>
                2. Select <b>Gurobi Exact</b> in the algorithm list when you want this license to be used.
            </div>
            """,
            unsafe_allow_html=True,
        )

    generate_clicked = st.button(
        "Generate Output Sequence",
        type="primary",
        use_container_width=True,
        disabled=bool(PRODUCT_LOAD_ERROR),
    )

    st.button(
        "Reset",
        type="secondary",
        use_container_width=True,
        on_click=reset_all,
    )

    if generate_clicked:
        selected_count = len(st.session_state.chosen_products)
        selected_algorithms = st.session_state.get("selected_algorithms", [])

        if selected_count == 0:
            st.session_state.result = {"error": "Please select at least one product."}
        elif selected_count > 15 and not selected_algorithms:
            st.session_state.result = {"error": "Please select at least one algorithm for optimization."}
        else:
            gurobi_license_config = None
            if st.session_state.get("use_gurobi_wls"):
                gurobi_license_config = {
                    "WLSACCESSID": st.session_state.get("gurobi_wls_accessid", ""),
                    "WLSSECRET": st.session_state.get("gurobi_wls_secret", ""),
                    "LICENSEID": st.session_state.get("gurobi_license_id", ""),
                }

            progress_placeholder = st.empty()
            progress_bar = st.progress(0)

            def ui_progress_callback(current_idx: int, total_count: int, algorithm_name: str) -> None:
                safe_total = max(1, total_count)
                progress_placeholder.info(
                    f"⏳ Optimizing... Running {algorithm_name} ({current_idx}/{safe_total})"
                )
                progress_bar.progress(min(current_idx - 1, safe_total) / safe_total)

            try:
                with st.spinner("Optimizing... Please wait while the sequence is being generated."):
                    st.session_state.result = run_optimizer(
                        selected_display_products=st.session_state.chosen_products,
                        problem_type=problem_type,
                        time_limit_hours=time_limit_hours,
                        selected_algorithms=selected_algorithms,
                        matrix_excel_path=CO_MATRIX_PATH,
                        progress_callback=ui_progress_callback,
                        gurobi_license_config=gurobi_license_config,
                    )
                progress_placeholder.success("✅ Optimization completed.")
                progress_bar.progress(1.0)
            except Exception as exc:
                progress_placeholder.empty()
                progress_bar.empty()
                st.session_state.result = {"error": str(exc)}

    st.markdown(
        """
        <div class="muted-note">
            The product list is loaded from <b>COMatrix.xlsx</b>, while the selected products, loop type, time limit, algorithm choices, and optional Gurobi WLS credentials all come directly from this user interface. When multiple algorithms are selected, the chosen time limit is now shared across them to keep the total runtime more manageable.
        </div>
        """,
        unsafe_allow_html=True,
    )

# -------------------------------------------------
# RIGHT PANEL
# -------------------------------------------------
with right_col:
    st.markdown('<div class="panel-title">Generated Output</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="panel-subtitle">Review the optimization results, including the recommended product sequence, total changeover performance, key summary metrics, and downloadable outputs.</div>',
        unsafe_allow_html=True,
    )

    result = st.session_state.get("result", None)

    if result is None:
        st.markdown(
            """
            <div class="empty-state">
                <b>No output yet.</b><br><br>
                Configure the settings on the left and click
                <b>Generate Output Sequence</b>.<br><br>
                This area is designed to present:
                <ul>
                    <li>ordered product sequence</li>
                    <li>summary metrics</li>
                    <li>preview table</li>
                    <li>downloadable results</li>
                </ul>
            </div>
            """,
            unsafe_allow_html=True,
        )

    elif "error" in result:
        st.markdown(
            f"""
            <div class="soft-alert">
                <b>Input needed.</b><br>
                {result["error"]}
            </div>
            """,
            unsafe_allow_html=True,
        )

    else:
        total_changeover_display = "-" if result["total_changeover"] is None else result["total_changeover"]
        best_algorithm_display = result.get("best_algorithm") or "-"

        st.markdown(
            f"""
            <div class="metric-grid">
                <div class="metric-card">
                    <div class="metric-label">Loop Type</div>
                    <div class="metric-value">{result["problem_type"]}</div>
                    <div class="metric-sub">Selected sequencing mode</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Selected Products</div>
                    <div class="metric-value">{result["selected_count"]}</div>
                    <div class="metric-sub">Items included in optimization run</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Total Changeover</div>
                    <div class="metric-value">{total_changeover_display}</div>
                    <div class="metric-sub">Best feasible result</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Runtime</div>
                    <div class="metric-value">{result["runtime_seconds"]} s</div>
                    <div class="metric-sub">Best method runtime</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            f"""
            <div class="section-header-row">
                <div class="section-label" style="margin:0;">Best Sequence</div>
                <div class="algo-pill">Best Algorithm: {best_algorithm_display}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown(
            f'<div class="sequence-box">{result["sequence_text"]}</div>',
            unsafe_allow_html=True,
        )

        st.markdown('<div class="section-label">Output table</div>', unsafe_allow_html=True)
        output_df = result["output_df"]
        st.dataframe(output_df, use_container_width=True, hide_index=True)

        st.download_button(
            label="Download output workbook (.xlsx)",
            data=result["output_excel_bytes"],
            file_name=result["output_excel_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.markdown(
            f"""
            <div class="note-box">
                <b>Note</b><br>
                {result["solver_note"]}<br><br>
                Workbook saved as: <b>{result["output_excel_name"]}</b>
            </div>
            """,
            unsafe_allow_html=True,
        )
