import os
import re
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# -- Default path config (edit these if the file moves) -----------------------
DEFAULT_FOLDER = os.path.dirname(os.path.abspath(__file__))
DEFAULT_FILE = "JIRA_Report_2026.xlsx"

# -- Page config --------------------------------------------------------------
st.set_page_config(
    page_title="Jira Issues Dashboard",
    page_icon="Adobe",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -- Theme -------------------------------------------------------------------
ADOBE_RED = "#FF0000"
ADOBE_RED_SOFT = "#FF5A5A"
ADOBE_CORAL = "#FF7A59"
ADOBE_PURPLE = "#9B5CFF"
ADOBE_BLUE = "#1473E6"
ADOBE_TEAL = "#00A3A3"
ADOBE_GOLD = "#E7B22A"
INK = "#1F1F1F"
TEXT = "#222222"
MUTED = "#6B7280"
SURFACE = "#FFFFFF"
SURFACE_2 = "#FFF7F7"
BORDER = "#E5E7EB"
GRID = "#EEF0F3"
SHADOW = "0 14px 40px rgba(17, 24, 39, 0.08)"

st.markdown(
    f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;600&display=swap');

  :root {{
    --bg: #f6f3f1;
    --surface: {SURFACE};
    --surface-2: {SURFACE_2};
    --text: {TEXT};
    --muted: {MUTED};
    --border: {BORDER};
    --accent: {ADOBE_RED};
    --accent-soft: {ADOBE_RED_SOFT};
    --shadow: {SHADOW};
  }}

  html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
  }}

  .stApp {{
    background:
      radial-gradient(circle at top left, rgba(255, 0, 0, 0.07), transparent 22%),
      radial-gradient(circle at top right, rgba(255, 122, 89, 0.10), transparent 20%),
      linear-gradient(180deg, #fffaf8 0%, var(--bg) 100%);
    color: var(--text);
  }}

  [data-testid="stSidebar"] {{
    background: linear-gradient(180deg, #ffffff 0%, #fff6f6 100%);
    border-right: 1px solid var(--border);
    box-shadow: inset -1px 0 0 rgba(255,255,255,0.6);
  }}

  [data-testid="stSidebar"] h1,
  [data-testid="stSidebar"] h2,
  [data-testid="stSidebar"] h3,
  [data-testid="stSidebar"] p,
  [data-testid="stSidebar"] label,
  [data-testid="stSidebar"] span {{
    color: var(--text) !important;
  }}

  .hero {{
    background: linear-gradient(135deg, rgba(255, 0, 0, 0.10), rgba(255, 122, 89, 0.08) 45%, rgba(155, 92, 255, 0.08));
    border: 1px solid rgba(255, 0, 0, 0.10);
    border-radius: 24px;
    padding: 20px 22px;
    box-shadow: var(--shadow);
    margin: 4px 0 18px 0;
  }}

  .hero-title {{
    font-size: 2rem;
    line-height: 1.05;
    font-weight: 800;
    letter-spacing: -0.03em;
    margin: 0;
    color: var(--text);
  }}

  .hero-subtitle {{
    margin-top: 8px;
    font-size: 0.98rem;
    color: var(--muted);
  }}

  .hero-chip {{
    display: inline-block;
    margin-top: 10px;
    padding: 6px 10px;
    border-radius: 999px;
    background: rgba(255, 0, 0, 0.08);
    color: var(--accent);
    border: 1px solid rgba(255, 0, 0, 0.12);
    font-size: 0.78rem;
    font-weight: 700;
    letter-spacing: 0.04em;
    text-transform: uppercase;
  }}

  .section-header {{
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 0.78rem;
    font-weight: 800;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: var(--accent);
    margin: 26px 0 14px 0;
  }}
  .section-header::before {{
    content: '';
    width: 28px;
    height: 3px;
    border-radius: 999px;
    background: linear-gradient(90deg, var(--accent), var(--accent-soft));
  }}

  [data-testid="metric-container"] {{
    background: rgba(255,255,255,0.92);
    border: 1px solid var(--border);
    border-radius: 20px;
    padding: 16px 16px 14px 16px;
    box-shadow: var(--shadow);
  }}
  [data-testid="metric-container"] label {{
    color: var(--muted) !important;
    font-size: 0.72rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    font-weight: 700;
  }}
  [data-testid="metric-container"] [data-testid="stMetricValue"] {{
    color: var(--text);
    font-family: 'IBM Plex Mono', monospace;
    font-size: 2rem;
    line-height: 1.1;
  }}
  [data-testid="metric-container"] div[data-testid="stMetricDelta"] {{
    color: var(--accent);
    font-weight: 700;
  }}

  [data-testid="stFileUploader"] {{
    background: rgba(255,255,255,0.85);
    border: 1.5px dashed rgba(255, 0, 0, 0.22);
    border-radius: 18px;
    padding: 8px;
    box-shadow: var(--shadow);
  }}

  [data-testid="stDataFrame"] {{
    border: 1px solid var(--border);
    border-radius: 18px;
    overflow: hidden;
    box-shadow: var(--shadow);
  }}

  [data-testid="stTabs"] button {{
    font-family: 'Inter', sans-serif;
    font-size: 0.85rem;
    font-weight: 700;
    letter-spacing: 0.02em;
    color: var(--muted);
    background: transparent;
  }}
  [data-testid="stTabs"] button[aria-selected="true"] {{
    color: var(--accent);
    border-bottom-color: var(--accent);
  }}

  .stAlert {{ border-radius: 16px; }}

  div[data-testid="stMarkdownContainer"] p {{ color: var(--muted); }}

  section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p,
  section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] span {{
    color: var(--text) !important;
  }}

  input, textarea, select {{
    border-radius: 14px !important;
  }}
</style>
""",
    unsafe_allow_html=True,
)

# -- Priority colours ---------------------------------------------------------
PRIORITY_COLORS = {
    "P1 - Relationship is at risk, interferes with core business function or loss of mission critical data": ADOBE_RED,
    "P2 - Relationship will be affected negatively or non core activities affected": ADOBE_GOLD,
    "P3 - Relationship could be affected negatively, tasks are more difficult but not impossible to complete": ADOBE_BLUE,
    "P4 - Interferes with recreational OR non business related use OR relationship unchanged": "#29A66C",
}
PRIORITY_SHORT = {
    "P1 - Relationship is at risk, interferes with core business function or loss of mission critical data": "P1",
    "P2 - Relationship will be affected negatively or non core activities affected": "P2",
    "P3 - Relationship could be affected negatively, tasks are more difficult but not impossible to complete": "P3",
    "P4 - Interferes with recreational OR non business related use OR relationship unchanged": "P4",
}

COLOR_MAP = {
    "P1": ADOBE_RED,
    "P2": ADOBE_GOLD,
    "P3": ADOBE_BLUE,
    "P4": "#29A66C",
}

VALIDITY_COLORS = {
    "Valid": ADOBE_TEAL,
    "Invalid": ADOBE_RED,
}

PLOTLY_LAYOUT = dict(
    template="simple_white",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter", color=TEXT, size=12),
    margin=dict(l=10, r=10, t=30, b=10),
    xaxis=dict(gridcolor=GRID, zerolinecolor=GRID, linecolor=BORDER),
    yaxis=dict(gridcolor=GRID, zerolinecolor=GRID, linecolor=BORDER),
    legend=dict(
        bgcolor="rgba(255,255,255,0.7)",
        bordercolor=BORDER,
        borderwidth=1,
        orientation="h",
    ),
)

# -- Helpers -----------------------------------------------------------------
def rewind_source(source):
    if hasattr(source, "seek"):
        try:
            source.seek(0)
        except Exception:
            pass

def load_single_sheet(source, sheet_name: str | None = None) -> pd.DataFrame:
    rewind_source(source)
    name = source if isinstance(source, str) else source.name
    if str(name).lower().endswith((".xlsx", ".xls")):
        if sheet_name is not None:
            rewind_source(source)
            df = pd.read_excel(source, sheet_name=sheet_name, engine="openpyxl")
        else:
            rewind_source(source)
            df = pd.read_excel(source, engine="openpyxl")
    else:
        rewind_source(source)
        df = pd.read_csv(source)
    df.columns = df.columns.str.strip()
    return df


def sheet_is_empty(source, sheet_name: str) -> bool:
    try:
        rewind_source(source)
        df = pd.read_excel(source, sheet_name=sheet_name, engine="openpyxl", nrows=5)
        return df.empty or len(df.columns) == 0
    except Exception:
        return True


def quarter_number(sheet_name: str):
    match = re.search(r"Q([1-4])", str(sheet_name), flags=re.IGNORECASE)
    return int(match.group(1)) if match else None


def year_from_sheet_name(sheet_name: str):
    match = re.search(r"(20\d{2})", str(sheet_name))
    return int(match.group(1)) if match else None


def quarter_label(sheet_name: str):
    q = quarter_number(sheet_name)
    year = year_from_sheet_name(sheet_name)
    if q:
        if year:
            return f"Q{q} {str(year)[-2:]}"
        return f"Q{q}"
    return str(sheet_name)


def build_period_map(source):
    rewind_source(source)
    name = source if isinstance(source, str) else source.name
    if not str(name).lower().endswith((".xlsx", ".xls")):
        return {"All data": {"type": "csv", "sheets": None, "label": os.path.basename(str(name))}}

    xls = pd.ExcelFile(source, engine="openpyxl")
    sheets = [s for s in xls.sheet_names if not sheet_is_empty(source, s)]
    quarter_sheets = [s for s in sheets if quarter_number(s)]
    quarter_sheets = sorted(quarter_sheets, key=lambda s: (year_from_sheet_name(s) or 0, quarter_number(s) or 0, s))

    period_map = {
        "All sheets": {"type": "excel", "sheets": sheets, "label": "All sheets"},
    }

    by_year = {}
    for s in quarter_sheets:
        by_year.setdefault(year_from_sheet_name(s), []).append(s)

    for year, year_sheets in sorted(by_year.items(), key=lambda kv: kv[0] or 0):
        if not year_sheets:
            continue
        year_suffix = str(year)[-2:] if year else ""
        label_prefix = f"FY{year_suffix}" if year_suffix else "FY"
        period_map[label_prefix] = {"type": "excel", "sheets": year_sheets, "label": label_prefix}

        q_nums = {quarter_number(s) for s in year_sheets}
        if {1, 2}.issubset(q_nums):
            period_map[f"H1 {label_prefix}"] = {
                "type": "excel",
                "sheets": [s for s in year_sheets if quarter_number(s) in {1, 2}],
                "label": f"H1 {label_prefix}",
            }
        if {3, 4}.issubset(q_nums):
            period_map[f"H2 {label_prefix}"] = {
                "type": "excel",
                "sheets": [s for s in year_sheets if quarter_number(s) in {3, 4}],
                "label": f"H2 {label_prefix}",
            }

    for s in quarter_sheets:
        period_map[quarter_label(s)] = {"type": "excel", "sheets": [s], "label": quarter_label(s)}

    if not period_map:
        period_map = {"All sheets": {"type": "excel", "sheets": sheets, "label": "All sheets"}}

    return period_map


def load_data(source, selected_sheets=None) -> pd.DataFrame:
    rewind_source(source)
    name = source if isinstance(source, str) else source.name
    if str(name).lower().endswith((".xlsx", ".xls")):
        if selected_sheets is None:
            rewind_source(source)
            df = pd.read_excel(source, engine="openpyxl")
            df["Source Sheet"] = "(first sheet)"
        else:
            if isinstance(selected_sheets, str):
                selected_sheets = [selected_sheets]
            frames = []
            for sheet_name in selected_sheets:
                rewind_source(source)
                frame = pd.read_excel(source, sheet_name=sheet_name, engine="openpyxl")
                frame.columns = frame.columns.str.strip()
                frame["Source Sheet"] = sheet_name
                frames.append(frame)
            df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    else:
        rewind_source(source)
        df = pd.read_csv(source)
        df.columns = df.columns.str.strip()
        df["Source Sheet"] = os.path.basename(str(name))

    df.columns = df.columns.str.strip()

    if "Support Tickets" in df.columns:
        df["Support Tickets"] = df["Support Tickets"].astype(str).str.strip()

    if "Priority" in df.columns:
        df["Priority_Short"] = df["Priority"].map(PRIORITY_SHORT).fillna(df["Priority"])

    if "Resolution Date" in df.columns:
        df["Resolution Date"] = pd.to_datetime(df["Resolution Date"], errors="coerce", utc=True)
        df["Resolution Month"] = df["Resolution Date"].dt.to_period("M").astype(str)

    return df


def find_column(df: pd.DataFrame, candidates):
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for candidate in candidates:
        key = str(candidate).strip().lower()
        if key in lookup:
            return lookup[key]
    return None


def detect_id_column(df: pd.DataFrame):
    candidates = [
        "NEO id",
        "NEO ID",
        "Neo ID",
        "neo id",
        "NEO_ID",
        "NEO",
        "Issue key",
        "Issue Key",
        "Key",
    ]
    col = find_column(df, candidates)
    if col:
        return col
    for c in df.columns:
        low = str(c).lower()
        if "neo" in low and "id" in low:
            return c
    return None


def apply_text_condition(frame: pd.DataFrame, column: str, operator: str, value: str) -> pd.DataFrame:
    if column not in frame.columns:
        return frame

    series = frame[column].astype(str).str.strip()
    value = str(value).strip()
    if not value:
        return frame

    op = operator.lower().strip()
    if op == "is":
        mask = series.str.lower() == value.lower()
    elif op == "contains":
        mask = series.str.contains(value, case=False, na=False)
    elif op == "is not":
        mask = series.str.lower() != value.lower()
    elif op == "starts with":
        mask = series.str.startswith(value, na=False)
    elif op == "ends with":
        mask = series.str.endswith(value, na=False)
    else:
        mask = series.str.lower() == value.lower()
    return frame[mask]


def chart_theme(fig):
    fig.update_layout(**PLOTLY_LAYOUT)
    fig.update_xaxes(showgrid=True)
    fig.update_yaxes(showgrid=True)
    return fig


def section(title: str):
    st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)


# -- Data loading (sidebar) ---------------------------------------------------
with st.sidebar:
    st.markdown("## Adobe-styled data source")

    default_path = os.path.join(DEFAULT_FOLDER, DEFAULT_FILE)
    auto_loaded = os.path.isfile(default_path)

    if auto_loaded:
        st.success(f"Auto-loaded: `{DEFAULT_FILE}`")
        st.caption(f"From: `{DEFAULT_FOLDER}`")

    st.markdown("---")
    st.markdown("**Override:** upload a different file")
    uploaded = st.file_uploader(
        "Upload CSV or Excel",
        type=["csv", "xlsx", "xls"],
        help="Drop a new weekly export here to override the auto-loaded file.",
    )

    st.markdown("---")
    st.markdown("**Or** pick from folder")
    folder_path = st.text_input("Folder path", value=DEFAULT_FOLDER)

    data_file = None
    if uploaded:
        data_file = uploaded
        st.success(f"Using uploaded: `{uploaded.name}`")
    elif folder_path and os.path.isdir(folder_path):
        all_files = sorted(
            [f for f in os.listdir(folder_path) if f.lower().endswith((".csv", ".xlsx", ".xls"))],
            reverse=True,
        )
        if all_files:
            default_idx = all_files.index(DEFAULT_FILE) if DEFAULT_FILE in all_files else 0
            chosen = st.selectbox("Pick a file", all_files, index=default_idx)
            data_file = os.path.join(folder_path, chosen)
        else:
            if auto_loaded:
                data_file = default_path
            else:
                st.warning("No CSV/Excel files found in that folder.")
    elif auto_loaded:
        data_file = default_path

    st.markdown("---")
    st.caption("Replace `JIRA_Report_2026.xlsx` in the folder and the dashboard will pick it up automatically.")

# -- Load / gate --------------------------------------------------------------
if data_file is None:
    st.markdown('<div class="hero">', unsafe_allow_html=True)
    st.markdown('<div class="hero-title">Jira Issues Dashboard</div>', unsafe_allow_html=True)
    st.markdown(
        f"<div class=\"hero-subtitle\">Place <b>{DEFAULT_FILE}</b> in <code>{DEFAULT_FOLDER}</code> or upload a file from the sidebar.</div>",
        unsafe_allow_html=True,
    )
    st.markdown('<div class="hero-chip">Ready when the data lands</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    st.stop()


period_map = build_period_map(data_file)
period_options = list(period_map.keys())
if len(period_options) > 1:
    period_choice = st.selectbox(
        "Reporting period",
        period_options,
        index=period_options.index("All sheets") if "All sheets" in period_options else 0,
        help="Switch between quarters, half-year, or full-year views.",
    )
else:
    period_choice = period_options[0]

period_info = period_map[period_choice]
df_raw = load_data(data_file, period_info["sheets"])

# -- Sidebar filters ----------------------------------------------------------
with st.sidebar:
    st.markdown("## Refine the view")

    all_projects = sorted(df_raw["Project key"].dropna().unique()) if "Project key" in df_raw.columns else []
    sel_projects = st.multiselect("Project", all_projects, default=all_projects)

    all_priorities = sorted(df_raw["Priority_Short"].dropna().unique()) if "Priority_Short" in df_raw.columns else []
    sel_priorities = st.multiselect("Priority", all_priorities, default=all_priorities)

    all_status = sorted(df_raw["Status"].dropna().unique()) if "Status" in df_raw.columns else []
    sel_status = st.multiselect("Status", all_status, default=all_status)

    all_types = sorted(df_raw["Type of Request"].dropna().unique()) if "Type of Request" in df_raw.columns else []
    sel_types = st.multiselect("Type of Request", all_types, default=all_types)

    all_validity = sorted(df_raw["Validity"].dropna().unique()) if "Validity" in df_raw.columns else []
    sel_validity = st.multiselect("Validity", all_validity, default=all_validity)

# -- Apply filters ------------------------------------------------------------
df = df_raw.copy()
if sel_projects and "Project key" in df.columns:
    df = df[df["Project key"].isin(sel_projects)]
if sel_priorities and "Priority_Short" in df.columns:
    df = df[df["Priority_Short"].isin(sel_priorities)]
if sel_status and "Status" in df.columns:
    df = df[df["Status"].isin(sel_status)]
if sel_types and "Type of Request" in df.columns:
    df = df[df["Type of Request"].isin(sel_types)]
if sel_validity and "Validity" in df.columns:
    df = df[df["Validity"].isin(sel_validity)]

# -- Header -------------------------------------------------------------------
st.markdown('<div class="hero">', unsafe_allow_html=True)
st.markdown('<div class="hero-title">P2E dashboard</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="hero-subtitle">Showing <b>{len(df):,}</b> of <b>{len(df_raw):,}</b> issues · period <b>{period_choice}</b> · last loaded {datetime.now().strftime("%d %b %Y %H:%M")}</div>',
    unsafe_allow_html=True,
)
st.markdown('<div class="hero-chip">Adobe P2E data dashboard</div>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

def build_volume_summary(frame: pd.DataFrame, period_choice: str) -> pd.DataFrame | None:
    if "Resolution Date" not in frame.columns:
        return None

    temp = frame.dropna(subset=["Resolution Date"]).copy()
    if temp.empty:
        return None

    quarter_mode = not str(period_choice).upper().startswith("Q")
    if quarter_mode:
        temp["Volume Period"] = temp["Resolution Date"].dt.to_period("Q")
        summary = temp.groupby("Volume Period").size().reset_index(name="Count")
        summary["Label"] = summary["Volume Period"].astype(str)
        summary["SortKey"] = summary["Volume Period"].dt.start_time
    else:
        temp["Volume Period"] = temp["Resolution Date"].dt.to_period("M")
        summary = temp.groupby("Volume Period").size().reset_index(name="Count")
        summary["Label"] = summary["Volume Period"].astype(str)
        summary["SortKey"] = summary["Volume Period"].dt.to_timestamp()

    return summary.sort_values("SortKey")


# -- Tabs ---------------------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs(["Overview", "Classification", "Engineering", "Raw Data"])

# -- TAB 1: Overview ----------------------------------------------------------
with tab1:
    col1, col2 = st.columns(2)

    with col1:
        section("Root cause resolution breakdown")
        if "Root Cause Resolution" in df.columns:
            rc = df["Root Cause Resolution"].value_counts().reset_index()
            rc.columns = ["Root Cause", "Count"]
            fig = px.bar(
                rc,
                x="Count",
                y="Root Cause",
                orientation="h",
                color="Count",
                color_continuous_scale=["#ECFDF5", ADOBE_TEAL],
                text="Count",
            )
            fig.update_traces(textposition="outside", marker_line_width=0)
            fig.update_layout(coloraxis_showscale=False, yaxis=dict(autorange="reversed"), yaxis_title="")
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_01")

    with col2:
        section("Component distribution")
        if "Component/s" in df.columns:
            comp_series = df["Component/s"].dropna().str.split(",").explode().str.strip()
            comp_counts = comp_series.value_counts().head(15).reset_index()
            comp_counts.columns = ["Component", "Count"]
            fig = px.bar(
                comp_counts,
                x="Count",
                y="Component",
                orientation="h",
                color="Count",
                color_continuous_scale=["#FFF3E8", ADOBE_GOLD],
                text="Count",
            )
            fig.update_traces(textposition="outside", marker_line_width=0)
            fig.update_layout(coloraxis_showscale=False, yaxis=dict(autorange="reversed"), yaxis_title="")
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_02")

    col3, col4 = st.columns(2)

    with col3:
        section("Version distribution")
        if "Version" in df.columns:
            vc = df["Version"].value_counts().reset_index()
            vc.columns = ["Version", "Count"]
            fig = px.bar(
                vc.head(15),
                x="Version",
                y="Count",
                text="Count",
                color="Count",
                color_continuous_scale=["#F6F7FB", ADOBE_RED],
            )
            fig.update_traces(textposition="outside", marker_line_width=0, textfont_size=10)
            fig.update_layout(coloraxis_showscale=False, xaxis_title="")
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_03")

    with col4:
        section("Resolution type")
        if "Resolution" in df.columns:
            res = df["Resolution"].value_counts().reset_index()
            res.columns = ["Resolution", "Count"]
            fig = px.pie(
                res,
                names="Resolution",
                values="Count",
                hole=0.5,
                color_discrete_sequence=[ADOBE_RED, ADOBE_CORAL, ADOBE_BLUE, ADOBE_PURPLE, ADOBE_TEAL, ADOBE_GOLD],
            )
            fig.update_traces(textinfo="percent+label", marker=dict(line=dict(color="white", width=2)))
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_04")

    col5, col6 = st.columns(2)

    with col5:
        section("Type of request")
        if "Type of Request" in df.columns:
            tc = df["Type of Request"].value_counts().reset_index()
            tc.columns = ["Type", "Count"]
            fig = px.bar(
                tc,
                x="Count",
                y="Type",
                orientation="h",
                color="Count",
                color_continuous_scale=["#F7E9E9", ADOBE_RED],
                text="Count",
            )
            fig.update_traces(marker_line_width=0)
            fig.update_layout(coloraxis_showscale=False, yaxis_title="")
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_05")

    with col6:
        section("Overall volume")
        volume_summary = build_volume_summary(df, period_choice)
        if volume_summary is not None and not volume_summary.empty:
            x_label = "Month" if str(period_choice).upper().startswith("Q") else "Quarter"
            fig = px.bar(
                volume_summary,
                x="Label",
                y="Count",
                text="Count",
                color="Count",
                color_continuous_scale=["#FFF5F5", ADOBE_RED],
            )
            fig.update_traces(textposition="outside", marker_line_width=0)
            fig.update_layout(coloraxis_showscale=False, xaxis_title=x_label, yaxis_title="Issues")
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_06")

    col7, col8 = st.columns(2)

    with col7:
        section("Validity distribution")
        if "Validity" in df.columns:
            validity = (
                df["Validity"].fillna("(blank)").astype(str).str.strip().value_counts().reset_index()
            )
            validity.columns = ["Validity", "Count"]
            fig = px.bar(
                validity,
                x="Validity",
                y="Count",
                text="Count",
                color="Validity",
                color_discrete_map=VALIDITY_COLORS,
            )
            fig.update_traces(textposition="outside", marker_line_width=0)
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_07")

    with col8:
        section("Validity pie chart")
        if "Validity" in df.columns:
            validity = (
                df["Validity"].fillna("(blank)").astype(str).str.strip().value_counts().reset_index()
            )
            validity.columns = ["Validity", "Count"]
            fig = px.pie(
                validity,
                names="Validity",
                values="Count",
                hole=0.45,
                color_discrete_map=VALIDITY_COLORS,
            )
            fig.update_traces(textinfo="percent+label", marker=dict(line=dict(color="white", width=2)))
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_08")

# -- TAB 2: Classification ----------------------------------------------------
with tab2:
    col1, col2 = st.columns(2)

    with col1:
        section("Validity distribution")
        if "Validity" in df.columns:
            validity = df["Validity"].fillna("(blank)").astype(str).str.strip().value_counts().reset_index()
            validity.columns = ["Validity", "Count"]
            fig = px.bar(
                validity,
                x="Validity",
                y="Count",
                text="Count",
                color="Validity",
                color_discrete_map=VALIDITY_COLORS,
            )
            fig.update_traces(textposition="outside", marker_line_width=0)
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_09")

    with col2:
        section("Status breakdown")
        if "Status" in df.columns:
            sc = df["Status"].value_counts().reset_index()
            sc.columns = ["Status", "Count"]
            fig = px.pie(
                sc,
                names="Status",
                values="Count",
                hole=0.6,
                color_discrete_sequence=[ADOBE_RED, ADOBE_CORAL, ADOBE_BLUE, ADOBE_PURPLE, ADOBE_TEAL],
            )
            fig.update_traces(textinfo="percent+label", textfont_size=11, marker=dict(line=dict(color="white", width=2)))
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_10")

    col3, col4 = st.columns(2)

    with col3:
        section("RFH/Bug distribution")
        rfh_col = find_column(df, ["Correct / Incorrect classification? RFH/Bug"])
        if rfh_col:
            rfh = df[rfh_col].fillna("(blank)").astype(str).str.strip().value_counts().reset_index()
            rfh.columns = ["RFH/Bug", "Count"]
            fig = px.pie(
                rfh,
                names="RFH/Bug",
                values="Count",
                hole=0.45,
                color_discrete_sequence=[ADOBE_RED, ADOBE_BLUE, ADOBE_CORAL, ADOBE_PURPLE],
            )
            fig.update_traces(textinfo="percent+label", marker=dict(line=dict(color="white", width=2)))
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_11")

    with col4:
        section("Scope for deflection")
        if "Scope for Deflection" in df.columns:
            defl = df["Scope for Deflection"].value_counts().reset_index()
            defl.columns = ["Scope", "Count"]
            fig = px.bar(
                defl,
                x="Scope",
                y="Count",
                text="Count",
                color="Scope",
                color_discrete_sequence=[ADOBE_RED, ADOBE_CORAL, ADOBE_BLUE, ADOBE_PURPLE],
            )
            fig.update_traces(textposition="outside", marker_line_width=0)
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_12")

# -- TAB 3: Engineering -------------------------------------------------------
with tab3:
    col1, col2 = st.columns(2)

    with col1:
        section("Issues per R&D engineer")
        if "R&D Engineer" in df.columns:
            ec = df["R&D Engineer"].value_counts().reset_index()
            ec.columns = ["Engineer", "Count"]
            fig = px.bar(
                ec.head(15),
                x="Count",
                y="Engineer",
                orientation="h",
                color="Count",
                color_continuous_scale=["#F6F7FB", ADOBE_BLUE],
                text="Count",
            )
            fig.update_traces(textposition="outside", marker_line_width=0)
            fig.update_layout(coloraxis_showscale=False, yaxis_title="", yaxis=dict(autorange="reversed"))
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_13")

    with col2:
        section("Priority mix per engineer")
        if "R&D Engineer" in df.columns and "Priority_Short" in df.columns:
            top_eng = df["R&D Engineer"].value_counts().head(10).index
            eng_df = df[df["R&D Engineer"].isin(top_eng)]
            cross = pd.crosstab(eng_df["R&D Engineer"], eng_df["Priority_Short"])
            fig = px.bar(cross, barmode="stack", color_discrete_map=COLOR_MAP)
            fig.update_layout(xaxis_title="", legend_title="Priority")
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_14")

    section("Export root cause rows")
    if "Root Cause Resolution" in df.columns:
        id_col = detect_id_column(df)
        root_values = sorted(df["Root Cause Resolution"].dropna().astype(str).unique().tolist())
        chosen_root = st.selectbox(
            "Choose a root cause value to export",
            ["(all)"] + root_values,
            key="root_cause_export_choice",
        )
        root_export = df.copy()
        if chosen_root != "(all)":
            root_export = root_export[root_export["Root Cause Resolution"].astype(str) == chosen_root]

        st.caption(f"{len(root_export):,} matching rows")
        export_cols = []
        if id_col:
            export_cols.append(id_col)
        for col in ["Root Cause Resolution", "Summary", "Status", "Reporter", "Affected Version", "Affect Version", "Question Answered"]:
            if col in root_export.columns and col not in export_cols:
                export_cols.append(col)
        if not export_cols:
            export_cols = list(root_export.columns)

        st.download_button(
            "Download matching CSV",
            root_export[export_cols].drop_duplicates().to_csv(index=False).encode("utf-8"),
            file_name="root_cause_rows.csv",
            mime="text/csv",
        )

    if "Creator" in df.columns:
        section("Issues created by")
        cc = df["Creator"].value_counts().reset_index()
        cc.columns = ["Creator", "Count"]
        fig = px.bar(
            cc.head(12),
            x="Count",
            y="Creator",
            orientation="h",
            color="Count",
            color_continuous_scale=["#F5F3FF", ADOBE_PURPLE],
            text="Count",
        )
        fig.update_traces(textposition="outside", marker_line_width=0)
        fig.update_layout(coloraxis_showscale=False, yaxis=dict(autorange="reversed"), yaxis_title="")
        chart_theme(fig)
        st.plotly_chart(fig, use_container_width=True, key="chart_15")

# -- TAB 4: Raw Data ----------------------------------------------------------
with tab4:
    section("Filtered issues")

    search = st.text_input("Search summaries", placeholder="Type to filter...")
    display_df = df.copy()
    if search and "Summary" in display_df.columns:
        display_df = display_df[
            display_df["Summary"].astype(str).str.contains(search, case=False, na=False)
        ]

    st.caption(f"Detected ID column: {detect_id_column(df) or 'none'}")

    section("Custom export builder")
    st.caption(
        "Pick the columns you want to filter, choose values from dropdowns, then export the matching rows as CSV. "
        "This supports combos like question answered = true, reporter = nikhila, and affected version = 8.9.2."
    )

    max_filters = min(8, max(1, len(df.columns)))
    filter_count = st.number_input(
        "How many filters do you want to add?",
        min_value=1,
        max_value=max_filters,
        value=min(3, max_filters),
        step=1,
        key="custom_filter_count",
    )

    def normalise_value(value):
        if pd.isna(value):
            return "(blank)"
        text_value = str(value).strip()
        return text_value if text_value else "(blank)"

    def build_options(frame: pd.DataFrame, column: str):
        series = frame[column].map(normalise_value)
        unique_values = sorted(series.dropna().astype(str).unique().tolist())
        if "(blank)" in unique_values:
            unique_values.remove("(blank)")
            unique_values = ["(blank)"] + unique_values
        return unique_values

    filter_specs = []
    filter_cols = st.columns(3)
    for i in range(int(filter_count)):
        with filter_cols[i % 3]:
            col_name = st.selectbox(
                f"Filter {i + 1} column",
                ["(none)"] + list(df.columns),
                key=f"custom_filter_col_{i}",
            )

            if col_name == "(none)":
                filter_specs.append((col_name, None, None))
                st.selectbox(
                    f"Filter {i + 1} value",
                    ["(choose a column first)"],
                    key=f"custom_filter_val_{i}",
                    disabled=True,
                )
                continue

            op_name = st.selectbox(
                f"Filter {i + 1} operator",
                ["is", "is not", "contains", "does not contain"],
                key=f"custom_filter_op_{i}",
            )

            options = build_options(df, col_name)
            if len(options) > 500:
                search_hint = st.text_input(
                    f"Filter {i + 1} value",
                    key=f"custom_filter_search_{i}",
                    placeholder="Type to narrow options...",
                )
                filtered_options = [opt for opt in options if search_hint.lower() in opt.lower()] if search_hint else options[:500]
                value_name = st.selectbox(
                    f"Filter {i + 1} value",
                    filtered_options if filtered_options else options[:500],
                    key=f"custom_filter_val_{i}",
                )
            else:
                value_name = st.selectbox(
                    f"Filter {i + 1} value",
                    options if options else ["(blank)"],
                    key=f"custom_filter_val_{i}",
                )

            filter_specs.append((col_name, op_name, value_name))

    def apply_dropdown_condition(frame: pd.DataFrame, column: str, operator: str, value: str) -> pd.DataFrame:
        if column not in frame.columns:
            return frame
        series = frame[column].map(normalise_value).astype(str)
        value = normalise_value(value)
        op = (operator or "is").lower().strip()
        if value == "(choose a column first)":
            return frame
        if op == "is":
            mask = series.str.lower() == value.lower()
        elif op == "is not":
            mask = series.str.lower() != value.lower()
        elif op == "contains":
            mask = series.str.contains(value, case=False, na=False)
        elif op == "does not contain":
            mask = ~series.str.contains(value, case=False, na=False)
        else:
            mask = series.str.lower() == value.lower()
        return frame[mask]

    export_df = display_df.copy()
    for col_name, op_name, value_name in filter_specs:
        if col_name != "(none)" and value_name not in (None, "", "(choose a column first)"):
            export_df = apply_dropdown_condition(export_df, col_name, op_name, value_name)

    id_col = detect_id_column(export_df)
    st.metric("Matching rows", len(export_df))

    export_cols = []
    if id_col:
        export_cols.append(id_col)
    for col_name, _, _ in filter_specs:
        if col_name != "(none)" and col_name not in export_cols:
            export_cols.append(col_name)
    for fallback_col in ["Summary", "Status", "Reporter", "Affected Version", "Affect Version", "Question Answered"]:
        if fallback_col in export_df.columns and fallback_col not in export_cols:
            export_cols.append(fallback_col)

    if not export_cols:
        export_cols = list(export_df.columns)

    st.dataframe(export_df[export_cols], use_container_width=True, height=420)

    st.download_button(
        "Download custom CSV",
        export_df[export_cols].drop_duplicates().to_csv(index=False).encode("utf-8"),
        "custom_filtered_issues.csv",
        "text/csv",
    )

    section("Filtered issues table")

    show_cols = [c for c in display_df.columns if c not in ["Priority_Short", "Resolution Month"]]
    st.dataframe(display_df[show_cols], use_container_width=True, height=500)

    csv_bytes = display_df[show_cols].to_csv(index=False).encode("utf-8")
    st.download_button("Download filtered CSV", csv_bytes, "filtered_issues.csv", "text/csv")
