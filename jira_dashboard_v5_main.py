import os
import re
import sys
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# -- Default path config (edit these if the file moves) -----------------------
DEFAULT_FOLDER = os.path.dirname(os.path.abspath(__file__))
DEFAULT_FILE = "JIRA_Report_Resolution1_2026.xlsx"

# -- Page config --------------------------------------------------------------
st.set_page_config(
    page_title="Jira Issues Dashboard",
    page_icon="Adobe",
    layout="wide",
    # Start collapsed — leadership lands on the dashboard charts immediately
    # without a sidebar of filter controls in their face. The sidebar's
    # expand button (top-left) gets a subtle bounce animation to invite
    # discovery without nagging.
    initial_sidebar_state="collapsed",
)

# -- Theme: Aurora -----------------------------------------------------------
# Warm parchment canvas + muted jewel-tone accents. Each chart gets its own
# dedicated accent so the eye can tag them by color. All numbers stay bold ink.
INK = "#1A1A1A"          # primary text, numbers
TEXT = "#1A1A1A"         # alias for compatibility
SUBTEXT = "#5C5A56"      # warm secondary text
MUTED = "#8A877F"        # captions, hints
SURFACE = "#FFFFFF"
SURFACE_2 = "#FBF8F4"    # warm parchment — the page background
BG = "#FBF8F4"
BORDER = "#E8E2D7"       # warm hairline that picks up the parchment
GRID = "#F0EBE0"         # chart gridlines
RULE = "#1A1A1A"
SHADOW = "0 1px 2px rgba(26, 26, 26, 0.04), 0 12px 32px rgba(26, 26, 26, 0.06)"

# Each chart's signature color. Muted jewel tones, ~60% saturation, designed
# to live together harmoniously. The 'deep' variant is ~20% darker for
# highlighting the largest bar in each chart.
CHART_INDIGO     = "#7C6FE8"; CHART_INDIGO_DEEP    = "#5849C9"  # validity
CHART_TERRACOTTA = "#E8896F"; CHART_TERRACOTTA_DEEP= "#C66143"  # type of request
CHART_EMERALD    = "#3F8B7E"; CHART_EMERALD_DEEP   = "#2A6F62"  # version
CHART_ROSE       = "#B5837A"; CHART_ROSE_DEEP      = "#8E5C53"  # root cause
CHART_HONEY      = "#D4A574"; CHART_HONEY_DEEP     = "#B07F4A"  # component
CHART_STEEL      = "#5B7CB8"; CHART_STEEL_DEEP     = "#3D5C95"  # volume / engineers
CHART_PLUM       = "#9B6FA0"; CHART_PLUM_DEEP      = "#754878"  # secondary use

ACCENT = CHART_INDIGO        # primary brand accent for buttons/links
ACCENT_SOFT = "#A89DEF"

# Categorical palette for pie charts that need many slices
CAT_PALETTE = [
    CHART_INDIGO, CHART_TERRACOTTA, CHART_EMERALD, CHART_HONEY,
    CHART_STEEL, CHART_ROSE, CHART_PLUM, "#8A877F",
]

st.markdown(
    f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&family=IBM+Plex+Mono:wght@400;500;600;700&family=Fraunces:opsz,wght@9..144,500;9..144,600&display=swap');

  :root {{
    --bg: {BG};
    --surface: {SURFACE};
    --surface-2: {SURFACE_2};
    --ink: {INK};
    --text: {TEXT};
    --subtext: {SUBTEXT};
    --muted: {MUTED};
    --border: {BORDER};
    --accent: {ACCENT};
    --accent-soft: {ACCENT_SOFT};
    --rule: {RULE};
    --shadow: {SHADOW};
  }}

  html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
    color: var(--ink);
  }}

  .stApp {{
    background:
      radial-gradient(ellipse 800px 400px at 15% -5%, rgba(124, 111, 232, 0.08), transparent 60%),
      radial-gradient(ellipse 700px 350px at 95% 5%, rgba(232, 137, 111, 0.06), transparent 60%),
      var(--bg);
    background-attachment: fixed;
    color: var(--ink);
  }}

  [data-testid="stSidebar"] {{
    background: var(--surface);
    border-right: 1px solid var(--border);
  }}

  /* When the sidebar is collapsed, the expand control sits in the top-left.
     Pulse it gently for the first few seconds to invite discovery without
     nagging. Animation runs 4 times then stops — by then the user has either
     noticed it or is happily reading the dashboard.

     Targets both element ids Streamlit uses across versions:
     - stSidebarCollapsedControl (older builds)
     - stSidebarCollapseButton    (newer builds)
     - And the icon path inside, in case the wrapper is non-targetable. */
  @keyframes sidebar-attention-bounce {{
    0%, 100% {{
      transform: translateX(0);
      box-shadow: 0 1px 2px rgba(124, 111, 232, 0.12);
    }}
    25% {{
      transform: translateX(4px);
      box-shadow: 0 2px 12px rgba(124, 111, 232, 0.45);
    }}
    50% {{
      transform: translateX(0);
      box-shadow: 0 1px 2px rgba(124, 111, 232, 0.12);
    }}
    75% {{
      transform: translateX(2px);
      box-shadow: 0 2px 8px rgba(124, 111, 232, 0.30);
    }}
  }}
  @keyframes sidebar-attention-glow {{
    0%, 100% {{
      background: rgba(124, 111, 232, 0.00);
    }}
    50% {{
      background: rgba(124, 111, 232, 0.18);
    }}
  }}

  /* Pulse the expand button after page load. animation-iteration-count: 4
     means the bounce plays four times (~5 seconds) and then settles. */
  [data-testid="stSidebarCollapsedControl"],
  [data-testid="stSidebarCollapseButton"],
  [data-testid="collapsedControl"] {{
    animation: sidebar-attention-bounce 1.2s ease-in-out 4 !important;
    animation-delay: 0.6s !important;
    border-radius: 8px !important;
    transition: all 0.2s ease;
  }}
  [data-testid="stSidebarCollapsedControl"]:hover,
  [data-testid="stSidebarCollapseButton"]:hover,
  [data-testid="collapsedControl"]:hover {{
    background: rgba(124, 111, 232, 0.10) !important;
    transform: translateX(2px);
  }}

  [data-testid="stSidebar"] h1,
  [data-testid="stSidebar"] h2,
  [data-testid="stSidebar"] h3,
  [data-testid="stSidebar"] p,
  [data-testid="stSidebar"] label,
  [data-testid="stSidebar"] span {{
    color: var(--ink) !important;
  }}

  /* HERO — premium magazine masthead with serif accent + rainbow hairline */
  .hero {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 32px 36px 28px 36px;
    box-shadow: var(--shadow);
    margin: 4px 0 24px 0;
    position: relative;
    overflow: hidden;
  }}
  .hero::before {{
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 4px;
    background: linear-gradient(90deg, {CHART_INDIGO} 0%, {CHART_TERRACOTTA} 35%, {CHART_HONEY} 65%, {CHART_EMERALD} 100%);
  }}

  .hero-title {{
    font-family: 'Fraunces', 'Inter', serif;
    font-size: 2.6rem;
    line-height: 1.05;
    font-weight: 600;
    letter-spacing: -0.025em;
    margin: 0;
    color: var(--ink);
  }}

  .hero-subtitle {{
    margin-top: 12px;
    font-size: 0.96rem;
    line-height: 1.6;
    color: var(--subtext);
    max-width: 820px;
  }}

  .hero-chip {{
    display: inline-block;
    margin-top: 14px;
    padding: 5px 12px;
    border-radius: 999px;
    background: rgba(124, 111, 232, 0.10);
    color: {CHART_INDIGO_DEEP};
    border: 1px solid rgba(124, 111, 232, 0.20);
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.10em;
    text-transform: uppercase;
  }}

  /* SECTION HEADER — small-caps with colored accent dot */
  .section-header {{
    display: flex;
    align-items: center;
    gap: 12px;
    font-size: 0.74rem;
    font-weight: 700;
    letter-spacing: 0.16em;
    text-transform: uppercase;
    color: var(--ink);
    margin: 28px 0 16px 0;
    font-family: 'Inter', sans-serif;
  }}
  .section-header::before {{
    content: '';
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: var(--accent);
  }}
  .section-header::after {{
    content: '';
    flex: 1;
    height: 1px;
    background: linear-gradient(90deg, var(--border), transparent);
  }}

  /* SLIM STAT STRIP — replaces the old duplicate hero */
  .stat-strip {{
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 12px;
    padding: 12px 18px;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    box-shadow: var(--shadow);
    margin: 0 0 20px 0;
    font-size: 0.86rem;
    color: var(--subtext);
  }}
  .stat-strip-item b {{
    font-family: 'IBM Plex Mono', monospace;
    color: var(--ink);
    font-weight: 700;
  }}
  .stat-strip-sep {{
    color: var(--border);
    font-weight: 700;
  }}

  /* KPI METRIC CARDS — colored stripe on the left, lifts on hover */
  [data-testid="metric-container"] {{
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 18px 20px 16px 22px;
    box-shadow: var(--shadow);
    position: relative;
    overflow: hidden;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
  }}
  [data-testid="metric-container"]:hover {{
    transform: translateY(-1px);
    box-shadow: 0 1px 2px rgba(26,26,26,0.04), 0 16px 40px rgba(26,26,26,0.10);
  }}
  [data-testid="metric-container"]::before {{
    content: '';
    position: absolute;
    left: 0; top: 0; bottom: 0;
    width: 3px;
    background: var(--accent);
  }}
  /* Cycle KPI accent colors across the 5-card row */
  [data-testid="stHorizontalBlock"] > div:nth-child(1) [data-testid="metric-container"]::before {{ background: {CHART_INDIGO}; }}
  [data-testid="stHorizontalBlock"] > div:nth-child(2) [data-testid="metric-container"]::before {{ background: {CHART_TERRACOTTA}; }}
  [data-testid="stHorizontalBlock"] > div:nth-child(3) [data-testid="metric-container"]::before {{ background: {CHART_EMERALD}; }}
  [data-testid="stHorizontalBlock"] > div:nth-child(4) [data-testid="metric-container"]::before {{ background: {CHART_HONEY}; }}
  [data-testid="stHorizontalBlock"] > div:nth-child(5) [data-testid="metric-container"]::before {{ background: {CHART_STEEL}; }}

  [data-testid="metric-container"] label {{
    color: var(--muted) !important;
    font-size: 0.7rem;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    font-weight: 700;
    font-family: 'Inter', sans-serif;
  }}
  [data-testid="metric-container"] [data-testid="stMetricValue"] {{
    color: var(--ink) !important;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 2.2rem;
    font-weight: 700;
    line-height: 1.05;
    letter-spacing: -0.02em;
  }}
  [data-testid="metric-container"] [data-testid="stMetricDelta"] {{
    display: none !important;
  }}

  [data-testid="stFileUploader"] {{
    background: var(--surface);
    border: 1.5px dashed var(--border);
    border-radius: 12px;
    padding: 8px;
  }}

  [data-testid="stDataFrame"] {{
    border: 1px solid var(--border);
    border-radius: 12px;
    overflow: hidden;
    box-shadow: var(--shadow);
  }}

  /* TABS — accent underline */
  [data-testid="stTabs"] {{
    border-bottom: 1px solid var(--border);
  }}
  [data-testid="stTabs"] button {{
    font-family: 'Inter', sans-serif;
    font-size: 0.88rem;
    font-weight: 600;
    letter-spacing: 0.01em;
    color: var(--muted);
    background: transparent;
    padding: 12px 6px;
    transition: color 0.15s ease;
  }}
  [data-testid="stTabs"] button:hover {{
    color: var(--ink);
  }}
  [data-testid="stTabs"] button[aria-selected="true"] {{
    color: var(--ink);
    border-bottom: 2px solid var(--accent) !important;
    font-weight: 700;
  }}

  .stAlert {{ border-radius: 12px; }}

  div[data-testid="stMarkdownContainer"] p {{ color: var(--subtext); }}

  section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p,
  section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] span {{
    color: var(--ink) !important;
  }}

  input, textarea, select {{
    border-radius: 10px !important;
  }}

  /* Action buttons (st.button) — clean outline style, indigo accent.
     White background + indigo border + indigo text. Reads as a clickable
     primary action without competing visually with chart colors. */
  .stButton > button {{
    background: {SURFACE} !important;
    color: {CHART_INDIGO_DEEP} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 10px !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
    letter-spacing: 0.01em !important;
    padding: 8px 16px !important;
    transition: all 0.15s ease;
    box-shadow: 0 1px 2px rgba(26, 26, 26, 0.04);
  }}
  /* Force the text color of any nested element to win against BaseWeb */
  .stButton > button p,
  .stButton > button span,
  .stButton > button div {{
    color: {CHART_INDIGO_DEEP} !important;
    font-weight: 600 !important;
  }}
  .stButton > button:hover {{
    background: rgba(124, 111, 232, 0.06) !important;
    border-color: {CHART_INDIGO} !important;
    color: {CHART_INDIGO_DEEP} !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(124, 111, 232, 0.10);
  }}
  .stButton > button:hover p,
  .stButton > button:hover span,
  .stButton > button:hover div {{
    color: {CHART_INDIGO_DEEP} !important;
  }}
  .stButton > button:active {{
    transform: translateY(0);
    background: rgba(124, 111, 232, 0.10) !important;
  }}

  /* Download buttons — filled solid indigo (these ARE the primary CTA).
     White text forced via !important on every nested element. */
  .stDownloadButton > button {{
    background: {CHART_INDIGO_DEEP} !important;
    color: #FFFFFF !important;
    border: 1px solid {CHART_INDIGO_DEEP} !important;
    border-radius: 10px !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
    letter-spacing: 0.01em !important;
    padding: 8px 16px !important;
    transition: all 0.15s ease;
    box-shadow: 0 2px 6px rgba(124, 111, 232, 0.20);
  }}
  .stDownloadButton > button p,
  .stDownloadButton > button span,
  .stDownloadButton > button div {{
    color: #FFFFFF !important;
    font-weight: 600 !important;
  }}
  .stDownloadButton > button:hover {{
    background: {CHART_INDIGO} !important;
    border-color: {CHART_INDIGO} !important;
    color: #FFFFFF !important;
    transform: translateY(-1px);
    box-shadow: 0 6px 16px rgba(124, 111, 232, 0.30);
  }}
  .stDownloadButton > button:hover p,
  .stDownloadButton > button:hover span,
  .stDownloadButton > button:hover div {{
    color: #FFFFFF !important;
  }}
  .stDownloadButton > button:active {{
    transform: translateY(0);
  }}
</style>
""",
    unsafe_allow_html=True,
)

# -- Priority colours ---------------------------------------------------------
# Aurora jewel-tone palette. P1 (most urgent) = saturated indigo for visual
# weight without screaming red. Stepped down through the spectrum to cooler
# tones for lower priorities.
PRIORITY_COLORS = {
    "P1 - Relationship is at risk, interferes with core business function or loss of mission critical data": CHART_TERRACOTTA_DEEP,
    "P2 - Relationship will be affected negatively or non core activities affected": CHART_HONEY,
    "P3 - Relationship could be affected negatively, tasks are more difficult but not impossible to complete": CHART_STEEL,
    "P4 - Interferes with recreational OR non business related use OR relationship unchanged": CHART_EMERALD,
}
PRIORITY_SHORT = {
    "P1 - Relationship is at risk, interferes with core business function or loss of mission critical data": "P1",
    "P2 - Relationship will be affected negatively or non core activities affected": "P2",
    "P3 - Relationship could be affected negatively, tasks are more difficult but not impossible to complete": "P3",
    "P4 - Interferes with recreational OR non business related use OR relationship unchanged": "P4",
}

COLOR_MAP = {
    "P1": CHART_TERRACOTTA_DEEP,
    "P2": CHART_HONEY,
    "P3": CHART_STEEL,
    "P4": CHART_EMERALD,
}

VALIDITY_COLORS = {
    "Valid": CHART_EMERALD,
    "Invalid": CHART_TERRACOTTA_DEEP,
    "(blank)": "#C8C2B5",
}

# Default Plotly layout — bold black labels everywhere, generous tick fonts.
# Numbers must be readable at a glance from across a conference room.
PLOTLY_LAYOUT = dict(
    template="simple_white",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter", color=INK, size=13, weight=600),
    margin=dict(l=10, r=10, t=20, b=10),
    xaxis=dict(
        gridcolor=GRID,
        zerolinecolor=GRID,
        linecolor=BORDER,
        tickfont=dict(color=INK, size=12, weight=600),
        title_font=dict(color=INK, size=12, weight=700),
    ),
    yaxis=dict(
        gridcolor=GRID,
        zerolinecolor=GRID,
        linecolor=BORDER,
        tickfont=dict(color=INK, size=12, weight=600),
        title_font=dict(color=INK, size=12, weight=700),
    ),
    legend=dict(
        bgcolor="rgba(255,255,255,0.9)",
        bordercolor=BORDER,
        borderwidth=1,
        orientation="h",
        font=dict(color=INK, size=12, weight=600),
    ),
    hoverlabel=dict(
        bgcolor=INK,
        bordercolor=INK,
        font=dict(color="#FFFFFF", size=12, family="Inter"),
    ),
)

# -- Helpers -----------------------------------------------------------------
def rewind_source(source):
    if hasattr(source, "seek"):
        try:
            source.seek(0)
        except Exception:
            pass


def is_quarter_sheet(sheet_name: str) -> bool:
    return bool(re.match(r"^Q[1-4]_20\d{2}$", str(sheet_name).strip(), flags=re.IGNORECASE))


def quarter_sheet_sort_key(sheet_name: str):
    match = re.match(r"^Q([1-4])_(20\d{2})$", str(sheet_name).strip(), flags=re.IGNORECASE)
    if not match:
        return (9999, 99, str(sheet_name))
    quarter = int(match.group(1))
    year = int(match.group(2))
    return (year, quarter, str(sheet_name))


def period_label(period) -> str:
    return f"Q{period.quarter} {period.year}"


def parse_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    for date_col in ["Creation Date", "Resolution Date"]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
            month_col = f"{date_col.split()[0]} Month"
            df[month_col] = df[date_col].dt.to_period("M").astype(str)
    return df


def sheet_to_quarter_label(sheet_name: str) -> str | None:
    """Convert a sheet name like 'Q1_2026' into a display label like 'Q1 2026'."""
    match = re.match(r"^Q([1-4])_(20\d{2})$", str(sheet_name).strip(), flags=re.IGNORECASE)
    if not match:
        return None
    return f"Q{match.group(1)} {match.group(2)}"


def quarter_label_sort_key(label: str):
    match = re.match(r"^Q([1-4])\s+(20\d{2})$", str(label).strip(), flags=re.IGNORECASE)
    if not match:
        return (9999, 99, str(label))
    return (int(match.group(2)), int(match.group(1)), str(label))


def _parse_issue_source(source) -> pd.DataFrame:
    """Inner parser — actually does the Excel/CSV reading and quarter
    bucketing. Both the cached and uncached entry points call this."""
    rewind_source(source)
    name = source if isinstance(source, str) else source.name
    is_excel = str(name).lower().endswith((".xlsx", ".xls"))

    if is_excel:
        xls = pd.ExcelFile(source, engine="openpyxl")
        sheet_names = [s for s in xls.sheet_names if is_quarter_sheet(s)]
        sheet_names = sorted(sheet_names, key=quarter_sheet_sort_key)

        frames = []
        for sheet_name in sheet_names:
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

    if df.empty:
        return df

    df.columns = df.columns.str.strip()

    if "Support Tickets" in df.columns:
        df["Support Tickets"] = df["Support Tickets"].astype(str).str.strip()

    if "Priority" in df.columns:
        priority_clean = df["Priority"].astype(str).str.strip()
        # First try exact match against the long descriptions we know about.
        # Then fall back to a prefix match — if the value starts with "P1",
        # "P2", etc., normalize it to the short code. This survives both
        # exact matches we've seen AND future cases like "P0" or slightly
        # reworded descriptions ("P1 - critical, mission-impact loss").
        def _short_priority(value: str) -> str:
            if value in PRIORITY_SHORT:
                return PRIORITY_SHORT[value]
            # Match anything starting with P followed by a digit
            match = re.match(r"^(P[0-9])\b", value, flags=re.IGNORECASE)
            if match:
                return match.group(1).upper()
            return value
        df["Priority_Short"] = priority_clean.apply(_short_priority)

    # Normalize categorical columns at load time so KPIs (which lowercase/strip
    # internally) and charts (which group by raw text) can never disagree.
    # Without this, a future export with "Bug ", "bug", and "Bug" would split
    # the chart into 3 categories while the KPI would still total them as one.
    def _title_normalize(value):
        if pd.isna(value):
            return value
        s = str(value).strip()
        if not s:
            return None
        return s

    for cat_col in ["Validity", "Type of Request", "Status"]:
        if cat_col in df.columns:
            df[cat_col] = df[cat_col].apply(_title_normalize)
            # Title-case-fix the most common values (case-insensitive map).
            # Leaves anything we don't recognize alone, so unexpected values
            # still appear (you'll see them, just consistent in casing).
            canon_map = {
                "valid": "Valid",
                "invalid": "Invalid",
                "bug": "Bug",
                "request for help": "Request for Help",
                "rfh": "Request for Help",
            }
            df[cat_col] = df[cat_col].apply(
                lambda v: canon_map.get(str(v).lower(), v) if pd.notna(v) else v
            )

    # Derive a display-friendly "Quarter Label" from the source sheet
    # (e.g., 'Q1_2026' -> 'Q1 2026'). This is the single source of truth
    # for which quarter a row belongs to — no date math involved.
    if "Source Sheet" in df.columns:
        df["Quarter Label"] = df["Source Sheet"].map(sheet_to_quarter_label)

    df = parse_date_columns(df)
    return df


@st.cache_data(show_spinner=False, ttl=3600)
def _parse_issue_source_from_path(path: str, mtime: float) -> pd.DataFrame:
    """Cacheable variant for string paths on disk. The cache key includes
    the file's modification time so editing the source Excel auto-invalidates
    the cache — no manual clear needed when new data lands.
    Streamlit Cloud has a ~1GB RAM limit; this prevents re-parsing the
    Excel on every rerun (which was contributing to OOM kills)."""
    return _parse_issue_source(path)


def load_issue_data(source) -> pd.DataFrame:
    """Public entry point. Routes string paths through the cached
    parser; UploadedFile objects (sidebar drag-and-drop) get parsed
    fresh because they're not safely hashable."""
    if isinstance(source, str):
        try:
            mtime = os.path.getmtime(source)
            return _parse_issue_source_from_path(source, mtime)
        except Exception:
            # If anything goes wrong (file gone, race, etc.), fall back
            # to the direct parser so the app keeps working
            pass
    return _parse_issue_source(source)


def build_reporting_period_options(frame: pd.DataFrame):
    """Build the period dropdown options directly from the sheets present in
    the file. Adds H1/H2 groupings per fiscal year automatically when the
    underlying quarters exist.
    """
    if "Quarter Label" not in frame.columns:
        return ["All Data"], {}

    quarter_labels = (
        frame["Quarter Label"]
        .dropna()
        .astype(str)
        .str.strip()
        .replace({"": None, "nan": None, "None": None})
        .dropna()
        .drop_duplicates()
        .tolist()
    )
    quarter_labels = sorted(quarter_labels, key=quarter_label_sort_key)

    # period_map: option label -> list of quarter labels that option includes.
    # "All Data" is handled with empty list (means: don't filter).
    period_map: dict[str, list[str]] = {"All Data": []}
    for q in quarter_labels:
        period_map[q] = [q]

    # Build H1/H2 groups per fiscal year if the constituent quarters exist
    years = sorted({quarter_label_sort_key(q)[0] for q in quarter_labels})
    half_options: list[str] = []
    for year in years:
        fy_short = f"FY{str(year)[2:]}"  # 2026 -> FY26
        h1_quarters = [f"Q{i} {year}" for i in (1, 2) if f"Q{i} {year}" in quarter_labels]
        h2_quarters = [f"Q{i} {year}" for i in (3, 4) if f"Q{i} {year}" in quarter_labels]
        if h1_quarters:
            label = f"H1 {fy_short}"
            period_map[label] = h1_quarters
            half_options.append(label)
        if h2_quarters:
            label = f"H2 {fy_short}"
            period_map[label] = h2_quarters
            half_options.append(label)

    options = ["All Data"] + quarter_labels + half_options
    return options, period_map


def apply_reporting_period(frame: pd.DataFrame, selected_period: str, period_map: dict) -> pd.DataFrame:
    if selected_period == "All Data" or "Quarter Label" not in frame.columns:
        return frame
    quarters = period_map.get(selected_period, [])
    if not quarters:
        return frame
    return frame[frame["Quarter Label"].isin(quarters)]


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
    """Apply the editorial monochrome layout to every figure.
    Forces all labels, ticks, and annotations to bold black ink so numbers
    are readable from across a conference room.
    """
    fig.update_layout(**PLOTLY_LAYOUT)
    fig.update_xaxes(showgrid=True, tickfont=dict(color=INK, size=12))
    fig.update_yaxes(showgrid=True, tickfont=dict(color=INK, size=12))
    # Make any data labels (text on bars/pie slices) bold black
    fig.update_traces(
        textfont=dict(color=INK, size=13, family="Inter"),
        selector=dict(type="bar"),
    )
    return fig


def bar_palette(values, base_color: str = None, top_color: str = None):
    """Return solid colors for a bar chart. The largest bar gets the deeper
    'top_color', everything else gets the regular 'base_color'.
    Solid colors at full saturation — no gradient that fades to invisible.
    """
    if base_color is None:
        base_color = CHART_INDIGO
    if top_color is None:
        top_color = CHART_INDIGO_DEEP
    if values is None or len(values) == 0:
        return [base_color]
    vals = list(values)
    max_idx = vals.index(max(vals))
    return [top_color if i == max_idx else base_color for i in range(len(vals))]


def section(title: str):
    st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)


def make_xlsx_bytes(frame: pd.DataFrame, sheet_name: str = "Issues") -> bytes:
    """Convert a dataframe to .xlsx bytes for download_button.
    Strips timezone info from datetime columns because Excel does not
    support tz-aware datetimes (pandas raises ValueError if any cell
    is tz-aware). Wall-clock time is preserved.
    """
    from io import BytesIO
    safe_sheet = re.sub(r"[:\\/?*\[\]]", "_", str(sheet_name))[:31] or "Issues"

    # Defensive copy + strip timezones from any datetime column.
    safe_frame = frame.copy()
    for col in safe_frame.columns:
        s = safe_frame[col]
        if isinstance(s.dtype, pd.DatetimeTZDtype):
            safe_frame[col] = s.dt.tz_localize(None)
        elif s.dtype == "object":
            try:
                first_val = s.dropna().head(1)
                if len(first_val) and getattr(first_val.iloc[0], "tzinfo", None) is not None:
                    safe_frame[col] = s.apply(
                        lambda v: v.replace(tzinfo=None) if hasattr(v, "tzinfo") and v is not None and getattr(v, "tzinfo", None) is not None else v
                    )
            except Exception:
                pass

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        safe_frame.to_excel(writer, sheet_name=safe_sheet, index=False)
    return buf.getvalue()


XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# -- Data loading (sidebar) ---------------------------------------------------
with st.sidebar:
    st.markdown("## Data source")

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
    st.caption("The dashboard reads only quarter sheets like Q1_2026, Q2_2026, etc. The stats sheet is ignored.")

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


try:
    df_source = load_issue_data(data_file)
except PermissionError as exc:
    st.markdown('<div class="hero">', unsafe_allow_html=True)
    st.markdown('<div class="hero-title">File is locked</div>', unsafe_allow_html=True)
    st.markdown(
        f"<div class=\"hero-subtitle\">The Excel file <code>{DEFAULT_FILE}</code> is currently locked by another process. "
        "On Windows this almost always means one of three things:"
        "<ul>"
        "<li>The file is open in Excel — close every Excel window (and check Task Manager for stray <code>EXCEL.EXE</code>)</li>"
        "<li>OneDrive is syncing or hasn't downloaded the file locally — right-click it in File Explorer and choose <b>Always keep on this device</b></li>"
        "<li>Another app has a handle on it — pause OneDrive sync from the tray for 2 hours and retry</li>"
        "</ul></div>",
        unsafe_allow_html=True,
    )
    st.markdown(f'<div class="hero-chip">PermissionError · {exc.errno}</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)
    if st.button("Retry"):
        st.rerun()
    st.stop()
except FileNotFoundError:
    st.error(f"Could not find the data file. Looked at: {data_file}")
    st.stop()
except Exception as exc:
    st.error(f"Failed to load data: {type(exc).__name__}: {exc}")
    st.stop()

# -- Intro and reporting controls --------------------------------------------
st.markdown('<div class="hero">', unsafe_allow_html=True)
st.markdown('<div class="hero-title">P2E dashboard</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="hero-subtitle">Tracking quality, classification, and resolution patterns for customer-reported issues — quarter by quarter. Use the dropdown below to focus on a specific quarter, half-year, or all data.</div>',
    unsafe_allow_html=True,
)
st.markdown('</div>', unsafe_allow_html=True)

# Resolution Date is the single reporting basis. No toggle.
reporting_date_col = "Resolution Date"

period_options, period_map = build_reporting_period_options(df_source)
if "reporting_period" not in st.session_state or st.session_state.reporting_period not in period_options:
    st.session_state.reporting_period = period_options[0]

# Minimalist period picker — just a clean compact selectbox.
# Constrained to 240px width via scoped CSS, with a small label above and
# a subtle indigo focus ring. No fancy widgets, no leaking panels — just
# a dropdown that does one thing well.
picker_container = st.container(key="period_picker")
with picker_container:
    reporting_period = st.selectbox(
        "Reporting period",
        period_options,
        key="reporting_period",
    )

st.markdown(
    f"""
<style>
  /* Constrain the picker width — never stretches */
  .st-key-period_picker {{
    max-width: 240px;
    margin: 4px 0 22px 0;
  }}

  /* Label above the dropdown — small, uppercase, muted */
  .st-key-period_picker label {{
    font-family: 'Inter', sans-serif !important;
    font-size: 0.68rem !important;
    font-weight: 700 !important;
    letter-spacing: 0.16em !important;
    text-transform: uppercase !important;
    color: {MUTED} !important;
    margin-bottom: 6px !important;
  }}

  /* The dropdown box itself */
  .st-key-period_picker [data-baseweb="select"] > div {{
    background: {SURFACE} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 10px !important;
    box-shadow: 0 1px 2px rgba(26, 26, 26, 0.04);
    transition: all 0.15s ease;
    min-height: 40px !important;
  }}

  /* Hover — subtle indigo tint */
  .st-key-period_picker [data-baseweb="select"] > div:hover {{
    border-color: {CHART_INDIGO} !important;
    box-shadow: 0 2px 8px rgba(124, 111, 232, 0.10);
  }}

  /* Focused — soft indigo glow */
  .st-key-period_picker [data-baseweb="select"] > div:focus-within {{
    border-color: {CHART_INDIGO} !important;
    box-shadow: 0 0 0 3px rgba(124, 111, 232, 0.15);
  }}

  /* The selected text inside the box */
  .st-key-period_picker [data-baseweb="select"] [data-baseweb="tag"],
  .st-key-period_picker [data-baseweb="select"] input,
  .st-key-period_picker [data-baseweb="select"] div[data-testid="stSelectboxVirtualDropdown"] {{
    color: {INK} !important;
    font-family: 'Inter', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.92rem !important;
  }}

  /* The tiny chevron icon on the right */
  .st-key-period_picker [data-baseweb="select"] svg {{
    color: {SUBTEXT} !important;
  }}
</style>
""",
    unsafe_allow_html=True,
)

# Apply period filter before sidebar filters -----------------
df_period = apply_reporting_period(df_source, reporting_period, period_map)

# -- Sidebar filters ----------------------------------------------------------
with st.sidebar:
    st.markdown("## Refine the view")

    all_projects = sorted(df_period["Project key"].dropna().unique()) if "Project key" in df_period.columns else []
    sel_projects = st.multiselect("Project", all_projects, default=all_projects)

    all_priorities = sorted(df_period["Priority_Short"].dropna().unique()) if "Priority_Short" in df_period.columns else []
    sel_priorities = st.multiselect("Priority", all_priorities, default=all_priorities)

    all_status = sorted(df_period["Status"].dropna().unique()) if "Status" in df_period.columns else []
    sel_status = st.multiselect("Status", all_status, default=all_status)

    all_types = sorted(df_period["Type of Request"].dropna().unique()) if "Type of Request" in df_period.columns else []
    sel_types = st.multiselect("Type of Request", all_types, default=all_types)

    all_validity = sorted(df_period["Validity"].dropna().unique()) if "Validity" in df_period.columns else []
    sel_validity = st.multiselect("Validity", all_validity, default=all_validity)

    # ─── Diagnostics (collapsed by default) ────────────────────────────
    # Tucked at the bottom of the sidebar so leadership won't see it
    # unless they explicitly expand it. Useful for spotting memory creep
    # on Streamlit Cloud (~1 GB free-tier limit) before it crashes the app.
    with st.expander("⚙ Diagnostics", expanded=False):
        try:
            import psutil
            proc = psutil.Process()
            mem_mb = proc.memory_info().rss / 1024 / 1024
            cpu_pct = proc.cpu_percent(interval=0.05)

            # Memory readout — color the value if we're approaching the
            # Streamlit Cloud free-tier limit (~1024 MB).
            if mem_mb > 800:
                mem_color = CHART_TERRACOTTA_DEEP
                mem_status = "⚠ approaching limit"
            elif mem_mb > 600:
                mem_color = CHART_HONEY_DEEP
                mem_status = "elevated"
            else:
                mem_color = CHART_EMERALD
                mem_status = "healthy"

            st.markdown(
                f'<div style="font-family:Inter,sans-serif;font-size:0.78rem;color:{MUTED};'
                f'letter-spacing:0.08em;text-transform:uppercase;font-weight:700;margin:4px 0 2px 0;">Memory</div>'
                f'<div style="font-family:\'IBM Plex Mono\',monospace;font-size:1.4rem;'
                f'font-weight:600;color:{mem_color};line-height:1.1;">{mem_mb:.0f} <span style="font-size:0.7rem;color:{MUTED};">MB</span></div>'
                f'<div style="font-family:Inter,sans-serif;font-size:0.72rem;color:{MUTED};margin-bottom:10px;">{mem_status} · ~1024 MB cap on free tier</div>',
                unsafe_allow_html=True,
            )

            # CPU readout
            st.markdown(
                f'<div style="font-family:Inter,sans-serif;font-size:0.78rem;color:{MUTED};'
                f'letter-spacing:0.08em;text-transform:uppercase;font-weight:700;margin:4px 0 2px 0;">CPU</div>'
                f'<div style="font-family:\'IBM Plex Mono\',monospace;font-size:1.4rem;'
                f'font-weight:600;color:{INK};line-height:1.1;">{cpu_pct:.1f}<span style="font-size:0.7rem;color:{MUTED};">%</span></div>'
                f'<div style="font-family:Inter,sans-serif;font-size:0.72rem;color:{MUTED};margin-bottom:10px;">last 50 ms</div>',
                unsafe_allow_html=True,
            )

            # Row count of the current filtered view
            st.markdown(
                f'<div style="font-family:Inter,sans-serif;font-size:0.78rem;color:{MUTED};'
                f'letter-spacing:0.08em;text-transform:uppercase;font-weight:700;margin:4px 0 2px 0;">Rows in view</div>'
                f'<div style="font-family:\'IBM Plex Mono\',monospace;font-size:1.4rem;'
                f'font-weight:600;color:{INK};line-height:1.1;">{len(df_period):,}</div>'
                f'<div style="font-family:Inter,sans-serif;font-size:0.72rem;color:{MUTED};margin-bottom:10px;">after period filter</div>',
                unsafe_allow_html=True,
            )

            # Cache stats — useful for confirming the data load is cached
            st.caption(
                f"Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro} · "
                f"PID {proc.pid}"
            )

            if st.button("Clear data cache", key="clear_cache_btn"):
                st.cache_data.clear()
                st.success("Cache cleared. The app will re-parse the source on the next interaction.")
        except ImportError:
            st.caption(
                "Install `psutil` in requirements.txt to enable live memory and CPU readouts."
            )
        except Exception as e:
            st.caption(f"Diagnostics unavailable: {e}")

# -- Apply sidebar filters ----------------------------------------------------
df = df_period.copy()
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

# -- Live stats strip (slim, after filters) ----------------------------------
st.markdown(
    f'<div class="stat-strip">'
    f'<span class="stat-strip-item"><b>{len(df):,}</b> of <b>{len(df_period):,}</b> issues</span>'
    f'<span class="stat-strip-sep">·</span>'
    f'<span class="stat-strip-item">period <b>{reporting_period}</b></span>'
    f'<span class="stat-strip-sep">·</span>'
    f'<span class="stat-strip-item">updated {datetime.now().strftime("%d %b %Y %H:%M")}</span>'
    f'</div>',
    unsafe_allow_html=True,
)

def build_volume_summary(frame: pd.DataFrame, date_col: str) -> pd.DataFrame | None:
    if date_col not in frame.columns:
        return None

    temp = frame.dropna(subset=[date_col]).copy()
    if temp.empty:
        return None

    temp["Volume Period"] = temp[date_col].dt.to_period("M")
    summary = temp.groupby("Volume Period").size().reset_index(name="Count")
    summary["Label"] = summary["Volume Period"].dt.to_timestamp().dt.strftime("%b %Y")
    summary["SortKey"] = summary["Volume Period"].dt.to_timestamp()

    return summary.sort_values("SortKey")


# -- Executive KPI strip ------------------------------------------------------
# Five at-a-glance metrics that senior leadership reads in the first 5 seconds.
# Numbers are rendered via st.metric which gets the editorial styling from CSS.
def safe_pct(num, den):
    return (num / den * 100) if den else 0

total_issues = len(df)
p1_count = int((df["Priority_Short"] == "P1").sum()) if "Priority_Short" in df.columns else 0
# "Resolved" KPI: match status values that mean "the issue is closed out".
# Uses a whole-word regex (\b) so "Not Resolved" and "Unresolved" do NOT
# get counted. Also explicitly excludes any status starting with "Not " or
# containing "unresolved" as a final safety net.
_status_series = df["Status"].astype(str) if "Status" in df.columns else pd.Series([], dtype=str)
_resolved_mask = _status_series.str.contains(r"\b(?:resolved|closed|done)\b", case=False, regex=True, na=False)
_neg_mask = _status_series.str.contains(r"\b(?:not\s+resolved|unresolved|reopened|in\s+progress|open)\b", case=False, regex=True, na=False)
resolved_count = int((_resolved_mask & ~_neg_mask).sum()) if "Status" in df.columns else 0
valid_count = int((df["Validity"].astype(str).str.strip().str.lower() == "valid").sum()) if "Validity" in df.columns else 0
# Bug count comes from the Type of Request column ("Bug" / "Request for Help"),
# NOT from the RFH/Bug classification correctness column.
bug_count = int(df["Type of Request"].astype(str).str.strip().str.lower().eq("bug").sum()) if "Type of Request" in df.columns else 0

# Render KPIs as st.metric (no delta) + a plain markdown caption underneath.
# Why not use st.metric's delta param? Streamlit auto-injects a green/red
# trend arrow that can't be reliably suppressed across versions. Rendering the
# percentage as our own caption gives us full control and zero icon battle.
def kpi_caption(pct: float) -> str:
    return (
        f'<div style="font-family:Inter,sans-serif;font-size:0.78rem;'
        f'color:{SUBTEXT};font-weight:600;margin:-6px 0 2px 0;letter-spacing:0.01em;">'
        f'{pct:.0f}% of total</div>'
    )

kpi_cols = st.columns(5)
with kpi_cols[0]:
    st.metric("Total issues", f"{total_issues:,}")
with kpi_cols[1]:
    st.metric("P1 critical", f"{p1_count:,}")
    st.markdown(kpi_caption(safe_pct(p1_count, total_issues)), unsafe_allow_html=True)
with kpi_cols[2]:
    st.metric("Resolved", f"{resolved_count:,}")
    st.markdown(kpi_caption(safe_pct(resolved_count, total_issues)), unsafe_allow_html=True)
with kpi_cols[3]:
    st.metric("Valid", f"{valid_count:,}")
    st.markdown(kpi_caption(safe_pct(valid_count, total_issues)), unsafe_allow_html=True)
with kpi_cols[4]:
    st.metric("Bugs", f"{bug_count:,}")
    st.markdown(kpi_caption(safe_pct(bug_count, total_issues)), unsafe_allow_html=True)

# -- Tabs ---------------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs(["Snapshot", "Quality", "Team", "Data explorer", "How we calculate"])

# -- TAB 1: Overview ----------------------------------------------------------
with tab1:
    # Row 1: Validity pie chart    |    Type of Request
    col1, col2 = st.columns(2)

    with col1:
        section("Validity distribution")
        if "Validity" in df.columns:
            validity = (
                df["Validity"].fillna("(blank)").astype(str).str.strip().value_counts().reset_index()
            )
            validity.columns = ["Validity", "Count"]
            total = int(validity["Count"].sum())
            valid_n = int(validity.loc[validity["Validity"].str.lower() == "valid", "Count"].sum())
            invalid_n = int(validity.loc[validity["Validity"].str.lower() == "invalid", "Count"].sum())
            valid_pct = (valid_n / total * 100) if total else 0
            # Full breakdown so the math (valid + invalid = total) is obvious.
            # Prevents "where did the missing rows go?" confusion.
            st.markdown(
                f'<div style="font-family:IBM Plex Mono,monospace;font-size:0.78rem;color:{MUTED};letter-spacing:0.08em;text-transform:uppercase;margin-bottom:8px">'
                f'<b style="color:{INK}">{valid_pct:.1f}%</b> valid &nbsp;·&nbsp; '
                f'<b style="color:{INK}">{valid_n:,}</b> valid &nbsp;·&nbsp; '
                f'<b style="color:{INK}">{invalid_n:,}</b> invalid &nbsp;·&nbsp; '
                f'<b style="color:{INK}">{total:,}</b> total</div>',
                unsafe_allow_html=True,
            )
            fig = px.pie(
                validity,
                names="Validity",
                values="Count",
                hole=0.55,
                color="Validity",
                color_discrete_map=VALIDITY_COLORS,
            )
            fig.update_traces(
                textinfo="percent+label",
                textfont=dict(color=INK, size=14, family="Inter"),
                marker=dict(line=dict(color="#FFFFFF", width=3)),
                pull=[0.02 if v.lower() == "valid" else 0 for v in validity["Validity"]],
            )
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_01_validity")

    with col2:
        section("Type of request")
        if "Type of Request" in df.columns and "Validity" in df.columns:
            # Cross-tab gives us the breakdown: each Type split into Valid + Invalid.
            # Stacked horizontal bars let leadership see both totals AND the
            # valid/invalid mix in a single glance.
            type_validity = (
                df.assign(_v=df["Validity"].fillna("(blank)").astype(str).str.strip())
                  .groupby(["Type of Request", "_v"]).size()
                  .reset_index(name="Count")
                  .rename(columns={"_v": "Validity"})
            )
            # Order types by total descending so largest shows on top
            type_totals = type_validity.groupby("Type of Request")["Count"].sum().sort_values(ascending=True)
            type_order = type_totals.index.tolist()

            stack_order = ["Valid", "Invalid", "(blank)"]
            available_validities = [v for v in stack_order if v in type_validity["Validity"].unique()]

            fig = go.Figure()
            for v in available_validities:
                subset = type_validity[type_validity["Validity"] == v].set_index("Type of Request").reindex(type_order).fillna(0)
                fig.add_trace(go.Bar(
                    name=v,
                    y=type_order,
                    x=subset["Count"],
                    orientation="h",
                    marker=dict(color=VALIDITY_COLORS.get(v, "#C8C2B5"), line=dict(width=0)),
                    text=[int(c) if c > 0 else "" for c in subset["Count"]],
                    textposition="inside",
                    insidetextanchor="middle",
                    textfont=dict(color="#FFFFFF", size=13, family="Inter", weight=700),
                    hovertemplate="<b>%{y}</b><br>" + v + ": <b>%{x}</b><extra></extra>",
                ))

            # Add the grand total as an outside label on each bar
            for i, t in enumerate(type_order):
                total = int(type_totals[t])
                fig.add_annotation(
                    x=total,
                    y=t,
                    text=f"<b>{total}</b>",
                    showarrow=False,
                    xanchor="left",
                    xshift=8,
                    font=dict(color=INK, size=13, family="Inter"),
                )

            fig.update_layout(
                barmode="stack",
                yaxis_title="",
                xaxis_title="",
                yaxis=dict(tickfont=dict(color=INK, size=13, family="Inter")),
                xaxis=dict(showgrid=True, gridcolor=GRID, range=[0, (max(type_totals) * 1.15 if len(type_totals) else 1)]),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                bargap=0.35,
            )
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_02_type")

    # Row 2: Version Distribution    |    Root Cause Resolution Breakdown
    col3, col4 = st.columns(2)

    with col3:
        section("Version distribution")
        version_col = find_column(df, [
            "Affects Version",
            "Affected Version",
            "Affect Version",
            "Affects Version/s",
            "Version",
        ])
        if version_col:
            # Count blanks/N-A explicitly so they show up as a (blank) bar
            # instead of being silently dropped. A row with NO version listed
            # is itself a data point — leadership wants to see it.
            raw = df[version_col].astype(str).str.strip()
            # Normalize what counts as "missing": NaN, empty, the strings
            # "nan"/"none", and the literal "N/A" / "NA" some teams type in.
            blank_mask = (
                df[version_col].isna()
                | raw.eq("")
                | raw.str.lower().isin(["nan", "none", "n/a", "na", "(blank)"])
            )
            blank_count = int(blank_mask.sum())

            # For the rows that DO have a version, explode comma-separated values
            non_blank = df.loc[~blank_mask, version_col].astype(str).str.strip()
            version_series = (
                non_blank
                .str.split(r"[,;\n]+", regex=True)
                .explode()
                .astype(str)
                .str.strip()
            )
            # Normalize again post-explode (e.g. trailing commas could yield empties)
            version_series = version_series[
                ~version_series.str.lower().isin(["", "nan", "none", "n/a", "na"])
            ]

            vc = version_series.value_counts().reset_index()
            vc.columns = ["Version", "Count"]

            # Take top 14 non-blank values + always reserve a slot for (blank)
            # if there are any. This way a small number of blanks never get
            # silently clipped out of the chart.
            top_n = 15 if blank_count == 0 else 14
            vc = vc.sort_values(["Count", "Version"], ascending=[False, True]).head(top_n)

            if blank_count > 0:
                vc = pd.concat(
                    [vc, pd.DataFrame([{"Version": "(blank)", "Count": blank_count}])],
                    ignore_index=True,
                )

            # Sort one more time so the chart displays in order with blank
            # placed naturally by its count
            vc = vc.sort_values(["Count", "Version"], ascending=[False, True])

            if not vc.empty:
                # Blanks get a muted gray; everything else uses the emerald palette
                # with the largest non-blank bar highlighted in deep emerald.
                non_blank_counts = vc.loc[vc["Version"] != "(blank)", "Count"]
                top_count = int(non_blank_counts.max()) if len(non_blank_counts) else 0
                colors = []
                for _, row in vc.iterrows():
                    if row["Version"] == "(blank)":
                        colors.append("#C8C2B5")  # neutral gray for missing
                    elif row["Count"] == top_count:
                        colors.append(CHART_EMERALD_DEEP)
                    else:
                        colors.append(CHART_EMERALD)

                version_fig = go.Figure(
                    data=[
                        go.Bar(
                            x=vc["Count"],
                            y=vc["Version"],
                            orientation="h",
                            text=vc["Count"],
                            textposition="outside",
                            textfont=dict(color=INK, size=13, family="Inter"),
                            marker=dict(color=colors, line=dict(width=0)),
                            cliponaxis=False,
                        )
                    ]
                )
                version_fig.update_layout(
                    template="simple_white",
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    margin=dict(l=10, r=10, t=20, b=10),
                    xaxis_title="",
                    yaxis_title="",
                    yaxis=dict(autorange="reversed", automargin=True, tickfont=dict(color=INK, size=12)),
                    xaxis=dict(tickfont=dict(color=INK, size=12), gridcolor=GRID),
                    height=max(340, 28 * len(vc) + 120),
                    font=dict(family="Inter", color=INK, size=13),
                )
                version_fig.update_xaxes(showgrid=True, gridcolor=GRID, zerolinecolor=GRID, linecolor=BORDER)
                version_fig.update_yaxes(showgrid=False, linecolor=BORDER)
                st.plotly_chart(version_fig, use_container_width=True, key="chart_03_version")
            else:
                st.info(f"No version values found in `{version_col}` for the current selection.")
        else:
            st.info("No version column detected.")

    with col4:
        section("Root cause resolution breakdown")
        if "Root Cause Resolution" in df.columns:
            rc = df["Root Cause Resolution"].value_counts().reset_index()
            rc.columns = ["Root Cause", "Count"]
            fig = px.bar(
                rc,
                x="Count",
                y="Root Cause",
                orientation="h",
                text="Count",
            )
            fig.update_traces(
                marker=dict(color=bar_palette(rc["Count"].tolist(), CHART_ROSE, CHART_ROSE_DEEP), line=dict(width=0)),
                textposition="outside",
                textfont=dict(color=INK, size=13, family="Inter"),
                cliponaxis=False,
            )
            fig.update_layout(yaxis=dict(autorange="reversed"), yaxis_title="", xaxis_title="")
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_04_rootcause")

    # Row 3: Component Distribution    |    Overall Volume
    col5, col6 = st.columns(2)

    with col5:
        section("Component distribution")
        if "Component/s" in df.columns:
            # Split components on common separators (comma, semicolon, newline).
            # Some Jira exports use one, some use another — being permissive
            # here means the chart stays accurate regardless of which export
            # convention the upstream script uses.
            comp_series = (
                df["Component/s"].dropna()
                .str.split(r"[,;\n]+", regex=True)
                .explode()
                .str.strip()
            )
            comp_series = comp_series[comp_series != ""]
            comp_counts = comp_series.value_counts().head(15).reset_index()
            comp_counts.columns = ["Component", "Count"]

            if not comp_counts.empty:
                # Bloomberg-style horizontal lollipop chart.
                # Each component is a thin connector line + a circular marker
                # at its count value. Reads as a clean ranked leaderboard —
                # the eye scans top-to-bottom with the dots as anchor points.
                # Top-1 gets a deeper accent so it visually pops.
                comp_sorted = comp_counts.sort_values("Count", ascending=True).reset_index(drop=True)
                top_count = int(comp_sorted["Count"].max())

                marker_colors = [
                    CHART_INDIGO_DEEP if int(c) == top_count else CHART_INDIGO
                    for c in comp_sorted["Count"]
                ]
                marker_sizes = [
                    16 if int(c) == top_count else 13
                    for c in comp_sorted["Count"]
                ]

                fig = go.Figure()
                # Thin connector lines from y-axis to each marker
                for i, row in comp_sorted.iterrows():
                    fig.add_shape(
                        type="line",
                        x0=0, x1=row["Count"],
                        y0=row["Component"], y1=row["Component"],
                        line=dict(color="rgba(124, 111, 232, 0.35)", width=1.5),
                        layer="below",
                    )
                # The lollipop "candy" — one marker per component
                fig.add_trace(go.Scatter(
                    x=comp_sorted["Count"],
                    y=comp_sorted["Component"],
                    mode="markers+text",
                    marker=dict(
                        color=marker_colors,
                        size=marker_sizes,
                        line=dict(color="#FFFFFF", width=2),
                    ),
                    text=comp_sorted["Count"],
                    textposition="middle right",
                    textfont=dict(color=INK, size=12, family="Inter", weight=700),
                    cliponaxis=False,
                    hovertemplate="<b>%{y}</b><br>%{x} issues<extra></extra>",
                ))
                fig.update_layout(
                    xaxis=dict(
                        showgrid=True, gridcolor=GRID,
                        zeroline=False,
                        range=[0, comp_sorted["Count"].max() * 1.15],
                        tickfont=dict(color=INK, size=11),
                    ),
                    yaxis=dict(
                        tickfont=dict(color=INK, size=12, family="Inter", weight=600),
                    ),
                    height=max(380, 26 * len(comp_sorted) + 80),
                    showlegend=False,
                    margin=dict(l=10, r=10, t=10, b=10),
                )
                chart_theme(fig)
                st.plotly_chart(fig, use_container_width=True, key="chart_05_component")

    with col6:
        section("Volume trend")
        volume_summary = build_volume_summary(df, reporting_date_col)
        if volume_summary is not None and not volume_summary.empty:
            # Area + line + markers — communicates trend at a glance.
            # Highlight the peak month with a deeper marker so the eye locks on it.
            peak_idx = int(volume_summary["Count"].idxmax())
            marker_colors = [
                CHART_STEEL_DEEP if i == peak_idx else CHART_STEEL
                for i in range(len(volume_summary))
            ]
            marker_sizes = [11 if i == peak_idx else 8 for i in range(len(volume_summary))]

            fig = go.Figure()
            # Soft fill under the line for visual weight
            fig.add_trace(go.Scatter(
                x=volume_summary["Label"],
                y=volume_summary["Count"],
                mode="lines",
                line=dict(color=CHART_STEEL, width=0),
                fill="tozeroy",
                fillcolor="rgba(91, 124, 184, 0.10)",
                hoverinfo="skip",
                showlegend=False,
            ))
            # The actual line + markers + value labels
            fig.add_trace(go.Scatter(
                x=volume_summary["Label"],
                y=volume_summary["Count"],
                mode="lines+markers+text",
                line=dict(color=CHART_STEEL_DEEP, width=2.5, shape="spline", smoothing=0.6),
                marker=dict(
                    color=marker_colors,
                    size=marker_sizes,
                    line=dict(color="#FFFFFF", width=2),
                ),
                text=volume_summary["Count"],
                textposition="top center",
                textfont=dict(color=INK, size=12, family="Inter", weight=700),
                hovertemplate="<b>%{x}</b><br>%{y} issues<extra></extra>",
                showlegend=False,
                cliponaxis=False,
            ))
            fig.update_layout(
                xaxis_title="",
                yaxis_title="",
                yaxis=dict(rangemode="tozero"),
            )
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_06_volume")

# -- TAB 2: Quality -----------------------------------------------------------
with tab2:
    col1, col2 = st.columns(2)

    with col1:
        section("Resolution time trend")
        # Calculates avg & median days from Creation Date → Resolution Date,
        # bucketed by resolution month. Tells leadership "are we getting faster?"
        if "Creation Date" in df.columns and "Resolution Date" in df.columns:
            res_df = df.dropna(subset=["Creation Date", "Resolution Date"]).copy()
            res_df["DaysToResolve"] = (
                res_df["Resolution Date"] - res_df["Creation Date"]
            ).dt.days
            res_df = res_df[res_df["DaysToResolve"] >= 0]

            if not res_df.empty:
                res_df["ResMonth"] = res_df["Resolution Date"].dt.to_period("M")
                trend = (
                    res_df.groupby("ResMonth")["DaysToResolve"]
                    .agg(mean_days="mean", median_days="median", count="size")
                    .reset_index()
                    .sort_values("ResMonth")
                )
                trend["Label"] = trend["ResMonth"].dt.to_timestamp().dt.strftime("%b %Y")

                fig = go.Figure()
                # Soft fill under median line
                fig.add_trace(go.Scatter(
                    x=trend["Label"],
                    y=trend["median_days"],
                    mode="lines",
                    line=dict(color=CHART_INDIGO, width=0),
                    fill="tozeroy",
                    fillcolor="rgba(124, 111, 232, 0.10)",
                    hoverinfo="skip",
                    showlegend=False,
                ))
                fig.add_trace(go.Scatter(
                    x=trend["Label"],
                    y=trend["median_days"],
                    mode="lines+markers+text",
                    name="Median days",
                    line=dict(color=CHART_INDIGO_DEEP, width=2.5, shape="spline", smoothing=0.6),
                    marker=dict(size=9, color=CHART_INDIGO_DEEP, line=dict(color="#FFFFFF", width=2)),
                    text=[f"{int(d)}d" for d in trend["median_days"]],
                    textposition="top center",
                    textfont=dict(color=INK, size=12, family="Inter", weight=700),
                    hovertemplate="<b>%{x}</b><br>Median: %{y:.0f} days<extra></extra>",
                    cliponaxis=False,
                ))
                # Mean as a lighter dashed reference line
                fig.add_trace(go.Scatter(
                    x=trend["Label"],
                    y=trend["mean_days"],
                    mode="lines",
                    name="Mean days",
                    line=dict(color=CHART_HONEY, width=1.8, dash="dot"),
                    hovertemplate="<b>%{x}</b><br>Mean: %{y:.0f} days<extra></extra>",
                ))
                fig.update_layout(
                    xaxis_title="",
                    yaxis_title="",
                    yaxis=dict(rangemode="tozero"),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                )
                chart_theme(fig)
                st.plotly_chart(fig, use_container_width=True, key="chart_10")
            else:
                st.info("Not enough date data to compute resolution time trend.")

    with col2:
        section("Classification accuracy")
        # How often was the original RFH-vs-Bug classification correct?
        # Broken down by Type of Request — tells you if bugs are harder to
        # classify than RFHs (or vice versa).
        rfh_col = find_column(df, ["Correct / Incorrect classification? RFH/Bug"])
        if rfh_col and "Type of Request" in df.columns:
            ct = (
                df.assign(
                    _t=df["Type of Request"].fillna("(blank)").astype(str).str.strip(),
                    _c=df[rfh_col].fillna("(blank)").astype(str).str.strip(),
                )
                .groupby(["_t", "_c"]).size()
                .reset_index(name="Count")
                .rename(columns={"_t": "Type", "_c": "Classification"})
            )
            type_totals = ct.groupby("Type")["Count"].sum().sort_values(ascending=True)
            type_order = type_totals.index.tolist()

            class_colors = {"Correct": CHART_EMERALD, "Incorrect": CHART_TERRACOTTA_DEEP, "(blank)": "#C8C2B5"}
            stack_order = ["Correct", "Incorrect", "(blank)"]
            available = [c for c in stack_order if c in ct["Classification"].unique()]

            fig = go.Figure()
            for c in available:
                subset = ct[ct["Classification"] == c].set_index("Type").reindex(type_order).fillna(0)
                fig.add_trace(go.Bar(
                    name=c,
                    y=type_order,
                    x=subset["Count"],
                    orientation="h",
                    marker=dict(color=class_colors.get(c, "#A89DEF"), line=dict(width=0)),
                    text=[int(v) if v > 0 else "" for v in subset["Count"]],
                    textposition="inside",
                    insidetextanchor="middle",
                    textfont=dict(color="#FFFFFF", size=12, family="Inter", weight=700),
                    hovertemplate="<b>%{y}</b><br>" + c + ": <b>%{x}</b><extra></extra>",
                ))
            for t in type_order:
                total = int(type_totals[t])
                fig.add_annotation(
                    x=total, y=t, text=f"<b>{total}</b>", showarrow=False,
                    xanchor="left", xshift=8,
                    font=dict(color=INK, size=13, family="Inter"),
                )
            fig.update_layout(
                barmode="stack",
                yaxis_title="", xaxis_title="",
                yaxis=dict(tickfont=dict(color=INK, size=13, family="Inter")),
                xaxis=dict(showgrid=True, gridcolor=GRID, range=[0, (max(type_totals) * 1.15 if len(type_totals) else 1)]),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                bargap=0.35,
            )
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_11")

    # Row 2: Scope for deflection alone (left half, right half empty for breathing room)
    col3, col4 = st.columns(2)

    with col3:
        section("Scope for deflection")
        if "Scope for Deflection" in df.columns:
            defl = df["Scope for Deflection"].value_counts().reset_index()
            defl.columns = ["Scope", "Count"]
            fig = px.bar(
                defl,
                x="Scope",
                y="Count",
                text="Count",
            )
            fig.update_traces(
                marker=dict(color=bar_palette(defl["Count"].tolist(), CHART_PLUM, CHART_PLUM_DEEP), line=dict(width=0)),
                textposition="outside",
                textfont=dict(color=INK, size=13, family="Inter"),
                cliponaxis=False,
            )
            fig.update_layout(xaxis_title="", yaxis_title="")
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_12")

# -- TAB 3: Team --------------------------------------------------------------
with tab3:
    col1, col2 = st.columns(2)

    with col1:
        section("Issues per R&D engineer")
        if "R&D Engineer" in df.columns:
            ec = df["R&D Engineer"].value_counts().reset_index()
            ec.columns = ["Engineer", "Count"]
            top_ec = ec.head(15)
            fig = px.bar(
                top_ec,
                x="Count",
                y="Engineer",
                orientation="h",
                text="Count",
            )
            fig.update_traces(
                marker=dict(color=bar_palette(top_ec["Count"].tolist(), CHART_INDIGO, CHART_INDIGO_DEEP), line=dict(width=0)),
                textposition="outside",
                textfont=dict(color=INK, size=13, family="Inter"),
                cliponaxis=False,
            )
            fig.update_layout(yaxis_title="", xaxis_title="", yaxis=dict(autorange="reversed"))
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
            "⬇ Download Excel",
            make_xlsx_bytes(root_export[export_cols].drop_duplicates(), sheet_name="Root_Cause"),
            file_name="root_cause_rows.xlsx",
            mime=XLSX_MIME,
            key="dl_root_cause_xlsx",
        )

    if "Creator" in df.columns:
        section("Issues created by")

        # Determine which dimensions are available to split on
        has_validity = "Validity" in df.columns
        has_type = "Type of Request" in df.columns
        split_options = []
        if has_validity:
            split_options.append("Valid / Invalid")
        if has_type:
            split_options.append("RFH / Bug")
        if not split_options:
            split_options = ["Total only"]

        # Small toggle above the chart — minimal radio so leadership can flip
        # views during the meeting without touching anything else.
        split_container = st.container(key="creator_split_picker")
        with split_container:
            split_mode = st.radio(
                "Split by",
                split_options,
                index=0,
                horizontal=True,
                key="creator_split_mode",
                label_visibility="collapsed",
            )

        # Top 12 creators by total volume
        top_creators = df["Creator"].value_counts().head(12).index.tolist()
        creator_df = df[df["Creator"].isin(top_creators)].copy()

        # Pick the split dimension and define color/order based on selection
        if split_mode == "Valid / Invalid":
            split_col = "Validity"
            stack_order = ["Valid", "Invalid", "(blank)"]
            split_colors = VALIDITY_COLORS
        elif split_mode == "RFH / Bug":
            split_col = "Type of Request"
            stack_order = ["Request for Help", "Bug", "(blank)"]
            split_colors = {
                "Request for Help": CHART_STEEL,
                "Bug": CHART_TERRACOTTA,
                "(blank)": "#C8C2B5",
            }
        else:
            split_col = None

        if split_col is None or split_col not in creator_df.columns:
            # Fallback — original simple bar with no split
            cc = creator_df["Creator"].value_counts().reset_index()
            cc.columns = ["Creator", "Count"]
            cc = cc.head(12)
            fig = px.bar(cc, x="Count", y="Creator", orientation="h", text="Count")
            fig.update_traces(
                marker=dict(color=bar_palette(cc["Count"].tolist(), CHART_PLUM, CHART_PLUM_DEEP), line=dict(width=0)),
                textposition="outside",
                textfont=dict(color=INK, size=13, family="Inter"),
                cliponaxis=False,
            )
            fig.update_layout(yaxis=dict(autorange="reversed"), yaxis_title="", xaxis_title="")
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_15")
        else:
            # Stacked bar: each Creator's bar split by the chosen dimension.
            # Same visual pattern as the Type of Request chart on the
            # Overview tab — leadership already knows how to read it.
            cross = (
                creator_df.assign(_s=creator_df[split_col].fillna("(blank)").astype(str).str.strip())
                .groupby(["Creator", "_s"]).size()
                .reset_index(name="Count")
                .rename(columns={"_s": "Split"})
            )

            # Order creators ascending by total so largest renders on top
            creator_totals = (
                cross.groupby("Creator")["Count"].sum().sort_values(ascending=True)
            )
            creator_order = creator_totals.index.tolist()

            available_splits = [s for s in stack_order if s in cross["Split"].unique()]
            # Plus any unexpected values that aren't in our predefined order
            for s in cross["Split"].unique():
                if s not in available_splits:
                    available_splits.append(s)

            fig = go.Figure()
            for s in available_splits:
                subset = (
                    cross[cross["Split"] == s]
                    .set_index("Creator")
                    .reindex(creator_order)
                    .fillna(0)
                )
                fig.add_trace(go.Bar(
                    name=s,
                    y=creator_order,
                    x=subset["Count"],
                    orientation="h",
                    marker=dict(color=split_colors.get(s, "#A89DEF"), line=dict(width=0)),
                    text=[int(c) if c > 0 else "" for c in subset["Count"]],
                    textposition="inside",
                    insidetextanchor="middle",
                    textfont=dict(color="#FFFFFF", size=12, family="Inter", weight=700),
                    hovertemplate="<b>%{y}</b><br>" + s + ": <b>%{x}</b><extra></extra>",
                ))

            # Grand total annotation outside each bar
            for creator in creator_order:
                total = int(creator_totals[creator])
                fig.add_annotation(
                    x=total,
                    y=creator,
                    text=f"<b>{total}</b>",
                    showarrow=False,
                    xanchor="left",
                    xshift=8,
                    font=dict(color=INK, size=13, family="Inter"),
                )

            fig.update_layout(
                barmode="stack",
                yaxis_title="",
                xaxis_title="",
                yaxis=dict(tickfont=dict(color=INK, size=12, family="Inter")),
                xaxis=dict(showgrid=True, gridcolor=GRID, range=[0, (max(creator_totals) * 1.15 if len(creator_totals) else 1)]),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                bargap=0.30,
                height=max(380, 32 * len(creator_order) + 80),
            )
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_15")

# Scoped CSS for the split-mode radio — make it look like a proper toggle
# (compact pill row instead of vertical radio buttons).
st.markdown(
    f"""
<style>
  .st-key-creator_split_picker {{
    margin: 4px 0 14px 0;
  }}
  .st-key-creator_split_picker [role="radiogroup"] {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 10px;
    padding: 3px;
    display: inline-flex !important;
    gap: 2px;
    box-shadow: 0 1px 2px rgba(26, 26, 26, 0.04);
  }}
  .st-key-creator_split_picker [role="radiogroup"] label {{
    margin: 0 !important;
    padding: 6px 14px !important;
    border-radius: 7px !important;
    cursor: pointer !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 0.82rem !important;
    font-weight: 600 !important;
    color: {SUBTEXT} !important;
    transition: all 0.15s ease;
  }}
  /* Beat BaseWeb's text-color overrides on nested p/span/div */
  .st-key-creator_split_picker [role="radiogroup"] label p,
  .st-key-creator_split_picker [role="radiogroup"] label span,
  .st-key-creator_split_picker [role="radiogroup"] label div {{
    color: {SUBTEXT} !important;
    font-weight: 600 !important;
  }}
  .st-key-creator_split_picker [role="radiogroup"] label:hover,
  .st-key-creator_split_picker [role="radiogroup"] label:hover p,
  .st-key-creator_split_picker [role="radiogroup"] label:hover span,
  .st-key-creator_split_picker [role="radiogroup"] label:hover div {{
    background: rgba(124, 111, 232, 0.08);
    color: {INK} !important;
  }}
  .st-key-creator_split_picker [role="radiogroup"] label:has(input:checked) {{
    background: linear-gradient(135deg, {CHART_INDIGO} 0%, {CHART_INDIGO_DEEP} 100%);
    color: #FFFFFF !important;
    box-shadow: 0 2px 6px rgba(124, 111, 232, 0.28);
  }}
  /* White text inside the selected pill, all nested elements */
  .st-key-creator_split_picker [role="radiogroup"] label:has(input:checked) p,
  .st-key-creator_split_picker [role="radiogroup"] label:has(input:checked) span,
  .st-key-creator_split_picker [role="radiogroup"] label:has(input:checked) div {{
    color: #FFFFFF !important;
    font-weight: 700 !important;
  }}
  /* Hide the actual radio circle — we only want the labels as pills */
  .st-key-creator_split_picker [role="radiogroup"] label > div:first-child {{
    display: none !important;
  }}
</style>
""",
    unsafe_allow_html=True,
)

# Scoped CSS for the Data explorer cards, divider, and match-mode toggle.
# Each section becomes a distinct surface so leadership doesn't mistake one
# tool's chrome for the other tool's controls.
st.markdown(
    f"""
<style>
  /* Section header for each explorer tool — colored dot + title + subtitle
     in one row, with a divider line beneath. No content-wrapping container —
     just a strong visual marker that this is the start of a new tool. */
  .explorer-section-header {{
    display: flex;
    align-items: baseline;
    gap: 12px;
    margin: 6px 0 16px 0;
    padding-bottom: 14px;
    border-bottom: 1px solid {BORDER};
  }}
  .explorer-card-dot {{
    display: inline-block;
    width: 10px;
    height: 10px;
    border-radius: 50%;
    flex-shrink: 0;
    transform: translateY(1px);
  }}
  .explorer-card-title {{
    font-family: 'Inter', sans-serif;
    font-size: 1.05rem;
    font-weight: 700;
    color: {INK};
    letter-spacing: -0.005em;
  }}
  .explorer-card-subtitle {{
    font-family: 'Inter', sans-serif;
    font-size: 0.82rem;
    font-weight: 500;
    color: {MUTED};
    margin-left: 2px;
  }}
  /* Divider between the two tools — quiet horizontal rule with breathing room */
  .explorer-divider {{
    height: 1px;
    background: linear-gradient(to right,
      transparent 0%, {BORDER} 12%, {BORDER} 88%, transparent 100%);
    margin: 36px 0 28px 0;
  }}

  /* Match-mode toggle pill — same visual grammar as the creator split picker */
  .st-key-match_mode_picker {{
    margin: 0 0 16px 0;
  }}
  .st-key-match_mode_picker [role="radiogroup"] {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 10px;
    padding: 3px;
    display: inline-flex !important;
    gap: 2px;
    box-shadow: 0 1px 2px rgba(26, 26, 26, 0.04);
  }}
  .st-key-match_mode_picker [role="radiogroup"] label {{
    margin: 0 !important;
    padding: 7px 16px !important;
    border-radius: 7px !important;
    cursor: pointer !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 0.82rem !important;
    font-weight: 600 !important;
    color: {SUBTEXT} !important;
    transition: all 0.15s ease;
  }}
  /* Force the text inside the label (paragraph/span/div Streamlit injects)
     to use the same color as the label — beats BaseWeb's inline JS overrides */
  .st-key-match_mode_picker [role="radiogroup"] label p,
  .st-key-match_mode_picker [role="radiogroup"] label span,
  .st-key-match_mode_picker [role="radiogroup"] label div {{
    color: {SUBTEXT} !important;
    font-weight: 600 !important;
  }}
  .st-key-match_mode_picker [role="radiogroup"] label:hover,
  .st-key-match_mode_picker [role="radiogroup"] label:hover p,
  .st-key-match_mode_picker [role="radiogroup"] label:hover span,
  .st-key-match_mode_picker [role="radiogroup"] label:hover div {{
    background: rgba(124, 111, 232, 0.08);
    color: {INK} !important;
  }}
  .st-key-match_mode_picker [role="radiogroup"] label:has(input:checked) {{
    background: linear-gradient(135deg, {CHART_INDIGO} 0%, {CHART_INDIGO_DEEP} 100%);
    color: #FFFFFF !important;
    box-shadow: 0 2px 6px rgba(124, 111, 232, 0.28);
  }}
  /* Selected pill — force white text on every nested element */
  .st-key-match_mode_picker [role="radiogroup"] label:has(input:checked) p,
  .st-key-match_mode_picker [role="radiogroup"] label:has(input:checked) span,
  .st-key-match_mode_picker [role="radiogroup"] label:has(input:checked) div {{
    color: #FFFFFF !important;
    font-weight: 700 !important;
  }}
  .st-key-match_mode_picker [role="radiogroup"] label > div:first-child {{
    display: none !important;
  }}
</style>
""",
    unsafe_allow_html=True,
)

# -- TAB 4: Data explorer -----------------------------------------------------
with tab4:
    # ─── Section 1: Filtered issues table ────────────────────────────────────
    # Shows exactly what the dashboard is currently filtered to via the
    # period dropdown + sidebar filters. No additional filtering happens here —
    # the table is a faithful view of the data driving the rest of the page.

    st.markdown(
        f'<div class="explorer-section-header">'
        f'<span class="explorer-card-dot" style="background:{CHART_INDIGO};"></span>'
        f'<span class="explorer-card-title">Filtered issues table</span>'
        f'<span class="explorer-card-subtitle">View what the dashboard is currently filtered to</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        f'<div style="font-family:Inter,sans-serif;font-size:0.88rem;color:{SUBTEXT};line-height:1.55;margin:0 0 14px 0;">'
        f'Showing all <b style="color:{INK}">{len(df):,}</b> issues matching the current period and sidebar filters. '
        f'Use the search box below to narrow further by summary text.'
        f'</div>',
        unsafe_allow_html=True,
    )

    table_search = st.text_input(
        "Search by summary text",
        placeholder="Type any keyword from the summary…",
        key="table_search",
    )
    table_df = df.copy()
    if table_search and "Summary" in table_df.columns:
        table_df = table_df[
            table_df["Summary"].astype(str).str.contains(table_search, case=False, na=False)
        ]

    table_show_cols = [c for c in table_df.columns if c not in ["Priority_Short", "Resolution Month", "Quarter Label", "Source Sheet"]]
    st.dataframe(table_df[table_show_cols], use_container_width=True, height=420)
    st.caption(f"{len(table_df):,} of {len(df):,} rows shown · ID column detected: `{detect_id_column(df) or 'none'}`")

    st.download_button(
        "⬇ Download Excel",
        make_xlsx_bytes(table_df[table_show_cols], sheet_name="Filtered_Issues"),
        file_name="filtered_issues.xlsx",
        mime=XLSX_MIME,
        key="dl_filtered_xlsx",
    )

    # Visual divider between the two tools
    st.markdown('<div class="explorer-divider"></div>', unsafe_allow_html=True)

    # ─── Section 2: Custom export builder ────────────────────────────────────
    # Natural-language-style filter rows. A "Match" toggle controls whether
    # multiple conditions combine with AND or OR.

    st.markdown(
        f'<div class="explorer-section-header">'
        f'<span class="explorer-card-dot" style="background:{CHART_TERRACOTTA};"></span>'
        f'<span class="explorer-card-title">Custom export builder</span>'
        f'<span class="explorer-card-subtitle">Build your own slice of the data</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        f'<div style="font-family:Inter,sans-serif;font-size:0.86rem;color:{SUBTEXT};margin:-4px 0 16px 0;line-height:1.55;">'
        f'Build a custom slice of the data with one or more conditions, then download as Excel. '
        f'Use <b style="color:{INK}">All conditions</b> to require every condition to match (AND), or '
        f'<b style="color:{INK}">Any condition</b> to match rows that satisfy at least one (OR).'
        f'</div>',
        unsafe_allow_html=True,
    )

    # Reusable helpers ────────────────────────────────────────────────────
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

    # Curate the column list — show meaningful filter columns first, hide
    # internal helper columns. Keeps the dropdown clean for the user.
    PRIORITY_COLUMNS_HINT = [
        "Priority_Short", "Validity", "Type of Request", "Status",
        "Project key", "R&D Engineer", "Creator", "Component/s",
        "Affects Version", "Fixed Version", "Resolution",
        "Root Cause Resolution", "Scope for Deflection",
        "Correct / Incorrect classification? RFH/Bug",
    ]
    HIDDEN_COLUMNS = {
        # Raw Priority has the long descriptions — Priority_Short shows the
        # P1/P2/P3/P4 codes which is what users actually want to filter by.
        # Priority_Short is displayed in the dropdown labeled simply "Priority".
        "Priority",
        # Internal helpers users shouldn't see
        "Resolution Month", "Quarter Label", "Source Sheet",
    }
    available_cols = [c for c in PRIORITY_COLUMNS_HINT if c in df.columns]
    other_cols = [c for c in df.columns if c not in available_cols and c not in HIDDEN_COLUMNS]
    column_choices = available_cols + other_cols

    # Friendlier display names — "Priority_Short" → "Priority"
    COLUMN_DISPLAY_NAMES = {
        "Priority_Short": "Priority",
        "Component/s": "Component",
        "Project key": "Project",
        "R&D Engineer": "R&D engineer",
        "Type of Request": "Type of request",
        "Affects Version": "Affects version",
        "Fixed Version": "Fixed version",
        "Root Cause Resolution": "Root cause",
        "Scope for Deflection": "Scope for deflection",
        "Correct / Incorrect classification? RFH/Bug": "Classification correctness",
    }
    def col_display(col: str) -> str:
        return COLUMN_DISPLAY_NAMES.get(col, col)

    # Filter row state — list of dicts. Each row is one condition.
    if "explorer_filters" not in st.session_state:
        st.session_state.explorer_filters = [
            {"column": "Priority_Short" if "Priority_Short" in df.columns else column_choices[0], "op": "is", "value": None},
        ]
    if "explorer_match_mode" not in st.session_state:
        st.session_state.explorer_match_mode = "All conditions (AND)"

    # Match-mode toggle — controls whether multiple conditions combine with
    # AND (every condition must match) or OR (any condition matches).
    # One global setting keeps the logic unambiguous; no per-row precedence
    # confusion. If a user ever needs grouped (A OR B) AND (C OR D) logic,
    # that's a future enhancement.
    match_picker = st.container(key="match_mode_picker")
    with match_picker:
        match_mode = st.radio(
            "Match mode",
            ["All conditions (AND)", "Any condition (OR)"],
            index=0 if st.session_state.explorer_match_mode == "All conditions (AND)" else 1,
            horizontal=True,
            key="explorer_match_mode",
            label_visibility="collapsed",
        )
    is_or_mode = match_mode.startswith("Any")

    # Action buttons row: Add condition / Clear all
    action_cols = st.columns([1, 1, 4])
    with action_cols[0]:
        if st.button("➕ Add condition", key="add_filter_btn", use_container_width=True):
            default_col = column_choices[0] if column_choices else None
            st.session_state.explorer_filters.append({"column": default_col, "op": "is", "value": None})
            st.rerun()
    with action_cols[1]:
        if st.button("✕ Clear all", key="clear_filters_btn", use_container_width=True):
            st.session_state.explorer_filters = [
                {"column": "Priority_Short" if "Priority_Short" in df.columns else column_choices[0], "op": "is", "value": None},
            ]
            st.rerun()

    # Render each filter row in a natural-language layout:
    # [WHERE Priority] [is] [P1]  [✕]
    OPERATORS = ["is", "is not", "contains", "does not contain"]

    for i, flt in enumerate(st.session_state.explorer_filters):
        row_cols = st.columns([0.6, 2.2, 1.5, 3, 0.8])

        # Connector word — "Where" for first row, "And" or "Or" for the rest
        # depending on the match mode the user picked above.
        with row_cols[0]:
            if i == 0:
                connector = "Where"
            else:
                connector = "Or" if is_or_mode else "And"
            connector_color = CHART_TERRACOTTA_DEEP if is_or_mode and i > 0 else MUTED
            st.markdown(
                f'<div style="font-family:Inter,sans-serif;font-size:0.85rem;font-weight:700;'
                f'color:{connector_color};padding:8px 0 0 0;text-align:right;">{connector}</div>',
                unsafe_allow_html=True,
            )

        # Column picker — uses display names so user sees "Priority" not "Priority_Short"
        with row_cols[1]:
            current_col = flt.get("column") or column_choices[0]
            if current_col not in column_choices:
                current_col = column_choices[0]
            col_choice = st.selectbox(
                "Column",
                column_choices,
                index=column_choices.index(current_col) if current_col in column_choices else 0,
                format_func=col_display,
                key=f"flt_col_{i}",
                label_visibility="collapsed",
            )
            flt["column"] = col_choice

        # Operator picker
        with row_cols[2]:
            current_op = flt.get("op") or "is"
            if current_op not in OPERATORS:
                current_op = "is"
            op_choice = st.selectbox(
                "Operator",
                OPERATORS,
                index=OPERATORS.index(current_op),
                key=f"flt_op_{i}",
                label_visibility="collapsed",
            )
            flt["op"] = op_choice

        # Value picker — dropdown for "is"/"is not", text input for "contains"
        with row_cols[3]:
            if op_choice in ("is", "is not"):
                opts = build_options(df, col_choice)
                if not opts:
                    opts = ["(blank)"]
                # Persist previous selection if still valid
                prev_val = flt.get("value")
                default_idx = opts.index(prev_val) if prev_val in opts else 0
                val = st.selectbox(
                    "Value",
                    opts,
                    index=default_idx,
                    key=f"flt_val_{i}",
                    label_visibility="collapsed",
                )
            else:
                val = st.text_input(
                    "Value",
                    value=str(flt.get("value") or ""),
                    placeholder="Type a substring…",
                    key=f"flt_val_text_{i}",
                    label_visibility="collapsed",
                )
            flt["value"] = val

        # Remove button — only show if there's more than 1 row
        with row_cols[4]:
            if len(st.session_state.explorer_filters) > 1:
                if st.button("✕", key=f"flt_remove_{i}", help="Remove this condition"):
                    st.session_state.explorer_filters.pop(i)
                    st.rerun()

    # Apply filters ─────────────────────────────────────────────────────
    def condition_mask(frame: pd.DataFrame, column: str, operator: str, value) -> pd.Series:
        """Return a boolean mask of rows matching one condition.
        Returning a mask (instead of a filtered frame) lets us combine
        multiple conditions with & (AND) or | (OR) at the call site —
        which is what the Match toggle controls.
        """
        # Default = "match every row" so an invalid spec passes through harmlessly
        all_true = pd.Series(True, index=frame.index)
        if column not in frame.columns or value is None or value == "":
            return all_true
        series = frame[column].map(normalise_value).astype(str)
        value_str = normalise_value(value) if operator in ("is", "is not") else str(value).strip()
        if not value_str:
            return all_true
        op = (operator or "is").lower().strip()
        if op == "is":
            return series.str.lower() == value_str.lower()
        if op == "is not":
            return series.str.lower() != value_str.lower()
        if op == "contains":
            return series.str.contains(value_str, case=False, na=False)
        if op == "does not contain":
            return ~series.str.contains(value_str, case=False, na=False)
        return series.str.lower() == value_str.lower()

    # Accumulate masks per condition, then combine based on Match mode.
    # AND mode → all masks must be True (start with all True, AND each in).
    # OR  mode → at least one mask must be True (start with all False, OR each in).
    base_df = df.copy()
    active_filters = []
    masks = []
    for flt in st.session_state.explorer_filters:
        col = flt.get("column")
        op = flt.get("op")
        val = flt.get("value")
        if col and val not in (None, "", "(choose a column first)"):
            masks.append(condition_mask(base_df, col, op, val))
            active_filters.append((col, op, val))

    if not masks:
        export_df = base_df
    elif is_or_mode:
        combined = pd.Series(False, index=base_df.index)
        for m in masks:
            combined = combined | m
        export_df = base_df[combined]
    else:
        combined = pd.Series(True, index=base_df.index)
        for m in masks:
            combined = combined & m
        export_df = base_df[combined]

    # Active filters badge strip — visual confirmation of what's applied
    if active_filters:
        badge_html = '<div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:14px 0 12px 0;">'
        for idx, (col, op, val) in enumerate(active_filters):
            # Connector word between badges — shows the user how their
            # conditions are combining (AND in muted gray, OR in terracotta).
            if idx > 0:
                if is_or_mode:
                    badge_html += (
                        f'<span style="font-family:Inter,sans-serif;font-size:0.72rem;'
                        f'font-weight:800;letter-spacing:0.08em;color:{CHART_TERRACOTTA_DEEP};">OR</span>'
                    )
                else:
                    badge_html += (
                        f'<span style="font-family:Inter,sans-serif;font-size:0.72rem;'
                        f'font-weight:700;letter-spacing:0.08em;color:{MUTED};">AND</span>'
                    )
            badge_html += (
                f'<span style="display:inline-flex;align-items:center;gap:6px;'
                f'padding:5px 12px;border-radius:999px;background:rgba(124,111,232,0.10);'
                f'color:{CHART_INDIGO_DEEP};border:1px solid rgba(124,111,232,0.20);'
                f'font-family:Inter,sans-serif;font-size:0.8rem;font-weight:600;">'
                f'{col_display(col)} <span style="color:{MUTED};font-weight:500">{op}</span> '
                f'<b>{val}</b></span>'
            )
        badge_html += "</div>"
        st.markdown(badge_html, unsafe_allow_html=True)

    # Result count + preview ────────────────────────────────────────────
    # Custom markdown instead of st.metric — st.metric interprets the third
    # argument as a delta and renders a green ↑ arrow icon, which here
    # incorrectly suggests the row count is "going up" relative to something.
    st.markdown(
        f'<div style="margin: 8px 0 14px 0;">'
        f'  <div style="font-family:Inter,sans-serif;font-size:0.74rem;font-weight:700;'
        f'letter-spacing:0.14em;text-transform:uppercase;color:{MUTED};margin-bottom:4px;">'
        f'    Matching rows'
        f'  </div>'
        f'  <div style="display:flex;align-items:baseline;gap:10px;">'
        f'    <span style="font-family:\'IBM Plex Mono\',monospace;font-size:2rem;'
        f'font-weight:600;color:{INK};line-height:1;">{len(export_df):,}</span>'
        f'    <span style="font-family:Inter,sans-serif;font-size:0.86rem;color:{SUBTEXT};">'
        f'      of {len(df):,} filtered'
        f'    </span>'
        f'  </div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # Column selector — let user pick what columns to include in the export
    id_col = detect_id_column(export_df)
    default_export_cols = []
    if id_col:
        default_export_cols.append(id_col)
    for col, _, _ in active_filters:
        if col not in default_export_cols:
            default_export_cols.append(col)
    for fallback in ["Summary", "Status", "Validity", "Type of Request", "Priority_Short",
                     "Creator", "R&D Engineer", "Creation Date", "Resolution Date"]:
        if fallback in export_df.columns and fallback not in default_export_cols:
            default_export_cols.append(fallback)

    available_export_cols = [c for c in export_df.columns if c not in HIDDEN_COLUMNS]
    chosen_export_cols = st.multiselect(
        "Columns to include in export",
        available_export_cols,
        default=[c for c in default_export_cols if c in available_export_cols],
        format_func=col_display,
        key="export_cols_picker",
    )
    if not chosen_export_cols:
        chosen_export_cols = available_export_cols

    # Preview
    st.dataframe(export_df[chosen_export_cols], use_container_width=True, height=380)

    # Download button — Excel only (management prefers it).
    try:
        st.download_button(
            "⬇ Download Excel",
            make_xlsx_bytes(export_df[chosen_export_cols].drop_duplicates(), sheet_name="Filtered_Issues"),
            file_name="custom_filtered_issues.xlsx",
            mime=XLSX_MIME,
            key="dl_custom_xlsx",
        )
    except Exception as e:
        st.error(f"Excel export failed: {e}")

# -- TAB 5: How we calculate --------------------------------------------------
with tab5:
    # Section 5 is documentation, not visualization. Each chart on the dashboard
    # gets a card explaining its source columns, the math, and a worked example
    # using current data so numbers feel familiar to leadership.

    st.markdown(
        f'<div style="font-family:Inter,sans-serif;font-size:0.95rem;line-height:1.65;'
        f'color:{SUBTEXT};max-width:780px;margin:8px 0 26px 0;">'
        f'A reference for how each chart and KPI on this dashboard is computed. '
        f'Every number you see is derived directly from the columns in the source '
        f'spreadsheet — no hidden adjustments, no smoothing. Click any card to expand it.'
        f'</div>',
        unsafe_allow_html=True,
    )

    # Helper: build a styled card body. Uses small caps section labels with
    # dark text in a clean two-column-ish inline layout.
    def _calc_card(what: str, source: str, how: str, example: str, edge: str = ""):
        parts = [
            ("What it shows", what),
            ("Source columns", source),
            ("How it's computed", how),
            ("Worked example", example),
        ]
        if edge:
            parts.append(("Edge cases", edge))

        html = ""
        for label, body in parts:
            html += (
                f'<div style="margin:14px 0 0 0;">'
                f'<div style="font-family:Inter,sans-serif;font-size:0.7rem;font-weight:700;'
                f'letter-spacing:0.16em;text-transform:uppercase;color:{MUTED};margin-bottom:6px;">'
                f'{label}</div>'
                f'<div style="font-family:Inter,sans-serif;font-size:0.92rem;line-height:1.6;'
                f'color:{INK};">{body}</div>'
                f'</div>'
            )
        st.markdown(html, unsafe_allow_html=True)

    # --- KPI Strip ----------------------------------------------------------
    st.markdown(
        f'<div style="font-family:Inter,sans-serif;font-size:0.74rem;font-weight:700;'
        f'letter-spacing:0.16em;text-transform:uppercase;color:{INK};margin:24px 0 12px 0;">'
        f'KPI strip · top of every page</div>',
        unsafe_allow_html=True,
    )

    with st.expander("Total issues", expanded=False):
        _calc_card(
            what="The total number of issues in the current view (after period and sidebar filters).",
            source="Every row in the loaded quarter sheets (e.g. <code>Q1_2026</code>, <code>Q2_2026</code>).",
            how="A simple row count: <code>len(filtered_dataframe)</code>.",
            example="If the period is set to <b>Q2 2026</b> and no sidebar filters are applied, this counts all 302 rows in the Q2_2026 sheet.",
        )

    with st.expander("P1 critical", expanded=False):
        _calc_card(
            what="Number of issues whose priority is P1, the highest tier (mission-critical).",
            source="<code>Priority</code> column. Long descriptions like 'P1 - Relationship is at risk…' are normalized to short codes (P1, P2, P3, P4) inside the dashboard.",
            how="Counts rows where <code>Priority_Short == 'P1'</code>. The percentage shown is <code>P1_count ÷ total_issues × 100</code>.",
            example="If 89 of 302 Q2 issues are tagged P1, this card shows <b>89</b> with caption <b>29% of total</b>.",
        )

    with st.expander("Resolved", expanded=False):
        _calc_card(
            what="Number of issues whose status indicates they are no longer open.",
            source="<code>Status</code> column.",
            how="Counts rows where the Status text contains any of: <i>Resolved</i>, <i>Closed</i>, or <i>Done</i> (case-insensitive). "
                "Because the source spreadsheet is pre-filtered to issues resolved in the chosen quarter, this number is "
                "typically equal to the total issue count — i.e., 100% of total.",
            example="In Q2 2026, all 302 issues meet the resolution criteria, so this shows <b>302</b> with caption <b>100% of total</b>.",
        )

    with st.expander("Valid", expanded=False):
        _calc_card(
            what="Number of issues marked as a legitimate concern after triage.",
            source="<code>Validity</code> column.",
            how="Counts rows where Validity (trimmed and case-insensitive) equals <code>'valid'</code>. "
                "Percentage is <code>valid_count ÷ total_issues × 100</code>.",
            example="In Q2 2026, 265 of 302 issues are tagged Valid → <b>265</b> with caption <b>88% of total</b>.",
            edge="Rows where Validity is blank, NaN, or any value other than 'Valid' or 'Invalid' are not counted.",
        )

    with st.expander("Bugs", expanded=False):
        _calc_card(
            what="Number of issues classified as a Bug rather than a Request for Help.",
            source="<code>Type of Request</code> column.",
            how="Counts rows where <code>Type of Request</code> (trimmed and case-insensitive) equals <code>'bug'</code>. "
                "Note: this uses Type of Request directly, NOT the 'Correct / Incorrect classification? RFH/Bug' "
                "column (which only contains 'Correct' or 'Incorrect' and tells you whether the categorization was right).",
            example="In Q2 2026, 97 of 302 issues are tagged Bug → <b>97</b> with caption <b>32% of total</b>.",
        )

    # --- Snapshot tab -------------------------------------------------------
    st.markdown(
        f'<div style="font-family:Inter,sans-serif;font-size:0.74rem;font-weight:700;'
        f'letter-spacing:0.16em;text-transform:uppercase;color:{INK};margin:32px 0 12px 0;">'
        f'Snapshot tab</div>',
        unsafe_allow_html=True,
    )

    with st.expander("Validity distribution (donut)", expanded=False):
        _calc_card(
            what="Share of issues classified as Valid vs Invalid, with a callout above showing the percentage and counts.",
            source="<code>Validity</code> column.",
            how="Groups rows by their Validity value and counts each group. The two donut segments represent these counts. "
                "The percentage in the callout is <code>valid_count ÷ total × 100</code>.",
            example="Q2 2026 has 265 Valid + 37 Invalid = 302 total → callout reads <b>87.7% valid · 265 valid · 37 invalid · 302 total</b>.",
            edge="Blank Validity values would show as a <code>(blank)</code> segment, but in your data this column is always populated.",
        )

    with st.expander("Type of request (stacked bar with Valid/Invalid breakdown)", expanded=False):
        _calc_card(
            what="Total issues per request type (Bug, Request for Help), with each bar split into Valid + Invalid segments.",
            source="<code>Type of Request</code> and <code>Validity</code> columns.",
            how="Cross-tabulates the two columns: for each Type, counts how many were Valid and how many were Invalid. "
                "Stacks the two counts into a single horizontal bar. The grand total appears outside each bar.",
            example="Q2 2026 'Bug' = 97 issues, all 97 Valid → bar is fully emerald, total <b>97</b>. "
                    "'Request for Help' = 205 issues = 168 Valid + 37 Invalid → bar shows emerald (168) + terracotta (37), total <b>205</b>.",
            edge="If any rows have blank Validity, they would appear as a gray <code>(blank)</code> segment.",
        )

    with st.expander("Version distribution (bar chart)", expanded=False):
        _calc_card(
            what="The 15 most-affected product versions, ranked by number of issues.",
            source="<code>Affects Version</code> column. (NOT <code>Fixed Version</code> — that's a different field.)",
            how="Splits cells on commas/semicolons (since one issue can affect multiple versions), explodes them into individual rows, "
                "counts occurrences of each version, and keeps the top 15. The largest bar is highlighted in deeper emerald.",
            example="An issue with <code>Affects Version = '7.4.3, 8.9.2'</code> contributes <b>+1</b> to the 7.4.3 count AND <b>+1</b> to the 8.9.2 count. "
                    "In Q2 2026, version 8.6.4 leads with 79 issues.",
            edge="Rows with blank Affects Version (or literal 'N/A' / 'NA' strings) are counted in a <code>(blank)</code> bar so they aren't silently dropped. "
                 "If there are blanks but they don't make the top 15, they're added back in as a 16th bar.",
        )

    with st.expander("Root cause resolution breakdown (bar chart)", expanded=False):
        _calc_card(
            what="How issues were ultimately resolved, ranked from most common to least.",
            source="<code>Root Cause Resolution</code> column.",
            how="Counts unique values in the column. The most common root cause appears at the top and is highlighted in deeper rose.",
            example="If 95 issues have Root Cause Resolution = 'User Misunderstanding' and 60 have 'Code Defect', "
                    "the chart ranks 'User Misunderstanding' first with a <b>95</b> bar.",
            edge="Rows where Root Cause Resolution is blank are excluded from the chart entirely (since 'unresolved cause' is not a meaningful category).",
        )

    with st.expander("Component distribution (lollipop chart)", expanded=False):
        _calc_card(
            what="The 15 product components most frequently associated with issues, ranked.",
            source="<code>Component/s</code> column.",
            how="Splits comma-separated values (one issue can touch multiple components), explodes into rows, counts each component, "
                "keeps the top 15. Each component is shown as a connector line + a marker dot at its count value. The top component is "
                "highlighted in deeper indigo with a slightly larger dot.",
            example="An issue with <code>Component/s = 'Workflows, Channel - SMS'</code> adds <b>+1</b> to both Workflows AND Channel - SMS counts.",
            edge="Rows with blank Component/s are excluded.",
        )

    with st.expander("Volume trend (line chart)", expanded=False):
        _calc_card(
            what="How many issues were resolved each month within the chosen reporting period.",
            source="<code>Resolution Date</code> column.",
            how="Buckets rows by the month of their Resolution Date (e.g., 'Mar 2026'), counts rows per bucket, plots them as a smooth line "
                "with markers. The peak month gets a slightly larger marker in deeper steel blue.",
            example="If Q2 2026 has 58 issues resolved in March, 83 in April, and 21 in May, the line plots three points at those values.",
            edge="Rows with no Resolution Date are excluded. The most recent month may appear lower than reality if data was exported mid-month.",
        )

    # --- Quality tab --------------------------------------------------------
    st.markdown(
        f'<div style="font-family:Inter,sans-serif;font-size:0.74rem;font-weight:700;'
        f'letter-spacing:0.16em;text-transform:uppercase;color:{INK};margin:32px 0 12px 0;">'
        f'Quality tab</div>',
        unsafe_allow_html=True,
    )

    with st.expander("Resolution time trend (line chart)", expanded=False):
        _calc_card(
            what="The median (and mean) number of days from Creation Date to Resolution Date, plotted by resolution month. "
                 "Tells you whether the team is getting faster at resolving issues over time.",
            source="<code>Creation Date</code> and <code>Resolution Date</code> columns.",
            how="For each issue: <code>days_to_resolve = Resolution Date − Creation Date</code> (in days). "
                "Then groups issues by the month they were resolved, computes the median and mean of days_to_resolve within each group, "
                "and plots both lines. Median is the solid indigo line; mean is the lighter dotted honey line.",
            example="An issue created Jan 5 and resolved Jan 20 contributes <b>15 days</b> to the January 2026 bucket. "
                    "If January has 92 such issues, the chart plots their median (e.g., 32 days) as the data point for Jan 2026.",
            edge="Issues with negative day differences (Resolution Date earlier than Creation Date — likely data entry errors) are excluded. "
                 "Issues missing either date are also excluded. Median is preferred over mean because a few outliers (e.g., issues resolved after 500+ days) "
                 "would otherwise distort the trend.",
        )

    with st.expander("Classification accuracy (stacked bar)", expanded=False):
        _calc_card(
            what="For each Type of Request (Bug, Request for Help), how often the original classification was correct vs incorrect.",
            source="<code>Type of Request</code> and <code>Correct / Incorrect classification? RFH/Bug</code> columns.",
            how="Cross-tabulates the two columns. For each Type, counts rows where the classification check returned 'Correct' "
                "vs 'Incorrect'. Stacks the two counts into a horizontal bar.",
            example="If 144 issues have Type = 'Bug', and within those 89 are marked 'Correct' and 55 'Incorrect', "
                    "the Bug bar is split: <b>89 emerald + 55 terracotta = 144 total</b>.",
            edge="The 'Correct / Incorrect classification?' column should ALWAYS be either 'Correct' or 'Incorrect' — "
                 "any blank or other value would render as a gray '(blank)' segment.",
        )

    with st.expander("Scope for deflection (bar chart)", expanded=False):
        _calc_card(
            what="Categories of issues that could potentially have been deflected (resolved without engineering intervention), "
                 "ranked by frequency.",
            source="<code>Scope for Deflection</code> column.",
            how="Counts unique values in the column.",
            example="If 35 issues are tagged 'Documentation gap' and 22 are 'User training', "
                    "the chart shows two bars at 35 and 22.",
            edge="Rows with blank Scope for Deflection are excluded — this chart only reflects issues that have been actively tagged.",
        )

    # --- Team tab -----------------------------------------------------------
    st.markdown(
        f'<div style="font-family:Inter,sans-serif;font-size:0.74rem;font-weight:700;'
        f'letter-spacing:0.16em;text-transform:uppercase;color:{INK};margin:32px 0 12px 0;">'
        f'Team tab</div>',
        unsafe_allow_html=True,
    )

    with st.expander("Issues per R&D engineer (bar chart)", expanded=False):
        _calc_card(
            what="The top 15 R&D engineers by total issue volume.",
            source="<code>R&D Engineer</code> column.",
            how="Counts unique values in the column, takes the top 15. The engineer with the highest count is highlighted in deeper indigo.",
            example="If Engineer A has 45 issues and Engineer B has 32, A is on top with a <b>45</b> bar.",
            edge="Rows where R&D Engineer is blank or unassigned are excluded.",
        )

    with st.expander("Priority mix per engineer (stacked bar)", expanded=False):
        _calc_card(
            what="For the top 10 engineers, how their issue load breaks down across priority levels P1, P2, P3, P4.",
            source="<code>R&D Engineer</code> and <code>Priority_Short</code> (derived from <code>Priority</code>).",
            how="Cross-tabulates engineer × priority. Each engineer gets a stacked bar with one segment per priority. "
                "Colors: P1 = terracotta-deep (most urgent), P2 = honey, P3 = steel blue, P4 = emerald.",
            example="If Engineer A has 45 total: 12 P1 + 18 P2 + 10 P3 + 5 P4, the stack shows those four colored segments adding up to 45.",
        )

    with st.expander("Root cause export", expanded=False):
        _calc_card(
            what="Not a chart — a data export tool. Lets you filter all rows by their Root Cause Resolution value and download the matching subset as Excel.",
            source="<code>Root Cause Resolution</code> column for filtering; whichever ID column is detected (NEO ID / Issue Key) plus context columns for export.",
            how="Filters the current dataframe by the chosen Root Cause value, picks a small set of relevant columns "
                "(ID, Summary, Status, etc.), and offers it as a downloadable Excel file.",
            example="Choose 'Code Defect' from the dropdown → an Excel file with all rows tagged that root cause is generated.",
        )

    with st.expander("Issues created by — split toggle (stacked bar)", expanded=False):
        _calc_card(
            what="Top 12 creators by total volume, with each creator's bar split by either Validity (Valid/Invalid) or Type of Request (RFH/Bug). "
                 "A small toggle above the chart switches between the two views.",
            source="<code>Creator</code> + either <code>Validity</code> OR <code>Type of Request</code> depending on toggle.",
            how="Picks the top 12 creators by row count. Cross-tabulates Creator × selected dimension. "
                "Each creator's bar is a stack of the two segments. Grand total is annotated outside each bar.",
            example="In <b>Valid/Invalid</b> mode: Creator A has 72 issues = 65 Valid + 7 Invalid → stack shows 65 emerald + 7 terracotta. "
                    "In <b>RFH/Bug</b> mode: same Creator A has 72 issues = 62 RFH + 10 Bug → stack shows 62 steel blue + 10 terracotta.",
            edge="Rows with blank Creator are excluded.",
        )

    # --- Period dropdown ----------------------------------------------------
    st.markdown(
        f'<div style="font-family:Inter,sans-serif;font-size:0.74rem;font-weight:700;'
        f'letter-spacing:0.16em;text-transform:uppercase;color:{INK};margin:32px 0 12px 0;">'
        f'Reporting period dropdown</div>',
        unsafe_allow_html=True,
    )

    with st.expander("How periods are determined", expanded=False):
        _calc_card(
            what="The dropdown controls which subset of rows the entire dashboard uses.",
            source="The <b>sheet name</b> each row was loaded from — stored internally as a <code>Quarter Label</code> column.",
            how="Each quarter sheet (e.g. <code>Q1_2026</code>) is loaded and tagged with a display label (<code>Q1 2026</code>). "
                "When the dropdown changes, the dashboard filters rows by their tagged label. <code>All Data</code> applies no filter. "
                "<code>H1 FY26</code> = Q1 + Q2 of fiscal year 26. <code>H2 FY26</code> = Q3 + Q4.",
            example="Selecting <code>Q2 2026</code> shows only the 302 rows that came from the Q2_2026 sheet.",
            edge="The dashboard does NOT do its own date math to figure out which quarter a row belongs to — "
                 "it trusts the sheet name. So if a row is in the Q2_2026 sheet, it's treated as Q2 2026 regardless "
                 "of what its Resolution Date says. This is intentional: your upstream Python script already does the "
                 "quarter-bucketing, and we don't want two systems disagreeing.",
        )

    # Footer note
    st.markdown(
        f'<div style="font-family:Inter,sans-serif;font-size:0.82rem;line-height:1.6;'
        f'color:{MUTED};max-width:780px;margin:32px 0 12px 0;padding:14px 18px;'
        f'background:rgba(124,111,232,0.05);border-left:3px solid {CHART_INDIGO};border-radius:6px;">'
        f'<b style="color:{INK}">A note on filters:</b> the sidebar filters (Project, Priority, Status, Type of Request, Validity) '
        f'apply on top of the period dropdown. Every chart and KPI on the dashboard reflects the combined effect of '
        f'all active filters — so if you select Q2 2026 in the dropdown and tick only "P1" in the Priority sidebar, '
        f'every number on the page is computed from just those P1 issues in Q2.'
        f'</div>',
        unsafe_allow_html=True,
    )

# Scoped CSS for the How we calculate tab — make expanders feel premium
st.markdown(
    f"""
<style>
  /* The expander headers — clean cards with hover lift */
  [data-testid="stExpander"] {{
    border: 1px solid {BORDER} !important;
    border-radius: 10px !important;
    background: {SURFACE};
    margin-bottom: 8px !important;
    box-shadow: 0 1px 2px rgba(26, 26, 26, 0.03);
    transition: all 0.15s ease;
  }}
  [data-testid="stExpander"]:hover {{
    border-color: {CHART_INDIGO} !important;
    box-shadow: 0 4px 12px rgba(124, 111, 232, 0.10);
  }}
  [data-testid="stExpander"] summary {{
    font-family: 'Inter', sans-serif !important;
    font-size: 0.96rem !important;
    font-weight: 600 !important;
    color: {INK} !important;
    padding: 14px 18px !important;
  }}
  [data-testid="stExpander"] summary:hover {{
    color: {CHART_INDIGO_DEEP} !important;
  }}
  /* Body content padding */
  [data-testid="stExpander"] [data-testid="stExpanderDetails"] {{
    padding: 0 18px 16px 18px !important;
  }}
  /* Inline code styling inside the cards */
  [data-testid="stExpander"] code {{
    background: rgba(124, 111, 232, 0.08);
    color: {CHART_INDIGO_DEEP};
    padding: 1px 6px;
    border-radius: 4px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.86em;
  }}
</style>
""",
    unsafe_allow_html=True,
)