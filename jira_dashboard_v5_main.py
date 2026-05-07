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
DEFAULT_FILE = "JIRA_P2E_2026.xlsx"

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

  /* TABS — brick-style buttons with breathing room above the KPI strip */
  [data-testid="stTabs"] {{
    margin-top: 32px !important;
    border-bottom: 1px solid var(--border);
  }}
  [data-testid="stTabs"] [role="tablist"] {{
    gap: 8px;
  }}
  [data-testid="stTabs"] button {{
    font-family: 'Inter', sans-serif;
    font-size: 0.9rem;
    font-weight: 700;
    letter-spacing: 0.005em;
    color: var(--ink);
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 10px 18px !important;
    transition: all 0.15s ease;
  }}
  [data-testid="stTabs"] button:hover {{
    background: rgba(124, 111, 232, 0.08);
    border-color: var(--accent);
  }}
  [data-testid="stTabs"] button[aria-selected="true"] {{
    background: var(--ink);
    color: #FFFFFF !important;
    border-color: var(--ink);
  }}
  [data-testid="stTabs"] button[aria-selected="true"] p {{
    color: #FFFFFF !important;
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


# ─── JIRA_Created loader ────────────────────────────────────────────────────
# This sheet is a separate dataset from the quarter sheets (Q1_2026, Q2_2026).
# It contains every Jira CREATED in each quarter (not just resolved ones), with
# a Customer column. Used by the Snapshot tab to show top customers by Jira
# count. Schema is intentionally different from the quarter sheets:
#   Quarter, Summary, Issue Key, Customer, Support Ticket, Creation Date, Affects Version
def _parse_jira_created(source) -> pd.DataFrame:
    """Read the optional 'JIRA_Created' sheet. Returns an empty DataFrame if
    the file or sheet is absent — the dashboard degrades gracefully when this
    dataset isn't available."""
    rewind_source(source)
    name = source if isinstance(source, str) else getattr(source, "name", "")
    if not str(name).lower().endswith((".xlsx", ".xls")):
        return pd.DataFrame()
    try:
        xls = pd.ExcelFile(source, engine="openpyxl")
        if "JIRA_Created" not in xls.sheet_names:
            return pd.DataFrame()
        rewind_source(source)
        df = pd.read_excel(source, sheet_name="JIRA_Created", engine="openpyxl")
    except Exception:
        return pd.DataFrame()
    df.columns = df.columns.str.strip()
    return df


@st.cache_data(show_spinner=False, ttl=3600)
def _parse_jira_created_from_path(path: str, mtime: float) -> pd.DataFrame:
    """Cached entry for string paths."""
    return _parse_jira_created(path)


def load_jira_created(source) -> pd.DataFrame:
    """Public entry point for the JIRA_Created sheet. Same caching pattern
    as load_issue_data."""
    if isinstance(source, str):
        try:
            mtime = os.path.getmtime(source)
            return _parse_jira_created_from_path(source, mtime)
        except Exception:
            pass
    return _parse_jira_created(source)


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
    Uses openpyxl which is already in our requirements. Sheet names are
    sanitized — Excel limits them to 31 chars and forbids :\\/?*[].

    Excel files do NOT support timezone-aware datetimes. pandas raises
    `ValueError: Excel does not support datetimes with timezones` if any
    cell is tz-aware. Local exports often have naive datetimes, but the
    same source file read in a different pandas/python version (e.g. on
    Streamlit Cloud) may produce tz-aware datetimes. We strip tz info
    defensively before writing to make the export environment-agnostic.
    """
    from io import BytesIO
    safe_sheet = re.sub(r"[:\\/?*\[\]]", "_", str(sheet_name))[:31] or "Issues"

    # Defensive copy + strip timezones from any datetime column.
    # `tz_localize(None)` drops the timezone but keeps the wall-clock value,
    # which is what users expect to see in Excel.
    safe_frame = frame.copy()
    for col in safe_frame.columns:
        s = safe_frame[col]
        # Standard tz-aware datetime64 columns. Use isinstance check rather
        # than the deprecated `is_datetime64tz_dtype` helper.
        if isinstance(s.dtype, pd.DatetimeTZDtype):
            safe_frame[col] = s.dt.tz_localize(None)
        # Object columns may contain mixed Python datetime objects with tzinfo
        elif s.dtype == "object":
            try:
                # Cheap probe: check first non-null value
                first_val = s.dropna().head(1)
                if len(first_val) and getattr(first_val.iloc[0], "tzinfo", None) is not None:
                    safe_frame[col] = s.apply(
                        lambda v: v.replace(tzinfo=None) if hasattr(v, "tzinfo") and v is not None and getattr(v, "tzinfo", None) is not None else v
                    )
            except Exception:
                # If anything goes sideways probing the column, leave it alone —
                # the writer will surface a clearer error than we'd raise here
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
    # Companion dataset: every Jira CREATED in each quarter (not just resolved).
    # Has a Customer column we use for the "Top customers" chart on Snapshot.
    # Loads independently of the main quarter sheets — failure to find this
    # sheet is non-fatal because we degrade gracefully.
    df_jira_created = load_jira_created(data_file)
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
                f'font-weight:600;color:{INK};line-height:1.1;">{cpu_pct:.0f}<span style="font-size:0.7rem;color:{MUTED};">%</span></div>'
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
valid_count = int((df["Validity"].astype(str).str.strip().str.lower() == "valid").sum()) if "Validity" in df.columns else 0
# Bug count comes from the Type of Request column ("Bug" / "Request for Help"),
# NOT from the RFH/Bug classification correctness column.
bug_count = int(df["Type of Request"].astype(str).str.strip().str.lower().eq("bug").sum()) if "Type of Request" in df.columns else 0
# RFH count uses the same Type of Request column. Match canonical "Request for Help"
# (already normalized at load_issue_data), with a forgiving lowercase strip just
# in case a future export reintroduces a variant.
rfh_count = int(df["Type of Request"].astype(str).str.strip().str.lower().eq("request for help").sum()) if "Type of Request" in df.columns else 0

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

# 5-tile KPI strip — Total → P1 critical → Valid → Bugs → RFH.
# We deliberately don't show a "Resolved" tile because the underlying source
# data is already filtered to resolved Jiras only. Showing 100% Resolved would
# be redundant and risk implying these are open-vs-closed counts.
kpi_cols = st.columns(5)
with kpi_cols[0]:
    st.metric("Total issues", f"{total_issues:,}")
with kpi_cols[1]:
    st.metric("P1 critical", f"{p1_count:,}")
    st.markdown(kpi_caption(safe_pct(p1_count, total_issues)), unsafe_allow_html=True)
with kpi_cols[2]:
    st.metric("Valid", f"{valid_count:,}")
    st.markdown(kpi_caption(safe_pct(valid_count, total_issues)), unsafe_allow_html=True)
with kpi_cols[3]:
    st.metric("Bugs", f"{bug_count:,}")
    st.markdown(kpi_caption(safe_pct(bug_count, total_issues)), unsafe_allow_html=True)
with kpi_cols[4]:
    st.metric("RFH", f"{rfh_count:,}")
    st.markdown(kpi_caption(safe_pct(rfh_count, total_issues)), unsafe_allow_html=True)

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
                f'<b style="color:{INK}">{valid_pct:.0f}%</b> valid &nbsp;·&nbsp; '
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
                texttemplate="%{label}<br>%{percent:.0%}",
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

    # ─── Top customers (new section, full-width row) ─────────────────────
    # Uses the JIRA_Created sheet (separate dataset from the quarter sheets).
    # That sheet contains every Jira created in the period along with which
    # Customer raised it. Question this answers: "Which customers are raising
    # the most P2E Jiras?" — useful for account prioritization conversations.
    if df_jira_created is not None and not df_jira_created.empty and "Customer" in df_jira_created.columns:
        section("Top customers by JIRA volume")

        # Apply the same period filter to the customer dataset so the chart
        # matches the rest of the page. JIRA_Created has a Quarter column
        # (Q1_2026 / Q2_2026 / etc.) that's structurally identical to the
        # Source Sheet column on the issue dataset — we map it to the same
        # display label and reuse apply_reporting_period.
        df_jc = df_jira_created.copy()
        if "Quarter" in df_jc.columns:
            df_jc["Source Sheet"] = df_jc["Quarter"].astype(str).str.strip()
            df_jc["Quarter Label"] = df_jc["Source Sheet"].map(sheet_to_quarter_label)

        try:
            df_jc_period = apply_reporting_period(df_jc, reporting_period, period_map)
        except Exception:
            # If for any reason the period filter doesn't match the JIRA_Created
            # dataset, fall back to showing all of it rather than crashing.
            df_jc_period = df_jc

        if df_jc_period.empty:
            st.caption("No customer data available for the selected period.")
        else:
            # Top 12 customers by Jira count
            customer_counts = (
                df_jc_period["Customer"]
                .dropna()
                .astype(str)
                .str.strip()
                .replace("", pd.NA)
                .dropna()
                .value_counts()
                .head(12)
                .reset_index()
            )
            customer_counts.columns = ["Customer", "Count"]

            # Horizontal bar chart in the indigo gradient — matches the
            # Engineer chart on the Team tab so they read as siblings.
            fig = go.Figure()
            counts_list = customer_counts["Count"].tolist()
            colors = bar_palette(counts_list, CHART_PLUM, CHART_PLUM_DEEP)

            fig.add_trace(go.Bar(
                x=counts_list,
                y=customer_counts["Customer"].tolist(),
                orientation="h",
                marker=dict(color=colors, line=dict(width=0)),
                text=[f"<b>{v}</b>" for v in counts_list],
                textposition="outside",
                textfont=dict(color=INK, size=12, family="Inter"),
                hovertemplate="<b>%{y}</b><br>%{x} jiras created<extra></extra>",
                cliponaxis=False,
            ))
            fig.update_layout(
                yaxis=dict(autorange="reversed"),
                xaxis_title="",
                yaxis_title="",
                height=420,
            )
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="chart_07_top_customers")

            st.caption(
                f"Showing top {len(customer_counts)} customers · "
                f"{int(df_jc_period['Customer'].notna().sum()):,} total Jiras created in this period · "
                f"data from the JIRA_Created sheet (separate from the resolved-issues view above)."
            )

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

        # ─── Advanced analysis gate ─────────────────────────────────
        # Leadership asked us to keep the default view simple — just total
        # issue count per creator. The deeper Validity / RFH-Bug breakdown
        # is still available behind an "Advanced analysis" checkbox so it's
        # one click away when needed but doesn't dominate the screen.
        # When the checkbox is OFF, the split toggle is hidden and the chart
        # falls back to the simplest mode ("Total only" if nothing else fits,
        # otherwise the first available split mode is irrelevant because the
        # branch below uses split_mode = "Total only" anyway).
        adv_container = st.container(key="creator_adv_toggle")
        with adv_container:
            show_advanced = st.checkbox(
                "Advanced analysis",
                value=False,
                key="creator_show_advanced",
                help="Reveals a deeper breakdown of issues by Validity (Valid/Invalid) or by Type of Request (RFH/Bug).",
            )

        if show_advanced and split_options != ["Total only"]:
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
        else:
            # Default view — no split, just total count per creator.
            split_mode = "Total only"

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

  /* Advanced analysis checkbox — sits above the split toggle, matches the
     Aurora visual language. Compact, indigo-accented, low chrome. */
  .st-key-creator_adv_toggle {{
    margin: 4px 0 6px 0;
  }}
  .st-key-creator_adv_toggle label {{
    font-family: 'Inter', sans-serif !important;
    font-size: 0.82rem !important;
    font-weight: 600 !important;
    color: {SUBTEXT} !important;
    cursor: pointer !important;
  }}
  .st-key-creator_adv_toggle label p,
  .st-key-creator_adv_toggle label span,
  .st-key-creator_adv_toggle label div {{
    color: {SUBTEXT} !important;
    font-weight: 600 !important;
  }}
  .st-key-creator_adv_toggle label:hover,
  .st-key-creator_adv_toggle label:hover p,
  .st-key-creator_adv_toggle label:hover span {{
    color: {INK} !important;
  }}
  /* Indigo tint for the checkbox box itself when checked */
  .st-key-creator_adv_toggle [data-baseweb="checkbox"] [data-checked="true"] {{
    background-color: {CHART_INDIGO_DEEP} !important;
    border-color: {CHART_INDIGO_DEEP} !important;
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
    # Single tool — Custom export builder. (Filtered issues table was removed
    # in v9.2 because leadership wanted a single, focused export tool.)
    # Above the builder, a small "How to use" card explains the workflow so
    # first-time viewers know what they're looking at.

    st.markdown(
        f'<div style="background:{SURFACE};border:1px solid {BORDER};border-radius:12px;'
        f'padding:18px 22px;margin:4px 0 24px 0;box-shadow:{SHADOW};">'
        f'<div style="font-family:Inter,sans-serif;font-size:0.74rem;font-weight:700;'
        f'letter-spacing:0.16em;text-transform:uppercase;color:{CHART_INDIGO_DEEP};margin-bottom:10px;">'
        f'How to use the export builder'
        f'</div>'
        f'<div style="font-family:Inter,sans-serif;font-size:0.9rem;color:{INK};line-height:1.65;">'
        f'<b>1.</b> Click <b>+ Add condition</b> to add a filter row. Choose a column, an operator (equals, contains, etc.), and a value. '
        f'Repeat for as many conditions as you need.<br>'
        f'<b>2.</b> Use the <b>Match</b> toggle to control how conditions combine — '
        f'<b style="color:{CHART_INDIGO_DEEP};">All conditions (AND)</b> requires every condition to match a row, '
        f'<b style="color:{CHART_TERRACOTTA_DEEP};">Any condition (OR)</b> matches rows that satisfy at least one.<br>'
        f'<b>3.</b> Pick which columns appear in the export from the column selector below the conditions.<br>'
        f'<b>4.</b> Click <b>Download Excel</b> to save your custom slice as a spreadsheet.'
        f'</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ─── Custom export builder ────────────────────────────────────────────────
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
        # WIP — the AI-generated root cause column is currently work-in-progress.
        # We're keeping it out of the dashboard UI until the team validates the
        # output quality. The column still loads from the source file so it's
        # available for analysis behind the scenes.
        "Root Cause Resolution AI",
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
    # Documentation tab. Reorganized in v9.2 from a flat list of expanders
    # into a card-per-tab layout. Each card contains a row of "brick"
    # buttons (one per chart/KPI). Clicking a brick reveals its
    # documentation inline within that same card while the other bricks
    # stay visible — so the reader can browse without losing context.
    #
    # State model: one session_state key per card ("calc_active_<group>").
    # Holds either None (nothing selected) or the slug of the selected
    # brick. Buttons rerun the page on click; we render the body inline
    # below the brick row when state is non-None.

    st.markdown(
        f'<div style="font-family:Inter,sans-serif;font-size:0.95rem;line-height:1.65;'
        f'color:{SUBTEXT};max-width:780px;margin:8px 0 26px 0;">'
        f'A reference for how each chart and KPI on this dashboard is computed. '
        f'Every number you see is derived directly from the columns in the source '
        f'spreadsheet — no hidden adjustments, no smoothing. Click any brick to read its details.'
        f'</div>',
        unsafe_allow_html=True,
    )

    # Helper: render the body of a calc card. Same layout as before — labelled
    # sections (What it shows / Source / How / Worked example / Edge cases)
    # but rendered inline rather than inside an expander.
    def _calc_body(what: str, source: str, how: str, example: str, edge: str = ""):
        parts = [
            ("What it shows", what),
            ("Source columns", source),
            ("How it's computed", how),
            ("Worked example", example),
        ]
        if edge:
            parts.append(("Edge cases", edge))

        html = (
            f'<div style="background:rgba(124,111,232,0.04);border:1px solid {BORDER};'
            f'border-radius:10px;padding:18px 22px;margin:14px 0 4px 0;">'
        )
        for label, body in parts:
            html += (
                f'<div style="margin:10px 0 0 0;">'
                f'<div style="font-family:Inter,sans-serif;font-size:0.7rem;font-weight:700;'
                f'letter-spacing:0.16em;text-transform:uppercase;color:{MUTED};margin-bottom:6px;">'
                f'{label}</div>'
                f'<div style="font-family:Inter,sans-serif;font-size:0.92rem;line-height:1.6;'
                f'color:{INK};">{body}</div>'
                f'</div>'
            )
        html += '</div>'
        st.markdown(html, unsafe_allow_html=True)

    # Helper: render one card. `group_key` namespaces session_state.
    # `bricks` is a list of (slug, label, body_dict) where body_dict is the
    # dict of args passed to _calc_body.
    def _calc_card(card_title: str, group_key: str, bricks: list):
        state_key = f"calc_active_{group_key}"
        if state_key not in st.session_state:
            st.session_state[state_key] = None

        # Card outer container — parchment-on-white surface
        st.markdown(
            f'<div style="font-family:Inter,sans-serif;font-size:0.74rem;font-weight:700;'
            f'letter-spacing:0.16em;text-transform:uppercase;color:{INK};margin:24px 0 10px 0;">'
            f'{card_title}</div>',
            unsafe_allow_html=True,
        )

        # Use a Streamlit container with a key so we can scope the brick CSS
        with st.container(key=f"calc_card_{group_key}"):
            # Render bricks as a row of buttons. We chunk them into rows of 3
            # for desktop readability — Streamlit's st.columns handles the wrap.
            chunk_size = 3
            for i in range(0, len(bricks), chunk_size):
                row = bricks[i:i + chunk_size]
                cols = st.columns(chunk_size)
                for j, (slug, label, _) in enumerate(row):
                    with cols[j]:
                        is_active = st.session_state[state_key] == slug
                        # Use button type="primary" to mark active brick;
                        # secondary for inactive. Aurora CSS below restyles both.
                        clicked = st.button(
                            label,
                            key=f"brickbtn_{group_key}_{slug}",
                            use_container_width=True,
                            type="primary" if is_active else "secondary",
                        )
                        if clicked:
                            # Toggle: clicking active brick collapses it
                            st.session_state[state_key] = None if is_active else slug
                            st.rerun()

            # Render the body of the active brick inline below the row
            active = st.session_state[state_key]
            if active is not None:
                # Find the matching brick body
                for slug, _label, body in bricks:
                    if slug == active:
                        _calc_body(**body)
                        break

    # ─── Card 1: KPI strip ───────────────────────────────────────────────
    _calc_card(
        "KPI strip · top of every page",
        "kpi",
        [
            ("total", "Total issues", dict(
                what="The total number of issues in the current view (after period and sidebar filters).",
                source="Every row in the loaded quarter sheets (e.g. <code>Q1_2026</code>, <code>Q2_2026</code>).",
                how="A simple row count: <code>len(filtered_dataframe)</code>.",
                example="If the period is set to <b>Q2 2026</b> and no sidebar filters are applied, this counts all 162 rows in the Q2_2026 sheet.",
            )),
            ("p1", "P1 critical", dict(
                what="Number of issues whose priority is P1, the highest tier (mission-critical).",
                source="<code>Priority</code> column. Long descriptions like 'P1 - Relationship is at risk…' are normalized to short codes (P1, P2, P3, P4) inside the dashboard.",
                how="Counts rows where <code>Priority_Short == 'P1'</code>. The percentage shown is <code>P1_count ÷ total_issues × 100</code>.",
                example="If 60 of 162 Q2 issues are tagged P1, this card shows <b>60</b> with caption <b>37% of total</b>.",
            )),
            ("valid", "Valid", dict(
                what="Number of issues marked as a legitimate concern after triage.",
                source="<code>Validity</code> column.",
                how="Counts rows where Validity (trimmed and case-insensitive) equals <code>'valid'</code>. "
                    "Percentage is <code>valid_count ÷ total_issues × 100</code>.",
                example="In Q2 2026, 139 of 162 issues are tagged Valid → <b>139</b> with caption <b>86% of total</b>.",
                edge="Rows where Validity is blank, NaN, or any value other than 'Valid' or 'Invalid' are not counted.",
            )),
            ("bugs", "Bugs", dict(
                what="Number of issues classified as a Bug rather than a Request for Help.",
                source="<code>Type of Request</code> column.",
                how="Counts rows where <code>Type of Request</code> (trimmed and case-insensitive) equals <code>'bug'</code>. "
                    "Note: this uses Type of Request directly, NOT the 'Correct / Incorrect classification? RFH/Bug' "
                    "column (which only contains 'Correct' or 'Incorrect' and tells you whether the categorization was right).",
                example="In Q2 2026, 50 of 162 issues are tagged Bug → <b>50</b> with caption <b>31% of total</b>.",
            )),
            ("rfh", "RFH", dict(
                what="Number of issues classified as a Request for Help (i.e., not a Bug).",
                source="<code>Type of Request</code> column.",
                how="Counts rows where <code>Type of Request</code> (trimmed and case-insensitive) equals <code>'request for help'</code>. "
                    "Together with the Bugs KPI, these two categories cover the full population of typed issues.",
                example="In Q2 2026, 112 of 162 issues are tagged Request for Help → <b>112</b> with caption <b>69% of total</b>.",
            )),
        ],
    )

    # ─── Card 2: Snapshot ───────────────────────────────────────────────
    _calc_card(
        "Snapshot tab",
        "snapshot",
        [
            ("validity", "Validity distribution (donut)", dict(
                what="Two-slice donut showing the share of Valid vs Invalid issues.",
                source="<code>Validity</code> column.",
                how="Counts each Validity value, then renders as a donut. The percentage shown on each slice is rounded to whole numbers (e.g. 88% rather than 87.7%).",
                example="265 Valid + 37 Invalid → 88% Valid slice (emerald) and 12% Invalid slice (terracotta).",
                edge="Validity values other than 'Valid' or 'Invalid' (blanks, NaN) are excluded from the donut.",
            )),
            ("type", "Type of request (stacked bar)", dict(
                what="Total issues per request type (Bug, Request for Help), with each bar split into Valid + Invalid segments.",
                source="<code>Type of Request</code> and <code>Validity</code> columns.",
                how="Cross-tabulates Type × Validity, then plots a stacked bar with Type on the x-axis. Stack order: Valid first (bottom), then Invalid (top). The y-axis is total count of issues for that type.",
                example="Bug: 60 Valid + 37 Invalid → emerald 60 stacked under terracotta 37 = total bar height 97.",
                edge="Rows with blank Type of Request are excluded.",
            )),
            ("version", "Version distribution (bar chart)", dict(
                what="Top affected versions ordered by issue count.",
                source="<code>Affects Version</code> column.",
                how="Counts how many issues mention each version, then keeps the top 8 by count. Numeric versioning is sorted naturally (so 8.6.4 > 8.6.10 lexically would be wrong — we sort by count instead).",
                example="If 8.6.4 has 76 issues, 8.9.2 has 38, and so on, the chart shows them ordered by count.",
                edge="A row can list multiple versions in one cell separated by commas — each one is counted independently.",
            )),
            ("rootcause", "Root cause resolution breakdown (bar chart)", dict(
                what="Distribution of root cause categories across all issues.",
                source="<code>Root Cause Resolution</code> column.",
                how="Counts each unique value, sorts by frequency, displays as a horizontal bar chart for readability of long category names.",
                example="If 'Configuration issue' appears 45 times and 'Product defect' appears 33 times, those become the top two bars.",
                edge="Rows with blank Root Cause Resolution are excluded.",
            )),
            ("component", "Component distribution (lollipop chart)", dict(
                what="Top components affected by issues, displayed as a Bloomberg-style lollipop chart for visual distinction.",
                source="<code>Component/s</code> column.",
                how="Splits cells with multiple components (separated by commas, semicolons, or newlines), counts each unique component, keeps the top 12.",
                example="Workflow - Campaigns: 35 issues, Channel - Email: 29 issues, Reporting: 26 issues, etc.",
                edge="A row listing 'A, B, C' contributes to all three components' counts. This avoids hiding components that travel together.",
            )),
            ("volume", "Volume trend (line chart)", dict(
                what="Issue creation volume over time, by month.",
                source="<code>Resolution Date</code> column.",
                how="Groups issues by month of resolution, counts per month, plots as a line chart with steel-blue area fill below.",
                example="Dec 2025: 25 issues, Jan 2026: 32 issues, Feb 2026: 28 issues — three points on the line.",
                edge="Rows with missing Resolution Date are excluded.",
            )),
            ("customers", "Top customers by JIRA volume (bar chart)", dict(
                what="Top 12 customers ranked by how many P2E Jiras they raised in the selected period.",
                source="<code>JIRA_Created</code> sheet (a separate dataset from the quarter sheets). Uses the <code>Customer</code> column.",
                how="Counts Jiras per Customer, sorts descending, keeps the top 12. The chart respects the same reporting period as the rest of the dashboard — selecting Q2 2026 shows top customers for Q2 only.",
                example="In Q2 2026: MICROSOFT - MSCOM (9), H&amp;M (8), MASHREQ BANK (7), KASIKORNBANK PUBLIC COMPANY (6).",
                edge="Customer entries are taken as-is from the source — different casings or formats (e.g. 'Microsoft - MSCOM' vs 'MICROSOFT - MSCOM') count as distinct customers because that's how the data team labels them upstream.",
            )),
        ],
    )

    # ─── Card 3: Quality ────────────────────────────────────────────────
    _calc_card(
        "Quality tab",
        "quality",
        [
            ("restime", "Resolution time trend (line chart)", dict(
                what="Median and mean resolution time per month, in days.",
                source="<code>Creation Date</code> and <code>Resolution Date</code> columns.",
                how="For each issue: <code>days_to_resolve = Resolution Date - Creation Date</code>. "
                    "Then groups by Resolution Month, computes both median and mean per month, plots two lines.",
                example="Dec 2025: median 42 days, mean 50 days. May 2026: median 10 days, mean 14 days. "
                    "Two trend lines descending from left to right.",
                edge="Issues with missing Creation Date or Resolution Date are excluded. "
                    "Negative or zero day counts (impossible by data) are clipped to 0 to avoid misleading the chart.",
            )),
            ("classacc", "Classification accuracy (stacked bar)", dict(
                what="Per Type of Request (Bug, RFH), shows what fraction were classified Correctly vs Incorrectly.",
                source="<code>Type of Request</code> and <code>Correct / Incorrect classification? RFH/Bug</code> columns.",
                how="Cross-tabulates Type × Correct/Incorrect, then plots stacked bar. Each bar's full height is the type's total; the green segment is Correct, red is Incorrect.",
                example="Bug: 60 Correct + 37 Incorrect = 62% accuracy. RFH: 183 Correct + 22 Incorrect = 89% accuracy.",
                edge="Rows where the classification cell is blank or has any value other than 'Correct'/'Incorrect' are excluded from this chart.",
            )),
            ("scope", "Scope for deflection (bar chart)", dict(
                what="Counts of issues categorized as deflectable (Yes), not deflectable (No), or partially (Yes/No).",
                source="<code>Scope for Deflection</code> column.",
                how="Direct value counts of the Scope for Deflection column, plotted as a single horizontal bar chart.",
                example="No: 223, Yes/No: 43, Yes: 36 — three bars in descending order.",
                edge="Blank values are excluded.",
            )),
        ],
    )

    # ─── Card 4: Team ────────────────────────────────────────────────
    _calc_card(
        "Team tab",
        "team",
        [
            ("rdeng", "Issues per R&D engineer (bar chart)", dict(
                what="Top 12 R&D engineers by number of assigned issues.",
                source="<code>R&D Engineer</code> column.",
                how="Counts how many issues each R&D engineer is assigned to, sorts descending, keeps the top 12.",
                example="Milos Manic: 23 issues, Camilo Medina: 15, Bastien Armand: 13.",
                edge="Rows with blank R&D Engineer are excluded.",
            )),
            ("primix", "Priority mix per engineer (stacked bar)", dict(
                what="For each engineer, the breakdown of their issues by Priority (P1/P2/P3/P4).",
                source="<code>R&D Engineer</code> and <code>Priority</code> columns.",
                how="Cross-tabulates engineer × priority short code, plots stacked bar. Bar lengths are total issues; segments show P1 (terracotta), P2, P3, P4 in indigo gradient.",
                example="An engineer with 23 issues = 8 P1 + 10 P2 + 4 P3 + 1 P4.",
                edge="Engineers shown are limited to the top 12 by total assigned issues (same as the bar chart above).",
            )),
            ("createdby", "Issues created by — Advanced toggle", dict(
                what="Top 12 creators by number of issues raised. By default shows total counts only. The 'Advanced analysis' checkbox above the chart reveals a Validity (Valid/Invalid) or RFH/Bug breakdown.",
                source="<code>Creator</code> column. Optional split: <code>Validity</code> or <code>Type of Request</code>.",
                how="Counts how many issues each Creator raised, sorts descending, keeps the top 12. With Advanced enabled, splits each creator's bar by the chosen dimension. Default mode is total only — clean, no split.",
                example="In <b>Total only</b> (default): Creator A has 72 issues → single indigo bar. "
                    "In <b>Valid/Invalid</b> mode: 65 Valid + 7 Invalid → stack shows 65 emerald + 7 terracotta. "
                    "In <b>RFH/Bug</b> mode: 62 RFH + 10 Bug → stack shows 62 steel blue + 10 terracotta.",
                edge="Rows with blank Creator are excluded. Advanced analysis is hidden by default to keep the leadership view simple.",
            )),
        ],
    )

    # ─── Card 5: Data explorer ───────────────────────────────────────
    _calc_card(
        "Data explorer tab",
        "explorer",
        [
            ("builder", "Custom export builder", dict(
                what="A flexible filter tool: build any number of conditions, combine them with AND or OR, pick the columns you want, and download as Excel.",
                source="Every column in the dataset is available as a filter or as an export column.",
                how="Each condition is <code>column &lt;operator&gt; value</code>. Operators include equals, not equals, contains, does not contain, is one of, is not one of. The 'Match' toggle (All / Any) controls whether all conditions must be true (AND) or just at least one (OR).",
                example="Two conditions — <code>Priority_Short equals P1</code> AND <code>Validity equals Invalid</code> — return 8 rows. Switching Match to Any (OR) returns 130 rows.",
                edge="If a column has only blank or duplicate values for the operator chosen, the dropdown will gracefully empty out and no rows match.",
            )),
        ],
    )

    # ─── Card 6: Reporting period dropdown ───────────────────────────
    _calc_card(
        "Reporting period dropdown",
        "period",
        [
            ("howperiods", "How periods are determined", dict(
                what="The dropdown controls which subset of rows the entire dashboard uses.",
                source="The <b>sheet name</b> each row was loaded from — stored internally as a <code>Quarter Label</code> column.",
                how="Each quarter sheet (e.g. <code>Q1_2026</code>) is loaded and tagged with a display label (<code>Q1 2026</code>). "
                    "When the dropdown changes, the dashboard filters rows by their tagged label. <code>All Data</code> applies no filter. "
                    "<code>H1 FY26</code> = Q1 + Q2 of fiscal year 26. <code>H2 FY26</code> = Q3 + Q4.",
                example="Selecting <code>Q2 2026</code> shows only the 162 rows that came from the Q2_2026 sheet.",
                edge="The dashboard does NOT do its own date math to figure out which quarter a row belongs to — "
                     "it trusts the sheet name. So if a row is in the Q2_2026 sheet, it's treated as Q2 2026 regardless "
                     "of what its Resolution Date says. This is intentional: your upstream Python script already does the "
                     "quarter-bucketing, and we don't want two systems disagreeing.",
            )),
        ],
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

# Scoped CSS for the calc bricks — make buttons feel like solid bricks,
# active state in dark ink + white text matching the navigation tabs.
st.markdown(
    f"""
<style>
  /* Brick buttons inside the calc cards */
  div[class*="st-key-calc_card_"] .stButton > button {{
    font-family: 'Inter', sans-serif !important;
    font-size: 0.9rem !important;
    font-weight: 600 !important;
    color: {INK} !important;
    background: {SURFACE} !important;
    border: 1px solid {BORDER} !important;
    border-radius: 8px !important;
    padding: 12px 16px !important;
    text-align: left !important;
    justify-content: flex-start !important;
    transition: all 0.15s ease !important;
    height: auto !important;
    white-space: normal !important;
    line-height: 1.4 !important;
  }}
  /* Force ink color on all nested elements (Streamlit wraps button text in <p>) */
  div[class*="st-key-calc_card_"] .stButton > button p,
  div[class*="st-key-calc_card_"] .stButton > button span,
  div[class*="st-key-calc_card_"] .stButton > button div {{
    color: {INK} !important;
    font-weight: 600 !important;
  }}
  div[class*="st-key-calc_card_"] .stButton > button:hover {{
    background: rgba(124, 111, 232, 0.08) !important;
    border-color: {CHART_INDIGO} !important;
    color: {INK} !important;
  }}
  div[class*="st-key-calc_card_"] .stButton > button:hover p,
  div[class*="st-key-calc_card_"] .stButton > button:hover span,
  div[class*="st-key-calc_card_"] .stButton > button:hover div {{
    color: {INK} !important;
  }}
  /* Active brick — solid dark ink background, pure white text */
  div[class*="st-key-calc_card_"] .stButton > button[kind="primary"] {{
    background: {INK} !important;
    color: #FFFFFF !important;
    border-color: {INK} !important;
    box-shadow: 0 2px 6px rgba(26, 26, 26, 0.2);
  }}
  /* Defeat BaseWeb's nested-element color overrides — force white on every
     descendant of the active button so the text contrasts the dark background */
  div[class*="st-key-calc_card_"] .stButton > button[kind="primary"] p,
  div[class*="st-key-calc_card_"] .stButton > button[kind="primary"] span,
  div[class*="st-key-calc_card_"] .stButton > button[kind="primary"] div {{
    color: #FFFFFF !important;
    font-weight: 700 !important;
  }}
  div[class*="st-key-calc_card_"] .stButton > button[kind="primary"]:hover {{
    background: {INK} !important;
    color: #FFFFFF !important;
  }}
  div[class*="st-key-calc_card_"] .stButton > button[kind="primary"]:hover p,
  div[class*="st-key-calc_card_"] .stButton > button[kind="primary"]:hover span,
  div[class*="st-key-calc_card_"] .stButton > button[kind="primary"]:hover div {{
    color: #FFFFFF !important;
  }}
  /* Inline code styling inside the calc body */
  div[class*="st-key-calc_card_"] code {{
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