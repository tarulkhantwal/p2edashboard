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
DEFAULT_FILE = "JIRA_P2E_2026_label.xlsx"

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

# Read the user's current view choice early, BEFORE rendering the period
# selector. The view selector itself renders later in the page, but its
# session_state key persists across reruns, so by the second render we know
# which view is active and can hide the period dropdown when irrelevant.
# First render fallback: assume "Resolved Jiras" (the default).
_current_view_mode = st.session_state.get("dashboard_view_mode", "Resolved Jiras")
_hide_period_picker = str(_current_view_mode).startswith("Created")

# Minimalist period picker — just a clean compact selectbox.
# Constrained to 240px width via scoped CSS, with a small label above and
# a subtle indigo focus ring. No fancy widgets, no leaking panels — just
# a dropdown that does one thing well.
# Hidden entirely in the Created Jiras view — that view doesn't honor the
# period filter, so showing a non-functional dropdown would mislead users.
picker_container = st.container(key="period_picker")
with picker_container:
    if not _hide_period_picker:
        reporting_period = st.selectbox(
            "Reporting period",
            period_options,
            key="reporting_period",
        )
    else:
        # Keep the variable defined for any downstream references, even
        # though Created Jiras view ignores it. Use whatever's in session_state.
        reporting_period = st.session_state.reporting_period

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

# -- View mode selector ------------------------------------------------------
# Top-level switch between two distinct dashboard experiences:
#   • Main dashboard — resolution-date-based, has 5 tabs, period-driven.
#   • Reports — creation-date-based, single page, independent of period.
#
# We surface this as a segmented control directly under the period selector
# so leadership sees both controls at the top of the page, side-by-side
# in the visual flow but logically distinct. Defaults to Main dashboard.
view_picker_container = st.container(key="view_picker")
with view_picker_container:
    view_mode = st.radio(
        "View",
        ["Resolved Jiras", "Created Jiras"],
        index=0,
        horizontal=True,
        key="dashboard_view_mode",
        help="Switch between the resolution-based dashboard (default) and the creation-based report.",
        label_visibility="collapsed",
    )

# Style the view picker so the two options sit at opposite ends of the
# header row — Resolved Jiras on the left, Created Jiras on the right.
# Both are part of the same radio group; CSS just spreads them apart
# using justify-content: space-between on the radiogroup container.
#
# Position: the picker sits below the title block with normal spacing
# (no negative margin), so it never overlaps the description text. This
# prevents the "going behind the text" issue caused by the previous
# negative-margin layout.
st.markdown(
    f"""
<style>
  /* The container itself just provides vertical spacing — no flex tricks here */
  .st-key-view_picker {{
    margin: 16px 0 20px 0;
  }}
  /* The radio group inside is what we spread: Resolved on left, Created on right */
  .st-key-view_picker [data-testid="stHorizontalBlock"],
  .st-key-view_picker [role="radiogroup"] {{
    display: flex !important;
    flex-direction: row !important;
    justify-content: space-between !important;
    align-items: center !important;
    width: 100% !important;
    gap: 4px;
  }}
  /* Inactive options — muted text, transparent background, no border */
  .st-key-view_picker [role="radiogroup"] label {{
    font-family: 'Inter', sans-serif !important;
    font-size: 0.78rem !important;
    font-weight: 500 !important;
    color: {MUTED} !important;
    background: transparent !important;
    border: 1px solid transparent !important;
    border-radius: 6px !important;
    padding: 5px 12px !important;
    transition: all 0.15s ease !important;
    cursor: pointer !important;
  }}
  .st-key-view_picker [role="radiogroup"] label p,
  .st-key-view_picker [role="radiogroup"] label span {{
    color: {MUTED} !important;
    font-weight: 500 !important;
  }}
  .st-key-view_picker [role="radiogroup"] label:hover {{
    color: {INK} !important;
    background: rgba(124, 111, 232, 0.06) !important;
  }}
  .st-key-view_picker [role="radiogroup"] label:hover p,
  .st-key-view_picker [role="radiogroup"] label:hover span {{
    color: {INK} !important;
  }}
  /* Hide the radio circle */
  .st-key-view_picker [role="radiogroup"] label > div:first-child {{
    display: none !important;
  }}
  /* Active option — subtle indigo underline, not a heavy filled button */
  .st-key-view_picker [role="radiogroup"] label[data-checked="true"],
  .st-key-view_picker [role="radiogroup"] label:has(input:checked) {{
    background: rgba(124, 111, 232, 0.08) !important;
    color: {INK} !important;
    border-color: rgba(124, 111, 232, 0.25) !important;
    font-weight: 600 !important;
  }}
  .st-key-view_picker [role="radiogroup"] label[data-checked="true"] p,
  .st-key-view_picker [role="radiogroup"] label[data-checked="true"] span,
  .st-key-view_picker [role="radiogroup"] label:has(input:checked) p,
  .st-key-view_picker [role="radiogroup"] label:has(input:checked) span {{
    color: {INK} !important;
    font-weight: 600 !important;
  }}
</style>
""",
    unsafe_allow_html=True,
)

# Boolean for downstream logic — cleaner than string comparisons everywhere.
is_reports_view = view_mode.startswith("Created")

# ─── Created Jiras view (lens-aware top-level view) ─────────────────────
# When the user picks 'Created Jiras' from the view picker, we render this
# minimal block (KPI strip + context badge) and call st.stop() to skip the
# rest of the script. The customer-centric breakdowns leadership uses live
# in the Customers tab on the Resolved view — same dashboard, different lens.
#
# Note: the deeper Created-Jira charts (monthly trend, quarter-over-quarter
# comparison, top customers, top versions) were retired in v9.6 when the
# Customers tab replaced them with lens-aware versions. This block now just
# surfaces the 5 leadership-relevant KPIs at the top of the Created view.
if is_reports_view:
    if df_jira_created is None or df_jira_created.empty:
        st.warning(
            "JIRA_Created sheet not found in the source file. The Created Jiras view "
            "needs that sheet to compute creation-based metrics. Please ensure "
            "the latest Excel file is being used."
        )
    else:
        # Defensive copy + ensure dates parse correctly
        df_jc = df_jira_created.copy()
        df_jc["Creation Date"] = pd.to_datetime(df_jc["Creation Date"], errors="coerce")
        df_jc = df_jc[df_jc["Creation Date"].notna()].copy()

        # ─── Headline KPI strip (5 tiles) ────────────────────────────────
        # Director-relevant cuts: how much came in, how much is open, how
        # much closed, P1 volume, and P1 still-open (the leadership signal).
        #
        # IMPORTANT — "Open" definition on this view:
        # We exclude Deferred Jiras from "Open" because Deferred = intentionally
        # not being worked on. Counting them as "Open" inflates the backlog
        # signal that leadership cares about. This matches how Yash's pipeline
        # treats Deferred in the quarter sheets (also excluded). The same
        # definition is used by the "Open Jiras by age" table below so both
        # numbers stay consistent.
        CLOSED_OR_DEFERRED = ["Resolved", "Closed", "Deferred"]

        total_created = len(df_jc)
        open_count = int((~df_jc["Status"].isin(CLOSED_OR_DEFERRED)).sum())
        closed_count = int(df_jc["Status"].isin(["Resolved", "Closed"]).sum())
        p1_critical = int((df_jc["Priority"] == "P1").sum())
        p1_still_open = int(((df_jc["Priority"] == "P1") &
                             (~df_jc["Status"].isin(CLOSED_OR_DEFERRED))).sum())

        # Context badge — indigo to distinguish from the Resolved view's emerald
        st.markdown(
            f'<div style="display:inline-flex;align-items:center;gap:8px;'
            f'background:rgba(124,111,232,0.10);border:1px solid {CHART_INDIGO};'
            f'border-radius:999px;padding:5px 14px;margin:0 0 14px 0;">'
            f'<span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:{CHART_INDIGO_DEEP};"></span>'
            f'<span style="font-family:Inter,sans-serif;font-size:0.74rem;font-weight:700;'
            f'letter-spacing:0.08em;text-transform:uppercase;color:{CHART_INDIGO_DEEP};">'
            f'Showing created Jiras · all quarters'
            f'</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

        # 5-tile KPI strip
        rcols = st.columns(5)

        def _kpi_tile(col, label, value, sub, label_color=MUTED):
            col.markdown(
                f'<div style="background:{SURFACE};border:1px solid {BORDER};border-radius:14px;padding:22px 26px;box-shadow:{SHADOW};">'
                f'<div style="font-family:Inter,sans-serif;font-size:0.78rem;color:{label_color};letter-spacing:0.12em;text-transform:uppercase;font-weight:700;margin-bottom:6px;">{label}</div>'
                f'<div style="font-family:\'IBM Plex Mono\',monospace;font-size:2.4rem;font-weight:600;color:{INK};line-height:1.1;">{value:,}</div>'
                f'<div style="font-family:Inter,sans-serif;font-size:0.78rem;color:{SUBTEXT};margin-top:4px;">{sub}</div>'
                f'</div>',
                unsafe_allow_html=True,
            )

        _kpi_tile(rcols[0], "Total created", total_created, "across all quarters")
        _kpi_tile(rcols[1], "Open", open_count,
                  f"{int(open_count/total_created*100) if total_created else 0}% of total")
        _kpi_tile(rcols[2], "Closed", closed_count,
                  f"{int(closed_count/total_created*100) if total_created else 0}% of total")
        _kpi_tile(rcols[3], "P1 critical", p1_critical,
                  f"{int(p1_critical/total_created*100) if total_created else 0}% of total",
                  label_color=CHART_TERRACOTTA_DEEP)
        _kpi_tile(rcols[4], "P1 still open", p1_still_open,
                  f"{int(p1_still_open/p1_critical*100) if p1_critical else 0}% of P1s",
                  label_color=CHART_TERRACOTTA_DEEP)

        # Footnote explaining the Open definition — important because it
        # differs from a naive "Status not Resolved/Closed" count by
        # excluding Deferred. Deferred items are intentionally not being
        # worked on, so counting them as Open inflates the backlog signal.
        deferred_count = int((df_jc["Status"] == "Deferred").sum())
        if deferred_count > 0:
            st.markdown(
                f'<div style="font-family:Inter,sans-serif;font-size:0.78rem;color:{MUTED};'
                f'margin:10px 0 0 0;text-align:left;">'
                f'<i>Open = actively being worked on. Excludes Deferred ({deferred_count}) '
                f'and closed work.</i>'
                f'</div>',
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height:32px'></div>", unsafe_allow_html=True)

        # ─── Custom export builder (Created lens) ────────────────────────
        # Mirrors the Data Explorer's query builder but operates on
        # JIRA_Created data (open + closed Jiras). Supports filtering by
        # any column including the new Labels column from Yash's pipeline.
        # Includes a derived "Age (days)" column = (as-of date − Creation Date)
        # so users can build queries like "Age > 60 days" to surface stale
        # open work.
        #
        # Implementation note: this is a leaner build than the Resolved-lens
        # Data Explorer (no separate column selector, no how-to card, just
        # the working filter UI + export). Helpers are inlined here because
        # the full Data Explorer's helpers are nested inside `with tab4:`
        # which doesn't render on this lens.
        section("Custom export builder")

        # ─── Inline helpers (kept here because tab4's are out of scope) ──
        _DATE_OPS_CB = ["on", "on or after", "on or before", "between"]
        _LABEL_OPS_CB = ["has any of", "has none of"]
        _NUMERIC_OPS_CB = ["equals", "greater than", "less than", "at least", "at most", "between"]
        _TEXT_OPS_CB = ["is", "is not", "contains", "does not contain"]

        def _is_date_col(frame, col):
            return col in frame.columns and pd.api.types.is_datetime64_any_dtype(frame[col])

        def _is_numeric_col(frame, col):
            return (col in frame.columns
                    and pd.api.types.is_numeric_dtype(frame[col])
                    and not _is_date_col(frame, col))

        def _is_labels_col(col):
            return col.strip().lower() == "labels"

        def _parse_labels_cell(val):
            if pd.isna(val):
                return []
            s = str(val).strip()
            if not s or s == "[]":
                return []
            try:
                import ast as _ast
                parsed = _ast.literal_eval(s)
                if isinstance(parsed, list):
                    return [str(x).strip() for x in parsed if str(x).strip()]
                return []
            except Exception:
                return []

        def _build_label_opts(frame, col):
            if col not in frame.columns:
                return []
            counter = {}
            for v in frame[col].dropna():
                for lbl in _parse_labels_cell(v):
                    counter[lbl] = counter.get(lbl, 0) + 1
            if not counter:
                return []
            sorted_items = sorted(counter.items(), key=lambda kv: (-kv[1], kv[0].lower()))
            return [f"{lbl} ({cnt})" for lbl, cnt in sorted_items]

        def _strip_count(s):
            if not s:
                return ""
            m = re.search(r"\s*\(\d+\)\s*$", str(s))
            return s[:m.start()].strip() if m else str(s).strip()

        def _cb_mask(frame, col, op, val):
            """Mini condition_mask for the Created-lens builder."""
            all_true = pd.Series(True, index=frame.index)
            if col not in frame.columns or val is None:
                return all_true
            opl = (op or "").lower().strip()
            # Date branch
            if opl in ("on", "on or after", "on or before", "between") and _is_date_col(frame, col):
                ser_d = pd.to_datetime(frame[col], errors="coerce").dt.date
                try:
                    if opl == "between":
                        if not isinstance(val, (list, tuple)) or len(val) != 2:
                            return all_true
                        a, b = pd.to_datetime(val[0]).date(), pd.to_datetime(val[1]).date()
                        if a > b: a, b = b, a
                        return (ser_d >= a) & (ser_d <= b)
                    tgt = pd.to_datetime(val).date()
                    if opl == "on": return ser_d == tgt
                    if opl == "on or after": return ser_d >= tgt
                    if opl == "on or before": return ser_d <= tgt
                except Exception:
                    return all_true
                return all_true
            # Labels branch
            if opl in ("has any of", "has none of") and _is_labels_col(col):
                if not isinstance(val, (list, tuple)) or not val:
                    return all_true
                sel = {_strip_count(v).lower() for v in val if v}
                if not sel:
                    return all_true
                row_match = frame[col].map(lambda c: bool({l.lower() for l in _parse_labels_cell(c)} & sel))
                return row_match if opl == "has any of" else ~row_match
            # Numeric branch
            if opl in ("equals", "greater than", "less than", "at least", "at most", "between"):
                try:
                    ser_n = pd.to_numeric(frame[col], errors="coerce")
                    if opl == "between":
                        if not isinstance(val, (list, tuple)) or len(val) != 2:
                            return all_true
                        lo, hi = float(val[0]), float(val[1])
                        if lo > hi: lo, hi = hi, lo
                        return (ser_n >= lo) & (ser_n <= hi)
                    t = float(val)
                    if opl == "equals": return ser_n == t
                    if opl == "greater than": return ser_n > t
                    if opl == "less than": return ser_n < t
                    if opl == "at least": return ser_n >= t
                    if opl == "at most": return ser_n <= t
                except (ValueError, TypeError):
                    return all_true
                return all_true
            # Text branch (is/is not/contains)
            if val == "":
                return all_true
            ser_t = frame[col].astype(str).str.strip()
            val_t = str(val).strip()
            if not val_t:
                return all_true
            if opl == "is": return ser_t.str.lower() == val_t.lower()
            if opl == "is not": return ser_t.str.lower() != val_t.lower()
            if opl == "contains": return ser_t.str.contains(val_t, case=False, na=False)
            if opl == "does not contain": return ~ser_t.str.contains(val_t, case=False, na=False)
            return all_true

        # Build a working frame with the derived Age column
        cb_df = df_jc.copy()
        if "Resolution Date" in cb_df.columns:
            cb_df["Resolution Date"] = pd.to_datetime(cb_df["Resolution Date"], errors="coerce")
        # As-of date for Age: stable across reloads (latest activity in dataset)
        _asof_candidates = [
            cb_df["Resolution Date"].dropna().max() if "Resolution Date" in cb_df.columns else pd.NaT,
            cb_df["Creation Date"].dropna().max(),
            pd.Timestamp.today(),
        ]
        cb_asof = next((d for d in _asof_candidates if pd.notna(d)), pd.Timestamp.today())
        cb_asof = pd.Timestamp(cb_asof).normalize()
        # Age = (Resolution Date OR as-of) − Creation Date. For open Jiras,
        # this keeps ticking until they're resolved.
        if "Creation Date" in cb_df.columns:
            _end = cb_df["Resolution Date"].fillna(cb_asof) if "Resolution Date" in cb_df.columns else cb_asof
            cb_df["Age (days)"] = (_end - cb_df["Creation Date"]).dt.days

        # ─── Match toggle (AND / OR) ─────────────────────────────────
        cb_match_mode = st.radio(
            "Match",
            ["All conditions (AND)", "Any condition (OR)"],
            horizontal=True,
            key="cb_jc_match_mode",
            label_visibility="collapsed",
        )
        cb_is_or = cb_match_mode.startswith("Any")

        # Initialize a default filter row
        if "cb_jc_filters" not in st.session_state:
            st.session_state.cb_jc_filters = [{"column": "Priority", "op": "is", "value": "P1"}]

        # Action buttons
        bcols = st.columns([1, 1, 6])
        with bcols[0]:
            if st.button("+ Add condition", key="cb_jc_add"):
                st.session_state.cb_jc_filters.append({"column": None, "op": None, "value": None})
                st.rerun()
        with bcols[1]:
            if st.button("✕ Clear all", key="cb_jc_clear"):
                st.session_state.cb_jc_filters = [{"column": "Priority", "op": "is", "value": "P1"}]
                st.rerun()

        # Available columns on JIRA_Created (filters out unhelpful ones)
        cb_skip = {"Summary"}  # too long-form to filter by usefully
        cb_columns = [c for c in cb_df.columns if c not in cb_skip]

        # Helper to build distinct-value options for "is/is not" picker
        def _cb_options(frame, col):
            if col not in frame.columns:
                return []
            vals = frame[col].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist()
            return sorted(vals)

        # ─── Render filter rows ───────────────────────────────────────
        for cb_i, cb_flt in enumerate(st.session_state.cb_jc_filters):
            cb_row = st.columns([0.6, 2.2, 1.6, 3, 0.6])

            with cb_row[0]:
                connector = "Where" if cb_i == 0 else ("Or" if cb_is_or else "And")
                color = CHART_TERRACOTTA_DEEP if cb_is_or and cb_i > 0 else MUTED
                st.markdown(
                    f'<div style="font-family:Inter,sans-serif;font-size:0.85rem;font-weight:700;'
                    f'color:{color};padding:8px 0 0 0;text-align:right;">{connector}</div>',
                    unsafe_allow_html=True,
                )

            # Column picker
            with cb_row[1]:
                current_col = cb_flt.get("column") or cb_columns[0]
                if current_col not in cb_columns:
                    current_col = cb_columns[0]
                col_choice = st.selectbox(
                    "Column",
                    cb_columns,
                    index=cb_columns.index(current_col),
                    key=f"cb_jc_col_{cb_i}",
                    label_visibility="collapsed",
                )
                cb_flt["column"] = col_choice

            # Determine column type for operator/value picker
            _is_date = _is_date_col(cb_df, col_choice)
            _is_labels = _is_labels_col(col_choice)
            _is_numeric = _is_numeric_col(cb_df, col_choice) and not _is_date
            if _is_date:
                cb_ops = _DATE_OPS_CB
            elif _is_labels:
                cb_ops = _LABEL_OPS_CB
            elif _is_numeric:
                cb_ops = _NUMERIC_OPS_CB
            else:
                cb_ops = _TEXT_OPS_CB

            # Operator picker
            with cb_row[2]:
                cur_op = cb_flt.get("op") or cb_ops[0]
                if cur_op not in cb_ops:
                    cur_op = cb_ops[0]
                op_choice = st.selectbox(
                    "Operator",
                    cb_ops,
                    index=cb_ops.index(cur_op),
                    key=f"cb_jc_op_{cb_i}",
                    label_visibility="collapsed",
                )
                cb_flt["op"] = op_choice

            # Value picker — branches by column type
            with cb_row[3]:
                if _is_date:
                    dseries = pd.to_datetime(cb_df[col_choice], errors="coerce").dropna()
                    if not dseries.empty:
                        d_min, d_max = dseries.min().date(), dseries.max().date()
                    else:
                        d_min = d_max = pd.Timestamp.today().date()
                    prev = cb_flt.get("value")
                    if op_choice == "between":
                        if isinstance(prev, (list, tuple)) and len(prev) == 2:
                            try:
                                st_d = pd.to_datetime(prev[0]).date()
                                en_d = pd.to_datetime(prev[1]).date()
                            except Exception:
                                st_d, en_d = d_min, d_max
                        else:
                            st_d, en_d = d_min, d_max
                        sub = st.columns(2)
                        with sub[0]:
                            sd = st.date_input("From", value=st_d, min_value=d_min, max_value=d_max,
                                               key=f"cb_jc_dstart_{cb_i}", label_visibility="collapsed")
                        with sub[1]:
                            ed = st.date_input("To", value=en_d, min_value=d_min, max_value=d_max,
                                               key=f"cb_jc_dend_{cb_i}", label_visibility="collapsed")
                        cb_val = (sd, ed)
                    else:
                        try:
                            sd_def = pd.to_datetime(prev).date() if prev else d_max
                        except Exception:
                            sd_def = d_max
                        cb_val = st.date_input("Date", value=sd_def, min_value=d_min, max_value=d_max,
                                               key=f"cb_jc_date_{cb_i}", label_visibility="collapsed")
                elif _is_numeric:
                    nseries = pd.to_numeric(cb_df[col_choice], errors="coerce").dropna()
                    n_min = float(nseries.min()) if not nseries.empty else 0.0
                    n_max = float(nseries.max()) if not nseries.empty else 100.0
                    prev = cb_flt.get("value")
                    if op_choice == "between":
                        if isinstance(prev, (list, tuple)) and len(prev) == 2:
                            try:
                                ns_def, ne_def = float(prev[0]), float(prev[1])
                            except Exception:
                                ns_def, ne_def = n_min, n_max
                        else:
                            ns_def, ne_def = n_min, n_max
                        sub = st.columns(2)
                        with sub[0]:
                            ns = st.number_input("From", value=ns_def, step=1.0,
                                                 key=f"cb_jc_nstart_{cb_i}", label_visibility="collapsed")
                        with sub[1]:
                            ne = st.number_input("To", value=ne_def, step=1.0,
                                                 key=f"cb_jc_nend_{cb_i}", label_visibility="collapsed")
                        cb_val = (ns, ne)
                    else:
                        try:
                            single_n = float(prev) if isinstance(prev, (int, float)) else n_min
                        except Exception:
                            single_n = n_min
                        cb_val = st.number_input(
                            "Value", value=single_n, step=1.0,
                            key=f"cb_jc_num_{cb_i}", label_visibility="collapsed",
                            help=f"Range in data: {int(n_min)} to {int(n_max)}",
                        )
                elif _is_labels:
                    label_opts = _build_label_opts(cb_df, col_choice)
                    prev = cb_flt.get("value")
                    if isinstance(prev, (list, tuple)):
                        default_sel = [v for v in prev if v in label_opts]
                    else:
                        default_sel = []
                    cb_val = st.multiselect(
                        "Labels", label_opts, default=default_sel,
                        key=f"cb_jc_labels_{cb_i}", label_visibility="collapsed",
                        placeholder="Type to search labels…",
                    )
                elif op_choice in ("is", "is not"):
                    opts = _cb_options(cb_df, col_choice)
                    if not opts:
                        opts = ["(blank)"]
                    prev = cb_flt.get("value")
                    idx = opts.index(prev) if prev in opts else 0
                    cb_val = st.selectbox(
                        "Value", opts, index=idx,
                        key=f"cb_jc_val_{cb_i}", label_visibility="collapsed",
                    )
                else:
                    cb_val = st.text_input(
                        "Value", value=str(cb_flt.get("value") or ""),
                        placeholder="Type a substring…",
                        key=f"cb_jc_text_{cb_i}", label_visibility="collapsed",
                    )
                cb_flt["value"] = cb_val

            # Remove button
            with cb_row[4]:
                if len(st.session_state.cb_jc_filters) > 1:
                    if st.button("✕", key=f"cb_jc_rm_{cb_i}", help="Remove this condition"):
                        st.session_state.cb_jc_filters.pop(cb_i)
                        st.rerun()

        # ─── Apply filters and show results ──────────────────────────
        cb_masks = []
        cb_active = []
        for cb_flt in st.session_state.cb_jc_filters:
            cc = cb_flt.get("column")
            oo = cb_flt.get("op")
            vv = cb_flt.get("value")
            if cc and vv not in (None, "", "(choose a column first)"):
                cb_masks.append(_cb_mask(cb_df, cc, oo, vv))
                cb_active.append((cc, oo, vv))

        if not cb_masks:
            cb_filtered = cb_df
        elif cb_is_or:
            cb_combined = pd.Series(False, index=cb_df.index)
            for m in cb_masks:
                cb_combined = cb_combined | m
            cb_filtered = cb_df[cb_combined]
        else:
            cb_combined = pd.Series(True, index=cb_df.index)
            for m in cb_masks:
                cb_combined = cb_combined & m
            cb_filtered = cb_df[cb_combined]

        # Match count
        st.markdown(
            f'<div style="font-family:Inter,sans-serif;font-size:0.8rem;color:{MUTED};margin:8px 0 6px 0;">'
            f'<b style="color:{INK};font-size:1.05rem;font-family:\'IBM Plex Mono\',monospace;">{len(cb_filtered)}</b>'
            f'<span style="margin-left:6px;">matching rows</span>'
            f'<span style="margin-left:6px;color:{SUBTEXT};">of {len(cb_df)} total</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

        if not cb_filtered.empty:
            # Build preview with Issue Key as hyperlink
            cb_preview = cb_filtered.copy()
            # Show useful columns by default, but respect what's available
            cb_default_cols = [c for c in [
                "Issue Key", "Summary", "Customer", "Priority", "Status", "Creation Date",
                "Resolution Date", "Age (days)", "Labels", "Affects Version"
            ] if c in cb_preview.columns]
            cb_shown_cols = cb_default_cols

            # Format dates as readable strings
            for dc in ["Creation Date", "Resolution Date"]:
                if dc in cb_preview.columns:
                    cb_preview[dc] = pd.to_datetime(cb_preview[dc], errors="coerce").dt.strftime("%Y-%m-%d")
                    cb_preview[dc] = cb_preview[dc].fillna("—")

            # Hyperlink Issue Key
            cb_col_config = {}
            if "Issue Key" in cb_shown_cols:
                cb_preview["Issue Key"] = cb_preview["Issue Key"].apply(
                    lambda k: f"https://jira.corp.adobe.com/browse/{k}" if pd.notna(k) and str(k).strip() else None
                )
                cb_col_config["Issue Key"] = st.column_config.LinkColumn(
                    "Issue Key",
                    help="Click to open this Jira in a new tab.",
                    display_text=r"https?://jira\.corp\.adobe\.com/browse/(.+)",
                )
            if "Age (days)" in cb_shown_cols:
                cb_col_config["Age (days)"] = st.column_config.NumberColumn(
                    "Age (days)",
                    help=f"Days since creation. For open Jiras, computed against {cb_asof.date()}.",
                    format="%d days",
                )

            st.dataframe(
                cb_preview[cb_shown_cols],
                use_container_width=True,
                height=400,
                column_config=cb_col_config,
                hide_index=True,
            )

            # Download
            try:
                st.download_button(
                    "⬇ Download Excel",
                    make_xlsx_bytes(
                        cb_filtered[[c for c in cb_shown_cols if c != "Issue Key" or True]].drop_duplicates(),
                        sheet_name="Filtered_Created_Jiras"
                    ),
                    file_name="custom_filtered_created_jiras.xlsx",
                    mime=XLSX_MIME,
                    key="cb_jc_download",
                )
            except Exception as e:
                st.error(f"Excel export failed: {e}")

        st.markdown("<div style='height:32px'></div>", unsafe_allow_html=True)

        # ─── Monthly creation trend ──────────────────────────────────────
        # Simple bar chart showing how many Jiras were created per month.
        # Helps leadership spot intake spikes and seasonality. The peak
        # month gets a deeper color treatment to draw the eye.
        section("Jiras created per month")

        # Group by month of creation, count per month
        df_monthly = df_jc.copy()
        df_monthly["Month"] = df_monthly["Creation Date"].dt.to_period("M")
        monthly_counts = df_monthly.groupby("Month").size().reset_index(name="Count")
        monthly_counts["Month Label"] = monthly_counts["Month"].dt.strftime("%b %Y")
        monthly_counts["Month Sort"] = monthly_counts["Month"].dt.to_timestamp()
        monthly_counts = monthly_counts.sort_values("Month Sort").reset_index(drop=True)

        if monthly_counts.empty:
            st.caption("No creation data available to plot.")
        else:
            # Highlight the peak month with the deep steel color, others muted
            peak_idx = monthly_counts["Count"].idxmax()
            colors = [
                CHART_STEEL_DEEP if i == peak_idx else CHART_STEEL
                for i in monthly_counts.index
            ]

            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=monthly_counts["Month Label"].tolist(),
                y=monthly_counts["Count"].tolist(),
                marker=dict(color=colors, line=dict(width=0)),
                text=[f"<b>{v}</b>" for v in monthly_counts["Count"].tolist()],
                textposition="outside",
                textfont=dict(color=INK, size=12, family="Inter"),
                hovertemplate="<b>%{x}</b><br>%{y} jiras created<extra></extra>",
                cliponaxis=False,
            ))
            fig.update_layout(
                xaxis_title="",
                yaxis_title="Jiras created",
                height=380,
                showlegend=False,
            )
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="created_monthly_trend")

            peak_month = monthly_counts.loc[peak_idx, "Month Label"]
            peak_count = int(monthly_counts.loc[peak_idx, "Count"])
            st.caption(
                f"Peak month: <b>{peak_month}</b> with {peak_count} Jiras created · "
                f"{int(monthly_counts['Count'].sum()):,} total Jiras across {len(monthly_counts)} months.",
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height:32px'></div>", unsafe_allow_html=True)

        # ─── Open Jiras by age ───────────────────────────────────────────
        # Interactive table showing currently-open Jiras sorted by Age.
        # Age = days since creation, calculated against a stable "as-of"
        # date (max Resolution Date in dataset) so numbers don't drift
        # when the same file is viewed on different days.
        #
        # Features:
        #   • Filtered to OPEN Jiras only (Status NOT IN Resolved/Closed)
        #   • Filters in collapsed expander: Customer, Priority, Status,
        #     Creation Date range (smart preset dropdown)
        #   • Pagination: 10 rows per page with Previous/Next buttons
        #   • Sorted by Age descending (oldest open Jira at the top)
        #
        # Date filter semantics: shortcut "Today" uses ACTUAL today (wall
        # clock) since users expect that. Age calculation continues to use
        # the dataset's max Resolution Date for stability.
        section("Open Jiras by age")

        open_df = df_jc.copy()
        open_df["Resolution Date"] = pd.to_datetime(open_df["Resolution Date"], errors="coerce")

        # Filter to open only
        # Filter to open only (uses the same CLOSED_OR_DEFERRED list defined
        # near the KPI strip above — keeps the "Open: 126" KPI consistent
        # with the row count of this table).
        open_df = open_df[~open_df["Status"].isin(CLOSED_OR_DEFERRED)].copy()

        # As-of date for Age calculation (stable across reloads)
        asof_candidates = [
            df_jc["Resolution Date"].dropna().max() if "Resolution Date" in df_jc.columns else pd.NaT,
            df_jc["Creation Date"].dropna().max(),
            pd.Timestamp.today(),
        ]
        asof_date = next((d for d in asof_candidates if pd.notna(d)), pd.Timestamp.today())
        asof_date = pd.Timestamp(asof_date).normalize()

        # Compute Age = as-of date − Creation Date (Resolution Date is null for all open)
        open_df["Age (days)"] = (asof_date - open_df["Creation Date"]).dt.days
        open_df = open_df[open_df["Age (days)"].notna() & (open_df["Age (days)"] >= 0)]

        if open_df.empty:
            st.caption("No open Jiras in the dataset.")
        else:
            # ─── Filters in expander ─────────────────────────────────
            with st.expander("Refine view", expanded=False):
                fcols = st.columns(2)

                # Customer filter (multi-select)
                all_customers = sorted(
                    open_df["Customer"].dropna().astype(str).str.strip()
                    .replace("", pd.NA).dropna().unique().tolist()
                )
                with fcols[0]:
                    sel_customers = st.multiselect(
                        "Customers",
                        all_customers,
                        default=[],
                        key="open_age_customers",
                        help="Leave empty to show all customers.",
                    )

                # Priority filter (multi-select)
                with fcols[1]:
                    sel_priorities = st.multiselect(
                        "Priorities",
                        ["P1", "P2", "P3", "P4"],
                        default=[],
                        key="open_age_priorities",
                        help="Leave empty to show all priorities.",
                    )

                # Status filter (multi-select from open statuses only)
                open_statuses_in_data = sorted(open_df["Status"].dropna().unique().tolist())
                sel_statuses = st.multiselect(
                    "Status",
                    open_statuses_in_data,
                    default=[],
                    key="open_age_statuses",
                    help="Filter by specific open-state statuses (New, In Progress, etc.).",
                )

                # ─── Date range query builder ───────────────────────
                # Smart preset dropdown that handles 95% of cases, with
                # an escape hatch to a custom range for power users.
                # "Today" uses the real wall-clock today since that's
                # what users expect from a date filter (vs the as-of
                # date used for Age, which is data-driven).
                st.markdown("**Creation Date range**")
                real_today = pd.Timestamp.today().normalize()

                date_range_options = [
                    "All time",
                    "Today",
                    "Yesterday",
                    "Last 7 days",
                    "Last 30 days",
                    "Last 90 days",
                    "Last 6 months",
                    "Year to date",
                    "Custom range...",
                ]
                date_choice = st.selectbox(
                    "Quick range",
                    date_range_options,
                    index=0,  # "All time" default
                    key="open_age_date_range",
                    label_visibility="collapsed",
                )

                # Compute the (from, to) date pair from the choice
                if date_choice == "All time":
                    date_from, date_to = None, None
                elif date_choice == "Today":
                    date_from = real_today
                    date_to = real_today
                elif date_choice == "Yesterday":
                    date_from = real_today - pd.Timedelta(days=1)
                    date_to = real_today - pd.Timedelta(days=1)
                elif date_choice == "Last 7 days":
                    date_from = real_today - pd.Timedelta(days=7)
                    date_to = real_today
                elif date_choice == "Last 30 days":
                    date_from = real_today - pd.Timedelta(days=30)
                    date_to = real_today
                elif date_choice == "Last 90 days":
                    date_from = real_today - pd.Timedelta(days=90)
                    date_to = real_today
                elif date_choice == "Last 6 months":
                    date_from = real_today - pd.DateOffset(months=6)
                    date_to = real_today
                elif date_choice == "Year to date":
                    date_from = pd.Timestamp(real_today.year, 1, 1)
                    date_to = real_today
                elif date_choice == "Custom range...":
                    # Two manual date pickers
                    dcols = st.columns(2)
                    # Sensible defaults: last 30 days
                    cust_default_from = real_today - pd.Timedelta(days=30)
                    cust_default_to = real_today
                    with dcols[0]:
                        date_from = st.date_input(
                            "From",
                            value=cust_default_from.date(),
                            key="open_age_custom_from",
                        )
                        date_from = pd.Timestamp(date_from)
                    with dcols[1]:
                        date_to = st.date_input(
                            "To",
                            value=cust_default_to.date(),
                            key="open_age_custom_to",
                        )
                        date_to = pd.Timestamp(date_to)
                else:
                    date_from, date_to = None, None

            # Apply all filters
            filtered = open_df.copy()
            if sel_customers:
                filtered = filtered[filtered["Customer"].isin(sel_customers)]
            if sel_priorities:
                filtered = filtered[filtered["Priority"].isin(sel_priorities)]
            if sel_statuses:
                filtered = filtered[filtered["Status"].isin(sel_statuses)]
            if date_from is not None:
                filtered = filtered[filtered["Creation Date"] >= date_from]
            if date_to is not None:
                # Include the entire "to" day
                filtered = filtered[filtered["Creation Date"] <= date_to + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)]

            # Sort by Age descending (longest-running first)
            filtered = filtered.sort_values("Age (days)", ascending=False).reset_index(drop=True)

            if filtered.empty:
                st.info("No open Jiras match the selected filters. Try relaxing the filters above.")
            else:
                # ─── Pagination ──────────────────────────────────────
                # 10 rows per page with Previous / Next buttons. Page state
                # lives in session_state so it survives reruns. Auto-reset
                # to page 1 when filters change (would require tracking
                # filter state — for now users manually go back to page 1
                # if they over-paginated; not worth the complexity yet).
                PAGE_SIZE = 10
                total_rows = len(filtered)
                total_pages = (total_rows - 1) // PAGE_SIZE + 1

                # Initialize / clamp page state
                if "open_age_page" not in st.session_state:
                    st.session_state.open_age_page = 1
                # Clamp in case filter result shrunk below current page
                if st.session_state.open_age_page > total_pages:
                    st.session_state.open_age_page = 1
                page = st.session_state.open_age_page

                start_idx = (page - 1) * PAGE_SIZE
                end_idx = start_idx + PAGE_SIZE
                page_slice = filtered.iloc[start_idx:end_idx].copy()

                # Build display columns with clickable Jira link
                JIRA_BASE_URL = "https://jira.corp.adobe.com/browse/"
                page_slice["Jira Link"] = page_slice["Issue Key"].apply(
                    lambda k: f"{JIRA_BASE_URL}{k}"
                )
                page_slice["Created"] = page_slice["Creation Date"].dt.strftime("%Y-%m-%d")

                show_cols = ["Issue Key", "Summary", "Customer", "Priority", "Status", "Created", "Age (days)", "Jira Link"]
                show_cols = [c for c in show_cols if c in page_slice.columns]

                st.dataframe(
                    page_slice[show_cols],
                    use_container_width=True,
                    height=400,
                    column_config={
                        "Jira Link": st.column_config.LinkColumn(
                            "Open in Jira",
                            display_text="🔗 View",
                        ),
                        "Age (days)": st.column_config.NumberColumn(
                            "Age (days)",
                            help=f"Days since creation, computed against {asof_date.date()} (latest known activity in the dataset).",
                            format="%d days",
                        ),
                    },
                    hide_index=True,
                )

                # Pagination controls
                pcols = st.columns([1, 2, 1])
                with pcols[0]:
                    if st.button("← Previous", key="open_age_prev", disabled=(page <= 1), use_container_width=True):
                        st.session_state.open_age_page = max(1, page - 1)
                        st.rerun()
                with pcols[1]:
                    st.markdown(
                        f'<div style="text-align:center;font-family:Inter,sans-serif;'
                        f'font-size:0.9rem;color:{SUBTEXT};padding:8px 0;">'
                        f'Page <b style="color:{INK}">{page}</b> of {total_pages} · '
                        f'<b style="color:{INK}">{total_rows}</b> open Jira(s) match the filters'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                with pcols[2]:
                    if st.button("Next →", key="open_age_next", disabled=(page >= total_pages), use_container_width=True):
                        st.session_state.open_age_page = min(total_pages, page + 1)
                        st.rerun()

                st.caption(
                    f"Showing rows {start_idx + 1}–{min(end_idx, total_rows)} of {total_rows} · "
                    f"Age computed as days since Creation Date, anchored to {asof_date.date()}."
                )

        # ─── Top 10 longest-open P1s by Age (visualization) ──────────────
        # Leadership signal chart: critical work that's been sitting too long.
        # Per Nikhila's ask — kept as a visualization (not a table) so the
        # message lands at a glance: "these P1s are hurting us right now."
        # Customer name shown alongside Issue Key so leadership can scan and
        # know who's waiting without needing a hover or a click.
        #
        # Data source: JIRA_Created (only source with open Jiras + Priority).
        # Filter: Status NOT IN (Resolved, Closed, Deferred) AND Priority = P1.
        # Sort: Age descending. Top 10.
        st.markdown("<div style='height:32px'></div>", unsafe_allow_html=True)
        section("Top 10 longest-open P1s")

        open_p1 = df_jc.copy()
        open_p1["Creation Date"] = pd.to_datetime(open_p1["Creation Date"], errors="coerce")
        if "Resolution Date" in open_p1.columns:
            open_p1["Resolution Date"] = pd.to_datetime(open_p1["Resolution Date"], errors="coerce")
        # Stable as-of date (same logic as the table above)
        p1_asof_candidates = [
            open_p1["Resolution Date"].dropna().max() if "Resolution Date" in open_p1.columns else pd.NaT,
            open_p1["Creation Date"].dropna().max(),
            pd.Timestamp.today(),
        ]
        p1_asof = next((d for d in p1_asof_candidates if pd.notna(d)), pd.Timestamp.today())
        p1_asof = pd.Timestamp(p1_asof).normalize()

        # Filter: open P1s
        open_p1 = open_p1[
            (open_p1["Priority"] == "P1") &
            (~open_p1["Status"].isin(["Resolved", "Closed", "Deferred"]))
        ].copy()
        open_p1["Age (days)"] = (p1_asof - open_p1["Creation Date"]).dt.days
        open_p1 = open_p1[open_p1["Age (days)"].notna() & (open_p1["Age (days)"] >= 0)]
        open_p1 = open_p1.sort_values("Age (days)", ascending=False).head(10)

        if open_p1.empty:
            st.caption("No open P1 Jiras to show — the critical backlog is clear ✅")
        else:
            # Build chart labels: "ISSUE-KEY · Customer (truncated)".
            # Truncation prevents long customer names (especially the
            # "X\r\nY" multi-customer cells) from breaking the layout.
            def _truncate(name, n=25):
                if not isinstance(name, str):
                    return "—"
                # Some Customer cells have \r\n joining multiple names — just
                # show the first one for the chart label, full name on hover
                first_line = name.split("\r\n")[0].split("\n")[0].strip()
                if len(first_line) > n:
                    return first_line[:n].rstrip() + "…"
                return first_line

            bar_labels = [
                f"{row['Issue Key']} · {_truncate(row.get('Customer', '—'))}"
                for _, row in open_p1.iterrows()
            ]
            bar_values = open_p1["Age (days)"].tolist()

            # Hover gives the full picture: full customer name, status, dates
            hover_texts = []
            for _, row in open_p1.iterrows():
                cust_full = str(row.get("Customer", "—")).replace("\r\n", " / ").replace("\n", " / ")
                created = row["Creation Date"].strftime("%Y-%m-%d") if pd.notna(row["Creation Date"]) else "—"
                hover_texts.append(
                    f"<b>{row['Issue Key']}</b><br>"
                    f"Customer: {cust_full}<br>"
                    f"Status: {row.get('Status', '—')}<br>"
                    f"Created: {created}<br>"
                    f"Age: <b>{int(row['Age (days)'])} days</b>"
                )

            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=bar_values,
                y=bar_labels,
                orientation="h",
                marker=dict(color=CHART_TERRACOTTA_DEEP, line=dict(width=0)),
                text=[f"<b>{int(v)} days</b>" for v in bar_values],
                textposition="outside",
                textfont=dict(color=INK, size=12, family="Inter"),
                hovertext=hover_texts,
                hoverinfo="text",
                cliponaxis=False,
            ))
            fig.update_layout(
                yaxis=dict(autorange="reversed", tickfont=dict(color=INK, size=12, family="Inter")),
                xaxis_title="Days open",
                yaxis_title="",
                height=460,
                showlegend=False,
                margin=dict(l=10, r=80, t=20, b=40),
            )
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="created_top10_open_p1")

            st.caption(
                f"Showing the {len(open_p1)} longest-open P1 Jiras · "
                f"Age = days since Creation Date, anchored to {p1_asof.date()} · "
                f"Customer name shown next to Issue Key. Hover for full details."
            )

    # Stop here — none of the main dashboard renders in Created Jiras view.
    st.stop()


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

# Context badge directly above the KPI strip — makes it unmissable that these
# numbers represent CLOSED work in the selected reporting period. Without this,
# leadership might mentally compare 304 (here) against 360 (Reports view) and
# assume it's the same population. The badge is small but does heavy lifting:
# it pre-empts a recurring source of confusion in cross-tab comparisons.
st.markdown(
    f'<div style="display:inline-flex;align-items:center;gap:8px;'
    f'background:rgba(63,139,126,0.10);border:1px solid {CHART_EMERALD};'
    f'border-radius:999px;padding:5px 14px;margin:0 0 14px 0;">'
    f'<span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:{CHART_EMERALD_DEEP};"></span>'
    f'<span style="font-family:Inter,sans-serif;font-size:0.74rem;font-weight:700;'
    f'letter-spacing:0.08em;text-transform:uppercase;color:{CHART_EMERALD_DEEP};">'
    f'Showing resolved Jiras · {reporting_period}'
    f'</span>'
    f'</div>',
    unsafe_allow_html=True,
)

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
tab1, tab2, tab3, tab_customer, tab4, tab5 = st.tabs(["Snapshot", "Quality", "Team", "Customers", "Data explorer", "How we calculate"])

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

    # ─── Top 10 longest-running Jiras ────────────────────────────────────
    # Moved here from Snapshot per leadership feedback (Nikhila): the Age
    # theme is more relevant to resolution-pattern analysis than to the
    # period-level Snapshot story, so it lives on Quality now.
    #
    # A spotlight view of the 10 P2E Jiras that took the longest to resolve
    # within the selected period. Resolution-based; respects the period
    # filter. No local filters — quick "what hurt us this period" view.
    # The Customers tab has the filter-capable version with open + closed.
    if not df.empty and "Creation Date" in df.columns and "Resolution Date" in df.columns:
        section("Top 10 longest-running Jiras")

        # Compute days to resolve. Negative or null days get filtered out
        # because they indicate data issues we can't draw conclusions from.
        lr = df.copy()
        lr["Creation Date"] = pd.to_datetime(lr["Creation Date"], errors="coerce")
        lr["Resolution Date"] = pd.to_datetime(lr["Resolution Date"], errors="coerce")
        lr["Days to Resolve"] = (lr["Resolution Date"] - lr["Creation Date"]).dt.days
        lr = lr[lr["Days to Resolve"].notna() & (lr["Days to Resolve"] >= 0)]

        if lr.empty:
            st.caption("No Jiras in the current period have valid date data for this calculation.")
        else:
            # Sort by Days to Resolve descending, take top 10
            lr = lr.sort_values("Days to Resolve", ascending=False).head(10)

            # Build display frame with clickable Jira link
            JIRA_BASE_URL = "https://jira.corp.adobe.com/browse/"
            display_lr = lr.copy()
            display_lr["Jira Link"] = display_lr["Issue Key"].apply(
                lambda k: f"{JIRA_BASE_URL}{k}"
            )
            display_lr["Created"] = display_lr["Creation Date"].dt.strftime("%Y-%m-%d")
            display_lr["Resolved"] = display_lr["Resolution Date"].dt.strftime("%Y-%m-%d")

            # Use Priority_Short if it exists (cleaner), otherwise raw Priority
            priority_col = "Priority_Short" if "Priority_Short" in display_lr.columns else "Priority"
            show_cols = ["Issue Key", "Summary", "Customer", priority_col, "Created", "Resolved", "Days to Resolve", "Jira Link"]
            show_cols = [c for c in show_cols if c in display_lr.columns]
            display_df = display_lr[show_cols].rename(columns={"Priority_Short": "Priority"})

            st.dataframe(
                display_df,
                use_container_width=True,
                height=400,
                column_config={
                    "Jira Link": st.column_config.LinkColumn(
                        "Open in Jira",
                        display_text="🔗 View",
                    ),
                    "Days to Resolve": st.column_config.NumberColumn(
                        "Days to Resolve",
                        help="Time between Creation Date and Resolution Date.",
                        format="%d days",
                    ),
                },
                hide_index=True,
            )

            st.caption(
                f"Showing the 10 longest-running Jiras in the selected period · "
                f"sorted by days from Creation Date to Resolution Date."
            )

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

# -- TAB: Customers -----------------------------------------------------------
# Customer-centric view requested by leadership. Four charts that answer the
# director's recurring questions:
#   1. Top customers — who's raising the most issues (lens-aware)
#   2. Priority breakdown per customer — how many P1s/P2s per customer
#   3. Resolution time by priority — are P1s getting faster service than P4s
#   4. Longest-running Jiras table — what's currently demanding attention
#
# Data source depends on the active lens:
#   • Resolved lens → df_source (the 371 P2E quarter-sheet rows)
#   • Created lens → df_jira_created (the 441 created Jiras)
# We pick the right source at the top of the tab to keep downstream code clean.
with tab_customer:
    # Choose the right dataset based on the active lens
    # Note: is_reports_view is the existing boolean from the view picker
    if is_reports_view:
        customer_df = df_jira_created.copy() if df_jira_created is not None else pd.DataFrame()
        lens_label = "Created Jiras"
        date_col = "Creation Date"
    else:
        customer_df = df.copy()  # df = period-filtered resolved data
        lens_label = "Resolved Jiras"
        date_col = "Resolution Date"

    if customer_df.empty or "Customer" not in customer_df.columns:
        st.warning(
            f"No customer data available for the {lens_label} view. "
            f"This view requires a Customer column on the underlying data."
        )
    else:
        # ─── Chart 1: Top customers ────────────────────────────────────
        # The classic "who is filing the most Jiras" chart. Lens-aware so
        # leadership can flip between resolution-based and creation-based
        # rankings without losing visual context.
        section(f"Top customers · {lens_label}")

        top_customers = (
            customer_df["Customer"]
            .dropna()
            .astype(str).str.strip()
            .replace("", pd.NA).dropna()
            .value_counts()
            .head(12)
            .reset_index()
        )
        top_customers.columns = ["Customer", "Count"]

        if not top_customers.empty:
            fig = go.Figure()
            counts_list = top_customers["Count"].tolist()
            colors = bar_palette(counts_list, CHART_PLUM, CHART_PLUM_DEEP)

            fig.add_trace(go.Bar(
                x=counts_list,
                y=top_customers["Customer"].tolist(),
                orientation="h",
                marker=dict(color=colors, line=dict(width=0)),
                text=[f"<b>{v}</b>" for v in counts_list],
                textposition="outside",
                textfont=dict(color=INK, size=12, family="Inter"),
                hovertemplate="<b>%{y}</b><br>%{x} jiras<extra></extra>",
                cliponaxis=False,
            ))
            fig.update_layout(
                yaxis=dict(autorange="reversed"),
                xaxis_title="",
                yaxis_title="",
                height=420,
            )
            chart_theme(fig)
            st.plotly_chart(fig, use_container_width=True, key="cust_top_customers")

            verb = "raised" if is_reports_view else "resolved"
            st.caption(
                f"Showing top {len(top_customers)} customers by Jiras {verb} · "
                f"{int(customer_df['Customer'].notna().sum()):,} total Jiras have a customer tagged."
            )

        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)

        # ─── Chart 2: Priority breakdown per customer ──────────────────
        # Director's stated ask: "how many P1s/P2s per customer." Stacked
        # horizontal bar with priority-colored segments. Filters live in
        # an expander to keep the default view clean.
        section(f"Priority breakdown per customer · {lens_label}")

        if "Priority" not in customer_df.columns and "Priority_Short" not in customer_df.columns:
            st.info("Priority data is not available on this dataset.")
        else:
            # Get short priority codes — JIRA_Created has Priority as P1/P2/P3/P4 directly,
            # quarter sheets have long descriptions that need shortening.
            if "Priority_Short" in customer_df.columns:
                customer_df["_pri"] = customer_df["Priority_Short"]
            else:
                def _sp(p):
                    if pd.isna(p): return None
                    s = str(p).strip()
                    m = re.match(r"^(P[0-9])\b", s, re.IGNORECASE)
                    return m.group(1).upper() if m else s
                customer_df["_pri"] = customer_df["Priority"].map(_sp)

            # ─── Filters in expander ─────────────────────────────────
            with st.expander("Refine view", expanded=False):
                fcols = st.columns(3)

                # Filter 1: Customer multi-select (default = top 10 by volume)
                all_customers = sorted(
                    customer_df["Customer"].dropna().astype(str).str.strip()
                    .replace("", pd.NA).dropna().unique().tolist()
                )
                default_customers = top_customers.head(10)["Customer"].tolist()
                with fcols[0]:
                    sel_customers = st.multiselect(
                        "Customers",
                        all_customers,
                        default=default_customers,
                        key="cust_priority_customers",
                        help="Pick specific customers to focus on. Defaults to top 10 by volume.",
                    )

                # Filter 2: Priority multi-select
                priority_options = ["P1", "P2", "P3", "P4"]
                with fcols[1]:
                    sel_priorities = st.multiselect(
                        "Priorities",
                        priority_options,
                        default=priority_options,
                        key="cust_priority_priorities",
                        help="Limit to specific priority levels.",
                    )

                # Filter 3: Type of Request (only meaningful when column exists)
                with fcols[2]:
                    if "Type of Request" in customer_df.columns:
                        type_options = ["Bug", "Request for Help"]
                        sel_types = st.multiselect(
                            "Type of Request",
                            type_options,
                            default=type_options,
                            key="cust_priority_types",
                            help="Filter by Bug vs Request for Help.",
                        )
                    else:
                        st.caption("Type of Request not available on this dataset")
                        sel_types = None

            # Apply filters
            chart_df = customer_df.copy()
            if sel_customers:
                chart_df = chart_df[chart_df["Customer"].isin(sel_customers)]
            if sel_priorities:
                chart_df = chart_df[chart_df["_pri"].isin(sel_priorities)]
            if sel_types is not None and sel_types and "Type of Request" in chart_df.columns:
                chart_df = chart_df[chart_df["Type of Request"].isin(sel_types)]

            if chart_df.empty:
                st.info("No data matches the selected filters. Adjust the filters above.")
            else:
                # Build cross-tab: customer × priority
                ct = pd.crosstab(chart_df["Customer"], chart_df["_pri"])
                # Order customers by total volume (descending)
                ct["__total"] = ct.sum(axis=1)
                ct = ct.sort_values("__total", ascending=False).head(10).drop(columns="__total")

                # Stacked horizontal bar — one trace per priority
                priority_colors = {
                    "P1": CHART_TERRACOTTA_DEEP,   # critical = red
                    "P2": CHART_INDIGO_DEEP,        # high = indigo
                    "P3": CHART_STEEL,              # medium = steel
                    "P4": CHART_PLUM,               # low = plum
                }

                fig = go.Figure()
                for pri in ["P1", "P2", "P3", "P4"]:
                    if pri not in ct.columns:
                        continue
                    fig.add_trace(go.Bar(
                        x=ct[pri].tolist(),
                        y=ct.index.tolist(),
                        name=pri,
                        orientation="h",
                        marker=dict(color=priority_colors.get(pri, MUTED), line=dict(width=0)),
                        hovertemplate=f"<b>%{{y}}</b><br>{pri}: %{{x}}<extra></extra>",
                    ))
                fig.update_layout(
                    barmode="stack",
                    yaxis=dict(autorange="reversed"),
                    xaxis_title="",
                    yaxis_title="",
                    height=460,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                )
                chart_theme(fig)
                st.plotly_chart(fig, use_container_width=True, key="cust_priority_breakdown")

                st.caption(
                    f"Each bar represents one customer; segments show priority distribution. "
                    f"{len(ct)} customer(s) shown · {int(ct.values.sum())} Jiras total in the filtered view."
                )

        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)

        # ─── Chart 3: Resolution time by priority ──────────────────────
        # Surfaces whether P1s actually get faster service than P4s.
        # Only meaningful on Resolved lens (open Jiras have no resolution time).
        section(f"Resolution time by priority · {lens_label}")

        if is_reports_view:
            # Created lens: show banner explaining this chart applies to closed work only
            closed_subset = customer_df[customer_df["Resolution Date"].notna()].copy()
            st.markdown(
                f'<div style="font-family:Inter,sans-serif;font-size:0.85rem;color:{SUBTEXT};'
                f'background:rgba(124,111,232,0.05);border-left:3px solid {CHART_INDIGO};'
                f'padding:12px 16px;border-radius:6px;margin:0 0 16px 0;">'
                f'Resolution time applies to closed Jiras only. Showing the {len(closed_subset)} '
                f'resolved subset of {len(customer_df)} created Jiras.'
                f'</div>',
                unsafe_allow_html=True,
            )
            rt_df = closed_subset
        else:
            rt_df = customer_df.copy()

        # Need both dates to compute days-to-resolve
        if "Creation Date" not in rt_df.columns or "Resolution Date" not in rt_df.columns:
            st.info("Resolution time requires both Creation Date and Resolution Date columns.")
        else:
            rt_df["Creation Date"] = pd.to_datetime(rt_df["Creation Date"], errors="coerce")
            rt_df["Resolution Date"] = pd.to_datetime(rt_df["Resolution Date"], errors="coerce")
            rt_df["_days"] = (rt_df["Resolution Date"] - rt_df["Creation Date"]).dt.days
            rt_df = rt_df[rt_df["_days"].notna() & (rt_df["_days"] >= 0)]

            # Compute Priority_Short if not present
            if "_pri" not in rt_df.columns:
                if "Priority_Short" in rt_df.columns:
                    rt_df["_pri"] = rt_df["Priority_Short"]
                else:
                    def _sp(p):
                        if pd.isna(p): return None
                        s = str(p).strip()
                        m = re.match(r"^(P[0-9])\b", s, re.IGNORECASE)
                        return m.group(1).upper() if m else s
                    rt_df["_pri"] = rt_df["Priority"].map(_sp)

            agg = rt_df.groupby("_pri")["_days"].agg(["median", "mean", "count"]).reset_index()
            agg = agg[agg["_pri"].isin(["P1", "P2", "P3", "P4"])]
            agg = agg.sort_values("_pri")

            if agg.empty:
                st.info("Not enough data to compute resolution time by priority.")
            else:
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=agg["_pri"].tolist(),
                    y=agg["median"].tolist(),
                    name="Median days",
                    marker=dict(color=CHART_EMERALD_DEEP, line=dict(width=0)),
                    text=[f"<b>{int(v)}</b>" for v in agg["median"].tolist()],
                    textposition="outside",
                    textfont=dict(color=INK, size=12, family="Inter"),
                    hovertemplate="<b>%{x}</b><br>Median: %{y:.1f} days<extra></extra>",
                ))
                fig.add_trace(go.Bar(
                    x=agg["_pri"].tolist(),
                    y=agg["mean"].tolist(),
                    name="Mean days",
                    marker=dict(color=CHART_STEEL, line=dict(width=0)),
                    text=[f"<b>{int(v)}</b>" for v in agg["mean"].tolist()],
                    textposition="outside",
                    textfont=dict(color=INK, size=12, family="Inter"),
                    hovertemplate="<b>%{x}</b><br>Mean: %{y:.1f} days<extra></extra>",
                ))
                fig.update_layout(
                    barmode="group",
                    xaxis_title="",
                    yaxis_title="Days",
                    height=360,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                )
                chart_theme(fig)
                st.plotly_chart(fig, use_container_width=True, key="cust_restime_by_priority")

                count_text = " · ".join(f"{r['_pri']}: {int(r['count'])} jiras" for _, r in agg.iterrows())
                st.caption(count_text)

        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)

        # ─── Chart 4: Longest-running Jiras table ─────────────────────
        # Director's "drill into specifics" view. Shows the top N
        # longest-running Jiras with full metadata. Filters let user
        # narrow by customer, status, and priority.
        section(f"Longest-running Jiras · {lens_label}")

        # The data source for this chart depends on the active lens:
        #   • Resolved lens → only Jiras with a Resolution Date (truly resolved)
        #   • Created lens → all of JIRA_Created (includes open + closed)
        # The local Status filter then operates within that already-scoped set.
        # Using JIRA_Created as the base for both (it has open + closed) and
        # filtering by Resolution Date for Resolved lens keeps the table
        # consistent with the rest of the lens's mental model.
        if df_jira_created is not None and not df_jira_created.empty and "Creation Date" in df_jira_created.columns:
            lr_df = df_jira_created.copy()
            lr_df["Creation Date"] = pd.to_datetime(lr_df["Creation Date"], errors="coerce")
            if "Resolution Date" in lr_df.columns:
                lr_df["Resolution Date"] = pd.to_datetime(lr_df["Resolution Date"], errors="coerce")
            else:
                lr_df["Resolution Date"] = pd.NaT

            # ─── Apply lens scope ────────────────────────────────────
            # On Resolved lens, restrict to Jiras that actually have a
            # Resolution Date. This prevents open Jiras (Status=New etc.)
            # from leaking into a view that's contextually about resolved
            # work. On Created lens, keep all rows so open backlog is visible.
            if not is_reports_view:
                lr_df = lr_df[lr_df["Resolution Date"].notna()].copy()

            # "As-of" date = max Resolution Date in dataset (per your earlier decision)
            asof_candidates = [
                lr_df["Resolution Date"].max(),
                lr_df["Creation Date"].max(),
                pd.Timestamp.today(),
            ]
            asof_date = next((d for d in asof_candidates if pd.notna(d)), pd.Timestamp.today())
            # Use last day of as-of month to keep numbers stable
            asof_date = pd.Timestamp(asof_date).normalize()

            # Compute days open: for closed Jiras = (resolution - creation), for open = (asof - creation)
            lr_df["_end_date"] = lr_df["Resolution Date"].fillna(asof_date)
            lr_df["Days Open"] = (lr_df["_end_date"] - lr_df["Creation Date"]).dt.days
            lr_df = lr_df[lr_df["Days Open"].notna() & (lr_df["Days Open"] >= 0)]

            # ─── Filters ──────────────────────────────────────────
            with st.expander("Refine view", expanded=False):
                lcols = st.columns(3)

                # Customer filter
                lr_customers = sorted(
                    lr_df["Customer"].dropna().astype(str).str.strip()
                    .replace("", pd.NA).dropna().unique().tolist()
                ) if "Customer" in lr_df.columns else []
                with lcols[0]:
                    lr_sel_customers = st.multiselect(
                        "Customers",
                        lr_customers,
                        default=[],
                        key="lr_customers",
                        help="Leave empty to show all customers.",
                    )

                # Status filter — lens-aware options
                with lcols[1]:
                    if is_reports_view:
                        # Created lens: All / Open / Closed
                        status_choice = st.selectbox(
                            "Status",
                            ["All", "Open", "Closed"],
                            index=0,
                            key="lr_status",
                            help="Open = Status not Resolved/Closed. Closed = Resolved or Closed.",
                        )
                    else:
                        # Resolved lens: All / Resolved / Closed
                        status_choice = st.selectbox(
                            "Status",
                            ["All", "Resolved", "Closed"],
                            index=0,
                            key="lr_status",
                            help="Filter by terminal status.",
                        )

                # Priority filter (default P1)
                with lcols[2]:
                    lr_sel_priorities = st.multiselect(
                        "Priorities",
                        ["P1", "P2", "P3", "P4"],
                        default=["P1"],
                        key="lr_priorities",
                        help="Default shows P1 only. Pick more to expand.",
                    )

            # Apply filters
            if lr_sel_customers:
                lr_df = lr_df[lr_df["Customer"].isin(lr_sel_customers)]
            if status_choice == "Open":
                lr_df = lr_df[~lr_df["Status"].isin(["Resolved", "Closed"])]
            elif status_choice == "Resolved":
                lr_df = lr_df[lr_df["Status"] == "Resolved"]
            elif status_choice == "Closed":
                lr_df = lr_df[lr_df["Status"] == "Closed"]
            if lr_sel_priorities:
                # Priority in JIRA_Created is already short (P1/P2/etc.)
                lr_df = lr_df[lr_df["Priority"].isin(lr_sel_priorities)]

            # Sort by Days Open descending, take top 15
            lr_df = lr_df.sort_values("Days Open", ascending=False).head(15)

            if lr_df.empty:
                st.info("No Jiras match the selected filters. Try relaxing the filters above.")
            else:
                # Build display columns + clickable Jira link
                JIRA_BASE_URL = "https://jira.corp.adobe.com/browse/"

                # Format dates as strings for display
                display_df = lr_df.copy()
                display_df["Jira Link"] = display_df["Issue Key"].apply(
                    lambda k: f"{JIRA_BASE_URL}{k}"
                )
                display_df["Created"] = display_df["Creation Date"].dt.strftime("%Y-%m-%d")
                display_df["Resolved"] = display_df["Resolution Date"].apply(
                    lambda d: d.strftime("%Y-%m-%d") if pd.notna(d) else "—"
                )

                show_cols = ["Issue Key", "Summary", "Customer", "Priority", "Status", "Created", "Resolved", "Days Open", "Jira Link"]
                show_cols = [c for c in show_cols if c in display_df.columns]

                st.dataframe(
                    display_df[show_cols],
                    use_container_width=True,
                    height=500,
                    column_config={
                        "Jira Link": st.column_config.LinkColumn(
                            "Open in Jira",
                            display_text="🔗 View",
                        ),
                        "Days Open": st.column_config.NumberColumn(
                            "Days Open",
                            help="Time between Creation Date and Resolution Date (or 'as-of' date for open Jiras).",
                            format="%d days",
                        ),
                    },
                    hide_index=True,
                )

                st.caption(
                    f"Showing top {len(display_df)} of filtered Jiras by Days Open · "
                    f"Data as of {asof_date.date()} · "
                    f"For open Jiras, 'Days Open' uses the as-of date as the endpoint."
                )
        else:
            st.info("JIRA_Created sheet is required for the longest-running Jiras table.")

# -- TAB 4: Data explorer -----------------------------------------------------
with tab4:
    # Derived columns for the query builder: Age (days) is computed from
    # Creation Date and Resolution Date. On the Resolved lens (where this
    # tab lives), every row has a Resolution Date by design — Age represents
    # "how long it took to resolve." Users can then build queries like
    # "Age > 90 days" to find unusually slow resolutions.
    #
    # We mutate df in place because Data Explorer is the last tab and the
    # column is harmless for the rest of the script (How we calculate tab
    # doesn't read from df). This avoids a separate derivation copy.
    if "Creation Date" in df.columns and "Resolution Date" in df.columns and "Age (days)" not in df.columns:
        _creation = pd.to_datetime(df["Creation Date"], errors="coerce")
        _resolution = pd.to_datetime(df["Resolution Date"], errors="coerce")
        df["Age (days)"] = (_resolution - _creation).dt.days

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
    # Date columns get their own operator set + a calendar value picker.
    # Detect a date column by dtype rather than name so this works for any
    # future date columns the data team adds.
    DATE_OPERATORS = ["on", "on or after", "on or before", "between"]
    # Labels column gets its own operator set since it's list-valued (each
    # cell is a Python list as string like "['Endava', 'P2EminingReviewed']").
    #   • "has any of"  → row matches if it has at least one of the picked labels (OR)
    #   • "has none of" → row matches if it has NONE of the picked labels (NOT)
    LABEL_OPERATORS = ["has any of", "has none of"]
    # Numeric columns (like the derived Age (days) column) get arithmetic
    # operators instead of text ones. Critical for "show me Jiras open > 30 days"
    # style queries that the director uses when scanning backlog.
    NUMERIC_OPERATORS = ["equals", "greater than", "less than", "at least", "at most", "between"]

    def is_date_column(frame: pd.DataFrame, column: str) -> bool:
        """True if the column has a datetime dtype (after parse_date_columns).
        This is what gates the date-aware operator set + calendar picker."""
        if column not in frame.columns:
            return False
        try:
            return pd.api.types.is_datetime64_any_dtype(frame[column])
        except Exception:
            return False

    def is_numeric_column(frame: pd.DataFrame, column: str) -> bool:
        """True if the column has a numeric dtype (int/float). Used to gate
        the numeric operator set + number-input value picker for derived
        columns like Age (days)."""
        if column not in frame.columns:
            return False
        try:
            return pd.api.types.is_numeric_dtype(frame[column])
        except Exception:
            return False

    def is_labels_column(column: str) -> bool:
        """True if the column is the Labels column (which contains list-as-string
        values like "['Endava', 'P2EminingReviewed']"). Gates the labels-specific
        operator set + searchable multi-select value picker."""
        return column.strip().lower() == "labels"

    def parse_labels_cell(val) -> list:
        """Parse a single Labels cell ('["Endava", "Cust-MSFT"]') into a Python list.
        Handles edge cases: empty list strings, NaN, malformed values."""
        if pd.isna(val):
            return []
        s = str(val).strip()
        if not s or s == "[]":
            return []
        try:
            import ast as _ast
            parsed = _ast.literal_eval(s)
            if isinstance(parsed, list):
                return [str(x).strip() for x in parsed if str(x).strip()]
            return []
        except Exception:
            return []

    def build_labels_options(frame: pd.DataFrame, column: str) -> list:
        """Return a sorted list of unique labels with their occurrence counts,
        formatted as 'Label (N)' for the multi-select. Sorted by count desc
        so the most common labels appear first. We keep the display format
        consistent for both selection and re-display when rebuilding the UI."""
        if column not in frame.columns:
            return []
        counter = {}
        for val in frame[column].dropna():
            for lbl in parse_labels_cell(val):
                counter[lbl] = counter.get(lbl, 0) + 1
        if not counter:
            return []
        # Sort by count desc, then alphabetical for ties
        sorted_items = sorted(counter.items(), key=lambda kv: (-kv[1], kv[0].lower()))
        return [f"{lbl} ({cnt})" for lbl, cnt in sorted_items]

    def strip_count_suffix(label_with_count: str) -> str:
        """Reverse the 'Label (N)' formatting to get back the raw label name.
        Used when applying the filter — we match against raw labels, not display strings."""
        if not label_with_count:
            return ""
        s = str(label_with_count)
        # Strip trailing " (N)" if present
        m = re.search(r"\s*\(\d+\)\s*$", s)
        if m:
            return s[:m.start()].strip()
        return s.strip()

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

        # Switch operator + value widgets based on the column type.
        # Four regimes: date column → DATE_OPERATORS, Labels column →
        # LABEL_OPERATORS, numeric column → NUMERIC_OPERATORS, anything
        # else → standard text OPERATORS.
        col_is_date = is_date_column(df, col_choice)
        col_is_labels = is_labels_column(col_choice)
        col_is_numeric = is_numeric_column(df, col_choice) and not col_is_date  # date is technically numeric
        if col_is_date:
            operators_for_col = DATE_OPERATORS
        elif col_is_labels:
            operators_for_col = LABEL_OPERATORS
        elif col_is_numeric:
            operators_for_col = NUMERIC_OPERATORS
        else:
            operators_for_col = OPERATORS

        # Operator picker
        with row_cols[2]:
            current_op = flt.get("op") or operators_for_col[0]
            if current_op not in operators_for_col:
                # User just switched between a text column and a date column;
                # reset to that operator-set's default rather than error.
                current_op = operators_for_col[0]
            op_choice = st.selectbox(
                "Operator",
                operators_for_col,
                index=operators_for_col.index(current_op),
                key=f"flt_op_{i}",
                label_visibility="collapsed",
            )
            flt["op"] = op_choice

        # Value picker — three branches:
        #   1. Date column → calendar picker(s). "between" shows two dates.
        #   2. is / is not → dropdown of distinct values.
        #   3. contains / does not contain → freeform text input.
        with row_cols[3]:
            if col_is_date:
                # Compute sensible min/max bounds from the data so the picker
                # opens on a relevant month. Falls back to today if data is empty.
                date_series = pd.to_datetime(df[col_choice], errors="coerce").dropna()
                if not date_series.empty:
                    data_min = date_series.min().date()
                    data_max = date_series.max().date()
                    default_date = data_max  # most recent — what users usually want
                else:
                    data_min = data_max = default_date = pd.Timestamp.today().date()

                # Persist previous selection across reruns when possible
                prev_val = flt.get("value")
                if op_choice == "between":
                    # "between" needs two dates → render in a 2-col split.
                    # Store the range as a tuple in flt["value"].
                    if isinstance(prev_val, (list, tuple)) and len(prev_val) == 2:
                        try:
                            start_default = pd.to_datetime(prev_val[0]).date()
                            end_default = pd.to_datetime(prev_val[1]).date()
                        except Exception:
                            start_default, end_default = data_min, data_max
                    else:
                        start_default, end_default = data_min, data_max

                    sub_cols = st.columns(2)
                    with sub_cols[0]:
                        start_date = st.date_input(
                            "From",
                            value=start_default,
                            min_value=data_min,
                            max_value=data_max,
                            key=f"flt_val_date_start_{i}",
                            label_visibility="collapsed",
                        )
                    with sub_cols[1]:
                        end_date = st.date_input(
                            "To",
                            value=end_default,
                            min_value=data_min,
                            max_value=data_max,
                            key=f"flt_val_date_end_{i}",
                            label_visibility="collapsed",
                        )
                    val = (start_date, end_date)
                else:
                    # Single-date picker for "on", "on or after", "on or before"
                    if isinstance(prev_val, (list, tuple)):
                        # User just switched from "between" — collapse to first
                        try:
                            single_default = pd.to_datetime(prev_val[0]).date()
                        except Exception:
                            single_default = default_date
                    elif prev_val:
                        try:
                            single_default = pd.to_datetime(prev_val).date()
                        except Exception:
                            single_default = default_date
                    else:
                        single_default = default_date
                    val = st.date_input(
                        "Date",
                        value=single_default,
                        min_value=data_min,
                        max_value=data_max,
                        key=f"flt_val_date_{i}",
                        label_visibility="collapsed",
                    )
            elif col_is_numeric:
                # Numeric column → use st.number_input (or two for "between").
                # Computes data min/max so the user knows the valid range.
                num_series = pd.to_numeric(df[col_choice], errors="coerce").dropna()
                if num_series.empty:
                    data_min_n, data_max_n = 0.0, 100.0
                else:
                    data_min_n = float(num_series.min())
                    data_max_n = float(num_series.max())
                prev_val = flt.get("value")
                if op_choice == "between":
                    # Two number inputs for a range
                    if isinstance(prev_val, (list, tuple)) and len(prev_val) == 2:
                        try:
                            start_default_n = float(prev_val[0])
                            end_default_n = float(prev_val[1])
                        except Exception:
                            start_default_n, end_default_n = data_min_n, data_max_n
                    else:
                        start_default_n, end_default_n = data_min_n, data_max_n
                    sub_cols = st.columns(2)
                    with sub_cols[0]:
                        start_n = st.number_input(
                            "From",
                            value=start_default_n,
                            step=1.0,
                            key=f"flt_val_num_start_{i}",
                            label_visibility="collapsed",
                        )
                    with sub_cols[1]:
                        end_n = st.number_input(
                            "To",
                            value=end_default_n,
                            step=1.0,
                            key=f"flt_val_num_end_{i}",
                            label_visibility="collapsed",
                        )
                    val = (start_n, end_n)
                else:
                    # Single number input for equals/greater than/less than/at least/at most
                    if isinstance(prev_val, (int, float)):
                        single_default_n = float(prev_val)
                    elif isinstance(prev_val, (list, tuple)):
                        # User switched from "between" — collapse to first
                        try:
                            single_default_n = float(prev_val[0])
                        except Exception:
                            single_default_n = data_min_n
                    else:
                        single_default_n = data_min_n
                    val = st.number_input(
                        "Value",
                        value=single_default_n,
                        step=1.0,
                        key=f"flt_val_num_{i}",
                        label_visibility="collapsed",
                        help=f"Range in data: {int(data_min_n)} to {int(data_max_n)}",
                    )
            elif col_is_labels:
                # Labels column → searchable multi-select of labels with counts.
                # Streamlit's st.multiselect supports type-to-search natively:
                # user types "end" and the dropdown filters to matching labels.
                # Display format is "Label (N)" so users see which spellings
                # are common vs rare. We strip the count when applying the filter.
                label_opts = build_labels_options(df, col_choice)
                if not label_opts:
                    st.caption("No labels found in the data.")
                    val = []
                else:
                    prev_val = flt.get("value")
                    # Persist previous selection across reruns when the options
                    # still contain them (defensive against data reloads).
                    if isinstance(prev_val, (list, tuple)):
                        default_selection = [v for v in prev_val if v in label_opts]
                    else:
                        default_selection = []
                    val = st.multiselect(
                        "Labels",
                        label_opts,
                        default=default_selection,
                        key=f"flt_val_labels_{i}",
                        label_visibility="collapsed",
                        placeholder="Type to search labels…",
                    )
            elif op_choice in ("is", "is not"):
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
        if column not in frame.columns or value is None:
            return all_true

        op = (operator or "is").lower().strip()

        # Date branch — runs whenever the column is datetime-typed AND the
        # operator is one of the date operators. Dates compare day-by-day
        # (timestamps normalized to date) so "on 2026-04-15" matches every
        # Jira created on that day regardless of time-of-day.
        if op in ("on", "on or after", "on or before", "between") and pd.api.types.is_datetime64_any_dtype(frame[column]):
            series_dt = pd.to_datetime(frame[column], errors="coerce").dt.date
            try:
                if op == "between":
                    if not isinstance(value, (list, tuple)) or len(value) != 2:
                        return all_true
                    start_d = pd.to_datetime(value[0]).date()
                    end_d = pd.to_datetime(value[1]).date()
                    if start_d > end_d:
                        start_d, end_d = end_d, start_d  # accept reversed input gracefully
                    return (series_dt >= start_d) & (series_dt <= end_d)
                target = pd.to_datetime(value).date()
                if op == "on":
                    return series_dt == target
                if op == "on or after":
                    return series_dt >= target
                if op == "on or before":
                    return series_dt <= target
            except Exception:
                # Bad date input → no-op, all rows pass through
                return all_true
            return all_true

        # Labels branch — runs when the column is Labels and the operator
        # is one of the label operators. Labels cells are list-as-string
        # ("['Endava', 'Cust-MSFT']") so we parse each row into a Python set
        # and check intersection with the user's selected labels (count
        # suffixes stripped first).
        if op in ("has any of", "has none of") and is_labels_column(column):
            if not isinstance(value, (list, tuple)) or not value:
                # No labels selected → don't filter anything (all rows pass)
                return all_true
            # Strip "(N)" suffix to recover the raw label name for matching
            selected_raw = {strip_count_suffix(v).lower() for v in value if v}
            if not selected_raw:
                return all_true

            def _row_has_any(cell_val):
                row_labels_lower = {lbl.lower() for lbl in parse_labels_cell(cell_val)}
                return bool(row_labels_lower & selected_raw)

            row_matches = frame[column].map(_row_has_any)
            if op == "has any of":
                return row_matches
            else:  # "has none of"
                return ~row_matches

        # Numeric branch — runs when the operator is one of the numeric ops.
        # Coerces both the column and the comparison value(s) to numbers,
        # handles NaN gracefully (excludes rows where the column value
        # can't be converted to a number).
        if op in ("equals", "greater than", "less than", "at least", "at most", "between"):
            try:
                series_num = pd.to_numeric(frame[column], errors="coerce")
                if op == "between":
                    if not isinstance(value, (list, tuple)) or len(value) != 2:
                        return all_true
                    low = float(value[0])
                    high = float(value[1])
                    if low > high:
                        low, high = high, low
                    return (series_num >= low) & (series_num <= high)
                target_num = float(value)
                if op == "equals":
                    return series_num == target_num
                if op == "greater than":
                    return series_num > target_num
                if op == "less than":
                    return series_num < target_num
                if op == "at least":
                    return series_num >= target_num
                if op == "at most":
                    return series_num <= target_num
            except (ValueError, TypeError):
                return all_true
            return all_true

        # Text branch (existing logic)
        if value == "":
            return all_true
        series = frame[column].map(normalise_value).astype(str)
        value_str = normalise_value(value) if operator in ("is", "is not") else str(value).strip()
        if not value_str:
            return all_true
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
        # Helper to render the value side of each badge. Handles three cases:
        #   • Date range tuple (2 dates) → "2026-04-01 to 2026-04-30"
        #   • Labels multi-select (list of "Label (N)" strings) → "Endava, Cust-MSFT"
        #   • Single date → "2026-04-15"
        #   • Anything else → string repr
        def format_badge_value(value):
            # Labels-style list — when multiple values are selected, show
            # them comma-separated with the count suffixes stripped for readability
            if isinstance(value, list) and len(value) >= 1:
                # Check if these look like labels (strip count suffix and join)
                stripped = [strip_count_suffix(v) for v in value]
                if len(stripped) <= 3:
                    return ", ".join(stripped)
                # Too many to show in a badge — abbreviate
                return f"{', '.join(stripped[:2])} +{len(stripped) - 2} more"
            # Date range (2 dates)
            if isinstance(value, tuple) and len(value) == 2:
                try:
                    a = pd.to_datetime(value[0]).strftime("%Y-%m-%d")
                    b = pd.to_datetime(value[1]).strftime("%Y-%m-%d")
                    return f"{a} to {b}"
                except Exception:
                    return str(value)
            try:
                # Single date object → format consistently
                if hasattr(value, "strftime"):
                    return value.strftime("%Y-%m-%d")
            except Exception:
                pass
            return str(value)

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
                f'<b>{format_badge_value(val)}</b></span>'
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
                     "Creator", "R&D Engineer", "Creation Date", "Resolution Date", "Age (days)"]:
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

    # Preview table.
    # If the user selected "Issue Key" (or its display variant) as one of
    # the export columns, render that column as a clickable hyperlink to
    # the Jira ticket. The cell still shows the Issue Key text — clicking
    # it opens the Jira ticket in a new tab. Keeps the table compact
    # (no separate link column) while making drill-down one click away.
    JIRA_BASE_URL = "https://jira.corp.adobe.com/browse/"
    preview_df = export_df[chosen_export_cols].copy()

    # Find which of the chosen columns IS the Issue Key column (case-insensitive)
    # so we can apply the LinkColumn config to it dynamically.
    issue_key_col = None
    for c in chosen_export_cols:
        if c.lower().replace(" ", "") in ("issuekey", "issue_key", "key"):
            issue_key_col = c
            break

    column_config = {}
    if issue_key_col and issue_key_col in preview_df.columns:
        # Build hyperlink URLs in-place so LinkColumn can render them.
        # The display text is the original key value; the underlying value
        # is the URL. Streamlit's LinkColumn lets us pass display_text as
        # a callable to keep the key visible.
        preview_df[issue_key_col] = preview_df[issue_key_col].apply(
            lambda k: f"{JIRA_BASE_URL}{k}" if pd.notna(k) and str(k).strip() else None
        )
        # display_text uses regex on the URL to extract the issue key portion
        column_config[issue_key_col] = st.column_config.LinkColumn(
            issue_key_col,
            help="Click to open this Jira in a new tab.",
            display_text=r"https?://jira\.corp\.adobe\.com/browse/(.+)",
        )

    # Format Age (days) nicely if it's in the selected columns
    if "Age (days)" in preview_df.columns:
        column_config["Age (days)"] = st.column_config.NumberColumn(
            "Age (days)",
            help="Days from Creation Date to Resolution Date. For Resolved-lens data, every row has a resolution date so this is always defined.",
            format="%d days",
        )

    st.dataframe(
        preview_df,
        use_container_width=True,
        height=380,
        column_config=column_config,
    )

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

    # ─── Live numbers computed from the currently filtered dataframe ──────
    # The "Worked example" sections used to be hardcoded ("265 Valid + 37
    # Invalid"), which became confusing when leadership saw a chart showing
    # 134 Valid but read an example with 265. We now compute these live so
    # the documentation always matches whatever the user is currently
    # viewing — period dropdown and sidebar filters all flow through.
    _ex_total = len(df)
    _ex_period = reporting_period

    def _ex_pct(num, den):
        return f"{(num / den * 100) if den else 0:.0f}"

    _ex_p1 = int((df["Priority_Short"] == "P1").sum()) if "Priority_Short" in df.columns else 0
    _ex_valid = int((df["Validity"].astype(str).str.strip().str.lower() == "valid").sum()) if "Validity" in df.columns else 0
    _ex_invalid = int((df["Validity"].astype(str).str.strip().str.lower() == "invalid").sum()) if "Validity" in df.columns else 0
    _ex_bug = int(df["Type of Request"].astype(str).str.strip().str.lower().eq("bug").sum()) if "Type of Request" in df.columns else 0
    _ex_rfh = int(df["Type of Request"].astype(str).str.strip().str.lower().eq("request for help").sum()) if "Type of Request" in df.columns else 0

    # Bug ↔ Validity cross for the Type of Request stacked bar example
    _ex_bug_valid = 0
    _ex_bug_invalid = 0
    if "Type of Request" in df.columns and "Validity" in df.columns:
        _bug_mask = df["Type of Request"].astype(str).str.strip().str.lower().eq("bug")
        _ex_bug_valid = int(((_bug_mask) & (df["Validity"].astype(str).str.strip().str.lower() == "valid")).sum())
        _ex_bug_invalid = int(((_bug_mask) & (df["Validity"].astype(str).str.strip().str.lower() == "invalid")).sum())

    # Classification accuracy figures (Quality tab)
    _class_col = "Correct / Incorrect classification? RFH/Bug"
    _ex_bug_correct = _ex_bug_incorrect = _ex_rfh_correct = _ex_rfh_incorrect = 0
    _ex_bug_acc = _ex_rfh_acc = 0
    if _class_col in df.columns and "Type of Request" in df.columns:
        _bm = df["Type of Request"].astype(str).str.strip().str.lower().eq("bug")
        _rm = df["Type of Request"].astype(str).str.strip().str.lower().eq("request for help")
        _cm = df[_class_col].astype(str).str.strip().str.lower().eq("correct")
        _im = df[_class_col].astype(str).str.strip().str.lower().eq("incorrect")
        _ex_bug_correct = int((_bm & _cm).sum())
        _ex_bug_incorrect = int((_bm & _im).sum())
        _ex_rfh_correct = int((_rm & _cm).sum())
        _ex_rfh_incorrect = int((_rm & _im).sum())
        _bt = _ex_bug_correct + _ex_bug_incorrect
        _rt = _ex_rfh_correct + _ex_rfh_incorrect
        _ex_bug_acc = int(_ex_bug_correct / _bt * 100) if _bt else 0
        _ex_rfh_acc = int(_ex_rfh_correct / _rt * 100) if _rt else 0

    # Top affects versions
    _ex_top_versions_text = ""
    if "Affects Version" in df.columns:
        _v = df["Affects Version"].dropna().astype(str).str.strip()
        _v = _v[_v != ""].value_counts().head(3)
        if not _v.empty:
            _ex_top_versions_text = ", ".join(f"{name} has {cnt}" for name, cnt in _v.items())

    # Top root cause
    _ex_top_root_text = ""
    if "Root Cause Resolution" in df.columns:
        _r = df["Root Cause Resolution"].dropna().astype(str).str.strip()
        _r = _r[_r != ""].value_counts().head(2)
        if len(_r) >= 1:
            parts = [f"'{name}' appears {cnt} times" for name, cnt in _r.items()]
            _ex_top_root_text = " and ".join(parts)

    # Top components (split by , ; \n)
    _ex_top_components_text = ""
    if "Component/s" in df.columns:
        _comps = []
        for v in df["Component/s"].dropna():
            for p in re.split(r"[,;\n]+", str(v)):
                p = p.strip()
                if p:
                    _comps.append(p)
        if _comps:
            _cs = pd.Series(_comps).value_counts().head(3)
            _ex_top_components_text = ", ".join(f"{name}: {cnt} issues" for name, cnt in _cs.items())

    # Scope for Deflection
    _ex_scope_text = ""
    if "Scope for Deflection" in df.columns:
        _s = df["Scope for Deflection"].dropna().astype(str).str.strip()
        _s = _s[_s != ""].value_counts()
        if not _s.empty:
            _ex_scope_text = ", ".join(f"{name}: {cnt}" for name, cnt in _s.items())

    # Top R&D engineers
    _ex_top_engineers_text = ""
    if "R&D Engineer" in df.columns:
        _e = df["R&D Engineer"].dropna().astype(str).str.strip()
        _e = _e[_e != ""].value_counts().head(3)
        if not _e.empty:
            _ex_top_engineers_text = ", ".join(f"{name}: {cnt}" for name, cnt in _e.items())

    # Top customers (from JIRA_Created — separate dataset)
    _ex_top_customers_text = ""
    try:
        if df_jira_created is not None and not df_jira_created.empty and "Customer" in df_jira_created.columns:
            _df_jc_local = df_jira_created.copy()
            if "Quarter" in _df_jc_local.columns:
                _df_jc_local["Source Sheet"] = _df_jc_local["Quarter"].astype(str).str.strip()
                _df_jc_local["Quarter Label"] = _df_jc_local["Source Sheet"].map(sheet_to_quarter_label)
            try:
                _df_jc_local = apply_reporting_period(_df_jc_local, reporting_period, period_map)
            except Exception:
                pass
            _c = _df_jc_local["Customer"].dropna().astype(str).str.strip()
            _c = _c[_c != ""].value_counts().head(4)
            if not _c.empty:
                _ex_top_customers_text = ", ".join(f"{name} ({cnt})" for name, cnt in _c.items())
    except Exception:
        pass

    # ─── Card 1: KPI strip ───────────────────────────────────────────────
    _calc_card(
        "KPI strip · top of every page",
        "kpi",
        [
            ("total", "Total issues", dict(
                what="The total number of issues in the current view (after period and sidebar filters).",
                source="Every row in the loaded quarter sheets (e.g. <code>Q1_2026</code>, <code>Q2_2026</code>).",
                how="A simple row count: <code>len(filtered_dataframe)</code>.",
                example=f"In the current view (<b>{_ex_period}</b>), this counts <b>{_ex_total:,}</b> rows.",
            )),
            ("p1", "P1 critical", dict(
                what="Number of issues whose priority is P1, the highest tier (mission-critical).",
                source="<code>Priority</code> column. Long descriptions like 'P1 - Relationship is at risk…' are normalized to short codes (P1, P2, P3, P4) inside the dashboard.",
                how="Counts rows where <code>Priority_Short == 'P1'</code>. The percentage shown is <code>P1_count ÷ total_issues × 100</code>.",
                example=f"In <b>{_ex_period}</b>, <b>{_ex_p1}</b> of {_ex_total:,} issues are tagged P1 → caption shows <b>{_ex_pct(_ex_p1, _ex_total)}% of total</b>.",
            )),
            ("valid", "Valid", dict(
                what="Number of issues marked as a legitimate concern after triage.",
                source="<code>Validity</code> column.",
                how="Counts rows where Validity (trimmed and case-insensitive) equals <code>'valid'</code>. "
                    "Percentage is <code>valid_count ÷ total_issues × 100</code>.",
                example=f"In <b>{_ex_period}</b>, <b>{_ex_valid}</b> of {_ex_total:,} issues are tagged Valid → caption shows <b>{_ex_pct(_ex_valid, _ex_total)}% of total</b>.",
                edge="Rows where Validity is blank, NaN, or any value other than 'Valid' or 'Invalid' are not counted.",
            )),
            ("bugs", "Bugs", dict(
                what="Number of issues classified as a Bug rather than a Request for Help.",
                source="<code>Type of Request</code> column.",
                how="Counts rows where <code>Type of Request</code> (trimmed and case-insensitive) equals <code>'bug'</code>. "
                    "Note: this uses Type of Request directly, NOT the 'Correct / Incorrect classification? RFH/Bug' "
                    "column (which only contains 'Correct' or 'Incorrect' and tells you whether the categorization was right).",
                example=f"In <b>{_ex_period}</b>, <b>{_ex_bug}</b> of {_ex_total:,} issues are tagged Bug → caption shows <b>{_ex_pct(_ex_bug, _ex_total)}% of total</b>.",
            )),
            ("rfh", "RFH", dict(
                what="Number of issues classified as a Request for Help (i.e., not a Bug).",
                source="<code>Type of Request</code> column.",
                how="Counts rows where <code>Type of Request</code> (trimmed and case-insensitive) equals <code>'request for help'</code>. "
                    "Together with the Bugs KPI, these two categories cover the full population of typed issues.",
                example=f"In <b>{_ex_period}</b>, <b>{_ex_rfh}</b> of {_ex_total:,} issues are tagged Request for Help → caption shows <b>{_ex_pct(_ex_rfh, _ex_total)}% of total</b>.",
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
                example=f"<b>{_ex_valid} Valid</b> + <b>{_ex_invalid} Invalid</b> → <b>{_ex_pct(_ex_valid, _ex_valid + _ex_invalid)}% Valid</b> slice (emerald) and <b>{_ex_pct(_ex_invalid, _ex_valid + _ex_invalid)}% Invalid</b> slice (terracotta).",
                edge="Validity values other than 'Valid' or 'Invalid' (blanks, NaN) are excluded from the donut.",
            )),
            ("type", "Type of request (stacked bar)", dict(
                what="Total issues per request type (Bug, Request for Help), with each bar split into Valid + Invalid segments.",
                source="<code>Type of Request</code> and <code>Validity</code> columns.",
                how="Cross-tabulates Type × Validity, then plots a stacked bar with Type on the x-axis. Stack order: Valid first (bottom), then Invalid (top). The y-axis is total count of issues for that type.",
                example=f"Bug: <b>{_ex_bug_valid} Valid</b> + <b>{_ex_bug_invalid} Invalid</b> → emerald {_ex_bug_valid} stacked under terracotta {_ex_bug_invalid} = total bar height <b>{_ex_bug_valid + _ex_bug_invalid}</b>.",
                edge="Rows with blank Type of Request are excluded.",
            )),
            ("version", "Version distribution (bar chart)", dict(
                what="Top affected versions ordered by issue count.",
                source="<code>Affects Version</code> column.",
                how="Counts how many issues mention each version, then keeps the top 8 by count. Numeric versioning is sorted naturally (so 8.6.4 > 8.6.10 lexically would be wrong — we sort by count instead).",
                example=(f"In the current view: {_ex_top_versions_text} — the chart shows them ordered by count." if _ex_top_versions_text else "Versions are sorted by count and the top 8 are shown."),
                edge="A row can list multiple versions in one cell separated by commas — each one is counted independently.",
            )),
            ("rootcause", "Root cause resolution breakdown (bar chart)", dict(
                what="Distribution of root cause categories across all issues.",
                source="<code>Root Cause Resolution</code> column.",
                how="Counts each unique value, sorts by frequency, displays as a horizontal bar chart for readability of long category names.",
                example=(f"In the current view: {_ex_top_root_text} — those become the top bars." if _ex_top_root_text else "The most common root cause categories rise to the top."),
                edge="Rows with blank Root Cause Resolution are excluded.",
            )),
            ("component", "Component distribution (lollipop chart)", dict(
                what="Top components affected by issues, displayed as a Bloomberg-style lollipop chart for visual distinction.",
                source="<code>Component/s</code> column.",
                how="Splits cells with multiple components (separated by commas, semicolons, or newlines), counts each unique component, keeps the top 12.",
                example=(f"In the current view: {_ex_top_components_text}." if _ex_top_components_text else "The 12 most-mentioned components are shown."),
                edge="A row listing 'A, B, C' contributes to all three components' counts. This avoids hiding components that travel together.",
            )),
            ("volume", "Volume trend (line chart)", dict(
                what="Issue creation volume over time, by month.",
                source="<code>Resolution Date</code> column.",
                how="Groups issues by month of resolution, counts per month, plots as a line chart with steel-blue area fill below.",
                example="Each month becomes a point on the line — for example, Jan 2026: 32 issues, Feb 2026: 28 issues. The chart's hover tooltip shows the exact count per month.",
                edge="Rows with missing Resolution Date are excluded.",
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
                example=f"In the current view — Bug: <b>{_ex_bug_correct} Correct + {_ex_bug_incorrect} Incorrect</b> = <b>{_ex_bug_acc}% accuracy</b>. RFH: <b>{_ex_rfh_correct} Correct + {_ex_rfh_incorrect} Incorrect</b> = <b>{_ex_rfh_acc}% accuracy</b>.",
                edge="Rows where the classification cell is blank or has any value other than 'Correct'/'Incorrect' are excluded from this chart.",
            )),
            ("scope", "Scope for deflection (bar chart)", dict(
                what="Counts of issues categorized as deflectable (Yes), not deflectable (No), or partially (Yes/No).",
                source="<code>Scope for Deflection</code> column.",
                how="Direct value counts of the Scope for Deflection column, plotted as a single horizontal bar chart.",
                example=(f"In the current view: {_ex_scope_text} — the chart shows these in descending order." if _ex_scope_text else "Bars are sorted from most-frequent value to least."),
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
                example=(f"In the current view: {_ex_top_engineers_text}." if _ex_top_engineers_text else "The 12 engineers with the most assigned issues are shown."),
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
                example=f"Selecting a single quarter shows only the rows that came from that quarter's sheet. Currently <b>{_ex_period}</b> is selected, showing <b>{_ex_total:,}</b> rows.",
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