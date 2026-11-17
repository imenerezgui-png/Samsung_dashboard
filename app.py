"""
Samsung Campaigns Dashboard — KPI overview, highlights & take-aways.
Samsung-blue theme, powered by Ykone.
"""

from __future__ import annotations

import base64
import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# Page config  (must be first Streamlit call)
# ─────────────────────────────────────────────────────────────────────────────
APP_DIR = Path(__file__).parent
DATA_FILE = APP_DIR / "Samsung-Campaigns-Aug-1-2026-Nov-5-2026.xlsx"
SAMSUNG_LOGO = APP_DIR / "samsung_logo.png"
YKONE_LOGO = APP_DIR / "ykone_logo.jpg"

st.set_page_config(
    page_title="Samsung Campaigns Dashboard",
    page_icon=str(SAMSUNG_LOGO) if SAMSUNG_LOGO.exists() else "📱",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# Theme — Samsung blue / white
# ─────────────────────────────────────────────────────────────────────────────
BLUE      = "#1428A0"   # Samsung brand blue
BLUE_DARK = "#0A1663"
BLUE_SOFT = "#3B52C4"
SKY       = "#5B7BD5"
ACCENT    = "#00B3E3"
BG        = "#F4F6FB"
CARD      = "#FFFFFF"
INK       = "#101936"
DIM       = "#5A6689"
GOOD      = "#0FA573"
WARN      = "#E2574C"

PALETTE = [BLUE, ACCENT, SKY, "#8A9BE0", "#0FA573", "#E2574C", "#F2A93B"]

st.markdown(
    f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=IBM+Plex+Sans:wght@400;500;600&display=swap');

html, body, [class*="css"], .stApp {{
    background: {BG};
    color: {INK};
    font-family: 'IBM Plex Sans', 'Segoe UI', sans-serif;
}}
h1, h2, h3, h4 {{
    font-family: 'Space Grotesk', 'Segoe UI', sans-serif !important;
    color: {INK};
}}
section[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, {BLUE_DARK} 0%, {BLUE} 100%) !important;
}}
section[data-testid="stSidebar"] * {{
    color: #FFFFFF !important;
}}
section[data-testid="stSidebar"] .stMultiSelect div[data-baseweb="select"] > div,
section[data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"] > div {{
    background: rgba(255,255,255,0.10) !important;
    border-color: rgba(255,255,255,0.35) !important;
}}
section[data-testid="stSidebar"] span[data-baseweb="tag"] {{
    background: {ACCENT} !important;
}}

.hero {{
    background: linear-gradient(120deg, {BLUE_DARK} 0%, {BLUE} 55%, {BLUE_SOFT} 100%);
    border-radius: 18px;
    padding: 1.6rem 2rem;
    margin-bottom: 1.2rem;
    overflow: hidden;
    box-shadow: 0 14px 34px rgba(20, 40, 160, 0.28);
}}
.hero-wordmark {{
    display: block;
    width: 100%;
    height: auto;
    margin-bottom: 1.1rem;
    filter: drop-shadow(0 3px 8px rgba(5, 10, 40, 0.35));
}}
.hero-row {{
    display: flex;
    flex-wrap: wrap;
    gap: 1rem;
    align-items: center;
    justify-content: space-between;
}}
.hero-title {{
    color: #FFFFFF;
    font-family: 'Space Grotesk', sans-serif;
    font-size: 1.9rem;
    font-weight: 700;
    letter-spacing: 0.02em;
    margin: 0;
}}
.hero-sub {{
    color: rgba(255,255,255,0.78);
    font-size: 0.95rem;
    margin-top: 0.3rem;
}}
.hero-logos {{
    display: flex;
    align-items: center;
    gap: 1.1rem;
    flex-shrink: 0;
}}
.hero-logos img.ykone {{
    height: 46px; border-radius: 10px; border: 2px solid rgba(255,255,255,0.65);
}}
.hero-x {{
    color: rgba(255,255,255,0.7); font-size: 1.3rem; font-weight: 600;
}}

.kpi-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
    gap: 0.9rem;
    margin-bottom: 0.4rem;
}}
.kpi-card {{
    background: {CARD};
    border-radius: 14px;
    padding: 1rem 1.1rem;
    border-top: 4px solid {BLUE};
    box-shadow: 0 6px 18px rgba(16, 25, 54, 0.07);
}}
.kpi-card.alt   {{ border-top-color: {ACCENT}; }}
.kpi-card.money {{ border-top-color: {GOOD}; }}
.kpi-label {{
    font-size: 0.72rem;
    letter-spacing: 0.09em;
    text-transform: uppercase;
    color: {DIM};
    margin-bottom: 0.35rem;
}}
.kpi-value {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 1.55rem;
    font-weight: 700;
    color: {INK};
    line-height: 1.1;
}}
.kpi-foot {{
    font-size: 0.74rem;
    color: {DIM};
    margin-top: 0.3rem;
}}

.insight-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(290px, 1fr));
    gap: 0.9rem;
    margin: 0.3rem 0 0.8rem 0;
}}
.insight-card {{
    background: {CARD};
    border-radius: 14px;
    padding: 1rem 1.15rem;
    border-left: 5px solid {BLUE};
    box-shadow: 0 6px 18px rgba(16, 25, 54, 0.07);
}}
.insight-card.good {{ border-left-color: {GOOD}; }}
.insight-card.warn {{ border-left-color: {WARN}; }}
.insight-card.sky  {{ border-left-color: {ACCENT}; }}
.insight-tag {{
    display: inline-block;
    font-size: 0.68rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: {BLUE};
    background: rgba(20, 40, 160, 0.08);
    border-radius: 20px;
    padding: 0.15rem 0.6rem;
    margin-bottom: 0.45rem;
}}
.insight-card.good .insight-tag {{ color: {GOOD}; background: rgba(15,165,115,0.10); }}
.insight-card.warn .insight-tag {{ color: {WARN}; background: rgba(226,87,76,0.10); }}
.insight-card.sky  .insight-tag {{ color: #0387AB; background: rgba(0,179,227,0.10); }}
.insight-title {{
    font-family: 'Space Grotesk', sans-serif;
    font-weight: 700;
    font-size: 1.0rem;
    margin-bottom: 0.3rem;
}}
.insight-body {{
    font-size: 0.86rem;
    color: {DIM};
    line-height: 1.45;
}}
.insight-body b {{ color: {INK}; }}

.section-title {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 1.12rem;
    font-weight: 700;
    color: {INK};
    margin: 1.1rem 0 0.5rem 0;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}}
.section-title::before {{
    content: "";
    display: inline-block;
    width: 10px; height: 10px;
    background: {BLUE};
    border-radius: 3px;
}}

div[data-testid="stDataFrame"] {{
    background: {CARD};
    border-radius: 12px;
    padding: 0.4rem;
    box-shadow: 0 6px 18px rgba(16, 25, 54, 0.07);
}}
div[data-testid="stPlotlyChart"] {{
    background: {CARD};
    border-radius: 16px;
    padding: 0.8rem 0.6rem 0.4rem 0.6rem;
    box-shadow: 0 6px 18px rgba(16, 25, 54, 0.07);
    overflow: hidden;
}}
div[data-testid="stPlotlyChart"] .js-plotly-plot,
div[data-testid="stPlotlyChart"] .plotly,
div[data-testid="stPlotlyChart"] svg {{
    border-radius: 12px;
}}
.gauge-card {{
    background: {CARD};
    border-radius: 14px;
    padding: 1rem 1.1rem 1.1rem 1.1rem;
    border-top: 4px solid {BLUE};
    box-shadow: 0 6px 18px rgba(16, 25, 54, 0.07);
    display: flex;
    flex-direction: column;
    align-items: center;
    height: 100%;
}}
.gauge-wrap {{
    position: relative;
    width: 100%;
    max-width: 260px;
    margin: 0.4rem auto 0 auto;
}}
.gauge-wrap svg {{ width: 100%; display: block; }}
.gauge-needle {{
    position: absolute;
    left: 50%;
    bottom: 9%;
    height: 56%;
    transform-origin: 50% 100%;
    filter: drop-shadow(0 2px 5px rgba(16, 25, 54, 0.35));
    transition: transform 0.6s ease;
}}
.gauge-value {{
    font-family: 'Space Grotesk', sans-serif;
    font-size: 1.6rem;
    font-weight: 700;
    color: {INK};
    margin-top: 0.5rem;
    line-height: 1.1;
}}
.gauge-sub {{
    font-size: 0.78rem;
    color: {DIM};
    text-align: center;
    margin-top: 0.2rem;
}}
.gauge-sub b {{ color: {BLUE}; }}

footer {{ visibility: hidden; }}
</style>
""",
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────────────────────────────────────
# Campaign-name parsing
# ─────────────────────────────────────────────────────────────────────────────
COUNTRY_MAP = {"TN": "Tunisia", "LYB": "Libya"}

OBJECTIVE_PATTERNS = [
    (r"awareness", "Awareness"),
    (r"reach", "Reach"),
    (r"engagement|\beng\b", "Engagement"),
    (r"conversion|checkout|purchase", "Conversion"),
    (r"lead gen|inscription|leads?\b.*-|messenger", "Lead Gen"),
    (r"thruplays?", "ThruPlays"),
    (r"visionnage|video views|views", "Video Views"),
]


def parse_country(name: str) -> str:
    m = re.match(r"\s*(TN|LYB)\b", name, flags=re.IGNORECASE)
    if m:
        return COUNTRY_MAP[m.group(1).upper()]
    return "Other"


def parse_budget(name: str) -> float:
    """Sum every amount attached to a $ sign (or USD) in the campaign name.

    Handles '2000$', '500 $', '+50$', '333 USD' and spaced thousands '1 066$'.
    """
    total = 0.0
    for m in re.finditer(r"(\d{1,3}(?:[ \u00a0]\d{3})+|\d+(?:[.,]\d+)?)\s*(?:\$|USD\b)", name, flags=re.IGNORECASE):
        raw = m.group(1).replace("\u00a0", " ").replace(" ", "").replace(",", ".")
        total += float(raw)
    return total


def parse_platform(name: str) -> str:
    has_fb = re.search(r"\bFB\b", name, flags=re.IGNORECASE) is not None
    has_ig = re.search(r"\bIG\b", name, flags=re.IGNORECASE) is not None
    if has_fb and has_ig:
        return "FB & IG"
    if has_fb:
        return "Facebook"
    if has_ig:
        return "Instagram"
    return "Mixed / N.C."


def parse_objective(name: str) -> str:
    low = name.lower()
    for pattern, label in OBJECTIVE_PATTERNS:
        if re.search(pattern, low):
            return label
    return "Other"


def clean_label(name: str) -> str:
    """Short readable label: drop country prefix and budget chunks."""
    label = re.sub(r"^\s*(TN|LYB)\s*-\s*", "", name, flags=re.IGNORECASE)
    label = re.sub(r"-?\s*\d[\d \u00a0.,]*[Kk]?\s*(?:\$|USD\b|DT\b)[^-]*", "", label, flags=re.IGNORECASE)
    label = re.sub(r"\s*-\s*$", "", label.strip())
    label = re.sub(r"\s{2,}", " ", label)
    return label or name


# ─────────────────────────────────────────────────────────────────────────────
# Data loading
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data() -> tuple[pd.DataFrame, pd.Series | None]:
    raw = pd.read_excel(DATA_FILE)
    raw.columns = [c.strip() for c in raw.columns]

    # The row without a campaign name is the account-level total (deduplicated reach).
    totals_mask = raw["Campaign name"].isna()
    account_totals = raw[totals_mask].iloc[0] if totals_mask.any() else None

    df = raw[~totals_mask].copy()
    df["Campaign name"] = df["Campaign name"].astype(str).str.strip()
    df["Country"] = df["Campaign name"].map(parse_country)
    df["Budget (USD)"] = df["Campaign name"].map(parse_budget)
    df["Platform"] = df["Campaign name"].map(parse_platform)
    df["Objective"] = df["Campaign name"].map(parse_objective)
    df["Campaign"] = df["Campaign name"].map(clean_label)

    num_cols = [
        "Reach", "Impressions", "Frequency", "Post engagements", "Post comments",
        "Post reactions", "Post shares", "Post saves", "15-second ThruPlays",
        "Ad recall lift", "CPM (cost per 1,000 impressions) (USD)", "CPC (all) (USD)",
    ]
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Derived metrics (spend reconstructed from CPM × impressions).
    df["Spend (USD)"] = df["Impressions"] * df["CPM (cost per 1,000 impressions) (USD)"] / 1000
    df["Clicks (est.)"] = (df["Spend (USD)"] / df["CPC (all) (USD)"]).where(df["CPC (all) (USD)"] > 0)
    df["Engagement rate %"] = (df["Post engagements"] / df["Impressions"] * 100).round(2)
    return df, account_totals


def fmt_num(value: float | None, decimals: int = 0) -> str:
    if value is None or pd.isna(value):
        return "—"
    if abs(value) >= 1_000_000:
        return f"{value / 1_000_000:,.1f}M"
    if abs(value) >= 10_000:
        return f"{value / 1_000:,.1f}K"
    return f"{value:,.{decimals}f}"


def fmt_money(value: float | None, decimals: int = 0) -> str:
    if value is None or pd.isna(value):
        return "—"
    return f"${value:,.{decimals}f}"


try:
    data, account_totals = load_data()
except FileNotFoundError:
    st.error(f"Data file not found: `{DATA_FILE.name}` — place the Excel export next to app.py.")
    st.stop()


def img_b64(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode()


# ─────────────────────────────────────────────────────────────────────────────
# Header — Samsung × Ykone
# ─────────────────────────────────────────────────────────────────────────────
period = ""
if account_totals is not None:
    start = pd.to_datetime(account_totals["Reporting starts"]).strftime("%b %d, %Y")
    end = pd.to_datetime(account_totals["Reporting ends"]).strftime("%b %d, %Y")
    period = f"{start} → {end}"

WORDMARK = APP_DIR / "samsung_wordmark.png"

wordmark_html = ""
if WORDMARK.exists():
    wordmark_html = (
        f'<img class="hero-wordmark" alt="SAMSUNG" '
        f'src="data:image/png;base64,{img_b64(WORDMARK)}"/>'
    )

logos_html = ""
if YKONE_LOGO.exists():
    logos_html = f'<img class="ykone" src="data:image/jpeg;base64,{img_b64(YKONE_LOGO)}" alt="Ykone"/>'

st.markdown(
    f"""
<div class="hero">
  {wordmark_html}
  <div class="hero-row">
    <div>
      <p class="hero-title">Samsung Campaigns Dashboard</p>
      <div class="hero-sub">Paid social performance · {period} · Tunisia &amp; Libya · by Ykone</div>
    </div>
    <div class="hero-logos">{logos_html}</div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────────────────────────────────────────
# Sidebar filters
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    if YKONE_LOGO.exists():
        st.image(str(YKONE_LOGO), width=110)
    st.markdown("### Filters")
    countries = st.multiselect(
        "Country", sorted(data["Country"].unique()), default=sorted(data["Country"].unique())
    )
    platforms = st.multiselect(
        "Platform", sorted(data["Platform"].unique()), default=sorted(data["Platform"].unique())
    )
    objectives = st.multiselect(
        "Objective", sorted(data["Objective"].unique()), default=sorted(data["Objective"].unique())
    )
    st.markdown("---")
    st.caption("Budgets, countries & objectives are parsed automatically from campaign names (TN → Tunisia, LYB → Libya, amounts in $).")

flt = data[
    data["Country"].isin(countries)
    & data["Platform"].isin(platforms)
    & data["Objective"].isin(objectives)
]

if flt.empty:
    st.warning("No campaign matches the current filters.")
    st.stop()

all_selected = len(flt) == len(data)

# ─────────────────────────────────────────────────────────────────────────────
# KPI cards
# ─────────────────────────────────────────────────────────────────────────────
tot_budget = flt["Budget (USD)"].sum()
tot_spend = flt["Spend (USD)"].sum()
tot_impr = flt["Impressions"].sum()
tot_eng = flt["Post engagements"].sum()
tot_thru = flt["15-second ThruPlays"].sum()
tot_recall = flt["Ad recall lift"].sum()

# Use the account-level deduplicated reach when nothing is filtered out.
if all_selected and account_totals is not None:
    tot_reach = account_totals["Reach"]
    reach_note = "deduplicated, account level"
else:
    tot_reach = flt["Reach"].sum()
    reach_note = "sum of campaigns"

w_cpm = tot_spend / tot_impr * 1000 if tot_impr else float("nan")
clicks = flt["Clicks (est.)"].sum()
w_cpc = tot_spend / clicks if clicks else float("nan")
eng_rate = tot_eng / tot_impr * 100 if tot_impr else float("nan")

kpis = [
    ("Campaigns", f"{len(flt)}", f"{flt['Country'].nunique()} market(s)", ""),
    ("Planned budget", fmt_money(tot_budget), "parsed from campaign names", "money"),
    ("Delivered spend", fmt_money(tot_spend), "CPM × impressions", "money"),
    ("Reach", fmt_num(tot_reach), reach_note, ""),
    ("Impressions", fmt_num(tot_impr), f"avg freq {flt['Frequency'].mean():.1f}", ""),
    ("Engagements", fmt_num(tot_eng), f"{eng_rate:.2f}% of impressions", "alt"),
    ("ThruPlays (15s)", fmt_num(tot_thru), "video completions", "alt"),
    ("Avg CPM", fmt_money(w_cpm, 3), "spend-weighted", "alt"),
    ("Avg CPC", fmt_money(w_cpc, 3), "spend-weighted", "alt"),
    ("Ad recall lift", fmt_num(tot_recall), "estimated people", ""),
]

cards = "".join(
    f'<div class="kpi-card {cls}"><div class="kpi-label">{label}</div>'
    f'<div class="kpi-value">{value}</div><div class="kpi-foot">{foot}</div></div>'
    for label, value, foot, cls in kpis
)

NEEDLE_IMG = APP_DIR / "s26fe_needle.png"


def render_gauge(spent: float, budget: float) -> str:
    """Speedometer of delivered spend vs total budget; the S26 Fe is the needle."""
    pct = (spent / budget * 100) if budget else 0.0
    frac = min(max(pct / 100, 0.0), 1.0)
    angle = -90 + frac * 180            # -90° = 0%, +90° = 100%
    arc_len = 3.14159 * 80              # semicircle radius 80
    dash = frac * arc_len
    needle = (
        f'<img class="gauge-needle" alt="Samsung S26 Fe needle" '
        f'style="transform: translateX(-50%) rotate({angle:.1f}deg);" '
        f'src="data:image/png;base64,{img_b64(NEEDLE_IMG)}"/>'
        if NEEDLE_IMG.exists() else ""
    )
    return f"""
<div class="gauge-card">
  <div class="kpi-label">Budget speedometer</div>
  <div class="gauge-wrap">
    <svg viewBox="0 0 200 112">
      <defs>
        <linearGradient id="gfill" x1="0%" y1="0%" x2="100%" y2="0%">
          <stop offset="0%" stop-color="{ACCENT}"/>
          <stop offset="100%" stop-color="{BLUE}"/>
        </linearGradient>
      </defs>
      <path d="M 20 100 A 80 80 0 0 1 180 100" fill="none" stroke="#E3E8F5"
            stroke-width="13" stroke-linecap="round"/>
      <path d="M 20 100 A 80 80 0 0 1 180 100" fill="none" stroke="url(#gfill)"
            stroke-width="13" stroke-linecap="round"
            stroke-dasharray="{dash:.1f} {arc_len:.1f}"/>
      <text x="16" y="112" font-size="9" font-weight="bold" fill="{BLUE}">0%</text>
      <text x="92" y="14" font-size="9" font-weight="bold" fill="{BLUE}">50%</text>
      <text x="166" y="112" font-size="9" font-weight="bold" fill="{BLUE}">100%</text>
      <circle cx="100" cy="100" r="5" fill="{BLUE}"/>
    </svg>
    {needle}
  </div>
  <div class="gauge-value">{fmt_money(spent)}</div>
  <div class="gauge-sub"><b>{pct:.1f}%</b> of the {fmt_money(budget)} total budget spent</div>
</div>
"""


col_kpi, col_gauge = st.columns([2.4, 1], gap="medium")
with col_kpi:
    st.markdown(f'<div class="kpi-grid">{cards}</div>', unsafe_allow_html=True)
with col_gauge:
    st.markdown(render_gauge(tot_spend, tot_budget), unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# Highlights & take-aways (rendered at the bottom of the page)
# ─────────────────────────────────────────────────────────────────────────────
def build_insights(d: pd.DataFrame) -> list[tuple[str, str, str, str]]:
    out: list[tuple[str, str, str, str]] = []

    top_eng = d.loc[d["Post engagements"].idxmax()]
    out.append((
        "good", "Top engagement",
        top_eng["Campaign"],
        f"<b>{fmt_num(top_eng['Post engagements'])}</b> engagements in {top_eng['Country']} "
        f"({top_eng['Engagement rate %']:.1f}% of impressions) — the strongest interaction driver of the flight.",
    ))

    paid = d[d["Spend (USD)"] > 0]
    if not paid.empty:
        best_cpm = paid.loc[paid["CPM (cost per 1,000 impressions) (USD)"].idxmin()]
        out.append((
            "sky", "Most efficient CPM",
            best_cpm["Campaign"],
            f"CPM of <b>{fmt_money(best_cpm['CPM (cost per 1,000 impressions) (USD)'], 3)}</b> in {best_cpm['Country']} — "
            f"cheapest visibility of the period ({fmt_num(best_cpm['Impressions'])} impressions).",
        ))
        best_cpc = paid[paid["CPC (all) (USD)"] > 0]
        if not best_cpc.empty:
            row = best_cpc.loc[best_cpc["CPC (all) (USD)"].idxmin()]
            out.append((
                "sky", "Cheapest clicks",
                row["Campaign"],
                f"CPC of <b>{fmt_money(row['CPC (all) (USD)'], 4)}</b> ({row['Country']}, {row['Platform']}) — "
                "a benchmark to replicate on traffic objectives.",
            ))

    by_country = d.groupby("Country").agg(
        budget=("Budget (USD)", "sum"), impr=("Impressions", "sum"), eng=("Post engagements", "sum")
    )
    if len(by_country) > 1:
        lead = by_country["impr"].idxmax()
        share = by_country.loc[lead, "impr"] / by_country["impr"].sum() * 100
        out.append((
            "", "Market split",
            f"{lead} leads delivery",
            f"<b>{lead}</b> concentrates <b>{share:.0f}%</b> of impressions with "
            f"{fmt_money(by_country.loc[lead, 'budget'])} planned budget vs "
            f"{fmt_money(by_country.drop(lead)['budget'].sum())} for the other market(s).",
        ))

    thru = d[d["15-second ThruPlays"].notna() & (d["15-second ThruPlays"] > 0)]
    if not thru.empty:
        row = thru.loc[thru["15-second ThruPlays"].idxmax()]
        out.append((
            "good", "Video champion",
            row["Campaign"],
            f"<b>{fmt_num(row['15-second ThruPlays'])}</b> 15-second ThruPlays — "
            "video creatives are retaining attention; prioritise this format for upcoming launches.",
        ))

    recall = d[d["Ad recall lift"] > 0]
    if not recall.empty:
        row = recall.loc[recall["Ad recall lift"].idxmax()]
        out.append((
            "", "Brand memory",
            row["Campaign"],
            f"Estimated <b>{fmt_num(row['Ad recall lift'])}</b> people likely to remember the ad — "
            "the flight is building brand salience beyond raw reach.",
        ))

    over = d[(d["Budget (USD)"] > 0) & (d["Spend (USD)"] > d["Budget (USD)"] * 1.15)]
    under = d[(d["Budget (USD)"] > 0) & (d["Spend (USD)"] < d["Budget (USD)"] * 0.85)]
    if not over.empty:
        row = over.loc[(over["Spend (USD)"] - over["Budget (USD)"]).idxmax()]
        out.append((
            "warn", "Watch pacing",
            row["Campaign"],
            f"Delivered spend <b>{fmt_money(row['Spend (USD)'])}</b> vs <b>{fmt_money(row['Budget (USD)'])}</b> parsed budget "
            "(+15% or more) — double-check budget caps or name labelling.",
        ))
    elif not under.empty:
        row = under.loc[(under["Budget (USD)"] - under["Spend (USD)"]).idxmax()]
        out.append((
            "warn", "Under-delivery",
            row["Campaign"],
            f"Only <b>{fmt_money(row['Spend (USD)'])}</b> of <b>{fmt_money(row['Budget (USD)'])}</b> planned budget delivered — "
            "room to push remaining budget before flight end.",
        ))

    return out[:6]


# ─────────────────────────────────────────────────────────────────────────────
# Charts
# ─────────────────────────────────────────────────────────────────────────────
CHART_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="IBM Plex Sans, Segoe UI, sans-serif", color=INK),
    margin=dict(l=10, r=10, t=42, b=10),
    legend=dict(orientation="h", y=-0.18, font=dict(color=BLUE, weight="bold")),
    barcornerradius=7,
)

AXIS_FONT = dict(color=BLUE, weight="bold")


def style_axes(fig: go.Figure) -> go.Figure:
    """Blue, bold tick labels and axis titles on every axis."""
    fig.update_xaxes(tickfont=AXIS_FONT, title_font=AXIS_FONT)
    fig.update_yaxes(tickfont=AXIS_FONT, title_font=AXIS_FONT)
    return fig

col_a, col_b = st.columns([1, 1])

with col_a:
    st.markdown('<div class="section-title">Planned budget by market</div>', unsafe_allow_html=True)
    by_country = flt.groupby("Country", as_index=False)["Budget (USD)"].sum()
    fig = px.pie(
        by_country, names="Country", values="Budget (USD)",
        hole=0.55, color_discrete_sequence=[BLUE, ACCENT, SKY],
    )
    fig.update_traces(textinfo="label+percent", textfont_size=13)
    fig.update_layout(**CHART_LAYOUT, showlegend=False, height=330)
    st.plotly_chart(fig, use_container_width=True)

with col_b:
    st.markdown('<div class="section-title">Budget vs delivered spend by objective</div>', unsafe_allow_html=True)
    by_obj = flt.groupby("Objective", as_index=False).agg(
        Budget=("Budget (USD)", "sum"), Spend=("Spend (USD)", "sum")
    ).sort_values("Budget", ascending=False)
    fig = go.Figure()
    fig.add_bar(x=by_obj["Objective"], y=by_obj["Budget"], name="Planned budget", marker_color=BLUE)
    fig.add_bar(x=by_obj["Objective"], y=by_obj["Spend"], name="Delivered spend", marker_color=ACCENT)
    fig.update_layout(**CHART_LAYOUT, barmode="group", height=330, yaxis_title="USD")
    st.plotly_chart(style_axes(fig), use_container_width=True)

st.markdown('<div class="section-title">Campaign delivery — impressions &amp; engagement rate</div>', unsafe_allow_html=True)
top = flt.sort_values("Impressions", ascending=False)
fig = go.Figure()
fig.add_bar(
    x=top["Campaign"], y=top["Impressions"], name="Impressions",
    marker_color=[BLUE if c == "Tunisia" else ACCENT for c in top["Country"]],
    customdata=top[["Country", "Budget (USD)"]],
    hovertemplate="<b>%{x}</b><br>%{customdata[0]}<br>Impressions: %{y:,.0f}<br>Budget: $%{customdata[1]:,.0f}<extra></extra>",
)
fig.add_scatter(
    x=top["Campaign"], y=top["Engagement rate %"], name="Engagement rate %",
    yaxis="y2", mode="lines+markers", line=dict(color=WARN, width=2),
)
fig.update_layout(
    **CHART_LAYOUT, height=460,
    yaxis=dict(title="Impressions"),
    yaxis2=dict(title="Engagement rate %", overlaying="y", side="right", showgrid=False,
                tickfont=AXIS_FONT, title_font=AXIS_FONT),
    xaxis=dict(tickangle=-38, tickfont=dict(size=10, color=BLUE, weight="bold")),
)
# Long angled labels: park the legend above the plot instead of below it.
fig.update_layout(legend=dict(orientation="h", y=1.12, x=0, font=dict(color=BLUE, weight="bold")))
st.plotly_chart(style_axes(fig), use_container_width=True)

col_c, col_d = st.columns([1, 1])

with col_c:
    st.markdown('<div class="section-title">Cost per 1,000 impressions — cheapest first</div>', unsafe_allow_html=True)
    cpm_col = "CPM (cost per 1,000 impressions) (USD)"
    eff = flt[flt[cpm_col] > 0].sort_values(cpm_col, ascending=True).head(12)
    def _short(s: str) -> str:
        return s if len(s) <= 34 else s[:20].rstrip() + "…" + s[-12:]

    short, seen = [], {}
    for name in eff["Campaign"]:
        lbl = _short(name)
        n = seen.get(lbl, 0)
        seen[lbl] = n + 1
        short.append(lbl + "\u200b" * n)  # keep duplicate labels as distinct categories
    fig = go.Figure()
    fig.add_bar(
        x=eff[cpm_col], y=short, orientation="h",
        marker_color=[BLUE if c == "Tunisia" else ACCENT for c in eff["Country"]],
        text=[f"${v:.3f}" for v in eff[cpm_col]],
        textposition="outside",
        textfont=dict(color=BLUE, weight="bold"),
        customdata=eff[["Country", "Impressions", "Campaign"]],
        hovertemplate="<b>%{customdata[2]}</b><br>%{customdata[0]}<br>CPM: $%{x:.3f}<br>Impressions: %{customdata[1]:,.0f}<extra></extra>",
    )
    fig.update_layout(
        **CHART_LAYOUT, height=380, showlegend=False,
        xaxis_title="CPM (USD) — lower is better",
        yaxis=dict(autorange="reversed"),
        xaxis=dict(range=[0, eff[cpm_col].max() * 1.3], tickangle=0),
    )
    fig.update_yaxes(tickfont=dict(size=10, color=BLUE, weight="bold"))
    fig.update_layout(margin=dict(r=52))
    fig.update_traces(cliponaxis=False)
    st.plotly_chart(style_axes(fig), use_container_width=True)

with col_d:
    st.markdown('<div class="section-title">Engagement mix</div>', unsafe_allow_html=True)
    mix = flt[["Post reactions", "Post comments", "Post shares", "Post saves"]].sum()
    fig = px.bar(
        x=mix.values, y=["Reactions", "Comments", "Shares", "Saves"],
        orientation="h", color=["Reactions", "Comments", "Shares", "Saves"],
        color_discrete_sequence=PALETTE,
    )
    fig.update_layout(**CHART_LAYOUT, height=360, showlegend=False, xaxis_title="Total", yaxis_title="")
    st.plotly_chart(style_axes(fig), use_container_width=True)

# ─────────────────────────────────────────────────────────────────────────────
# Detail table
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Campaign details</div>', unsafe_allow_html=True)
detail = flt[[
    "Campaign", "Country", "Platform", "Objective", "Budget (USD)", "Spend (USD)",
    "Reach", "Impressions", "Frequency", "Post engagements", "Engagement rate %",
    "15-second ThruPlays", "Ad recall lift",
    "CPM (cost per 1,000 impressions) (USD)", "CPC (all) (USD)",
]].sort_values("Impressions", ascending=False)
st.dataframe(
    detail,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Budget (USD)": st.column_config.NumberColumn(format="$%.0f"),
        "Spend (USD)": st.column_config.NumberColumn(format="$%.0f"),
        "Frequency": st.column_config.NumberColumn(format="%.2f"),
        "CPM (cost per 1,000 impressions) (USD)": st.column_config.NumberColumn("CPM (USD)", format="$%.3f"),
        "CPC (all) (USD)": st.column_config.NumberColumn("CPC (USD)", format="$%.4f"),
    },
)

st.markdown('<div class="section-title">Highlights &amp; take-aways</div>', unsafe_allow_html=True)
insights_html = "".join(
    f'<div class="insight-card {cls}"><span class="insight-tag">{tag}</span>'
    f'<div class="insight-title">{title}</div><div class="insight-body">{body}</div></div>'
    for cls, tag, title, body in build_insights(flt)
)
st.markdown(f'<div class="insight-grid">{insights_html}</div>', unsafe_allow_html=True)

st.caption("Samsung Campaigns Dashboard · data: Meta export Aug 1 – Nov 5, 2026 · crafted by Ykone.")
