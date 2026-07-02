# -*- coding: utf-8 -*-
"""
Executive Company Intelligence Dashboard
=========================================
A self-adapting Streamlit dashboard that reads an Excel workbook, auto-detects
its structure (companies, products, categories/tags, cities, regions, status,
dates, ids), cleans it, and renders an executive-grade analytics experience.

Run with:  streamlit run app.py
"""

from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# --------------------------------------------------------------------------------------
# PAGE CONFIG & GLOBAL STYLE
# --------------------------------------------------------------------------------------

st.set_page_config(
    page_title="قطاع الشركات",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

PALETTE = {
    "ink": "#0B2540",        # deep navy - primary text / headers
    "bg": "#F7F9FB",         # page background
    "card": "#7DBCFF",
    "accent": "#C9A24B",     # muted gold - executive accent
    "accent2": "#1B6F8C",    # teal - secondary accent
    "positive": "#2E8B57",
    "negative": "#1A6078",
    "warning": "#D98E04",
    "muted": "#6B7B8C",
    "grid": "#1B6F8C",
}

CHART_SEQUENCE = ["#1B6F8C", "#C9A24B", "#0B2540", "#5FA8B8", "#D98E04", "#7C9885", "#B0413E", "#005EFF"]

CUSTOM_CSS = f"""
<style>
    html, body, [class*="css"] {{
        font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
    }}
    .stApp {{
        background-color: {PALETTE['bg']};
    }}
    #MainMenu, footer {{visibility: visible;}}

    .dash-header {{
    padding: 1.4rem 1.8rem;
    background: linear-gradient(135deg, #1B6F8C 0%, #0B2540 100%);
    border-radius: 14px;
    margin-bottom: 1.4rem;
    box-shadow: 0 6px 18px rgba(192,57,43,0.25); /* red-tinted shadow */
    }}

    .dash-header h1 {{
        color: white;
        font-size: 1.65rem;
        font-weight: 700;
        margin: 0;
        letter-spacing: 0.2px;
    }}
    .dash-header p {{
        color: #CBD8E4;
        margin: 0.25rem 0 0 0;
        font-size: 0.92rem;
    }}

    .kpi-card {{
        background: {PALETTE['card']};
        border-radius: 12px;
        padding: 1rem 1.1rem;
        border: 1px solid {PALETTE['grid']};
        box-shadow: 0 2px 10px rgba(11,37,64,0.05);
        height: 108px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: transform 0.15s ease;
    }}
    .kpi-card:hover {{ transform: translateY(-2px); box-shadow: 0 6px 16px rgba(11,37,64,0.1); }}
    .kpi-top {{ display:flex; align-items:center; justify-content:space-between; }}
    .kpi-icon {{ font-size: 1.3rem; }}
    .kpi-title {{ color:{PALETTE['muted']}; font-size:0.76rem; font-weight:600; text-transform:uppercase; letter-spacing:0.4px; }}
    .kpi-value {{ color:{PALETTE['ink']}; font-size:1.65rem; font-weight:700; line-height:1.1; }}
    .kpi-delta-pos {{ color:{PALETTE['positive']}; font-size:0.78rem; font-weight:600; }}
    .kpi-delta-neg {{ color:{PALETTE['negative']}; font-size:0.78rem; font-weight:600; }}
    .kpi-delta-neu {{ color:{PALETTE['muted']}; font-size:0.78rem; font-weight:600; }}

    .section-title {{
        color:{PALETTE['ink']};
        font-size:1.05rem;
        font-weight:700;
        margin: 1.1rem 0 0.6rem 0;
        border-left: 4px solid {PALETTE['accent']};
        padding-left: 0.6rem;
    }}

    .insight-card {{
        background:{PALETTE['card']};
        border-radius: 10px;
        padding: 0.85rem 1rem;
        border: 1px solid {PALETTE['grid']};
        border-left: 4px solid var(--accent-color, {PALETTE['accent2']});
        margin-bottom: 0.6rem;
        font-size: 0.92rem;
        color: {PALETTE['ink']};
    }}
    .insight-tag {{
        display:inline-block;
        font-size: 0.68rem;
        font-weight:700;
        text-transform:uppercase;
        letter-spacing:0.4px;
        padding: 0.1rem 0.5rem;
        border-radius: 999px;
        margin-bottom: 0.35rem;
    }}

    div[data-testid="stMetricValue"] {{ color: {PALETTE['ink']}; }}
    section[data-testid="stSidebar"] *{{ 
    background-color: #0B2540; /* dark navy */
    border-right: 1px solid #1B6F8C; }}
    /* ⭐ ADVANCED ANALYTICS — ALL TEXT INSIDE TABS */
    div[data-testid="stTabs"] *{{
        color: #1A6078 !important;
    }}

    /* ⭐ DATA EXPLORER — TABLE TEXT */
    div[data-testid="stDataFrame"] *{{
        color: #1A6078 !important;
    }}

    /* ⭐ DATA EXPLORER — MARKDOWN TEXT */
    div[data-testid="stMarkdown"]  {{
        color: #FFFFFF !important;
    }}

    /* ⭐ DATA EXPLORER — SEARCH BOX TEXT */
    div[data-testid="stTextInput"] *{{
        color: #1A6078 !important;
    }}

    /* ⭐ DATA EXPLORER — DROPDOWN TEXT */
    div[data-testid="stSelectbox"] *{{
        color: #1A6078 !important;
    }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# --------------------------------------------------------------------------------------
# COLUMN DETECTION VOCABULARY (bilingual, extend freely — nothing here is a hard column name)
# --------------------------------------------------------------------------------------

KEYWORDS = {
    "id": ["id", "no.", "no", "رقم", "#", "index", "الرقم", "التسلسل"],
    "company": ["company", "vendor", "client", "شركة", "الشركة", "اسم السجل", "المؤسسة",
                "organization", "supplier", "manufacturer", "المصنع", "الجهة"],
    "product": ["product", "منتج", "المنتج", "service", "خدمة", "solution", "الحل", "app", "تطبيق"],
    "category": ["category", "tag", "type", "class", "segment", "وسم", "الوسوم", "تصنيف",
                 "التصنيف", "نوع", "فئة", "قطاع", "sector"],
    "status": ["status", "state", "حالة", "الحالة", "نشط", "active", "flag"],
    "city": ["city", "مدينة", "المدينة", "town"],
    "region": ["region", "منطقة", "المنطقة", "province", "area", "المنطقه"],
    "date": ["date", "تاريخ", "day", "month", "year", "created", "updated", "issued", "expiry",
             "expire", "اصدار", "انتهاء", "التحديث"],
    "count": ["count", "عدد", "number", "qty", "quantity", "total", "المجموع"],
}

DATE_ISSUE_HINTS = ["issue", "اصدار", "start", "created", "من"]
DATE_EXPIRY_HINTS = ["expiry", "expire", "انتهاء", "end", "الى", "حتى"]


def _normalize(text: str) -> str:
    text = str(text)
    text = unicodedata.normalize("NFKC", text)
    return text.strip().lower()


def _keyword_score(col_name: str, keywords: List[str]) -> int:
    norm = _normalize(col_name)
    score = 0
    for kw in keywords:
        if _normalize(kw) in norm:
            score += 1
    return score


# --------------------------------------------------------------------------------------
# DATA STRUCTURES
# --------------------------------------------------------------------------------------

@dataclass
class CleaningReport:
    original_rows: int = 0
    final_rows: int = 0
    empty_rows_removed: int = 0
    duplicate_rows_removed: int = 0
    missing_cells: int = 0
    total_cells: int = 0

    @property
    def missing_pct(self) -> float:
        return round(100 * self.missing_cells / self.total_cells, 2) if self.total_cells else 0.0


@dataclass
class ColumnMap:
    id: List[str] = field(default_factory=list)
    company: List[str] = field(default_factory=list)
    product: List[str] = field(default_factory=list)
    category: List[str] = field(default_factory=list)
    status: List[str] = field(default_factory=list)
    city: List[str] = field(default_factory=list)
    region: List[str] = field(default_factory=list)
    date: List[str] = field(default_factory=list)
    numeric: List[str] = field(default_factory=list)
    other_categorical: List[str] = field(default_factory=list)

    def first(self, role: str) -> Optional[str]:
        vals = getattr(self, role, [])
        return vals[0] if vals else None


# --------------------------------------------------------------------------------------
# CORE ENGINE
# --------------------------------------------------------------------------------------

class DataProcessor:
    """Loads, cleans and classifies an arbitrary Excel workbook."""

    def __init__(self, file):
        self.file = file
        self.sheets_raw: Dict[str, pd.DataFrame] = {}
        self.sheets_clean: Dict[str, pd.DataFrame] = {}
        self.reports: Dict[str, CleaningReport] = {}
        self.colmaps: Dict[str, ColumnMap] = {}

    # -- loading -------------------------------------------------------------------
    def load(self) -> List[str]:
        xl = pd.ExcelFile(self.file)
        valid_sheets = []
        for name in xl.sheet_names:
            try:
                df = xl.parse(name)
                if df is None or df.shape[0] == 0 or df.shape[1] == 0:
                    continue
                # drop fully unnamed / empty columns
                df = df.loc[:, ~df.columns.astype(str).str.contains(r"^Unnamed", na=False) | df.notna().any()]
                self.sheets_raw[name] = df
                valid_sheets.append(name)
            except Exception:
                continue
        if not valid_sheets:
            raise ValueError("No readable sheets with data were found in this workbook.")
        return valid_sheets

    # -- cleaning --------------------------------------------------------------------
    def clean(self, sheet_name: str) -> pd.DataFrame:
        df = self.sheets_raw[sheet_name].copy()
        report = CleaningReport(original_rows=len(df))

        # trim column names
        df.columns = [str(c).strip() for c in df.columns]
        
        # trim whitespace on string cells
        for col in df.columns:
            if df[col].dtype == object or pd.api.types.is_string_dtype(df[col]):
                df[col] = df[col].ffill()
                df[col] = df[col].apply(lambda v: v.strip() if isinstance(v, str) else v)
                df[col] = df[col].replace(r"^\s*$", np.nan, regex=True)
                df[col] = df[col].replace(["_", "-", "N/A", "n/a", "NA", "None", "null"], np.nan)

        # remove fully empty rows
        before = len(df)
        df = df.dropna(how="all")
        report.empty_rows_removed = before - len(df)

        # remove duplicate rows
        before = len(df)
        df = df.drop_duplicates()
        report.duplicate_rows_removed = before - len(df)

        # attempt date parsing on likely date columns
        for col in df.columns:
            if _keyword_score(col, KEYWORDS["date"]) > 0:
                parsed = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
                # only adopt parsed version if a meaningful fraction actually parsed
                if parsed.notna().sum() >= max(1, int(0.15 * df[col].notna().sum() + 0.0001)):
                    df[col] = parsed

        # attempt numeric coercion on count-like columns
        for col in df.columns:
            if _keyword_score(col, KEYWORDS["count"]) > 0 and not pd.api.types.is_numeric_dtype(df[col]):
                coerced = pd.to_numeric(df[col], errors="coerce")
                if coerced.notna().sum() >= 0.5 * df[col].notna().sum():
                    df[col] = coerced

        report.final_rows = len(df)
        report.total_cells = df.shape[0] * df.shape[1]
        report.missing_cells = int(df.isna().sum().sum())

        df = df.reset_index(drop=True)
        self.sheets_clean[sheet_name] = df
        self.reports[sheet_name] = report
        return df

    # -- classification --------------------------------------------------------------
    def classify(self, sheet_name: str) -> ColumnMap:
        df = self.sheets_clean[sheet_name]
        cmap = ColumnMap()
        n = len(df)

        for col in df.columns:
            series = df[col]
            is_numeric = pd.api.types.is_numeric_dtype(series)
            is_datetime = pd.api.types.is_datetime64_any_dtype(series)
            nunique = series.nunique(dropna=True)
            cardinality_ratio = nunique / n if n else 0

            if is_datetime:
                cmap.date.append(col)
                continue

            # Numeric columns can only ever be an id or a measure (count) —
            # e.g. "Number of Products" must not be misread as a "product" text column.
            role_pool = ["id", "count"] if is_numeric else list(KEYWORDS.keys())
            scores = {role: _keyword_score(col, KEYWORDS[role]) for role in role_pool}
            best_role = max(scores, key=scores.get)
            best_score = scores[best_role]

            if best_score > 0:
                if best_role == "id":
                    cmap.id.append(col)
                elif best_role == "company":
                    cmap.company.append(col)
                elif best_role == "product":
                    cmap.product.append(col)
                elif best_role == "category":
                    cmap.category.append(col)
                elif best_role == "status":
                    cmap.status.append(col)
                elif best_role == "city":
                    cmap.city.append(col)
                elif best_role == "region":
                    cmap.region.append(col)
                elif best_role == "date":
                    cmap.date.append(col)
                elif best_role == "count" and is_numeric:
                    cmap.numeric.append(col)
                continue

            # fallback heuristics when no keyword match
            if is_numeric:
                if cardinality_ratio > 0.9 and nunique == n:
                    cmap.id.append(col)
                else:
                    cmap.numeric.append(col)
            else:
                if 0 < nunique <= max(15, int(0.3 * n)):
                    cmap.other_categorical.append(col)

        self.colmaps[sheet_name] = cmap
        return cmap


def classify_date_role(col_name: str) -> str:
    norm = _normalize(col_name)
    if any(_normalize(h) in norm for h in DATE_EXPIRY_HINTS):
        return "expiry"
    if any(_normalize(h) in norm for h in DATE_ISSUE_HINTS):
        return "issue"
    return "generic"


# --------------------------------------------------------------------------------------
# FILTERS
# --------------------------------------------------------------------------------------

def build_sidebar_filters(df: pd.DataFrame, cmap: ColumnMap) -> pd.DataFrame:
    st.sidebar.markdown("### 🔎 Filters")
    filtered = df.copy()

    filter_roles = {
        "Company": cmap.company,
        "Product": cmap.product,
        "Category / Tag": cmap.category,
        "Status": cmap.status,
        "City": cmap.city,
        "Region": cmap.region,
    }

    any_filter = False
    for label, cols in filter_roles.items():
        for col in cols:
            options = sorted([v for v in filtered[col].dropna().unique().tolist()])
            if not options or len(options) > 300:
                continue
            any_filter = True
            selected = st.sidebar.multiselect(f"{label}: {col}", options, default=[], key=f"filt_{col}")
            if selected:
                filtered = filtered[filtered[col].isin(selected)]

    # date range filters
    for col in cmap.date:
        valid_dates = filtered[col].dropna()
        if valid_dates.empty:
            continue
        any_filter = True
        min_d, max_d = valid_dates.min().date(), valid_dates.max().date()
        if min_d == max_d:
            continue
        date_range = st.sidebar.date_input(f"📅 {col}", value=(min_d, max_d), min_value=min_d, max_value=max_d, key=f"date_{col}")
        if isinstance(date_range, tuple) and len(date_range) == 2:
            start, end = date_range
            filtered = filtered[(filtered[col].isna()) | ((filtered[col].dt.date >= start) & (filtered[col].dt.date <= end))]

    if not any_filter:
        st.sidebar.caption("No categorical or date fields detected to filter on.")

    st.sidebar.markdown("---")
    if st.sidebar.button("♻️ Reset all filters"):
        for key in list(st.session_state.keys()):
            if key.startswith("filt_") or key.startswith("date_"):
                del st.session_state[key]
        st.rerun()

    return filtered


# --------------------------------------------------------------------------------------
# KPI ENGINE
# --------------------------------------------------------------------------------------

def render_kpi_card(col, icon: str, title: str, value: str, delta: Optional[str] = None, delta_type: str = "neu"):
    delta_html = f'<div class="kpi-delta-{delta_type}">{delta}</div>' if delta else "<div>&nbsp;</div>"
    col.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-top">
                <span class="kpi-icon">{icon}</span>
            </div>
            <div>
                <div class="kpi-title">{title}</div>
                <div class="kpi-value">{value}</div>
            </div>
            {delta_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def compute_and_render_kpis(df: pd.DataFrame, cmap: ColumnMap, report: CleaningReport):
    kpis = []

    if cmap.company:
        kpis.append(("🏢", "Total Companies", f"{df[cmap.company[0]].nunique():,}"))
    # if cmap.product:
    #     kpis.append(("📦", "Total Products", f"{df[cmap.product[0]].nunique():,}"))
    if cmap.category:
        kpis.append(("🌟", "Badges", f"{df[cmap.category[0]].nunique():,}"))
    if cmap.city:
        kpis.append(("🏙️", "Total Cities", f"{df[cmap.city[0]].nunique():,}"))
    if cmap.region:
        kpis.append(("🗺️", "Total Regions", f"{df[cmap.region[0]].nunique():,}"))
    kpis.append(("📄", "Total Products", f"{len(df):,}"))

    if cmap.status:
        col = cmap.status[0]
        vals = df[col].astype(str).str.lower()
        active_mask = vals.str.contains("active|نشط|approved|ملتزم|complete", na=False)
        inactive_mask = vals.str.contains("inactive|غير|reject|rejected|مرفوض", na=False)
        if active_mask.any() or inactive_mask.any():
            kpis.append(("✅", "Active-like Records", f"{int(active_mask.sum()):,}"))
            kpis.append(("⛔", "Inactive / Rejected", f"{int(inactive_mask.sum()):,}"))

    # kpis.append(("🧮", "Missing Data", f"{report.missing_pct}%"))
    # kpis.append(("🧬", "Duplicates Removed", f"{report.duplicate_rows_removed:,}"))

    if cmap.date:
        max_dates = [df[c].max() for c in cmap.date if df[c].notna().any()]
        if max_dates:
            last_update = max(max_dates)
            kpis.append(("🕒", "Most Recent Date", last_update.strftime("%Y-%m-%d")))

    # render grid, 4 per row
    st.markdown('<div class="section-title">Executive KPIs</div>', unsafe_allow_html=True)
    for i in range(0, len(kpis), 4):
        row = kpis[i:i + 4]
        cols = st.columns(4)
        for c, (icon, title, value) in zip(cols, row):
            render_kpi_card(c, icon, title, value)

    return kpis


# --------------------------------------------------------------------------------------
# CHART ENGINE
# --------------------------------------------------------------------------------------

def _styled_fig(fig, height=380):
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=45, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=PALETTE["ink"], size=13),  # darker text everywhere
        title_font=dict(size=16, color=PALETTE["ink"]),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(color=PALETTE["ink"], size=12)  # darker legend
        ),
        colorway=CHART_SEQUENCE,
    )

    fig.update_xaxes(
        gridcolor=PALETTE["grid"],
        zeroline=False,
        tickfont=dict(color=PALETTE["ink"], size=12),  # darker axis labels
        title_font=dict(color=PALETTE["ink"], size=13)
    )

    fig.update_yaxes(
        gridcolor=PALETTE["grid"],
        zeroline=False,
        tickfont=dict(color=PALETTE["ink"], size=12),  # darker axis labels
        title_font=dict(color=PALETTE["ink"], size=13)
    )

    fig.update_traces(
        hoverlabel=dict(
            bgcolor="white",
            font_size=12,
            font_color=PALETTE["ink"]  # darker hover text
        )
    )

    return fig




def render_charts(df: pd.DataFrame, cmap: ColumnMap):
    st.markdown('<div class="section-title">Visual Analytics</div>', unsafe_allow_html=True)
    charts_rendered = 0

    row_slots = []

    # --- Category / tag distribution ---
    if cmap.category:
        col = cmap.category[0]
        counts = df[col].dropna().astype(str).value_counts().reset_index()
        counts.columns = [col, "Count"]
        fig = px.bar(counts, x=col, y="Count", title=f"Distribution by Badges", text="Count")
        fig.update_layout(
        xaxis_title="Badges",
        )

        fig.update_traces(marker_color=PALETTE["accent2"], textposition="outside")
        row_slots.append(("bar", fig))

        fig2 = px.pie(counts, names=col, values="Count", title=f"Badges Share", hole=0.45)
        row_slots.append(("pie", fig2))

    # --- Region ---
    if cmap.region:
        col = cmap.region[0]
        counts = df[col].dropna().astype(str).value_counts().reset_index()
        counts.columns = [col, "Count"]
        fig = px.bar(counts, x=col, y="Count", title=f"Records by {col}", text="Count")
        fig.update_layout(
        xaxis_title="Badges",
        )
        fig.update_traces(marker_color=PALETTE["ink"], textposition="outside")
        row_slots.append(("bar", fig))

    # --- City (top 10) ---
    # if cmap.city:
    #     col = cmap.city[0]
    #     counts = df[col].dropna().astype(str).value_counts().head(10).reset_index()
    #     counts.columns = [col, "Count"]
    #     fig = px.bar(counts, x="Count", y=col, orientation="h", title=f"Top 10 Product Name", text="Count")
    #     fig.update_traces(marker_color=PALETTE["accent"], textposition="outside")
    #     fig.update_layout(yaxis=dict(categoryorder="total ascending"))
    #     row_slots.append(("bar", fig))

    # --- Top companies (by product count or record count) ---
    if cmap.company:
        col = cmap.company[0]
        if cmap.numeric:
            num_col = cmap.numeric[0]
            grp = df.groupby(col)[num_col].sum(numeric_only=True).sort_values(ascending=False).head(10).reset_index()
            fig = px.bar(grp, x=num_col, y=col, orientation="h", title=f"Top 10 Companies by Number of Products", text=num_col)
        else:
            grp = df[col].dropna().astype(str).value_counts().head(10).reset_index()
            grp.columns = [col, "Records"]
            fig = px.bar(grp, x="Records", y=col, orientation="h", title="Top 10 Companies by Record Count", text="Records")
        fig.update_traces(marker_color=PALETTE["accent2"], textposition="outside")
        fig.update_layout(yaxis=dict(categoryorder="total ascending"))
        row_slots.append(("bar", fig))

    # --- Top products ---
    # if cmap.product:
    #     col = cmap.product[0]
    #     counts = df[col].dropna().astype(str).value_counts().head(10).reset_index()
    #     counts.columns = [col, "Count"]
    #     fig = px.bar(counts, x="Count", y=col, orientation="h", title=f"Top 10 {col}", text="Count")
    #     fig.update_traces(marker_color=PALETTE["ink"], textposition="outside")
    #     fig.update_layout(yaxis=dict(categoryorder="total ascending"))
    #     row_slots.append(("bar", fig))

    # --- Status distribution ---
    if cmap.status and cmap.status[0] not in cmap.category:
        col = cmap.status[0]
        counts = df[col].dropna().astype(str).value_counts().reset_index()
        counts.columns = [col, "Count"]
        fig = px.pie(counts, names=col, values="Count", title=f"{col} Distribution", hole=0.45)
        row_slots.append(("pie", fig))

    # --- Treemap: company x category, if both exist ---
    if cmap.company and cmap.category:
        c1, c2 = cmap.company[0], cmap.category[0]
        sub = df[[c1, c2]].dropna()
        if not sub.empty:
            grp = sub.groupby([c2, c1]).size().reset_index(name="Count")
            fig = px.treemap(grp, path=[c2, c1], values="Count", title=f"Company Name within Badges (Treemap)",
                              color="Count", color_continuous_scale=["#EAF2F5", PALETTE["accent2"]])
            row_slots.append(("wide", fig))

    # --- Sunburst: region > city, if both exist ---
    if cmap.region and cmap.city:
        r, c = cmap.region[0], cmap.city[0]
        sub = df[[r, c]].dropna()
        if not sub.empty:
            grp = sub.groupby([r, c]).size().reset_index(name="Count")
            fig = px.sunburst(grp, path=[r, c], values="Count", title=f"{r} → {c} Breakdown",
                               color="Count", color_continuous_scale=["#F5EFDD", PALETTE["accent"]])
            row_slots.append(("wide", fig))

    # --- Time trend ---
    if cmap.date:
        for dcol in cmap.date:
            role = classify_date_role(dcol)
            ts = df[dcol].dropna()
            if ts.empty or ts.dt.to_period("M").nunique() < 2:
                continue
            monthly = ts.dt.to_period("M").value_counts().sort_index()
            monthly.index = monthly.index.astype(str)
            fig = px.line(x=monthly.index, y=monthly.values, markers=True,
                          title=f"Trend Over Time — Badge Issue Date", labels={"x": "Month", "y": "Records"})
            fig.update_traces(line_color=PALETTE["accent2"], line_width=3)
            row_slots.append(("wide", fig))
            break  # one trend chart is usually enough for the headline view

    # --- Numeric distributions (histogram + boxplot) ---
    for ncol in cmap.numeric[:2]:
        series = df[ncol].dropna()
        if series.nunique() <= 1:
            continue
        fig = px.histogram(df, x=ncol, title=f"Distribution of Number of Products", nbins=20)
        fig.update_traces(marker_color=PALETTE["accent2"])
        row_slots.append(("bar", fig))

        # fig_box = px.box(df, y=ncol, title=f"Boxplot / Outliers — {ncol}", points="outliers")
        # fig_box.update_traces(marker_color=PALETTE["negative"])
        # row_slots.append(("bar", fig_box))

    # --- Correlation matrix if 2+ numeric columns ---
    if len(cmap.numeric) >= 2:
        corr = df[cmap.numeric].corr(numeric_only=True)
        fig = px.imshow(corr, text_auto=".2f", title="Correlation Matrix (Numeric Fields)",
                         color_continuous_scale=["#B0413E", "#000000", PALETTE["accent2"]], zmin=-1, zmax=1)
        row_slots.append(("wide", fig))

    # --- Missing values heatmap ---
    miss = df.isna()
    if miss.values.any():
        miss_pct = (miss.mean() * 100).sort_values(ascending=False)
        miss_pct = miss_pct[miss_pct > 0]
        if not miss_pct.empty:
            fig = px.bar(x=miss_pct.values, y=miss_pct.index, orientation="h",
                         title="Missing Data by Column (%)", labels={"x": "% Missing", "y": ""})
            fig.update_traces(marker_color=PALETTE["warning"])
            row_slots.append(("wide", fig))

    # --- render in responsive pairs, "wide" charts get full width ---
    i = 0
    while i < len(row_slots):
        kind, fig = row_slots[i]
        if kind == "wide":
            st.plotly_chart(_styled_fig(fig, height=460), use_container_width=True)
            i += 1
            charts_rendered += 1
            continue
        # try to pair with next non-wide chart
        if i + 1 < len(row_slots) and row_slots[i + 1][0] != "wide":
            c1, c2 = st.columns(2)
            c1.plotly_chart(_styled_fig(fig), use_container_width=True)
            c2.plotly_chart(_styled_fig(row_slots[i + 1][1]), use_container_width=True)
            i += 2
            charts_rendered += 2
        else:
            st.plotly_chart(_styled_fig(fig), use_container_width=True)
            i += 1
            charts_rendered += 1

    if charts_rendered == 0:
        st.info("No chartable fields were confidently detected in this dataset yet. Try the Data Explorer tab to inspect raw columns.")


# --------------------------------------------------------------------------------------
# ADVANCED ANALYTICS
# --------------------------------------------------------------------------------------

def render_advanced_analytics(df: pd.DataFrame, cmap: ColumnMap):
    st.markdown('<div class="section-title">Advanced Analytics</div>', unsafe_allow_html=True)
    tabs = st.tabs(["🏆 Leaders", "⚠️ Gaps & Risks", "🔁 Duplicates & Quality", "📈 Outliers"])

    with tabs[0]:
        cols = st.columns(2)
        shown = False
        if cmap.company:
            col = cmap.company[0]
            top = df[col].dropna().astype(str).value_counts().head(10)
            if not top.empty:
                cols[0].markdown("**Top companies by record count**")
                cols[0].dataframe(top.rename_axis(col).reset_index(name="Records"), use_container_width=True, hide_index=True)
                shown = True
        if cmap.category:
            col = cmap.category[0]
            top = df[col].dropna().astype(str).value_counts()
            if not top.empty:
                cols[1].markdown("**Largest categories / tags**")
                cols[1].dataframe(top.rename_axis(col).reset_index(name="Count"), use_container_width=True, hide_index=True)
                shown = True
        if not shown:
            st.caption("No company or category fields detected.")

    with tabs[1]:
        found_any = False
        if cmap.company and cmap.product:
            comp_col, prod_col = cmap.company[0], cmap.product[0]
            all_companies = set(df[comp_col].dropna().astype(str))
            companies_with_products = set(df.loc[df[prod_col].notna(), comp_col].dropna().astype(str))
            without = all_companies - companies_with_products
            if without:
                st.markdown(f"**Companies without a listed product** ({len(without)})")
                st.dataframe(pd.DataFrame(sorted(without), columns=[comp_col]), use_container_width=True, hide_index=True)
                found_any = True
        if cmap.category:
            col = cmap.category[0]
            counts = df[col].dropna().astype(str).value_counts()
            low = counts[counts <= max(1, int(0.05 * counts.sum()))]
            if not low.empty:
                st.markdown("**Categories with low representation**")
                st.dataframe(low.rename_axis(col).reset_index(name="Count"), use_container_width=True, hide_index=True)
                found_any = True
        if cmap.date:
            for dcol in cmap.date:
                if classify_date_role(dcol) == "expiry":
                    upcoming = df[df[dcol].notna() & (df[dcol] <= pd.Timestamp.now() + pd.Timedelta(days=60)) & (df[dcol] >= pd.Timestamp.now() - pd.Timedelta(days=3650))]
                    if not upcoming.empty:
                        st.markdown(f"**Records expiring within 60 days — `{dcol}`** ({len(upcoming)})")
                        show_cols = [c for c in [cmap.first("company"), cmap.first("product"), dcol] if c]
                        st.dataframe(upcoming[show_cols].sort_values(dcol), use_container_width=True, hide_index=True)
                        found_any = True
        if not found_any:
            st.caption("No notable gaps or expiry risks detected with current data.")

    with tabs[2]:
        report = st.session_state.get("_active_report")
        c1, c2, c3 = st.columns(3)
        if report:
            c1.metric("Duplicate rows removed", f"{report.duplicate_rows_removed:,}")
            c2.metric("Empty rows removed", f"{report.empty_rows_removed:,}")
            c3.metric("Missing data", f"{report.missing_pct}%")
        miss_by_col = (df.isna().mean() * 100).round(1)
        miss_by_col = miss_by_col[miss_by_col > 0].sort_values(ascending=False)
        if not miss_by_col.empty:
            st.markdown("**Missing values by column**")
            st.dataframe(miss_by_col.rename_axis("Column").reset_index(name="% Missing"), use_container_width=True, hide_index=True)
        else:
            st.caption("No missing values remain in the cleaned dataset.")

    with tabs[3]:
        if cmap.numeric:
            for ncol in cmap.numeric[:4]:
                series = df[ncol].dropna()
                if series.nunique() <= 1:
                    continue
                q1, q3 = series.quantile(0.25), series.quantile(0.75)
                iqr = q3 - q1
                lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
                outliers = df[(df[ncol] < lower) | (df[ncol] > upper)]
                st.markdown(f"**{ncol}** — {len(outliers)} potential outlier(s) (IQR method)")
                if not outliers.empty:
                    show_cols = [c for c in [cmap.first("company"), cmap.first("product"), ncol] if c]
                    st.dataframe(outliers[show_cols], use_container_width=True, hide_index=True)
        else:
            st.caption("No numeric fields available for outlier detection.")


# --------------------------------------------------------------------------------------
# INSIGHTS ENGINE
# --------------------------------------------------------------------------------------

def generate_insights(df: pd.DataFrame, cmap: ColumnMap, report: CleaningReport) -> List[dict]:
    insights = []

    def add(kind, text):
        insights.append({"kind": kind, "text": text})

    n = len(df)
    if n == 0:
        return insights

    if cmap.category:
        col = cmap.category[0]
        counts = df[col].dropna().astype(str).value_counts()
        if not counts.empty:
            top_cat, top_val = counts.index[0], counts.iloc[0]
            add("observation", f"“{top_cat}” is the leading value in **{col}**, covering {top_val} of {n} records ({round(100*top_val/n,1)}%).")
            if len(counts) > 1:
                smallest = counts.index[-1]
                add("opportunity", f"“{smallest}” is the least represented value in **{col}** with only {counts.iloc[-1]} record(s) — a potential focus area for growth.")

    if cmap.company:
        col = cmap.company[0]
        n_companies = df[col].nunique()
        add("observation", f"The dataset covers **{n_companies} unique companies** across {n} records.")
        dup_companies = df[col].dropna().astype(str).value_counts()
        multi = dup_companies[dup_companies > 1]
        if not multi.empty:
            add("observation", f"{len(multi)} companies appear more than once, suggesting multiple products or repeated engagements.")

    if cmap.status:
        col = cmap.status[0]
        vals = df[col].astype(str).str.lower()
        rejected = vals.str.contains("reject|مرفوض", na=False).sum()
        pending = vals.str.contains("تحت التقييم|pending|review|evaluat", na=False).sum()
        if rejected:
            add("risk", f"{rejected} record(s) carry a rejected / declined status in **{col}** and may need follow-up.")
        if pending:
            add("recommendation", f"{pending} record(s) are still under evaluation in **{col}** — prioritizing these could accelerate throughput.")

    if cmap.date:
        for dcol in cmap.date:
            role = classify_date_role(dcol)
            if role == "expiry":
                soon = df[df[dcol].notna() & (df[dcol] <= pd.Timestamp.now() + pd.Timedelta(days=60)) & (df[dcol] >= pd.Timestamp.now())]
                if not soon.empty:
                    add("risk", f"{len(soon)} record(s) have entries in **{dcol}** expiring within the next 60 days — renewal action recommended.")
            if role == "issue" or role == "generic":
                ts = df[dcol].dropna()
                if ts.dt.to_period("M").nunique() >= 3:
                    recent = ts[ts >= ts.max() - pd.Timedelta(days=90)].shape[0]
                    older = ts.shape[0] - recent
                    if recent > older:
                        add("trend", f"Activity in **{dcol}** has accelerated — most entries fall within the last 90 days.")

    if report.missing_pct > 15:
        add("risk", f"Overall missing data is {report.missing_pct}%, which may reduce the reliability of downstream analysis.")
    elif report.missing_pct > 0:
        add("observation", f"Data completeness is strong — only {report.missing_pct}% of cells are missing after cleaning.")
    else:
        add("observation", "No missing values remain after cleaning — the dataset is fully populated.")

    if report.duplicate_rows_removed > 0:
        add("observation", f"{report.duplicate_rows_removed} duplicate row(s) were detected and removed during cleaning.")

    if cmap.region:
        col = cmap.region[0]
        n_regions = df[col].nunique()
        if n_regions >= 2:
            top_region = df[col].value_counts().idxmax()
            add("observation", f"**{top_region}** is the most represented region across {n_regions} region(s) tracked.")

    if not insights:
        add("observation", "The dataset was processed successfully; add more structured fields (dates, categories, regions) to unlock deeper insights.")

    return insights


INSIGHT_STYLE = {
    "observation": ("🔍", PALETTE["accent2"]),
    "trend": ("📈", PALETTE["positive"]),
    "risk": ("⚠️", PALETTE["negative"]),
    "opportunity": ("💡", PALETTE["accent"]),
    "recommendation": ("✅", PALETTE["ink"]),
}


def render_insights(df: pd.DataFrame, cmap: ColumnMap, report: CleaningReport):
    st.markdown('<div class="section-title">Executive Insights</div>', unsafe_allow_html=True)
    insights = generate_insights(df, cmap, report)
    col_a, col_b = st.columns(2)
    for idx, ins in enumerate(insights):
        icon, color = INSIGHT_STYLE.get(ins["kind"], ("🔍", PALETTE["accent2"]))
        target = col_a if idx % 2 == 0 else col_b
        target.markdown(
            f"""
            <div class="insight-card" style="--accent-color:{color}">
                <span class="insight-tag" style="background:{color}22;color:{color}">{icon} {ins['kind']}</span><br/>
                {ins['text']}
            </div>
            """,
            unsafe_allow_html=True,
        )


# --------------------------------------------------------------------------------------
# DATA EXPLORER
# --------------------------------------------------------------------------------------

def render_data_explorer(df: pd.DataFrame):
    st.markdown('<div class="section-title">Data Explorer</div>', unsafe_allow_html=True)

    search = st.text_input("🔍 Search across all columns", "")
    display_df = df.copy()
    if search:
        mask = display_df.apply(lambda row: row.astype(str).str.contains(search, case=False, na=False).any(), axis=1)
        display_df = display_df[mask]

    all_cols = list(display_df.columns)
    selected_cols = st.multiselect("Columns to display", all_cols, default=all_cols)

    st.dataframe(display_df[selected_cols], use_container_width=True, height=420)
    st.caption(f"Showing {len(display_df):,} of {len(df):,} records.")

    c1, c2 = st.columns(2)
    csv_bytes = display_df[selected_cols].to_csv(index=False).encode("utf-8-sig")
    c1.download_button("⬇️ Export CSV", csv_bytes, file_name="export.csv", mime="text/csv", use_container_width=True)

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        display_df[selected_cols].to_excel(writer, index=False, sheet_name="Data")
    c2.download_button("⬇️ Export Excel", excel_buffer.getvalue(), file_name="export.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)


# --------------------------------------------------------------------------------------
# MAIN APP
# --------------------------------------------------------------------------------------

def main():
    st.markdown(
        """
        <div class="dash-header">
            <h1>Compliance Assessment Dashboard</h1>
            <p>Automated multi-sheet analysis · Data cleaning · KPI monitoring · Business insights</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.sidebar.markdown("## ⚙️ Data Source")
    default_path = "data.xlsx"
    uploaded = st.sidebar.file_uploader("Upload an Excel workbook (.xlsx)", type=["xlsx", "xls"])

    source = uploaded
    source_label = uploaded.name if uploaded else None
    if source is None:
        try:
            with open(default_path, "rb") as f:
                source = io.BytesIO(f.read())
            source_label = default_path
        except FileNotFoundError:
            source = None

    if source is None:
        st.warning("📁 Please upload an Excel file to begin, or place `data.xlsx` next to `app.py`.")
        st.stop()

    st.sidebar.caption(f"Source: **{source_label}**")

    try:
        processor = DataProcessor(source)
        sheet_names = processor.load()
    except Exception as e:
        st.error(f"⚠️ Could not read this Excel file. Details: {e}")
        st.stop()

    st.sidebar.markdown("## 📑 Worksheet")
    active_sheet = st.sidebar.selectbox("Select sheet to analyze", sheet_names, index=0)

    try:
        clean_df = processor.clean(active_sheet)
        cmap = processor.classify(active_sheet)
        report = processor.reports[active_sheet]
    except Exception as e:
        st.error(f"⚠️ Error while processing sheet '{active_sheet}': {e}")
        st.stop()

    if clean_df.empty:
        st.info("This sheet contains no usable data after cleaning.")
        st.stop()

    st.session_state["_active_report"] = report

    filtered_df = build_sidebar_filters(clean_df, cmap)
    st.sidebar.markdown("---")
    st.sidebar.caption(f"Rows: {len(filtered_df):,} / {len(clean_df):,} shown")
    st.sidebar.caption(f"Cleaning: −{report.empty_rows_removed} empty · −{report.duplicate_rows_removed} duplicate rows")

    if filtered_df.empty:
        st.warning("No records match the current filter selection. Adjust filters in the sidebar.")
        st.stop()

    tab_overview, tab_analytics, tab_insights, tab_data = st.tabs(
        ["🏠 Overview", "📊 Advanced Analytics", "🧠 Executive Insights", "🗂️ Data Explorer"]
    )

    with tab_overview:
        compute_and_render_kpis(filtered_df, cmap, report)
        render_charts(filtered_df, cmap)

    with tab_analytics:
        render_advanced_analytics(filtered_df, cmap)

    with tab_insights:
        render_insights(filtered_df, cmap, report)

    with tab_data:
        render_data_explorer(filtered_df)

    st.markdown(
        f"<div style='text-align:center;color:{PALETTE['muted']};font-size:0.78rem;padding:1.2rem 0 0.4rem 0;'>"
        f"Generated automatically · {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
