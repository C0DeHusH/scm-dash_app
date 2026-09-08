from pathlib import Path

code = r'''"""
MUTI MC SCM Executive Control Tower
Dash-only implementation.

This version removes Streamlit completely while preserving the source
dashboard's presentation, KPI logic, workbook processing, trend charts,
network/branch drilldowns, Pareto tables, and procurement placeholder tab.
"""

import base64
import functools
import hashlib
import io
import os
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP

import dash
from dash import Dash, Input, Output, State, dcc, html, no_update
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go


# =========================================================
# 0. CONFIGURATION
# =========================================================

APP_TITLE = "SCM Executive Control Tower"
LOCAL_CACHE_FILE = "persistent_scm_data.xlsx"

XLSX_SIGNATURES = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")
XLS_SIGNATURE = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"

COLORS = {
    "navy": "#0b1220",
    "indigo": "#6366f1",
    "blue": "#0ea5e9",
    "green": "#10b981",
    "amber": "#f59e0b",
    "red": "#f43f5e",
    "purple": "#7c3aed",
    "border": "rgba(148,163,184,0.18)",
    "muted": "#94a3b8",
}


# =========================================================
# 1. EXCEL / ROUNDING UTILITIES
# =========================================================

def detect_excel_engine(file_bytes):
    """Detect openpyxl vs xlrd from workbook bytes."""
    if not file_bytes:
        raise ValueError("The uploaded workbook is empty.")

    signature = bytes(file_bytes[:8])

    if any(signature.startswith(sig) for sig in XLSX_SIGNATURES):
        return "openpyxl"

    if signature.startswith(XLS_SIGNATURE):
        return "xlrd"

    raise ValueError(
        "Not a valid Excel workbook. Please upload a genuine .xlsx or .xls file."
    )


def round_half_up(value, ndigits=0):
    """Business rounding using ROUND_HALF_UP."""
    if pd.isna(value):
        return np.nan

    quantizer = Decimal("1").scaleb(-ndigits)
    rounded = Decimal(str(float(value))).quantize(
        quantizer, rounding=ROUND_HALF_UP
    )
    return int(rounded) if ndigits == 0 else float(rounded)


def round_series_half_up(series):
    numeric = pd.to_numeric(series, errors="coerce")
    return numeric.map(
        lambda value: round_half_up(value) if pd.notna(value) else np.nan
    )


def save_local_cache(file_bytes):
    with open(LOCAL_CACHE_FILE, "wb") as handle:
        handle.write(file_bytes)


def load_local_cache():
    if not os.path.exists(LOCAL_CACHE_FILE):
        return None

    with open(LOCAL_CACHE_FILE, "rb") as handle:
        return handle.read()


# =========================================================
# 2. WORKBOOK PROCESSING
# =========================================================

def process_excel_file(file_bytes):
    """
    Expected workbook sheets:
      Raw_Data
      KPI_YTD_Input
      KPI_Weekly_Input

    The KPI parser dynamically locates date columns so January and
    other leading dates are not accidentally skipped.
    """
    excel_engine = detect_excel_engine(file_bytes)

    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine=excel_engine)
    except ImportError as exc:
        package = "openpyxl" if excel_engine == "openpyxl" else "xlrd"
        raise ValueError(
            f"Excel engine '{excel_engine}' is unavailable. "
            f"Install '{package}'."
        ) from exc

    sheet_map = {
        str(sheet).lower().strip().replace(" ", "_"): sheet
        for sheet in xls.sheet_names
    }

    def get_sheet(target):
        if target not in sheet_map:
            raise ValueError(
                f"Missing sheet '{target}'. Found: {xls.sheet_names}"
            )
        return sheet_map[target]

    # -------------------------
    # Raw data
    # -------------------------
    raw_df = pd.read_excel(
        xls,
        sheet_name=get_sheet("raw_data"),
    )

    raw_df.columns = (
        raw_df.columns.astype(str)
        .str.lower()
        .str.strip()
        .str.replace(" ", "_", regex=False)
    )

    if "class" in raw_df.columns:
        raw_df.rename(columns={"class": "pareto_class"}, inplace=True)

    required = [
        "area",
        "branch",
        "pareto_class",
        "stock_status",
        "remaining_inventory",
        "suggested_transfer",
        "doi",
        "model",
    ]

    missing = [column for column in required if column not in raw_df.columns]
    if missing:
        raise ValueError(
            "Raw_Data is missing required column(s): "
            + ", ".join(missing)
        )

    for column in ["remaining_inventory", "suggested_transfer", "doi"]:
        raw_df[column] = (
            round_series_half_up(
                pd.to_numeric(raw_df[column], errors="coerce").fillna(0)
            )
            .fillna(0)
            .astype(int)
        )

    for column in ["pareto_class", "stock_status", "area", "branch", "model"]:
        raw_df[column] = raw_df[column].fillna("").astype(str).str.strip()

    # -------------------------
    # KPI sheets
    # -------------------------
    def parse_header_date(value):
        if pd.isna(value):
            return pd.NaT

        if isinstance(value, (pd.Timestamp, datetime, date, np.datetime64)):
            parsed = pd.to_datetime(value, errors="coerce")

        elif isinstance(value, (int, float, np.integer, np.floating)):
            numeric_value = float(value)
            if 20000 <= numeric_value <= 60000:
                parsed = pd.Timestamp("1899-12-30") + pd.to_timedelta(
                    numeric_value, unit="D"
                )
            else:
                return pd.NaT

        else:
            text = str(value).strip()
            if not text:
                return pd.NaT
            parsed = pd.to_datetime(text, errors="coerce")

        if pd.isna(parsed):
            return pd.NaT

        parsed = pd.Timestamp(parsed)

        if parsed.year < 2000 or parsed.year > 2100:
            return pd.NaT

        return parsed.normalize()

    def parse_kpi(sheet_name):
        df_raw = pd.read_excel(
            xls,
            sheet_name=get_sheet(sheet_name),
            header=None,
        )

        empty = pd.DataFrame(
            columns=[
                "period",
                "after_po",
                "per_branch",
                "class_a_out",
                "before_po",
                "overall_doi",
                "class_a_doi",
            ]
        )

        if df_raw.empty or df_raw.shape[1] < 2:
            return empty

        best_date_columns = []
        best_dates = []

        # Scan the first ten rows for the date header.
        for row_idx in range(min(10, len(df_raw))):
            row_columns = []
            row_dates = []

            for col_idx in range(df_raw.shape[1]):
                parsed = parse_header_date(df_raw.iat[row_idx, col_idx])
                if pd.notna(parsed):
                    row_columns.append(col_idx)
                    row_dates.append(parsed)

            if len(row_columns) > len(best_date_columns):
                best_date_columns = row_columns
                best_dates = row_dates

        if not best_date_columns:
            raise ValueError(
                f"No valid KPI date headers found in sheet '{sheet_name}'."
            )

        kpis = {
            "MUTI MC : Stock Outrate - Overall after PO Balance": "after_po",
            "MUTI MC : Stock Outrate - Per Branch": "per_branch",
            "Overall Class A Stock Out Rate": "class_a_out",
            "MUTI MC : Stock Outrate - Overall (Before PO Balance)": "before_po",
            "MUTI MC : DoI": "overall_doi",
            "MC Class A Doi": "class_a_doi",
        }

        data = {"period": best_dates}
        first_col = df_raw[0].fillna("").astype(str)

        for kpi_name, column_name in kpis.items():
            indexes = df_raw.index[
                first_col.str.contains(kpi_name, regex=False, na=False)
            ].tolist()

            if indexes:
                values = df_raw.loc[indexes[0], best_date_columns].values
                data[column_name] = pd.to_numeric(values, errors="coerce")
            else:
                data[column_name] = [np.nan] * len(best_dates)

        clean = pd.DataFrame(data)
        clean["period"] = pd.to_datetime(clean["period"], errors="coerce")
        clean = clean.dropna(subset=["period"])

        metric_columns = list(kpis.values())
        clean = clean.dropna(subset=metric_columns, how="all")

        def last_valid(series):
            valid = series.dropna()
            return valid.iloc[-1] if not valid.empty else np.nan

        clean = (
            clean.sort_values("period")
            .groupby("period", as_index=False)
            .agg({column: last_valid for column in metric_columns})
        )

        # YTD = one point per calendar month.
        if "ytd" in sheet_name.lower() and not clean.empty:
            clean["month_year"] = clean["period"].dt.to_period("M")
            clean = (
                clean.sort_values("period")
                .drop_duplicates("month_year", keep="last")
                .drop(columns=["month_year"])
            )

        # Weekly view uses actual dated observations only.
        return (
            clean.sort_values("period")
            .drop_duplicates("period", keep="last")
            .reset_index(drop=True)
        )

    return (
        raw_df,
        parse_kpi("kpi_ytd_input"),
        parse_kpi("kpi_weekly_input"),
    )


@functools.lru_cache(maxsize=1)
def get_cached_data(file_hash):
    bytes_data = load_local_cache()

    if bytes_data:
        return process_excel_file(bytes_data)

    return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


# =========================================================
# 3. BUSINESS / CHART LOGIC
# =========================================================

def calculate_stockout_rate(df, pareto_class=None):
    subset = (
        df[df["pareto_class"].str.casefold() == pareto_class.casefold()]
        if pareto_class
        else df
    )

    if len(subset) == 0:
        return 0

    stockouts = (
        subset["stock_status"]
        .fillna("")
        .astype(str)
        .str.lower()
        .str.strip()
        .eq("stockout")
        .sum()
    )

    return round_half_up((stockouts / len(subset)) * 100)


def prepare_chart_series(df, y_col):
    if df.empty or y_col not in df.columns:
        return pd.DataFrame(columns=["period", y_col])

    chart_df = df[["period", y_col]].copy()
    chart_df["period"] = pd.to_datetime(chart_df["period"], errors="coerce")
    chart_df[y_col] = pd.to_numeric(chart_df[y_col], errors="coerce")

    return (
        chart_df.dropna()
        .sort_values("period")
        .drop_duplicates("period", keep="last")
        .reset_index(drop=True)
    )


def infer_percentage_scale(series):
    numeric = pd.to_numeric(series, errors="coerce").dropna()

    if numeric.empty:
        return 1.0

    return 0.01 if numeric.abs().max() > 1.5 else 1.0


def create_styled_line_chart(
    df,
    y_col,
    title,
    subtitle,
    line_color,
    is_weekly,
    is_percentage=True,
    fill=False,
):
    chart_df = prepare_chart_series(df, y_col)
    fig = go.Figure()

    if chart_df.empty:
        fig.add_annotation(
            text="No valid data available",
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
            showarrow=False,
            font=dict(color=COLORS["muted"]),
        )
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            height=365,
            xaxis=dict(visible=False),
            yaxis=dict(visible=False),
        )
        return fig

    plot_y = chart_df[y_col].copy()

    if is_percentage:
        scale = infer_percentage_scale(plot_y)
        plot_y = (
            round_series_half_up((plot_y * scale) * 100.0) / 100.0
        )
        text_labels = [f"{value * 100:.0f}%" for value in plot_y]
        tick_format = ".0%"
        default_ceiling = 0.10
    else:
        plot_y = round_series_half_up(plot_y).astype(int)
        text_labels = [f"{value:,.0f}" for value in plot_y]
        tick_format = ",.0f"
        default_ceiling = 10

    if is_weekly:
        chart_x = chart_df["period"].dt.strftime("%d %b %Y")
        tick_text = chart_df["period"].dt.strftime("%d %b")
        category_array = chart_x.tolist()
    else:
        chart_x = chart_df["period"].dt.strftime("%b")
        tick_text = chart_x.tolist()
        category_array = [
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        ]

    hover_dates = chart_df["period"].dt.strftime("%d %b %Y")
    hover_format = ":.0%" if is_percentage else ":,.0f"

    fillcolor = None
    if fill:
        r = int(line_color[1:3], 16)
        g = int(line_color[3:5], 16)
        b = int(line_color[5:7], 16)
        fillcolor = f"rgba({r},{g},{b},0.08)"

    text_positions = [
        "top center" if i % 2 == 0 else "bottom center"
        for i in range(len(chart_df))
    ]

    fig.add_trace(
        go.Scatter(
            x=chart_x,
            y=plot_y,
            customdata=hover_dates,
            name="Actual",
            mode="lines+markers+text",
            text=text_labels,
            textposition=text_positions,
            textfont=dict(size=10, family="Arial"),
            line=dict(
                shape="linear" if is_weekly else "spline",
                width=3.2,
                color=line_color,
            ),
            marker=dict(
                size=8,
                color=line_color,
                line=dict(width=2, color="#ffffff"),
            ),
            fill="tozeroy" if fill else "none",
            fillcolor=fillcolor,
            hovertemplate=(
                f"<b>%{{customdata}}</b><br>{title}: "
                f"<b>%{{y{hover_format}}}</b><extra></extra>"
            ),
            cliponaxis=False,
        )
    )

    trend_y = None
    trend_direction = "Stable"

    if len(chart_df) >= 2:
        trend_x = (
            chart_df["period"] - chart_df["period"].min()
        ).dt.total_seconds().to_numpy(dtype=float) / 86400.0

        trend_source_y = pd.to_numeric(
            plot_y, errors="coerce"
        ).to_numpy(dtype=float)

        valid = np.isfinite(trend_x) & np.isfinite(trend_source_y)

        if valid.sum() >= 2 and np.ptp(trend_x[valid]) > 0:
            slope, intercept = np.polyfit(
                trend_x[valid],
                trend_source_y[valid],
                1,
            )
            fitted_y = slope * trend_x + intercept

            actual_span = max(
                float(np.nanmax(trend_source_y[valid]))
                - float(np.nanmin(trend_source_y[valid])),
                0.0,
            )

            gap = max(
                actual_span * 0.18,
                0.010 if is_percentage else 1.0,
            )
            amplitude = max(
                actual_span * 0.10,
                0.006 if is_percentage else 0.65,
            )

            fitted_range = np.ptp(fitted_y[valid])
            if fitted_range > 0:
                normalized_fit = (
                    fitted_y - np.nanmin(fitted_y[valid])
                ) / fitted_range
            else:
                normalized_fit = np.full_like(fitted_y, 0.5)

            trend_y = (
                float(np.nanmax(trend_source_y[valid]))
                + gap
                + normalized_fit * amplitude
            )

            fitted_delta = float(fitted_y[-1] - fitted_y[0])
            flat_threshold = max(
                actual_span * 0.03,
                0.001 if is_percentage else 0.10,
            )

            if fitted_delta > flat_threshold:
                trend_direction = "Upward"
            elif fitted_delta < -flat_threshold:
                trend_direction = "Downward"

            symbol = {
                "Upward": "↑",
                "Downward": "↓",
                "Stable": "→",
            }[trend_direction]

            fig.add_trace(
                go.Scatter(
                    x=chart_x,
                    y=trend_y,
                    name=f"Trend {symbol}",
                    mode="lines",
                    line=dict(
                        width=2.4,
                        dash="dash",
                        color="rgba(100,116,139,0.95)",
                    ),
                    hovertemplate=(
                        f"Trend: <b>{trend_direction} "
                        f"{symbol}</b><extra></extra>"
                    ),
                )
            )

    y_max = float(pd.to_numeric(plot_y, errors="coerce").max())

    if trend_y is not None:
        y_max = max(y_max, float(np.nanmax(trend_y)))

    y_max = (
        max(y_max * 1.16, 0.05 if is_percentage else 5)
        if y_max > 0
        else default_ceiling
    )

    latest_text = chart_df["period"].max().strftime("%d %b %Y")

    xaxis = dict(
        type="category",
        showgrid=False,
        categoryorder="array",
        categoryarray=category_array,
        linecolor="rgba(148,163,184,0.18)",
    )

    if is_weekly:
        xaxis.update(
            tickmode="array",
            tickvals=chart_x.tolist(),
            ticktext=tick_text.tolist(),
        )
    else:
        xaxis.update(
            tickmode="array",
            tickvals=chart_x.tolist(),
            ticktext=tick_text,
        )

    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=365,
        margin=dict(t=82, b=44, l=46, r=20),
        hovermode="closest",
        hoverlabel=dict(
            bgcolor="#0f172a",
            bordercolor="rgba(148,163,184,0.28)",
            font=dict(color="#f8fafc", size=11),
        ),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.015,
            xanchor="right",
            x=0.99,
            font=dict(size=10, color="#94a3b8"),
        ),
        title=dict(
            text=(
                f"{title}<br>"
                f"<span style='font-size:10px;color:#94a3b8;'>"
                f"{subtitle} • LATEST {latest_text.upper()}</span>"
            ),
            x=0.02,
            y=0.96,
            font=dict(size=16),
        ),
        xaxis=xaxis,
        yaxis=dict(
            showgrid=True,
            gridcolor="rgba(148,163,184,0.14)",
            zeroline=False,
            tickformat=tick_format,
            range=[0, y_max],
        ),
    )

    return fig


# =========================================================
# 4. UI COMPONENTS
# =========================================================

def section_heading(title, subtitle=""):
    return html.Div(
        className="section-heading",
        children=[
            html.Span(className="dot"),
            html.Span(title, className="title"),
            html.Span(subtitle, className="subtitle"),
        ],
    )


def info_chip(label):
    return html.Span(label, className="info-chip")


def metric_card(title, value, note="", accent="indigo"):
    accent_color = COLORS.get(accent, COLORS["indigo"])

    return html.Div(
        className="metric-card-base",
        style={"--accent": accent_color},
        children=[
            html.Div(title, className="metric-title"),
            html.Div(
                f"{value}%",
                className="metric-value-sm",
            ),
            html.Div(note, className="metric-footnote"),
        ],
    )


def chart_card(figure):
    return html.Div(
        className="chart-card",
        children=dcc.Graph(
            figure=figure,
            config={
                "displayModeBar": False,
                "responsive": True,
            },
            style={"width": "100%"},
        ),
    )


def empty_message(message):
    return html.Div(
        message,
        className="empty-message",
    )


# =========================================================
# 5. DASH APPLICATION
# =========================================================

app = Dash(
    __name__,
    suppress_callback_exceptions=True,
    title=APP_TITLE,
)

server = app.server


# =========================================================
# 6. STYLES
# =========================================================

app.index_string = """
<!DOCTYPE html>
<html>
<head>
    {%metas%}
    <title>{%title%}</title>
    {%favicon%}
    {%css%}
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        :root {
            --scm-navy: #0b1220;
            --scm-indigo: #6366f1;
            --scm-blue: #0ea5e9;
            --scm-green: #10b981;
            --scm-amber: #f59e0b;
            --scm-red: #f43f5e;
            --scm-border: rgba(148,163,184,0.18);
            --scm-muted: #94a3b8;
        }

        * { box-sizing: border-box; }

        html, body {
            margin: 0;
            padding: 0;
            background: #0b1220;
            color: #f8fafc;
            font-family: Inter, ui-sans-serif, system-ui, -apple-system,
                         BlinkMacSystemFont, "Segoe UI", sans-serif;
            overflow-x: hidden;
        }

        body { min-width: 0; }

        .app-shell {
            width: 100%;
            min-height: 100vh;
            padding: clamp(8px, 1.4vw, 20px);
            background:
                radial-gradient(circle at 80% -10%,
                    rgba(99,102,241,0.08), transparent 34%),
                #0b1220;
        }

        .hero-shell {
            width: 100%;
            min-width: 0;
            margin: 0 0 18px 0;
            padding: clamp(20px, 1.65vw, 28px)
                     clamp(22px, 2.15vw, 34px);
            border: 1px solid rgba(99,102,241,0.34);
            border-left: 4px solid #6366f1;
            border-radius: 18px;
            background: #0b1220;
            box-shadow: 0 10px 28px rgba(2,6,23,0.14);
        }

        .hero-kicker {
            color: #a5b4fc;
            font-size: clamp(.62rem,.66vw,.72rem);
            font-weight: 850;
            line-height: 1.4;
            letter-spacing: .12em;
            text-transform: uppercase;
            margin-bottom: 10px;
        }

        .hero-title {
            color: #f8fafc;
            font-size: clamp(1.65rem,2.10vw,2.45rem);
            font-weight: 900;
            line-height: 1.13;
            letter-spacing: -.025em;
            margin-bottom: 12px;
        }

        .hero-subtitle {
            max-width: 1180px;
            color: #cbd5e1;
            font-size: clamp(.80rem,.86vw,.94rem);
            font-weight: 500;
            line-height: 1.55;
        }

        .hero-link {
            display: inline-flex;
            align-items: center;
            margin-top: 18px;
            padding: 8px 18px;
            background: rgba(99,102,241,.12);
            color: #a5b4fc;
            border: 1px solid rgba(99,102,241,.4);
            border-radius: 8px;
            font-size: .72rem;
            font-weight: 850;
            text-decoration: none;
            letter-spacing: .05em;
            text-transform: uppercase;
            transition: all .2s ease;
        }

        .hero-link:hover {
            background: rgba(99,102,241,.25);
            color: #fff;
            border-color: rgba(99,102,241,.8);
            transform: translateY(-1px);
        }

        .tabs {
            display: flex;
            gap: 8px;
            width: 100%;
            border-bottom: 1px solid rgba(148,163,184,.16);
            margin-bottom: 18px;
        }

        .tab-btn {
            border: 0;
            background: transparent;
            color: #94a3b8;
            padding: 11px 16px;
            border-radius: 10px 10px 0 0;
            font-size: .82rem;
            font-weight: 850;
            cursor: pointer;
            transition: .18s ease;
        }

        .tab-btn:hover {
            color: #f8fafc;
            background: rgba(99,102,241,.08);
        }

        .tab-btn.active {
            color: #fff;
            background: rgba(99,102,241,.15);
            box-shadow: inset 0 -2px 0 #6366f1;
        }

        .tab-btn.inactive { color: #94a3b8; }

        .section-heading {
            display: grid;
            grid-template-columns: auto minmax(0,1fr) minmax(0,auto);
            align-items: center;
            gap: 9px;
            min-width: 0;
            margin: 16px 0 10px;
            padding: 3px 2px;
        }

        .section-heading .dot {
            width: 9px;
            height: 9px;
            border-radius: 50%;
            background: #6366f1;
            box-shadow: 0 0 0 4px rgba(99,102,241,.11);
        }

        .section-heading .title {
            font-size: 1.12rem;
            font-weight: 900;
            letter-spacing: -.012em;
        }

        .section-heading .subtitle {
            color: #94a3b8;
            font-size: .75rem;
            font-weight: 550;
            text-align: right;
            overflow-wrap: anywhere;
        }

        .control-row {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            align-items: end;
            margin-bottom: 12px;
        }

        .control-group {
            min-width: 210px;
        }

        .control-label {
            display: block;
            color: #94a3b8;
            font-size: .62rem;
            font-weight: 850;
            letter-spacing: .09em;
            text-transform: uppercase;
            margin-bottom: 5px;
        }

        .dash-dropdown .Select-control,
        .dash-dropdown .Select-menu-outer {
            border-radius: 10px !important;
        }

        .dash-dropdown .Select-control {
            min-height: 40px;
            background: #111827 !important;
            border: 1px solid rgba(148,163,184,.22) !important;
            color: #f8fafc !important;
        }

        .dash-dropdown .Select-value-label,
        .dash-dropdown .Select-placeholder {
            color: #e2e8f0 !important;
        }

        .dash-dropdown .Select-menu-outer {
            background: #111827 !important;
            border: 1px solid rgba(148,163,184,.22) !important;
            color: #f8fafc !important;
            z-index: 9999 !important;
        }

        .dash-dropdown .VirtualizedSelectOption {
            color: #e2e8f0;
        }

        .dash-dropdown .VirtualizedSelectFocusedOption {
            background: rgba(99,102,241,.22);
        }

        .sync-shell {
            margin-left: auto;
            text-align: right;
        }

        .sync-caption {
            color: #94a3b8;
            font-size: .62rem;
            font-weight: 850;
            letter-spacing: .09em;
            text-transform: uppercase;
            margin-bottom: 5px;
        }

        .sync-button {
            border: 1px solid rgba(99,102,241,.50);
            background: linear-gradient(135deg,
                rgba(79,70,229,.98), rgba(37,99,235,.96));
            color: #fff;
            min-height: 42px;
            padding: 0 16px;
            border-radius: 11px;
            font-weight: 850;
            cursor: pointer;
            box-shadow: 0 8px 20px rgba(37,99,235,.18);
            transition: transform .15s ease, box-shadow .15s ease;
        }

        .sync-button:hover {
            transform: translateY(-1px);
            box-shadow: 0 11px 24px rgba(37,99,235,.24);
        }

        .upload-box {
            border: 1px dashed rgba(99,102,241,.48);
            background: rgba(99,102,241,.035);
            border-radius: 14px;
            padding: 16px;
            margin-bottom: 10px;
        }

        .upload-note {
            color: #94a3b8;
            font-size: .78rem;
            line-height: 1.5;
            margin-bottom: 12px;
        }

        .chart-grid {
            display: grid;
            grid-template-columns: repeat(2,minmax(0,1fr));
            gap: 12px;
            width: 100%;
            min-width: 0;
            padding: 12px;
            border: 1px solid rgba(148,163,184,.18);
            border-radius: 15px;
            background: rgba(148,163,184,.012);
        }

        .chart-card {
            min-width: 0;
            border: 1px solid rgba(148,163,184,.18);
            border-radius: 15px;
            background: rgba(148,163,184,.012);
            box-shadow: 0 6px 18px rgba(15,23,42,.035);
            overflow: hidden;
        }

        .metric-grid {
            display: grid;
            grid-template-columns: repeat(4,minmax(0,1fr));
            gap: 12px;
            width: 100%;
        }

        .metric-card-base {
            position: relative;
            box-sizing: border-box;
            width: 100%;
            min-width: 0;
            min-height: 120px;
            padding: 15px 17px;
            border: 1px solid rgba(148,163,184,.18);
            border-radius: 15px;
            background: rgba(148,163,184,.025);
            box-shadow: 0 7px 20px rgba(15,23,42,.045);
            overflow: hidden;
        }

        .metric-card-base::before {
            content: "";
            position: absolute;
            left: 0;
            top: 0;
            bottom: 0;
            width: 3px;
            background: var(--accent,#6366f1);
        }

        .metric-title {
            color: #94a3b8;
            font-size: .67rem;
            font-weight: 850;
            text-transform: uppercase;
            letter-spacing: .075em;
        }

        .metric-value-sm {
            font-size: clamp(1.72rem,1.85vw,2.05rem);
            font-weight: 900;
            line-height: 1.05;
            margin: 9px 0 6px;
            letter-spacing: -.025em;
        }

        .metric-footnote {
            color: #94a3b8;
            font-size: .69rem;
            line-height: 1.35;
        }

        .info-strip {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            gap: 6px;
            width: 100%;
            margin: 6px 0 12px;
            padding: 8px 10px;
            border: 1px solid rgba(148,163,184,.18);
            border-radius: 12px;
            background: rgba(148,163,184,.018);
        }

        .info-chip {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            min-width: 0;
            padding: 4px 8px;
            border: 1px solid rgba(148,163,184,.18);
            border-radius: 999px;
            font-size: .67rem;
            font-weight: 700;
            color: #94a3b8;
            line-height: 1.25;
        }

        .panel {
            width: 100%;
            min-width: 0;
            padding: 12px;
            border: 1px solid rgba(148,163,184,.18);
            border-radius: 15px;
            background: rgba(148,163,184,.012);
            box-shadow: 0 6px 18px rgba(15,23,42,.035);
        }

        .two-column {
            display: grid;
            grid-template-columns: repeat(2,minmax(0,1fr));
            gap: 12px;
        }

        .pareto-panel {
            margin-bottom: 18px;
            padding: 12px;
            border: 1px solid rgba(148,163,184,.18);
            border-radius: 15px;
            background: rgba(148,163,184,.012);
        }

        .pareto-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
            width: 100%;
            padding: 9px 0;
            border-bottom: 2px solid;
            margin-bottom: 10px;
        }

        .pareto-count {
            white-space: nowrap;
            border: 1px solid rgba(128,128,128,.22);
            padding: 4px 8px;
            border-radius: 999px;
            font-size: .68rem;
            font-weight: 850;
        }

        .pareto-html-shell {
            width: 100%;
            max-width: 100%;
            min-width: 0;
            overflow-x: auto;
            border: 1px solid rgba(148,163,184,.18);
            border-radius: 12px;
        }

        .pareto-html-table {
            width: 100%;
            min-width: 680px;
            border-collapse: collapse;
            font-size: .78rem;
        }

        .pareto-html-table th {
            padding: .62rem .56rem;
            text-align: left;
            font-size: .67rem;
            font-weight: 850;
            letter-spacing: .05em;
            text-transform: uppercase;
            color: #94a3b8;
            background: rgba(148,163,184,.055);
            border-bottom: 1px solid rgba(148,163,184,.18);
        }

        .pareto-html-table td {
            padding: .60rem .56rem;
            border-bottom: 1px solid rgba(148,163,184,.11);
            vertical-align: middle;
        }

        .pareto-html-table tr:last-child td {
            border-bottom: none;
        }

        .pareto-html-table tr:hover {
            background: rgba(99,102,241,.035);
        }

        .pareto-status {
            display: inline-flex;
            align-items: center;
            max-width: 100%;
            padding: 3px 7px;
            border-radius: 999px;
            border: 1px solid rgba(148,163,184,.20);
            font-size: .67rem;
            font-weight: 800;
            line-height: 1.2;
        }

        .pareto-status.stockout {
            color: #f87171;
            background: rgba(248,113,113,.08);
            border-color: rgba(248,113,113,.22);
        }

        .empty-message {
            padding: 22px;
            color: #94a3b8;
            font-size: .8rem;
            border: 1px dashed rgba(148,163,184,.20);
            border-radius: 12px;
            text-align: center;
        }

        .alert {
            width: 100%;
            margin-bottom: 12px;
            padding: 12px 14px;
            border-radius: 12px;
            font-size: .78rem;
            font-weight: 750;
        }

        .alert-success {
            color: #34d399;
            background: rgba(16,185,129,.08);
            border: 1px solid rgba(16,185,129,.22);
        }

        .alert-error {
            color: #fb7185;
            background: rgba(244,63,94,.08);
            border: 1px solid rgba(244,63,94,.22);
        }

        .procurement-card {
            min-height: 250px;
            padding: 18px;
            border: 1px solid rgba(148,163,184,.18);
            border-radius: 15px;
            background: rgba(148,163,184,.025);
        }

        .procurement-title {
            font-size: 1rem;
            font-weight: 900;
            letter-spacing: .04em;
        }

        .procurement-text {
            color: #94a3b8;
            font-size: .78rem;
            line-height: 1.7;
        }

        .hidden { display: none !important; }

        hr {
            border: none;
            border-top: 1px solid rgba(148,163,184,.18);
            margin: 22px 0;
        }

        @media (max-width: 900px) {
            .chart-grid,
            .two-column {
                grid-template-columns: 1fr;
            }

            .metric-grid {
                grid-template-columns: repeat(2,minmax(0,1fr));
            }

            .section-heading {
                grid-template-columns: auto minmax(0,1fr);
            }

            .section-heading .subtitle {
                grid-column: 2;
                text-align: left;
            }

            .sync-shell {
                margin-left: 0;
                text-align: left;
            }
        }

        @media (max-width: 560px) {
            .metric-grid {
                grid-template-columns: 1fr;
            }

            .app-shell {
                padding: 7px;
            }

            .hero-shell {
                padding: 18px;
            }
        }
    </style>
</head>
<body>
    {%app_entry%}
    <footer>
        {%config%}
        {%scripts%}
        {%renderer%}
    </footer>
</body>
</html>
"""


# =========================================================
# 7. LAYOUT
# =========================================================

app.layout = html.Div(
    className="app-shell",
    children=[
        dcc.Store(id="file-hash", data="initial"),
        dcc.Store(id="active-tab", data="inventory"),
        dcc.Store(id="upload-data-store"),
        html.Div(id="upload-alert"),

        # Hero
        html.Div(
            className="hero-shell",
            children=[
                html.Div(
                    "Supply Chain Management • Executive Analytics",
                    className="hero-kicker",
                ),
                html.Div(
                    "MUTI MC SCM Executive Control Tower",
                    className="hero-title",
                ),
                html.Div(
                    "Inventory visibility, Pareto risk prioritization, "
                    "stockout trends, Days of Inventory, and branch-level "
                    "action monitoring.",
                    className="hero-subtitle",
                ),
                html.A(
                    "Launch Delivery Requirements Plan (DRP) ↗",
                    href="https://scmdrp.streamlit.app/",
                    target="_blank",
                    className="hero-link",
                ),
            ],
        ),

        # Tabs
        html.Div(
            className="tabs",
            children=[
                html.Button(
                    "📊 Inventory Control Tower",
                    id="btn-inventory",
                    n_clicks=0,
                    className="tab-btn active",
                ),
                html.Button(
                    "📦 Procurements",
                    id="btn-procurements",
                    n_clicks=0,
                    className="tab-btn inactive",
                ),
            ],
        ),

        # Inventory tab
        html.Div(
            id="inventory-wrapper",
            children=[
                html.Div(
                    className="control-row",
                    children=[
                        html.Div(
                            [
                                html.Div(
                                    "TIMEFRAME",
                                    className="control-label",
                                ),
                                dcc.Dropdown(
                                    id="timeframe-drop",
                                    options=[
                                        {
                                            "label": "Year-to-Date (YTD)",
                                            "value": "Year-to-Date (YTD)",
                                        },
                                        {
                                            "label": "Weekly View",
                                            "value": "Weekly View",
                                        },
                                    ],
                                    value="Year-to-Date (YTD)",
                                    clearable=False,
                                    className="dash-dropdown",
                                ),
                            ],
                            className="control-group",
                        ),
                        html.Div(
                            className="sync-shell",
                            children=[
                                html.Div(
                                    "Latest Workbook",
                                    className="sync-caption",
                                ),
                                dcc.Upload(
                                    id="upload-data",
                                    children=html.Button(
                                        "Data Sync",
                                        className="sync-button",
                                    ),
                                    multiple=False,
                                    accept=".xlsx,.xls",
                                ),
                            ],
                        ),
                    ],
                ),

                html.Div(
                    id="kpi-info-strip",
                    className="info-strip",
                ),

                html.Div(
                    id="charts-grid",
                    className="chart-grid",
                ),

                section_heading(
                    "Network Scope",
                    "Filter the operational view without changing "
                    "the network-level trend history",
                ),

                html.Div(
                    [
                        html.Div(
                            "AREA",
                            className="control-label",
                        ),
                        dcc.Dropdown(
                            id="network-drop",
                            clearable=False,
                            className="dash-dropdown",
                        ),
                    ],
                    className="control-group",
                ),

                section_heading(
                    "Performance Overview",
                    "Current raw-data stockout profile",
                ),

                html.Div(
                    id="performance-cards",
                    className="metric-grid",
                ),

                section_heading(
                    "Stock Out Rate per Area",
                    "Average = mean of Class A/B/C rates • "
                    "Class A = A Out ÷ A Total",
                ),

                html.Div(
                    id="area-bar-charts",
                    className="two-column",
                ),

                section_heading(
                    "Class A Branch Stockout Ranking",
                    "Top branch risks and zero-stockout leaders",
                ),

                html.Div(
                    id="branch-rankings",
                    className="two-column",
                ),

                html.Hr(),

                section_heading(
                    "Branch-Level Drilldown",
                    "Select a branch to review stockout risk "
                    "and Pareto action models",
                ),

                html.Div(
                    [
                        html.Div(
                            "BRANCH",
                            className="control-label",
                        ),
                        dcc.Dropdown(
                            id="branch-drop",
                            clearable=False,
                            className="dash-dropdown",
                        ),
                    ],
                    className="control-group",
                ),

                html.Div(
                    id="branch-cards",
                    className="metric-grid",
                ),

                html.Div(id="pareto-tables"),
            ],
        ),

        # Procurement tab
        html.Div(
            id="procurements-wrapper",
            className="hidden",
            children=[
                section_heading(
                    "Procurements Control",
                    "Manage purchase orders, incoming stock allocations, "
                    "and supplier lead times",
                ),
                html.Div(
                    className="two-column",
                    children=[
                        html.Div(
                            className="procurement-card",
                            style={
                                "borderLeft": "4px solid #6366f1"
                            },
                            children=[
                                html.Div(
                                    "🏍️ MOTORCYCLE UNITS",
                                    className="procurement-title",
                                    style={"color": "#818cf8"},
                                ),
                                html.Hr(),
                                html.Div(
                                    [
                                        html.B(
                                            "Pipeline Visibility",
                                            style={"color": "#cbd5e1"},
                                        ),
                                        html.Br(),
                                        "• Supplier Lead Times",
                                        html.Br(),
                                        "• Incoming Allocations",
                                        html.Br(),
                                        "• Backorder Tracking",
                                        html.Br(),
                                        html.Br(),
                                        html.I(
                                            "(Procurement data integration pending)",
                                            style={"opacity": ".7"},
                                        ),
                                    ],
                                    className="procurement-text",
                                ),
                            ],
                        ),
                        html.Div(
                            className="procurement-card",
                            style={
                                "borderLeft": "4px solid #10b981"
                            },
                            children=[
                                html.Div(
                                    "⚙️ SPARE PARTS",
                                    className="procurement-title",
                                    style={"color": "#34d399"},
                                ),
                                html.Hr(),
                                html.Div(
                                    [
                                        html.B(
                                            "Replenishment Status",
                                            style={"color": "#cbd5e1"},
                                        ),
                                        html.Br(),
                                        "• Active Purchase Orders",
                                        html.Br(),
                                        "• Critical Shortages",
                                        html.Br(),
                                        "• Parts Delivery Schedule",
                                        html.Br(),
                                        html.Br(),
                                        html.I(
                                            "(Procurement data integration pending)",
                                            style={"opacity": ".7"},
                                        ),
                                    ],
                                    className="procurement-text",
                                ),
                            ],
                        ),
                    ],
                ),
            ],
        ),
    ],
)


# =========================================================
# 8. CALLBACK: UPLOAD / DATA SYNC
# =========================================================

@app.callback(
    Output("upload-alert", "children"),
    Output("file-hash", "data"),
    Input("upload-data", "contents"),
    State("upload-data", "filename"),
    prevent_initial_call=True,
)
def handle_upload(contents, filename):
    if not contents:
        return no_update, no_update

    try:
        _, content_string = contents.split(",", 1)
        decoded = base64.b64decode(content_string)

        # Validate before replacing the current workbook.
        process_excel_file(decoded)

        save_local_cache(decoded)
        get_cached_data.cache_clear()

        file_hash = hashlib.sha256(decoded).hexdigest()

        return (
            html.Div(
                f"Data successfully synchronized: {filename or 'workbook'}",
                className="alert alert-success",
            ),
            file_hash,
        )

    except Exception as exc:
        return (
            html.Div(
                f"Data Sync failed: {exc}",
                className="alert alert-error",
            ),
            no_update,
        )


# =========================================================
# 9. CALLBACK: TABS
# =========================================================

@app.callback(
    Output("inventory-wrapper", "className"),
    Output("procurements-wrapper", "className"),
    Output("btn-inventory", "className"),
    Output("btn-procurements", "className"),
    Input("btn-inventory", "n_clicks"),
    Input("btn-procurements", "n_clicks"),
)
def toggle_tabs(inv_clicks, proc_clicks):
    ctx = dash.callback_context

    if not ctx.triggered:
        return "", "hidden", "tab-btn active", "tab-btn inactive"

    trigger = ctx.triggered[0]["prop_id"].split(".")[0]

    if trigger == "btn-procurements":
        return (
            "hidden",
            "",
            "tab-btn inactive",
            "tab-btn active",
        )

    return (
        "",
        "hidden",
        "tab-btn active",
        "tab-btn inactive",
    )


# =========================================================
# 10. CALLBACK: KPI TREND CHARTS
# =========================================================

@app.callback(
    Output("charts-grid", "children"),
    Output("kpi-info-strip", "children"),
    Input("file-hash", "data"),
    Input("timeframe-drop", "value"),
)
def update_kpi_charts(file_hash, timeframe):
    _, kpi_ytd, kpi_weekly = get_cached_data(file_hash)

    kpi_data = (
        kpi_weekly
        if timeframe == "Weekly View"
        else kpi_ytd
    )

    if kpi_data.empty:
        return (
            [empty_message("No KPI workbook data available.")],
            [
                info_chip(f"View: {timeframe}"),
                info_chip("Latest KPI: No KPI date"),
                info_chip("Actual data dates only"),
            ],
        )

    is_weekly = timeframe == "Weekly View"

    latest = pd.to_datetime(
        kpi_data["period"],
        errors="coerce",
    ).max()

    latest_label = (
        latest.strftime("%d %b %Y")
        if pd.notna(latest)
        else "No KPI date"
    )

    chart_specs = [
        (
            "class_a_doi",
            "MC Class A DoI",
            "CLASS A DAYS OF INVENTORY",
            "#7c3aed",
            False,
            True,
        ),
        (
            "overall_doi",
            "Days of Inventory",
            "OVERALL INVENTORY COVERAGE",
            "#2563eb",
            False,
            True,
        ),
        (
            "per_branch",
            "Per Branch OOS",
            "STOCKOUT RATE",
            "#0ea5e9",
            True,
            False,
        ),
        (
            "class_a_out",
            "Overall Class A Rate",
            "CLASS A STOCKOUT RATE",
            "#f43f5e",
            True,
            False,
        ),
        (
            "before_po",
            "Overall Before PO Balance",
            "STOCKOUT RATE BEFORE PO",
            "#f59e0b",
            True,
            False,
        ),
        (
            "after_po",
            "Overall After PO Balance",
            "STOCKOUT RATE AFTER PO",
            "#10b981",
            True,
            True,
        ),
    ]

    cards = []

    for column, title, subtitle, color, percentage, fill in chart_specs:
        figure = create_styled_line_chart(
            kpi_data,
            column,
            title,
            subtitle,
            color,
            is_weekly=is_weekly,
            is_percentage=percentage,
            fill=fill,
        )
        cards.append(chart_card(figure))

    return (
        cards,
        [
            info_chip(f"View: {timeframe}"),
            info_chip(f"Latest KPI: {latest_label}"),
            info_chip("Actual data dates only"),
        ],
    )


# =========================================================
# 11. CALLBACK: AREA FILTER
# =========================================================

@app.callback(
    Output("network-drop", "options"),
    Output("network-drop", "value"),
    Input("file-hash", "data"),
)
def update_network_drop(file_hash):
    raw, _, _ = get_cached_data(file_hash)

    if raw.empty:
        return [], None

    areas = ["All Areas"] + sorted(
        [
            str(area)
            for area in raw["area"].dropna().unique()
            if str(area).strip()
        ]
    )

    return (
        [{"label": area, "value": area} for area in areas],
        "All Areas",
    )


# =========================================================
# 12. CALLBACK: PERFORMANCE CARDS
# =========================================================

@app.callback(
    Output("performance-cards", "children"),
    Input("file-hash", "data"),
    Input("network-drop", "value"),
)
def update_performance(file_hash, area):
    raw, _, _ = get_cached_data(file_hash)

    if raw.empty:
        return [empty_message("No raw-data workbook available.")]

    if area and area != "All Areas":
        raw = raw[raw["area"] == area]

    rates = [
        (
            "Class A Rate",
            calculate_stockout_rate(raw, "Class A"),
            "Highest-priority Pareto",
            "red",
        ),
        (
            "Class B Rate",
            calculate_stockout_rate(raw, "Class B"),
            "Medium-priority Pareto",
            "amber",
        ),
        (
            "Class C Rate",
            calculate_stockout_rate(raw, "Class C"),
            "Lower-priority Pareto",
            "green",
        ),
    ]

    average = round_half_up(
        sum(item[1] for item in rates) / 3
    )

    rates.append(
        (
            "Average Rate",
            average,
            "Average of A, B and C",
            "blue",
        )
    )

    return [
        metric_card(title, value, note, accent)
        for title, value, note, accent in rates
    ]


# =========================================================
# 13. CALLBACK: AREA CHARTS
# =========================================================

@app.callback(
    Output("area-bar-charts", "children"),
    Input("file-hash", "data"),
)
def update_area_charts(file_hash):
    raw, _, _ = get_cached_data(file_hash)

    if raw.empty:
        return [empty_message("No raw-data workbook available.")]

    area_rates = []

    for area, area_df in raw.groupby(
        "area",
        sort=True,
        dropna=True,
    ):
        if not str(area).strip():
            continue

        average_rate = round_half_up(
            (
                calculate_stockout_rate(area_df, "Class A")
                + calculate_stockout_rate(area_df, "Class B")
                + calculate_stockout_rate(area_df, "Class C")
            )
            / 3
        )

        class_a_df = area_df[
            area_df["pareto_class"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.casefold()
            == "class a"
        ]

        total_a = len(class_a_df)

        out_a = (
            class_a_df["stock_status"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.casefold()
            .eq("stockout")
            .sum()
        )

        class_a_rate = (
            round_half_up((out_a / total_a) * 100)
            if total_a > 0
            else 0
        )

        area_rates.append(
            {
                "Area": area,
                "Average": average_rate,
                "ClassA": class_a_rate,
                "Count": out_a,
                "Total": total_a,
            }
        )

    data = pd.DataFrame(area_rates)

    if data.empty:
        return [empty_message("No area data available.")]

    data = data.sort_values(
        "Average",
        ascending=False,
    )

    fig_average = px.bar(
        data,
        x="Area",
        y="Average",
        text="Average",
        title="Average Stock Out Rate per Area",
    )

    fig_average.update_traces(
        marker_color="#6366f1",
        texttemplate="%{text:.0f}%",
        textposition="outside",
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Average: <b>%{y:.0f}%</b><extra></extra>"
        ),
    )

    fig_class_a = px.bar(
        data,
        x="Area",
        y="ClassA",
        text="ClassA",
        title="Class A Stock Out Rate per Area",
        custom_data=["Count", "Total"],
    )

    fig_class_a.update_traces(
        marker_color="#f43f5e",
        texttemplate="%{text:.0f}%",
        textposition="outside",
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Rate: <b>%{y:.0f}%</b><br>"
            "Count: <b>%{customdata[0]}</b><br>"
            "Total: <b>%{customdata[1]}</b>"
            "<extra></extra>"
        ),
    )

    for figure in [fig_average, fig_class_a]:
        figure.update_layout(
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            height=350,
            margin=dict(t=50, b=30, l=30, r=20),
            showlegend=False,
        )

    return [
        chart_card(fig_average),
        chart_card(fig_class_a),
    ]


# =========================================================
# 14. CALLBACK: BRANCH RANKINGS
# =========================================================

@app.callback(
    Output("branch-rankings", "children"),
    Input("file-hash", "data"),
    Input("network-drop", "value"),
)
def update_rankings(file_hash, area):
    raw, _, _ = get_cached_data(file_hash)

    if raw.empty:
        return [empty_message("No raw-data workbook available.")]

    if area and area != "All Areas":
        raw = raw[raw["area"] == area]

    class_a = raw[
        raw["pareto_class"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
        == "class a"
    ].copy()

    if class_a.empty:
        return [empty_message("No Class A branch data available.")]

    class_a["_is_out"] = (
        class_a["stock_status"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
        .eq("stockout")
        .astype(int)
    )

    summary = (
        class_a.groupby(
            ["area", "branch"],
            as_index=False,
        )
        .agg(
            OutCount=("_is_out", "sum"),
            TotalCount=("_is_out", "size"),
        )
    )

    summary["Rate"] = summary.apply(
        lambda row: (
            round_half_up(
                (row["OutCount"] / row["TotalCount"]) * 100
            )
            if row["TotalCount"] > 0
            else 0
        ),
        axis=1,
    )

    summary["Display"] = (
        summary["branch"]
        + (
            " • " + summary["area"]
            if area == "All Areas"
            else ""
        )
    )

    high_risk = (
        summary[summary["Rate"] > 0]
        .sort_values(
            ["Rate", "OutCount", "branch"],
            ascending=[False, False, True],
        )
        .head(10)
    )

    zero_risk = (
        summary[summary["Rate"] == 0]
        .sort_values(
            ["TotalCount", "branch"],
            ascending=[False, True],
        )
        .head(10)
    )

    fig_high = px.bar(
        high_risk,
        x="Rate",
        y="Display",
        orientation="h",
        text="Rate",
        title="Top Highest Class A Stockout Risk",
        custom_data=["OutCount"],
    )

    fig_high.update_traces(
        marker_color="#f43f5e",
        texttemplate="%{text:.0f}%",
        textposition="outside",
        hovertemplate=(
            "<b>%{y}</b><br>"
            "Rate: <b>%{x:.0f}%</b><extra></extra>"
        ),
    )

    fig_high.update_layout(
        yaxis={"categoryorder": "total ascending"}
    )

    fig_zero = px.bar(
        zero_risk,
        x="TotalCount",
        y="Display",
        orientation="h",
        title="Top Branches with 0% Class A Rate",
    )

    fig_zero.update_traces(
        marker_color="#10b981",
        texttemplate="0% OOS",
        textposition="outside",
        hovertemplate=(
            "<b>%{y}</b><br>"
            "Total Covered: <b>%{x}</b><extra></extra>"
        ),
    )

    fig_zero.update_layout(
        yaxis={"categoryorder": "total ascending"}
    )

    for figure in [fig_high, fig_zero]:
        figure.update_layout(
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            height=350,
            margin=dict(t=50, b=30, l=150, r=40),
            showlegend=False,
        )

    return [
        chart_card(fig_high),
        chart_card(fig_zero),
    ]


# =========================================================
# 15. CALLBACK: BRANCH DROPDOWN
# =========================================================

@app.callback(
    Output("branch-drop", "options"),
    Output("branch-drop", "value"),
    Input("file-hash", "data"),
    Input("network-drop", "value"),
)
def update_branch_drop(file_hash, area):
    raw, _, _ = get_cached_data(file_hash)

    if raw.empty:
        return [], None

    if area and area != "All Areas":
        raw = raw[raw["area"] == area]

    branches = ["All Branches"] + sorted(
        [
            str(branch)
            for branch in raw["branch"].dropna().unique()
            if str(branch).strip()
        ]
    )

    return (
        [{"label": branch, "value": branch} for branch in branches],
        "All Branches",
    )


# =========================================================
# 16. CALLBACK: BRANCH KPI CARDS
# =========================================================

@app.callback(
    Output("branch-cards", "children"),
    Input("file-hash", "data"),
    Input("network-drop", "value"),
    Input("branch-drop", "value"),
)
def update_branch_cards(file_hash, area, branch):
    raw, _, _ = get_cached_data(file_hash)

    if raw.empty:
        return [empty_message("No raw-data workbook available.")]

    if area and area != "All Areas":
        raw = raw[raw["area"] == area]

    if branch and branch != "All Branches":
        raw = raw[raw["branch"] == branch]

    rates = [
        (
            "CLASS A RATE",
            calculate_stockout_rate(raw, "Class A"),
            "red",
        ),
        (
            "CLASS B RATE",
            calculate_stockout_rate(raw, "Class B"),
            "amber",
        ),
        (
            "CLASS C RATE",
            calculate_stockout_rate(raw, "Class C"),
            "green",
        ),
    ]

    average = round_half_up(
        sum(item[1] for item in rates) / 3
    )

    rates.append(
        (
            "BRANCH AVERAGE",
            average,
            "blue",
        )
    )

    return [
        metric_card(title, value, "", accent)
        for title, value, accent in rates
    ]


# =========================================================
# 17. CALLBACK: PARETO ACTION TABLES
# =========================================================

@app.callback(
    Output("pareto-tables", "children"),
    Input("file-hash", "data"),
    Input("network-drop", "value"),
    Input("branch-drop", "value"),
)
def update_pareto(file_hash, area, branch):
    raw, _, _ = get_cached_data(file_hash)

    if raw.empty:
        return [empty_message("No raw-data workbook available.")]

    if area and area != "All Areas":
        raw = raw[raw["area"] == area]

    if branch and branch != "All Branches":
        raw = raw[raw["branch"] == branch]

    def make_table(p_class, color):
        data = raw[
            raw["pareto_class"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.casefold()
            == p_class.casefold()
        ].copy()

        count = len(data)

        if count == 0:
            return html.Div(
                f"No {p_class} items active.",
                className="empty-message",
            )

        data = data.sort_values(
            ["suggested_transfer", "doi"],
            ascending=[False, True],
        ).reset_index(drop=True)

        rows = []

        for index, row in data.iterrows():
            status = str(row["stock_status"])

            status_class = (
                "pareto-status stockout"
                if status.lower() == "stockout"
                else "pareto-status"
            )

            rows.append(
                html.Tr(
                    [
                        html.Td(
                            index + 1,
                            style={
                                "textAlign": "center",
                                "width": "7%",
                            },
                        ),
                        html.Td(
                            str(row["model"]),
                            style={"width": "35%"},
                        ),
                        html.Td(
                            html.Span(
                                status,
                                className=status_class,
                            ),
                            style={"width": "16%"},
                        ),
                        html.Td(
                            f"{int(row['remaining_inventory']):,}",
                            style={
                                "textAlign": "right",
                                "width": "14%",
                            },
                        ),
                        html.Td(
                            f"{int(row['suggested_transfer']):,}",
                            style={
                                "textAlign": "right",
                                "width": "14%",
                            },
                        ),
                        html.Td(
                            f"{int(row['doi']):,}",
                            style={
                                "textAlign": "right",
                                "width": "14%",
                            },
                        ),
                    ]
                )
            )

        return html.Div(
            className="pareto-panel",
            children=[
                html.Div(
                    className="pareto-header",
                    style={"borderColor": color},
                    children=[
                        html.Span(
                            p_class.upper(),
                            style={"color": color},
                        ),
                        html.Span(
                            f"● {count} Items",
                            className="pareto-count",
                            style={
                                "color": color,
                                "borderColor": color,
                            },
                        ),
                    ],
                ),
                html.Div(
                    className="pareto-html-shell",
                    children=[
                        html.Table(
                            className="pareto-html-table",
                            children=[
                                html.Thead(
                                    html.Tr(
                                        [
                                            html.Th(
                                                "Rank",
                                                style={
                                                    "textAlign": "center"
                                                },
                                            ),
                                            html.Th("Model"),
                                            html.Th("Status"),
                                            html.Th(
                                                "Inventory",
                                                style={
                                                    "textAlign": "right"
                                                },
                                            ),
                                            html.Th(
                                                "Transfer",
                                                style={
                                                    "textAlign": "right"
                                                },
                                            ),
                                            html.Th(
                                                "DOI",
                                                style={
                                                    "textAlign": "right"
                                                },
                                            ),
                                        ]
                                    )
                                ),
                                html.Tbody(rows),
                            ],
                        )
                    ],
                ),
            ],
        )

    return [
        make_table("Class A", "#f87171"),
        make_table("Class B", "#fbbf24"),
        make_table("Class C", "#4ade80"),
    ]


# =========================================================
# 18. INITIAL DATA LOAD
# =========================================================

def initialize_from_existing_cache():
    """
    Load the existing local workbook at startup if present.
    The UI remains usable without a workbook and displays
    an empty-state message until Data Sync is performed.
    """
    cached = load_local_cache()

    if not cached:
        return

    try:
        process_excel_file(cached)
        get_cached_data.cache_clear()
    except Exception:
        # Do not crash the web server because of an old/invalid cache.
        pass


initialize_from_existing_cache()


# =========================================================
# 19. RUN
# =========================================================

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=8050)
'''

path = Path("/mnt/data/muti_scm_dashboard_dash.py")
path.write_text(code, encoding="utf-8")

requirements = """dash>=2.18
pandas>=2.2
numpy>=1.26
plotly>=5.24
openpyxl>=3.1
xlrd>=2.0.1
"""

req_path = Path("/mnt/data/requirements.txt")
req_path.write_text(requirements, encoding="utf-8")

print(f"Created: {path}")
print(f"Created: {req_path}")
print(f"Python lines: {len(code.splitlines())}")
