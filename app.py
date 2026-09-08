import dash
from dash import dcc, html, Input, Output, State, no_update
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import os
import io
import base64
import functools
import hashlib
from decimal import Decimal, ROUND_HALF_UP
from datetime import date, datetime

# =========================================================
# 0. UTILITIES & DATA PROCESSING
# =========================================================
LOCAL_CACHE_FILE = "persistent_scm_data.xlsx"
XLSX_SIGNATURES = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")
XLS_SIGNATURE = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"

def detect_excel_engine(file_bytes):
    signature = bytes(file_bytes[:8])
    if any(signature.startswith(sig) for sig in XLSX_SIGNATURES): return "openpyxl"
    if signature.startswith(XLS_SIGNATURE): return "xlrd"
    raise ValueError("Not a valid Excel workbook.")

def round_half_up(value, ndigits=0):
    if pd.isna(value): return np.nan
    quantizer = Decimal("1").scaleb(-ndigits)
    rounded = Decimal(str(float(value))).quantize(quantizer, rounding=ROUND_HALF_UP)
    return int(rounded) if ndigits == 0 else float(rounded)

def round_series_half_up(series):
    numeric = pd.to_numeric(series, errors="coerce")
    return numeric.map(lambda value: round_half_up(value) if pd.notna(value) else np.nan)

def save_local_cache(file_bytes):
    with open(LOCAL_CACHE_FILE, "wb") as f: f.write(file_bytes)

def load_local_cache():
    if os.path.exists(LOCAL_CACHE_FILE):
        with open(LOCAL_CACHE_FILE, "rb") as f: return f.read()
    return None

def process_excel_file(file_bytes):
    excel_engine = detect_excel_engine(file_bytes)
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine=excel_engine)
    sheet_map = {str(s).lower().strip().replace(" ", "_"): s for s in xls.sheet_names}
    
    def get_sheet(target): return sheet_map[target]

    raw_df = pd.read_excel(xls, sheet_name=get_sheet("raw_data"))
    raw_df.columns = raw_df.columns.astype(str).str.lower().str.strip().str.replace(" ", "_", regex=False)
    if "class" in raw_df.columns: raw_df.rename(columns={"class": "pareto_class"}, inplace=True)

    for col in ["remaining_inventory", "suggested_transfer", "doi"]:
        raw_df[col] = round_series_half_up(pd.to_numeric(raw_df[col], errors="coerce").fillna(0)).fillna(0).astype(int)

    raw_df["pareto_class"] = raw_df["pareto_class"].astype(str).str.strip()
    raw_df["stock_status"] = raw_df["stock_status"].fillna("").astype(str).str.strip()
    raw_df["area"] = raw_df["area"].fillna("").astype(str).str.strip()
    raw_df["branch"] = raw_df["branch"].fillna("").astype(str).str.strip()

    def parse_kpi(sheet_name):
        df_raw = pd.read_excel(xls, sheet_name=get_sheet(sheet_name), header=None)
        if df_raw.empty or df_raw.shape[1] < 2: return pd.DataFrame()
        
        # Determine Date Row
        best_date_row = None
        best_date_columns = []
        best_dates = []
        
        for row_idx in range(min(10, len(df_raw))):
            r_cols, r_dates = [], []
            for col_idx in range(df_raw.shape[1]):
                val = df_raw.iat[row_idx, col_idx]
                parsed = pd.to_datetime(val, errors='coerce')
                if pd.notna(parsed) and 2000 <= parsed.year <= 2100:
                    r_cols.append(col_idx)
                    r_dates.append(parsed)
            if len(r_cols) > len(best_date_columns):
                best_date_row, best_date_columns, best_dates = row_idx, r_cols, r_dates
                
        if not best_date_columns: return pd.DataFrame()

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
        for kpi, col in kpis.items():
            idx = df_raw.index[first_col.str.contains(kpi, regex=False, na=False)].tolist()
            data[col] = pd.to_numeric(df_raw.loc[idx[0], best_date_columns].values, errors="coerce") if idx else [np.nan]*len(best_dates)
            
        clean_df = pd.DataFrame(data).dropna(subset=["period"])
        metric_cols = list(kpis.values())
        
        def last_valid(series):
            non_null = series.dropna()
            return non_null.iloc[-1] if not non_null.empty else np.nan
            
        clean_df = clean_df.sort_values("period").groupby("period", as_index=False).agg({col: last_valid for col in metric_cols})
        
        if "ytd" in sheet_name.lower() and not clean_df.empty:
            clean_df["month_year"] = clean_df["period"].dt.to_period("M")
            clean_df = clean_df.sort_values("period").drop_duplicates(subset=["month_year"], keep="last").drop(columns=["month_year"])
            
        return clean_df.sort_values("period").reset_index(drop=True)

    return raw_df, parse_kpi("kpi_ytd_input"), parse_kpi("kpi_weekly_input")

@functools.lru_cache(maxsize=1)
def get_cached_data(file_hash):
    bytes_data = load_local_cache()
    if bytes_data: return process_excel_file(bytes_data)
    return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# =========================================================
# 1. CHART LOGIC
# =========================================================
def calculate_stockout_rate(df, pareto_class=None):
    subset = df[df["pareto_class"] == pareto_class] if pareto_class else df
    if len(subset) == 0: return 0
    stockouts = (subset["stock_status"].fillna("").astype(str).str.lower().str.strip() == "stockout").sum()
    return round_half_up((stockouts / len(subset)) * 100)

def prepare_chart_series(df, y_col):
    if df.empty or y_col not in df.columns: return pd.DataFrame(columns=["period", y_col])
    chart_df = df[["period", y_col]].copy()
    chart_df["period"] = pd.to_datetime(chart_df["period"], errors="coerce")
    chart_df[y_col] = pd.to_numeric(chart_df[y_col], errors="coerce")
    return chart_df.dropna().sort_values("period").drop_duplicates(subset=["period"], keep="last").reset_index(drop=True)

def infer_percentage_scale(series):
    numeric = pd.to_numeric(series, errors="coerce").dropna()
    if numeric.empty: return 1.0
    return 0.01 if numeric.abs().max() > 1.5 else 1.0

def create_styled_line_chart(df, y_col, title, subtitle, line_color, is_weekly, is_percentage=True, fill=False):
    chart_df = prepare_chart_series(df, y_col)
    fig = go.Figure()

    if chart_df.empty:
        fig.add_annotation(text="No valid data available", x=0.5, y=0.5, xref="paper", yref="paper", showarrow=False, font=dict(color="#94a3b8"))
        fig.update_layout(template="plotly", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=350, xaxis=dict(visible=False), yaxis=dict(visible=False))
        return fig

    plot_y = chart_df[y_col].copy()
    if is_percentage:
        scale = infer_percentage_scale(plot_y)
        plot_y = round_series_half_up((plot_y * scale) * 100.0) / 100.0
        text_labels = [f"{v * 100:.0f}%" for v in plot_y]
        tick_format = ".0%"
        default_ceiling = 0.10
    else:
        plot_y = round_series_half_up(plot_y).astype(int)
        text_labels = [f"{v:,.0f}" for v in plot_y]
        tick_format = ",.0f"
        default_ceiling = 10

    if is_weekly:
        chart_x = chart_df["period"].dt.strftime("%d %b %Y")
        tick_text = chart_df["period"].dt.strftime("%d %b")
        category_array = chart_x.tolist()
    else:
        chart_x = chart_df["period"].dt.strftime("%b")
        tick_text = chart_x.tolist()
        category_array = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

    hover_dates = chart_df["period"].dt.strftime("%d %b %Y")
    hover_template = f"<b>%{{customdata}}</b><br>{title}: <b>%{{y{':.0%' if is_percentage else ':,.0f'}}}</b><extra></extra>"
    text_positions = ["top center" if i % 2 == 0 else "bottom center" for i in range(len(chart_df))]

    fillcolor = f"rgba({int(line_color[1:3], 16)}, {int(line_color[3:5], 16)}, {int(line_color[5:7], 16)}, 0.08)" if fill else None

    fig.add_trace(go.Scatter(
        x=chart_x, y=plot_y, customdata=hover_dates, name="Actual", mode="lines+markers+text",
        text=text_labels, textposition=text_positions, textfont=dict(size=10, family="Arial"),
        line=dict(shape="linear" if is_weekly else "spline", width=3.2, color=line_color),
        marker=dict(size=8, color=line_color, line=dict(width=2, color="#ffffff")),
        fill="tozeroy" if fill else "none", fillcolor=fillcolor, hovertemplate=hover_template, cliponaxis=False
    ))

    # Trend Guide Calculation
    trend_y = None
    trend_direction = "Stable"
    if len(chart_df) >= 2:
        trend_x = (chart_df["period"] - chart_df["period"].min()).dt.total_seconds().to_numpy(dtype=float) / 86400.0
        trend_source_y = pd.to_numeric(plot_y, errors="coerce").to_numpy(dtype=float)
        valid = np.isfinite(trend_x) & np.isfinite(trend_source_y)

        if valid.sum() >= 2 and np.ptp(trend_x[valid]) > 0:
            slope, intercept = np.polyfit(trend_x[valid], trend_source_y[valid], 1)
            fitted_y = slope * trend_x + intercept
            
            actual_span = max(float(np.nanmax(trend_source_y[valid])) - float(np.nanmin(trend_source_y[valid])), 0.0)
            gap = max(actual_span * 0.18, 0.010 if is_percentage else 1.0)
            amplitude = max(actual_span * 0.10, 0.006 if is_percentage else 0.65)
            
            fitted_range = np.ptp(fitted_y[valid])
            normalized_fit = (fitted_y - np.nanmin(fitted_y[valid])) / fitted_range if fitted_range > 0 else np.full_like(fitted_y, 0.5)
            trend_y = float(np.nanmax(trend_source_y[valid])) + gap + (normalized_fit * amplitude)
            
            fitted_delta = float(fitted_y[-1] - fitted_y[0])
            flat_threshold = max(actual_span * 0.03, 0.001 if is_percentage else 0.10)
            trend_direction = "Upward" if fitted_delta > flat_threshold else "Downward" if fitted_delta < -flat_threshold else "Stable"
            
            dir_sym = {"Upward": "↑", "Downward": "↓", "Stable": "→"}[trend_direction]
            
            fig.add_trace(go.Scatter(
                x=chart_x, y=trend_y, name=f"Trend {dir_sym}", mode="lines",
                line=dict(width=2.4, dash="dash", color="rgba(100,116,139,0.95)"),
                hovertemplate=f"Trend: <b>{trend_direction} {dir_sym}</b><extra></extra>"
            ))

    y_max = max(pd.to_numeric(plot_y, errors="coerce").max(), np.nanmax(trend_y) if trend_y is not None else 0)
    y_max = max(y_max * 1.16, 0.05 if is_percentage else 5) if y_max > 0 else default_ceiling

    latest_text = chart_df["period"].max().strftime("%d %b %Y")
    
    xaxis_config = dict(type="category", showgrid=False, categoryorder="array", categoryarray=category_array, linecolor="rgba(148,163,184,0.18)")
    if is_weekly: xaxis_config.update(tickmode="array", tickvals=chart_x.tolist(), ticktext=tick_text.tolist())
    else: xaxis_config.update(tickmode="array", tickvals=chart_x.tolist(), ticktext=tick_text)

    fig.update_layout(
        template="plotly", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        height=365, margin=dict(t=82, b=44, l=46, r=20),
        hovermode="closest", hoverlabel=dict(bgcolor="#0f172a", bordercolor="rgba(148,163,184,0.28)", font=dict(color="#f8fafc", size=11)),
        showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=1.015, xanchor="right", x=0.99, font=dict(size=10, color="#94a3b8")),
        title=dict(text=f"{title}<br><span style='font-size:10px; color:#94a3b8;'>{subtitle} • LATEST {latest_text.upper()}</span>", x=0.02, y=0.96, font=dict(size=16)),
        xaxis=xaxis_config, yaxis=dict(showgrid=True, gridcolor="rgba(148,163,184,0.14)", zeroline=False, tickformat=tick_format, range=[0, y_max])
    )
    return fig

# =========================================================
# 2. DASH APPLICATION INITIALIZATION
# =========================================================
app = dash.Dash(__name__, suppress_callback_exceptions=True, external_scripts=[{'src': 'https://cdn.tailwindcss.com'}])
server = app.server
app.title = "SCM Executive Control Tower"

def section_heading(title, subtitle=""):
    return html.Div(className="flex items-center gap-2 mb-4 mt-6", children=[
        html.Div(className="w-2 h-2 rounded-full bg-indigo-500 shadow-[0_0_0_4px_rgba(99,102,241,0.11)]"),
        html.Span(title, className="text-lg font-black tracking-tight text-white"),
        html.Span(subtitle, className="text-slate-400 text-xs font-semibold ml-auto text-right")
    ])

# =========================================================
# 3. LAYOUT
# =========================================================
app.layout = html.Div(className="w-full max-w-none px-4 py-4 bg-[#0b1220] min-h-screen font-sans", children=[
    dcc.Store(id='file-hash', data="initial"),
    
    # HERO SECTION
    html.Div(className="border border-indigo-500/30 border-l-4 border-l-indigo-500 rounded-2xl bg-[#0b1220] p-6 lg:p-8 shadow-2xl mb-8", children=[
        html.Div("Supply Chain Management • Executive Analytics", className="text-indigo-300 text-xs font-extrabold tracking-widest uppercase mb-2"),
        html.Div("MUTI MC SCM Executive Control Tower", className="text-slate-50 text-2xl lg:text-3xl font-black mb-3 tracking-tight"),
        html.Div("Inventory visibility, Pareto risk prioritization, stockout trends, and branch-level action monitoring.", className="text-slate-400 text-sm mb-5"),
    ]),

    # TABS
    html.Div(className="flex space-x-4 border-b border-slate-700/50 mb-6", children=[
        html.Button("📊 Inventory Control Tower", id="btn-inventory", n_clicks=0, className="tab-btn active"),
        html.Button("📦 Procurements", id="btn-procurements", n_clicks=0, className="tab-btn inactive"),
    ]),

    html.Div(id='upload-alert'),

    # TAB 1: INVENTORY CONTENT
    html.Div(id="inventory-wrapper", children=[
        html.Div(className="flex justify-between items-end mb-2", children=[
            section_heading("MUTI MC Trends", ""),
            html.Div(className="text-right", children=[
                html.Div("Latest Workbook", className="text-slate-400 text-[10px] font-bold uppercase tracking-widest mb-1"),
                dcc.Upload(id='upload-data', children=html.Button("Data Sync", className="bg-indigo-600 hover:bg-indigo-500 text-white font-bold py-2 px-4 rounded-lg shadow cursor-pointer text-sm"), multiple=False)
            ])
        ]),
        
        dcc.Dropdown(id='timeframe-drop', options=['Year-to-Date (YTD)', 'Weekly View'], value='Year-to-Date (YTD)', className="w-64 mb-4 text-black text-sm"),
        html.Div(id='charts-grid', className="grid grid-cols-1 lg:grid-cols-2 gap-4 border border-slate-700/50 rounded-xl p-4 bg-slate-900/20"),
        
        section_heading("Network Scope", "Filter the operational view without changing the network-level trend history"),
        dcc.Dropdown(id='network-drop', className="w-64 mb-4 text-black text-sm"),
        
        section_heading("Performance Overview", "Current raw-data stockout profile"),
        html.Div(id='performance-cards', className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8"),
        
        section_heading("Stock Out Rate per Area", "Average = mean of Class A/B/C rates • Class A = A Out ÷ A Total"),
        html.Div(id='area-bar-charts', className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-8 border border-slate-700/50 rounded-xl p-4 bg-slate-900/20"),
        
        section_heading("Class A Branch Stockout Ranking", "Top branch risks and zero-stockout leaders"),
        html.Div(id='branch-rankings', className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-8 border border-slate-700/50 rounded-xl p-4 bg-slate-900/20"),
        
        html.Hr(className="border-slate-700/50 my-8"),
        
        section_heading("Branch-Level Drilldown", "Select a branch to review stockout risk and Pareto action models"),
        dcc.Dropdown(id='branch-drop', className="w-64 mb-4 text-black text-sm"),
        html.Div(id='branch-cards', className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8"),
        html.Div(id='pareto-tables')
    ]),

    # TAB 2: PROCUREMENTS CONTENT
    html.Div(id="procurements-wrapper", className="hidden", children=[
        section_heading("Procurements Control", "Manage purchase orders, incoming stock allocations, and supplier lead times"),
        html.Div(className="grid grid-cols-1 md:grid-cols-2 gap-4 w-full", children=[
            html.Div(className="metric-card border-l-4 border-l-indigo-500", style={"minHeight": "250px"}, children=[
                html.Div("🏍️ MOTORCYCLE UNITS", className="text-indigo-400 font-black text-lg"),
                html.Hr(className="border-slate-700/50 my-3"),
                html.Div(className="text-slate-400 text-sm leading-relaxed", children=[
                    html.B("Pipeline Visibility", className="text-slate-300"), html.Br(),
                    "• Supplier Lead Times", html.Br(), "• Incoming Allocations", html.Br(), "• Backorder Tracking", html.Br(), html.Br(),
                    html.I("(Procurement data integration pending)", className="opacity-70")
                ])
            ]),
            html.Div(className="metric-card border-l-4 border-l-emerald-500", style={"minHeight": "250px"}, children=[
                html.Div("⚙️ SPARE PARTS", className="text-emerald-400 font-black text-lg"),
                html.Hr(className="border-slate-700/50 my-3"),
                html.Div(className="text-slate-400 text-sm leading-relaxed", children=[
                    html.B("Replenishment Status", className="text-slate-300"), html.Br(),
                    "• Active Purchase Orders", html.Br(), "• Critical Shortages", html.Br(), "• Parts Delivery Schedule", html.Br(), html.Br(),
                    html.I("(Procurement data integration pending)", className="opacity-70")
                ])
            ])
        ])
    ])
])

# =========================================================
# 4. CALLBACKS
# =========================================================
@app.callback(
    Output('upload-alert', 'children'), Output('file-hash', 'data'),
    Input('upload-data', 'contents'), prevent_initial_call=True
)
def handle_upload(contents):
    if contents:
        _, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        save_local_cache(decoded)
        get_cached_data.cache_clear() 
        alert = html.Div("Data successfully synchronized.", className="p-4 mb-4 text-emerald-400 bg-emerald-900/20 rounded-lg border border-emerald-500/30 text-sm font-bold")
        return alert, hashlib.sha256(decoded).hexdigest()
    return no_update, no_update

@app.callback(
    Output('inventory-wrapper', 'className'), Output('procurements-wrapper', 'className'),
    Output('btn-inventory', 'className'), Output('btn-procurements', 'className'),
    Input('btn-inventory', 'n_clicks'), Input('btn-procurements', 'n_clicks')
)
def toggle_tabs(inv_clicks, proc_clicks):
    ctx = dash.callback_context
    if not ctx.triggered or ctx.triggered[0]["prop_id"].split(".")[0] == "btn-inventory":
        return "", "hidden", "tab-btn active", "tab-btn inactive"
    return "hidden", "w-full", "tab-btn inactive", "tab-btn active"

@app.callback(
    Output('charts-grid', 'children'),
    Input('file-hash', 'data'), Input('timeframe-drop', 'value')
)
def update_kpi_charts(file_hash, timeframe):
    _, kpi_ytd, kpi_weekly = get_cached_data(file_hash)
    if kpi_ytd.empty and kpi_weekly.empty: return []
    
    kpi_data = kpi_weekly if timeframe == "Weekly View" else kpi_ytd
    is_weekly = timeframe == "Weekly View"

    figs = [
        create_styled_line_chart(kpi_data, "class_a_doi", "MC Class A DoI", "CLASS A DAYS OF INVENTORY", "#7c3aed", is_weekly, False, True),
        create_styled_line_chart(kpi_data, "overall_doi", "Days of Inventory", "OVERALL INVENTORY COVERAGE", "#2563eb", is_weekly, False, True),
        create_styled_line_chart(kpi_data, "per_branch", "Per Branch OOS", "STOCKOUT RATE", "#0ea5e9", is_weekly, True),
        create_styled_line_chart(kpi_data, "class_a_out", "Overall Class A Rate", "CLASS A STOCKOUT RATE", "#f43f5e", is_weekly, True),
        create_styled_line_chart(kpi_data, "before_po", "Overall Before PO Balance", "STOCKOUT RATE BEFORE PO", "#f59e0b", is_weekly, True),
        create_styled_line_chart(kpi_data, "after_po", "Overall After PO Balance", "STOCKOUT RATE AFTER PO", "#10b981", is_weekly, True, True),
    ]
    return [dcc.Graph(figure=fig, config={'displayModeBar': False}) for fig in figs]

@app.callback(
    Output('network-drop', 'options'), Output('network-drop', 'value'),
    Input('file-hash', 'data')
)
def update_network_drop(file_hash):
    raw, _, _ = get_cached_data(file_hash)
    if raw.empty: return [], None
    areas = ["All Areas"] + sorted([a for a in raw["area"].dropna().unique() if str(a).strip()])
    return [{'label': a, 'value': a} for a in areas], "All Areas"

@app.callback(
    Output('performance-cards', 'children'),
    Input('file-hash', 'data'), Input('network-drop', 'value')
)
def update_performance(file_hash, area):
    raw, _, _ = get_cached_data(file_hash)
    if raw.empty: return []
    if area and area != "All Areas": raw = raw[raw['area'] == area]
    
    rates = {
        "Class A Rate": (calculate_stockout_rate(raw, "Class A"), "Highest-priority Pareto"),
        "Class B Rate": (calculate_stockout_rate(raw, "Class B"), "Medium-priority Pareto"),
        "Class C Rate": (calculate_stockout_rate(raw, "Class C"), "Lower-priority Pareto"),
    }
    avg = round_half_up(sum(r[0] for r in rates.values()) / 3)
    rates["Average Rate"] = (avg, "Average of A, B and C")

    return [
        html.Div(className='metric-card-base', children=[
            html.Div(title, className='metric-title'),
            html.Div(f"{val}%", className='metric-value-sm'),
            html.Div(note, className='metric-footnote')
        ]) for title, (val, note) in rates.items()
    ]

@app.callback(
    Output('area-bar-charts', 'children'),
    Input('file-hash', 'data')
)
def update_area_charts(file_hash):
    raw, _, _ = get_cached_data(file_hash)
    if raw.empty: return []
    
    area_rates = []
    for area, a_df in raw.groupby("area", sort=True, dropna=True):
        if not str(area).strip(): continue
        avg_rate = round_half_up((calculate_stockout_rate(a_df, "Class A") + calculate_stockout_rate(a_df, "Class B") + calculate_stockout_rate(a_df, "Class C")) / 3)
        class_a_df = a_df[a_df["pareto_class"].fillna("").astype(str).str.strip().str.casefold() == "class a"]
        total_a = len(class_a_df)
        out_a = (class_a_df["stock_status"].fillna("").astype(str).str.strip().str.casefold() == "stockout").sum()
        a_rate = round_half_up((out_a / total_a) * 100) if total_a > 0 else 0
        area_rates.append({"Area": area, "Average": avg_rate, "ClassA": a_rate, "Count": out_a, "Total": total_a})

    df = pd.DataFrame(area_rates).sort_values("Average", ascending=False)
    if df.empty: return []

    fig_avg = px.bar(df, x="Area", y="Average", text="Average", title="Average Stock Out Rate per Area")
    fig_avg.update_traces(marker_color="#6366f1", texttemplate="%{text:.0f}%", textposition="outside", hovertemplate="<b>%{x}</b><br>Average: <b>%{y:.0f}%</b><extra></extra>")
    
    fig_a = px.bar(df, x="Area", y="ClassA", text="ClassA", title="Class A Stock Out Rate per Area", custom_data=["Count", "Total"])
    fig_a.update_traces(marker_color="#f43f5e", texttemplate="%{text:.0f}%", textposition="outside", hovertemplate="<b>%{x}</b><br>Rate: <b>%{y:.0f}%</b><br>Count: <b>%{customdata[0]}</b><extra></extra>")

    for f in [fig_avg, fig_a]:
        f.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=350, margin=dict(t=50, b=30, l=30, r=20), showlegend=False)
        
    return [dcc.Graph(figure=fig_avg, config={'displayModeBar': False}), dcc.Graph(figure=fig_a, config={'displayModeBar': False})]

@app.callback(
    Output('branch-rankings', 'children'),
    Input('file-hash', 'data'), Input('network-drop', 'value')
)
def update_rankings(file_hash, area):
    raw, _, _ = get_cached_data(file_hash)
    if raw.empty: return []
    if area and area != "All Areas": raw = raw[raw['area'] == area]
    
    b_class_a = raw[raw["pareto_class"].fillna("").astype(str).str.strip().str.casefold() == "class a"].copy()
    if b_class_a.empty: return []
    
    b_class_a["_is_out"] = (b_class_a["stock_status"].fillna("").astype(str).str.strip().str.casefold() == "stockout").astype(int)
    summary = b_class_a.groupby(["area", "branch"], as_index=False).agg(OutCount=("_is_out", "sum"), TotalCount=("_is_out", "size"))
    summary["Rate"] = summary.apply(lambda r: round_half_up((r["OutCount"] / r["TotalCount"]) * 100) if r["TotalCount"] > 0 else 0, axis=1)
    summary["Display"] = summary["branch"] + ("  •  " + summary["area"] if area == "All Areas" else "")

    high_risk = summary[summary["Rate"] > 0].sort_values(["Rate", "OutCount", "branch"], ascending=[False, False, True]).head(10)
    zero_risk = summary[summary["Rate"] == 0].sort_values(["TotalCount", "branch"], ascending=[False, True]).head(10)

    fig_high = px.bar(high_risk, x="Rate", y="Display", orientation="h", text="Rate", title="Top Highest Class A Stockout Risk", custom_data=["OutCount"])
    fig_high.update_traces(marker_color="#f43f5e", texttemplate="%{text:.0f}%", textposition="outside", hovertemplate="<b>%{y}</b><br>Rate: <b>%{x:.0f}%</b><extra></extra>")
    fig_high.update_layout(yaxis={'categoryorder':'total ascending'})

    fig_zero = px.bar(zero_risk, x="TotalCount", y="Display", orientation="h", title="Top Branches with 0% Class A Rate")
    fig_zero.update_traces(marker_color="#10b981", texttemplate="0% OOS", textposition="outside", hovertemplate="<b>%{y}</b><br>Total Covered: <b>%{x}</b><extra></extra>")
    fig_zero.update_layout(yaxis={'categoryorder':'total ascending'})

    for f in [fig_high, fig_zero]:
        f.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", height=350, margin=dict(t=50, b=30, l=150, r=40), showlegend=False)

    return [dcc.Graph(figure=fig_high, config={'displayModeBar': False}), dcc.Graph(figure=fig_zero, config={'displayModeBar': False})]

@app.callback(
    Output('branch-drop', 'options'), Output('branch-drop', 'value'),
    Input('file-hash', 'data'), Input('network-drop', 'value')
)
def update_branch_drop(file_hash, area):
    raw, _, _ = get_cached_data(file_hash)
    if raw.empty: return [], None
    if area and area != "All Areas": raw = raw[raw['area'] == area]
    branches = ["All Branches"] + sorted([b for b in raw["branch"].dropna().unique() if str(b).strip()])
    return [{'label': b, 'value': b} for b in branches], "All Branches"

@app.callback(
    Output('branch-cards', 'children'),
    Input('file-hash', 'data'), Input('network-drop', 'value'), Input('branch-drop', 'value')
)
def update_branch_cards(file_hash, area, branch):
    raw, _, _ = get_cached_data(file_hash)
    if raw.empty: return []
    if area and area != "All Areas": raw = raw[raw['area'] == area]
    if branch and branch != "All Branches": raw = raw[raw['branch'] == branch]

    rates = [
        ("CLASS A RATE", calculate_stockout_rate(raw, "Class A"), "text-red-400"),
        ("CLASS B RATE", calculate_stockout_rate(raw, "Class B"), "text-yellow-400"),
        ("CLASS C RATE", calculate_stockout_rate(raw, "Class C"), "text-green-400"),
    ]
    avg = round_half_up(sum(r[1] for r in rates) / 3)
    rates.append(("BRANCH AVERAGE", avg, "text-blue-400"))

    return [
        html.Div(className='metric-card-base', children=[
            html.Div(title, className='metric-title'),
            html.Div(f"{val}%", className=f'metric-value-sm {color}')
        ]) for title, val, color in rates
    ]

@app.callback(
    Output('pareto-tables', 'children'),
    Input('file-hash', 'data'), Input('network-drop', 'value'), Input('branch-drop', 'value')
)
def update_pareto(file_hash, area, branch):
    raw, _, _ = get_cached_data(file_hash)
    if raw.empty: return []
    if area and area != "All Areas": raw = raw[raw['area'] == area]
    if branch and branch != "All Branches": raw = raw[raw['branch'] == branch]

    def make_table(p_class, color):
        df = raw[raw["pareto_class"] == p_class].copy()
        count = len(df)
        if count == 0: return html.Div(f"No {p_class} items active.", className="text-slate-400 text-sm mb-6")
        
        df = df.sort_values(["suggested_transfer", "doi"], ascending=[False, True]).reset_index(drop=True)
        
        rows = []
        for i, row in df.iterrows():
            status = str(row['stock_status'])
            st_class = "pareto-status stockout" if status.lower() == "stockout" else "pareto-status text-slate-300"
            rows.append(html.Tr([
                html.Td(i + 1, className="text-center w-[7%]"),
                html.Td(str(row['model']), className="w-[35%]"),
                html.Td(html.Span(status, className=st_class), className="w-[16%]"),
                html.Td(f"{int(row['remaining_inventory']):,}", className="text-right w-[14%]"),
                html.Td(f"{int(row['suggested_transfer']):,}", className="text-right w-[14%]"),
                html.Td(f"{int(row['doi']):,}", className="text-right w-[14%]")
            ]))
            
        return html.Div(className="mb-8 border border-slate-700/50 rounded-xl p-4 bg-slate-900/20", children=[
            html.Div(className="flex justify-between items-center mb-3 pb-2 border-b", style={"borderColor": color}, children=[
                html.Span(p_class.upper(), style={"color": color}, className="text-lg font-black tracking-widest"),
                html.Span(f"● {count} Items", style={"color": color, "borderColor": color}, className="text-xs font-bold border px-3 py-1 rounded-full")
            ]),
            html.Div(className="pareto-html-shell", children=[
                html.Table(className="pareto-html-table", children=[
                    html.Thead(html.Tr([
                        html.Th("Rank", className="text-center"), html.Th("Model"), html.Th("Status"),
                        html.Th("Inventory", className="text-right"), html.Th("Transfer", className="text-right"), html.Th("DOI", className="text-right")
                    ])),
                    html.Tbody(rows)
                ])
            ])
        ])

    return [make_table(c, col) for c, col in [("Class A", "#f87171"), ("Class B", "#fbbf24"), ("Class C", "#4ade80")]]

if __name__ == "__main__":
    app.run_server(debug=True, port=8050)
