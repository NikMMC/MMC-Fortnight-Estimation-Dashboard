import base64
import io
import pandas as pd
from dash import Dash, dcc, html
from dash.dependencies import Input, Output, State
import plotly.express as px


# =========================================
# THEME
# =========================================
COLORS = {
    "bg": "#0B1120",
    "panel": "#111827",
    "border": "#1F2937",
    "text": "#E5E7EB",
    "muted": "#9CA3AF",
    "accent": "#F59E0B",
}

BAR_COLOR = "#F59E0B"


# =========================================
# DATA PARSER
# =========================================
def build_data_from_excel(excel_bytes):
    xls = pd.ExcelFile(io.BytesIO(excel_bytes))

    rfq = pd.read_excel(xls, "RFQ_Clean")
    submitted = pd.read_excel(xls, "Submitted_Clean")
    orders = pd.read_excel(xls, "Orders_Clean")
    todo = pd.read_excel(xls, "Quotes_To_Do_Clean")

    submitted["quote_value"] = pd.to_numeric(submitted["quote_value"], errors="coerce")
    orders["order_value"] = pd.to_numeric(orders["order_value"], errors="coerce")

    def agg(df, val=None):
        if val:
            return df.groupby("job_no").agg(
                job_description=("job_description", "first"),
                project=("project", "first"),
                **{val: (val, "sum")}
            ).reset_index()
        return df.groupby("job_no").agg(
            job_description=("job_description", "first"),
            project=("project", "first"),
        ).reset_index()

    a_rfq = agg(rfq)
    a_sub = agg(submitted, "quote_value")
    a_ord = agg(orders, "order_value")
    a_todo = agg(todo)

    all_jobs = sorted(
        set(a_rfq.job_no) |
        set(a_sub.job_no) |
        set(a_ord.job_no) |
        set(a_todo.job_no)
    )

    jobs = pd.DataFrame({"job_no": all_jobs})
    jobs = jobs.merge(a_sub, on="job_no", how="left")
    jobs = jobs.merge(a_ord, on="job_no", how="left")

    def pick(r, df, col):
        hit = df[df.job_no == r.job_no]
        return hit.iloc[0][col] if not hit.empty else None

    jobs["job_description"] = jobs.apply(
        lambda r:
            pick(r, a_sub, "job_description") or
            pick(r, a_ord, "job_description") or
            pick(r, a_rfq, "job_description") or
            pick(r, a_todo, "job_description"),
        axis=1)

    jobs["project"] = jobs.apply(
        lambda r:
            pick(r, a_sub, "project") or
            pick(r, a_ord, "project") or
            pick(r, a_rfq, "project") or
            pick(r, a_todo, "project"),
        axis=1)

    jobs["has_rfq"] = jobs.job_no.isin(a_rfq.job_no)
    jobs["has_quote"] = jobs.job_no.isin(a_sub.job_no)
    jobs["has_order"] = jobs.job_no.isin(a_ord.job_no)
    jobs["has_todo"] = jobs.job_no.isin(a_todo.job_no)

    def status(r):
        if r.has_order: return "Order Received"
        if r.has_quote: return "Submitted Only"
        if r.has_todo: return "Quote To-Do"
        if r.has_rfq:  return "RFQ Only"
        return "Unclassified"

    jobs["status"] = jobs.apply(status, axis=1)

    q90 = jobs.quote_value.quantile(0.9)

    def risk(r):
        if pd.isna(r.quote_value): return "Low"
        if r.has_order: return "Low"
        if r.quote_value >= q90: return "High"
        return "Medium"

    jobs["risk_level"] = jobs.apply(risk, axis=1)

    kpi = {
        "rfq_count": rfq.job_no.nunique(),
        "submitted_count": submitted.job_no.nunique(),
        "order_count": orders.job_no.nunique(),
        "total_quote_value": submitted.quote_value.sum(),
        "total_order_value": orders.order_value.sum(),
    }

    if kpi["total_quote_value"] > 0:
        kpi["hit_rate_value"] = kpi["total_order_value"] / kpi["total_quote_value"]
    else:
        kpi["hit_rate_value"] = None

    return rfq, submitted, orders, todo, jobs, kpi


# =========================================
# KPI CARD
# =========================================
def make_kpi_card(title, value, fmt=None):
    if value is None:
        disp = "-"
    else:
        if fmt == "currency": disp = f"${value:,.0f}"
        elif fmt == "percent": disp = f"{value*100:0.1f}%"
        else: disp = f"{value:,}"

    return html.Div(
        [
            html.Div(title, style={"color": COLORS["muted"], "fontSize": "12px"}),
            html.Div(disp, style={"color": COLORS["text"], "fontSize": "22px", "fontWeight": "600"}),
        ],
        style={
            "padding": "8px",
            "backgroundColor": COLORS["panel"],
            "borderRadius": "8px",
            "border": f"1px solid {COLORS['border']}",
            "minWidth": "130px",
            "marginRight": "8px",
        }
    )


# =========================================
# RISK TABLE (ONLY INTERNAL SCROLL)
# =========================================
def risk_style(level):
    if level == "High": return {"backgroundColor": "#7F1D1D", "color": "#fff"}
    if level == "Medium": return {"backgroundColor": "#78350F", "color": "#fff"}
    if level == "Low": return {"backgroundColor": "#064E3B", "color": "#fff"}
    return {"color": "#fff"}


def make_risk_table(df):
    if df.empty:
        return html.Div("No data", style={"color": COLORS["muted"]})

    cols = ["job_no", "project", "quote_value", "order_value", "status", "risk_level"]

    header = html.Thead(
        html.Tr([
            html.Th(c.title(), style={
                "padding": "4px",
                "backgroundColor": "#020617",
                "color": COLORS["muted"],
                "position": "sticky",
                "top": 0
            }) for c in cols
        ])
    )

    rows = []
    for _, r in df.iterrows():
        row_cells = []
        for c in cols:
            v = r[c]
            if c in ["quote_value", "order_value"]:
                v = "-" if pd.isna(v) else f"${v:,.0f}"

            style = {
                "padding": "4px",
                "color": COLORS["text"],
                "borderBottom": f"1px solid {COLORS['border']}"
            }
            if c == "risk_level":
                style.update(risk_style(r[c]))

            row_cells.append(html.Td(v, style=style))

        rows.append(html.Tr(row_cells))

    return html.Div(
        html.Table([header, html.Tbody(rows)], style={"borderCollapse": "collapse"}),
        style={
            "height": "220px",
            "overflowY": "auto",
            "border": f"1px solid {COLORS['border']}",
            "backgroundColor": COLORS["panel"],
        }
    )


# =========================================
# DASH APP
# =========================================
app = Dash(__name__)
server = app.server

app.layout = html.Div(
    style={
        "height": "100vh",
        "overflow": "hidden",
        "display": "flex",
        "flexDirection": "column",
        "backgroundColor": COLORS["bg"],
    },
    children=[

        dcc.Store(id="store-jobs"),
        dcc.Store(id="store-kpi"),

        # -----------------------------
        # HEADER
        # -----------------------------
        html.Div(
            style={
                "height": "56px",
                "flexShrink": 0,
                "backgroundColor": COLORS["panel"],
                "borderBottom": f"1px solid {COLORS['border']}",
                "padding": "8px 16px",
                "display": "flex",
                "justifyContent": "space-between",
                "alignItems": "center",
            },
            children=[
                html.Div([
                    html.H2("MMC Fortnight Estimation Dashboard",
                            style={"color": COLORS["text"], "margin": 0, "fontSize": "20px"}),
                    html.Div("RFQ → Quote → Order Summary",
                             style={"color": COLORS["muted"], "fontSize": "11px"})
                ]),
                dcc.Upload(
                    id="upload-data",
                    children=html.Div([
                        "Drag or ",
                        html.Span("Select File", style={"color": COLORS["accent"]})
                    ]),
                    style={
                        "width": "180px",
                        "height": "32px",
                        "lineHeight": "32px",
                        "border": f"1px dashed {COLORS['border']}",
                        "borderRadius": "6px",
                        "textAlign": "center",
                        "color": COLORS["muted"],
                        "fontSize": "12px"
                    }
                ),
            ]
        ),

        # -----------------------------
        # BODY
        # -----------------------------
        html.Div(
            style={
                "flex": 1,
                "display": "flex",
                "overflow": "hidden",
                "minHeight": 0,
            },
            children=[

                # SIDEBAR — fully static height
                html.Div(
                    style={
                        "width": "210px",
                        "backgroundColor": "#020617",
                        "padding": "8px",
                        "borderRight": f"1px solid {COLORS['border']}",
                        "overflow": "hidden",
                        "flexShrink": 0,
                    },
                    children=[
                        html.H4("Filters", style={"color": COLORS["text"], "marginBottom": "6px"}),
                        html.Label("Status", style={"color": COLORS["muted"], "fontSize": "11px"}),
                        dcc.Dropdown(id="filter-status", multi=True,
                                     style={"marginBottom": "8px", "fontSize": "12px"}),

                        html.Label("Risk Level", style={"color": COLORS["muted"], "fontSize": "11px"}),
                        dcc.Dropdown(
                            id="filter-risk",
                            multi=True,
                            options=[{"label": x, "value": x} for x in ["High", "Medium", "Low"]],
                            style={"marginBottom": "8px", "fontSize": "12px"},
                        ),

                        html.Div(id="upload-status",
                                 style={"color": COLORS["muted"], "fontSize": "11px",
                                        "marginTop": "6px"}),
                    ]
                ),

                # MAIN PANEL (no scroll)
                html.Div(
                    style={
                        "flex": 1,
                        "padding": "6px",
                        "display": "flex",
                        "flexDirection": "column",
                        "overflow": "hidden",
                        "minHeight": 0
                    },
                    children=[

                        # KPI ROW
                        html.Div(
                            id="kpi-row",
                            style={
                                "display": "flex",
                                "flexShrink": 0,
                                "height": "48px",
                                "marginBottom": "4px"
                            }
                        ),

                        # 3 CHARTS ROW
                        html.Div(
                            style={
                                "display": "flex",
                                "flex": 1,
                                "gap": "6px",
                                "minHeight": 0,
                            },
                            children=[
                                html.Div(dcc.Graph(id="chart-quotes"), style={"flex": 1}),
                                html.Div(dcc.Graph(id="chart-orders"), style={"flex": 1}),
                                html.Div(dcc.Graph(id="chart-status"), style={"flex": 1}),
                            ]
                        ),

                        # RISK TABLE
                        html.Div(
                            style={"marginTop": "4px", "flexShrink": 0},
                            children=[
                                html.H4("Job Risk & Status", style={"color": COLORS["text"], "margin": 0}),
                                html.Div(id="risk-table")
                            ]
                        )
                    ]
                ),
            ]
        )
    ]
)


# =========================================
# FILE UPLOAD CALLBACK
# =========================================
@app.callback(
    [
        Output("store-jobs", "data"),
        Output("store-kpi", "data"),
        Output("upload-status", "children"),
    ],
    Input("upload-data", "contents"),
    State("upload-data", "filename")
)
def load_file(contents, filename):
    if contents is None:
        return None, None, "Upload Excel to begin"

    try:
        decoded = base64.b64decode(contents.split(",")[1])
        _, submitted, _, _, jobs, kpi = build_data_from_excel(decoded)
    except Exception as e:
        return None, None, f"Error: {e}"

    return jobs.to_json(orient="split"), kpi, f"Loaded {filename}"


# =========================================
# MAIN DASHBOARD REFRESH
# =========================================
@app.callback(
    [
        Output("kpi-row", "children"),
        Output("chart-quotes", "figure"),
        Output("chart-orders", "figure"),
        Output("chart-status", "figure"),
        Output("risk-table", "children"),
        Output("filter-status", "options"),
    ],
    [
        Input("store-jobs", "data"),
        Input("store-kpi", "data"),
        Input("filter-status", "value"),
        Input("filter-risk", "value"),
    ]
)
def update(jobs_json, kpi, f_status, f_risk):

    import plotly.graph_objs as go

    # If no data yet → empty figures
    if jobs_json is None:
        empty = go.Figure()
        empty.update_layout(
            paper_bgcolor=COLORS["panel"],
            plot_bgcolor=COLORS["panel"],
            font_color=COLORS["text"],
            height=170
        )
        return [], empty, empty, empty, html.Div(), []

    jobs = pd.read_json(jobs_json, orient="split")

    # KPI CARDS
    kpi_cards = [
        make_kpi_card("RFQs", kpi["rfq_count"]),
        make_kpi_card("Quotes", kpi["submitted_count"]),
        make_kpi_card("Orders", kpi["order_count"]),
        make_kpi_card("Quoted Value", kpi["total_quote_value"], "currency"),
        make_kpi_card("Order Value", kpi["total_order_value"], "currency"),
        make_kpi_card("Hit Rate", kpi["hit_rate_value"], "percent"),
    ]

    # FILTER DATA
    df = jobs.copy()
    if f_status:
        df = df[df.status.isin(f_status)]
    if f_risk:
        df = df[df.risk_level.isin(f_risk)]

    # CHART 1: QUOTES
    fig_q = px.bar(
        df.dropna(subset=["quote_value"]),
        x="job_no", y="quote_value",
        title="Quoted Value",
        color_discrete_sequence=[BAR_COLOR]
    )
    fig_q.update_yaxes(tickformat="$,.0f")
    fig_q.update_layout(
        template="plotly_dark",
        paper_bgcolor=COLORS["panel"],
        plot_bgcolor=COLORS["panel"],
        font_color=COLORS["text"],
        height=170
    )

    # CHART 2: ORDERS
    fig_o = px.bar(
        df.dropna(subset=["order_value"]),
        x="job_no", y="order_value",
        title="Order Value",
        color_discrete_sequence=[BAR_COLOR]
    )
    fig_o.update_yaxes(tickformat="$,.0f")
    fig_o.update_layout(
        template="plotly_dark",
        paper_bgcolor=COLORS["panel"],
        plot_bgcolor=COLORS["panel"],
        font_color=COLORS["text"],
        height=170
    )

    # CHART 3: STATUS
    status_counts = df.status.value_counts().reset_index()
    status_counts.columns = ["status", "count"]
    fig_s = px.bar(
        status_counts,
        x="status", y="count",
        title="Status Distribution",
        color_discrete_sequence=[BAR_COLOR]
    )
    fig_s.update_layout(
        template="plotly_dark",
        paper_bgcolor=COLORS["panel"],
        plot_bgcolor=COLORS["panel"],
        font_color=COLORS["text"],
        height=170
    )

    # RISK TABLE
    table = make_risk_table(df)

    # STATUS FILTER OPTIONS
    status_options = [
        {"label": s, "value": s}
        for s in sorted(jobs.status.unique())
    ]

    return kpi_cards, fig_q, fig_o, fig_s, table, status_options


# =========================================
# RUN
# =========================================
if __name__ == "__main__":
    app.run(debug=True)
