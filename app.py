import base64
import io

import pandas as pd
from dash import Dash, dcc, html, dash_table
from dash.dependencies import Input, Output, State
import plotly.express as px


# ---------- Helper functions ----------

def build_data_from_excel(excel_bytes: bytes):
    """Read the MMC_Fortnight_PowerBI_Data.xlsx structure and build clean dataframes + KPIs."""
    xls = pd.ExcelFile(io.BytesIO(excel_bytes))

    rfq = pd.read_excel(xls, "RFQ_Clean")
    submitted = pd.read_excel(xls, "Submitted_Clean")
    orders = pd.read_excel(xls, "Orders_Clean")
    todo = pd.read_excel(xls, "Quotes_To_Do_Clean")
    # Jobs_Analytics sheet is available but we recompute to be safe

    # Ensure numeric types
    if "quote_value" in submitted.columns:
        submitted["quote_value"] = pd.to_numeric(submitted["quote_value"], errors="coerce")
    else:
        submitted["quote_value"] = pd.NA

    if "order_value" in orders.columns:
        orders["order_value"] = pd.to_numeric(orders["order_value"], errors="coerce")
    else:
        orders["order_value"] = pd.NA

    # --- Aggregate by job ---
    def aggregate_by_job(df, value_col=None):
        if value_col and value_col in df.columns:
            agg = df.groupby("job_no").agg(
                job_description=("job_description", "first"),
                project=("project", "first"),
                **{value_col: (value_col, "sum")},
                count=("job_description", "count"),
            ).reset_index()
        else:
            agg = df.groupby("job_no").agg(
                job_description=("job_description", "first"),
                project=("project", "first"),
                count=("job_description", "count"),
            ).reset_index()
        return agg

    rfq_jobs = aggregate_by_job(rfq)
    sub_jobs = aggregate_by_job(submitted, "quote_value")
    ord_jobs = aggregate_by_job(orders, "order_value")
    todo_jobs = aggregate_by_job(todo)

    all_job_nos = sorted(
        set(rfq_jobs["job_no"]) |
        set(sub_jobs["job_no"]) |
        set(ord_jobs["job_no"]) |
        set(todo_jobs["job_no"])
    )
    jobs = pd.DataFrame({"job_no": all_job_nos})

    # Merge numeric fields
    jobs = jobs.merge(sub_jobs[["job_no", "quote_value"]], on="job_no", how="left")
    jobs = jobs.merge(ord_jobs[["job_no", "order_value"]], on="job_no", how="left")

    # Get a primary description / project
    def pick_field(row, table, field):
        match = table[table["job_no"] == row["job_no"]]
        if not match.empty:
            return match.iloc[0][field]
        return None

    jobs["job_description"] = jobs.apply(
        lambda r: pick_field(r, sub_jobs, "job_description")
        or pick_field(r, ord_jobs, "job_description")
        or pick_field(r, rfq_jobs, "job_description")
        or pick_field(r, todo_jobs, "job_description"),
        axis=1,
    )
    jobs["project"] = jobs.apply(
        lambda r: pick_field(r, sub_jobs, "project")
        or pick_field(r, ord_jobs, "project")
        or pick_field(r, rfq_jobs, "project")
        or pick_field(r, todo_jobs, "project"),
        axis=1,
    )

    # Flags / status
    jobs["has_rfq"] = jobs["job_no"].isin(rfq_jobs["job_no"])
    jobs["has_quote"] = jobs["job_no"].isin(sub_jobs["job_no"])
    jobs["has_order"] = jobs["job_no"].isin(ord_jobs["job_no"])
    jobs["has_todo"] = jobs["job_no"].isin(todo_jobs["job_no"])

    def derive_status(row):
        if row["has_order"]:
            return "Order Received"
        if row["has_quote"]:
            return "Submitted Only"
        if row["has_todo"]:
            return "Quote To-Do"
        if row["has_rfq"]:
            return "RFQ Only"
        return "Unclassified"

    jobs["status"] = jobs.apply(derive_status, axis=1)

    # Risk based on quote_value and order absence
    q90 = jobs["quote_value"].quantile(0.9) if not jobs["quote_value"].dropna().empty else None

    def risk_level(row):
        if pd.isna(row["quote_value"]) or row["quote_value"] <= 0:
            return "Low"
        if row["has_order"]:
            return "Low"
        if q90 and row["quote_value"] >= q90:
            return "High"
        return "Medium"

    jobs["risk_level"] = jobs.apply(risk_level, axis=1)

    # --- KPIs ---
    kpi = {}
    kpi["rfq_count"] = rfq["job_no"].nunique()
    kpi["submitted_count"] = submitted["job_no"].nunique()
    kpi["order_count"] = orders["job_no"].nunique()
    kpi["todo_count"] = todo["job_no"].nunique()

    kpi["total_quote_value"] = submitted["quote_value"].sum(skipna=True)
    kpi["total_order_value"] = orders["order_value"].sum(skipna=True)

    kpi["hit_rate_count"] = (
        kpi["order_count"] / kpi["submitted_count"]
        if kpi["submitted_count"] else None
    )
    kpi["hit_rate_value"] = (
        kpi["total_order_value"] / kpi["total_quote_value"]
        if kpi["total_quote_value"] else None
    )

    return rfq, submitted, orders, todo, jobs, kpi


def make_kpi_card(title, value, fmt=None, id=None):
    """Return a simple KPI card as html.Div."""
    if value is None:
        disp = "-"
    else:
        if fmt == "currency":
            disp = f"${value:,.0f}"
        elif fmt == "percent":
            disp = f"{value*100:0.1f}%"
        else:
            disp = f"{value:,}"

    return html.Div(
        id=id,
        className="kpi-card",
        children=[
            html.Div(title, className="kpi-title"),
            html.Div(disp, className="kpi-value"),
        ],
        style={
            "backgroundColor": "white",
            "padding": "12px 16px",
            "margin": "8px",
            "borderRadius": "6px",
            "border": "1px solid #E0E0E0",
            "boxShadow": "0 1px 2px rgba(0,0,0,0.05)",
            "minWidth": "150px",
            "flex": "1",
        },
    )


# ---------- Dash app ----------

app = Dash(__name__)
app.title = "MMC Fortnight Estimation Dashboard"

app.layout = html.Div(
    style={"backgroundColor": "#F9FAFB", "minHeight": "100vh", "padding": "16px"},
    children=[
        html.H2("MMC Fortnight Estimation Dashboard", style={"marginBottom": "8px"}),
        html.Div(
            "Upload the MMC_Fortnight_PowerBI_Data.xlsx file to refresh the dashboard.",
            style={"color": "#555", "marginBottom": "16px"},
        ),

        dcc.Upload(
            id="upload-data",
            children=html.Div([
                "Drag and Drop or ",
                html.A("Select Excel File", style={"color": "#0067B1", "fontWeight": "bold"})
            ]),
            style={
                "width": "100%",
                "height": "60px",
                "lineHeight": "60px",
                "borderWidth": "1px",
                "borderStyle": "dashed",
                "borderRadius": "6px",
                "borderColor": "#A0AEC0",
                "textAlign": "center",
                "backgroundColor": "white",
                "marginBottom": "16px",
            },
            multiple=False,
        ),
        html.Div(id="upload-status", style={"marginBottom": "16px", "color": "#555"}),

        # Hidden stores for data
        dcc.Store(id="store-rfq"),
        dcc.Store(id="store-submitted"),
        dcc.Store(id="store-orders"),
        dcc.Store(id="store-todo"),
        dcc.Store(id="store-jobs"),
        dcc.Store(id="store-kpi"),

        # KPI row
        html.Div(id="kpi-row", style={"display": "flex", "flexWrap": "wrap"}),

        # Middle charts
        html.Div(
            style={"display": "flex", "flexWrap": "wrap", "marginTop": "16px"},
            children=[
                html.Div(
                    dcc.Graph(id="chart-quotes-by-job"),
                    style={"flex": "1", "minWidth": "300px", "paddingRight": "8px"},
                ),
                html.Div(
                    dcc.Graph(id="chart-orders-by-job"),
                    style={"flex": "1", "minWidth": "300px", "padding": "0 4px"},
                ),
                html.Div(
                    dcc.Graph(id="chart-status-dist"),
                    style={"flex": "1", "minWidth": "300px", "paddingLeft": "8px"},
                ),
            ],
        ),

        # Bottom area: risk table + filters
        html.Div(
            style={"display": "flex", "flexWrap": "wrap", "marginTop": "16px"},
            children=[
                html.Div(
                    children=[
                        html.H4("Job Risk & Status", style={"marginBottom": "8px"}),
                        dash_table.DataTable(
                            id="risk-table",
                            columns=[],
                            data=[],
                            page_size=10,
                            style_table={"overflowX": "auto"},
                            style_cell={
                                "fontFamily": "Segoe UI",
                                "fontSize": "12px",
                                "padding": "4px 6px",
                                "border": "1px solid #E2E8F0",
                                "textAlign": "left",
                                "minWidth": "80px",
                            },
                            style_header={
                                "backgroundColor": "#EDF2F7",
                                "fontWeight": "bold",
                            },
                            style_data_conditional=[
                                {
                                    "if": {
                                        "filter_query": "{risk_level} = 'High'",
                                        "column_id": "risk_level",
                                    },
                                    "backgroundColor": "#F8D7DA",
                                    "color": "#721C24",
                                },
                                {
                                    "if": {
                                        "filter_query": "{risk_level} = 'Medium'",
                                        "column_id": "risk_level",
                                    },
                                    "backgroundColor": "#FFF3CD",
                                    "color": "#856404",
                                },
                                {
                                    "if": {
                                        "filter_query": "{risk_level} = 'Low'",
                                        "column_id": "risk_level",
                                    },
                                    "backgroundColor": "#D4EDDA",
                                    "color": "#155724",
                                },
                            ],
                        ),
                    ],
                    style={"flex": "3", "minWidth": "400px", "paddingRight": "8px"},
                ),
                html.Div(
                    children=[
                        html.H4("Filters", style={"marginBottom": "8px"}),
                        html.Label("Status"),
                        dcc.Dropdown(
                            id="filter-status",
                            options=[],
                            multi=True,
                            placeholder="All statuses",
                            style={"marginBottom": "12px"},
                        ),
                        html.Label("Risk level"),
                        dcc.Dropdown(
                            id="filter-risk",
                            options=[
                                {"label": "High", "value": "High"},
                                {"label": "Medium", "value": "Medium"},
                                {"label": "Low", "value": "Low"},
                            ],
                            multi=True,
                            placeholder="All risk levels",
                        ),
                    ],
                    style={"flex": "1", "minWidth": "220px", "paddingLeft": "8px"},
                ),
            ],
        ),
    ],
)


# ---------- Callbacks ----------

@app.callback(
    [
        Output("store-rfq", "data"),
        Output("store-submitted", "data"),
        Output("store-orders", "data"),
        Output("store-todo", "data"),
        Output("store-jobs", "data"),
        Output("store-kpi", "data"),
        Output("upload-status", "children"),
    ],
    Input("upload-data", "contents"),
    State("upload-data", "filename"),
)
def handle_upload(contents, filename):
    # Nothing uploaded yet
    if contents is None:
        return [None, None, None, None, None, None, "No file uploaded yet."]

    print(">>> Upload event received:", filename)

    try:
        content_type, content_string = contents.split(",")
    except Exception as e:
        msg = f"Upload error: could not parse contents ({e})"
        print(">>>", msg)
        return [None, None, None, None, None, None, msg]

    try:
        decoded = base64.b64decode(content_string)
    except Exception as e:
        msg = f"Upload error: could not decode file ({e})"
        print(">>>", msg)
        return [None, None, None, None, None, None, msg]

    try:
        rfq, submitted, orders, todo, jobs, kpi = build_data_from_excel(decoded)

        rfq_json = rfq.to_json(date_format="iso", orient="split")
        submitted_json = submitted.to_json(date_format="iso", orient="split")
        orders_json = orders.to_json(date_format="iso", orient="split")
        todo_json = todo.to_json(date_format="iso", orient="split")
        jobs_json = jobs.to_json(date_format="iso", orient="split")

        status_msg = (
            f"Loaded file: {filename} "
            f"• RFQs: {kpi.get('rfq_count', 0)} "
            f"• Quotes: {kpi.get('submitted_count', 0)} "
            f"• Orders: {kpi.get('order_count', 0)}"
        )
        print(">>>", status_msg)

        return (
            rfq_json,
            submitted_json,
            orders_json,
            todo_json,
            jobs_json,
            kpi,
            status_msg,
        )
    except Exception as e:
        msg = f"Error reading Excel file: {e}"
        print(">>>", msg)
        return [None, None, None, None, None, None, msg]


@app.callback(
    [
        Output("kpi-row", "children"),
        Output("chart-quotes-by-job", "figure"),
        Output("chart-orders-by-job", "figure"),
        Output("chart-status-dist", "figure"),
        Output("risk-table", "data"),
        Output("risk-table", "columns"),
        Output("filter-status", "options"),
    ],
    [
        Input("store-jobs", "data"),
        Input("store-kpi", "data"),
        Input("filter-status", "value"),
        Input("filter-risk", "value"),
    ],
)
def update_dashboard(jobs_json, kpi, status_filter, risk_filter):
    import plotly.graph_objs as go

    # If no data uploaded yet
    if jobs_json is None or kpi is None:
        empty_fig = go.Figure()
        empty_fig.update_layout(
            xaxis={"visible": False},
            yaxis={"visible": False},
            annotations=[
                {
                    "text": "Upload an Excel file to see data",
                    "xref": "paper",
                    "yref": "paper",
                    "showarrow": False,
                    "font": {"size": 14},
                }
            ],
        )
        return ([], empty_fig, empty_fig, empty_fig, [], [], [])

    jobs = pd.read_json(jobs_json, orient="split")

    # KPI cards
    kpi_cards = [
        make_kpi_card("RFQs (jobs)", kpi.get("rfq_count"), id="kpi-rfqs"),
        make_kpi_card("Quotes submitted (jobs)", kpi.get("submitted_count"), id="kpi-quotes"),
        make_kpi_card("Orders received (jobs)", kpi.get("order_count"), id="kpi-orders"),
        make_kpi_card("Total quoted value", kpi.get("total_quote_value"), fmt="currency", id="kpi-quote-val"),
        make_kpi_card("Total order value", kpi.get("total_order_value"), fmt="currency", id="kpi-order-val"),
        make_kpi_card("Hit rate (value)", kpi.get("hit_rate_value"), fmt="percent", id="kpi-hit-rate"),
    ]

    # Filter jobs
    filtered = jobs.copy()
    if status_filter:
        filtered = filtered[filtered["status"].isin(status_filter)]
    if risk_filter:
        filtered = filtered[filtered["risk_level"].isin(risk_filter)]

    # Quotes by job
    if "quote_value" in filtered.columns:
        df_quotes = filtered.dropna(subset=["quote_value"])
        fig_quotes = px.bar(
            df_quotes,
            x="job_no",
            y="quote_value",
            title="Quoted value by job",
        )
    else:
        fig_quotes = go.Figure()

    # Orders by job
    if "order_value" in filtered.columns:
        df_orders = filtered.dropna(subset=["order_value"])
        fig_orders = px.bar(
            df_orders,
            x="job_no",
            y="order_value",
            title="Order value by job",
        )
    else:
        fig_orders = go.Figure()

    # Status distribution
    status_counts = filtered["status"].value_counts().reset_index()
    status_counts.columns = ["status", "count"]
    fig_status = px.bar(
        status_counts,
        x="status",
        y="count",
        title="Job status distribution",
    )

    # Risk table
    table_cols = [
        "job_no",
        "job_description",
        "project",
        "quote_value",
        "order_value",
        "status",
        "risk_level",
    ]
    table_cols = [c for c in table_cols if c in filtered.columns]
    table_data = filtered[table_cols].to_dict("records")
    columns = [{"name": c.replace("_", " ").title(), "id": c} for c in table_cols]

    # Status dropdown options
    status_options = [
        {"label": s, "value": s}
        for s in sorted(jobs["status"].dropna().unique())
    ]

    return (
        kpi_cards,
        fig_quotes,
        fig_orders,
        fig_status,
        table_data,
        columns,
        status_options,
    )


if __name__ == "__main__":
    app.run(debug=True)
