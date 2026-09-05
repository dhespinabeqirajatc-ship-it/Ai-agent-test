import asyncio
import json
import os
import re
from datetime import datetime
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st

from copilot import CopilotClient
from copilot.session import PermissionHandler
from dotenv import load_dotenv

from pptx import Presentation
from pptx.util import Inches, Pt

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from workiva import (
    discover_available_quarters,
    find_quarter_source_smart,
    get_sheet_data,
    workiva_sheet_to_dataframe,
)


load_dotenv()


# ============================================================
# PAGE
# ============================================================

st.set_page_config(
    page_title="Workiva AI Agent",
    layout="wide",
)

APP_NAME = "Workiva AI Agent"


# ============================================================
# BASIC HELPERS
# ============================================================

def now_text():

    return (
        datetime.now()
        .astimezone()
        .strftime(
            "%Y-%m-%d %H:%M:%S %Z"
        )
    )


def friendly_error(error):

    text = str(error)

    lower = text.lower()

    if "401" in text:

        return (
            "Workiva authentication failed. "
            "Check the Workiva credentials, "
            "region and API version."
        )

    if "403" in text:

        return (
            "The Workiva connection does not "
            "have permission to read this resource."
        )

    if "404" in text:

        return (
            "The requested Workiva spreadsheet "
            "or sheet could not be found."
        )

    if "429" in text:

        return (
            "Workiva is temporarily rate-limiting "
            "requests. Wait briefly and try again."
        )

    if (
        "timeout" in lower
        or "timed out" in lower
    ):

        return (
            "The Workiva request took too long. "
            "Try again or use Refresh Workiva."
        )

    if "github_token" in lower:

        return (
            "The AI connection is not configured. "
            "Check GITHUB_TOKEN in .env or "
            "Streamlit Secrets."
        )

    return (
        "The operation could not be completed. "
        f"Technical detail: {text}"
    )


# ============================================================
# SESSION STATE
# ============================================================

DEFAULTS = {

    "app_open":
        False,

    "agent_plan":
        None,

    "agent_request":
        (
            "Build the best management "
            "dashboard for this dataset."
        ),

    "chat_messages":
        [],

    "comparison_result":
        None,

    "audit_log":
        [],

    "last_refreshed":
        None,

    "current_source_key":
        None,

    "pdf_bytes":
        None,

    "pptx_bytes":
        None,
}


for key, value in DEFAULTS.items():

    if key not in st.session_state:

        st.session_state[
            key
        ] = value


def log_event(
    action,
    detail="",
):

    st.session_state[
        "audit_log"
    ].append(
        {
            "time":
                now_text(),

            "action":
                action,

            "detail":
                detail,
        }
    )


# ============================================================
# DATA CLEANING
# ============================================================

def clean_dataframe(data):

    data = (
        data.copy()
        .replace(
            "",
            pd.NA,
        )
    )

    data = (
        data
        .dropna(
            axis=0,
            how="all",
        )
        .dropna(
            axis=1,
            how="all",
        )
    )

    return data


# ============================================================
# DATA PROFILING
# ============================================================

def profile_dataframe(data):

    data = (
        clean_dataframe(
            data
        )
    )

    profile = {

        "rows":
            len(data),

        "columns":
            len(
                data.columns
            ),

        "numeric":
            [],

        "categories":
            [],

        "dates":
            [],

        "identifiers":
            [],

        "details":
            {},
    }


    for column in data.columns:

        series = (
            data[column]
            .dropna()
        )

        if series.empty:

            continue


        name = str(
            column
        )

        lower = (
            name
            .strip()
            .lower()
        )


        looks_like_id = (

            lower == "id"

            or lower.endswith(
                "_id"
            )

            or lower.endswith(
                " id"
            )

            or lower.endswith(
                " code"
            )
        )


        numeric_version = (
            pd.to_numeric(
                series,
                errors="coerce",
            )
        )

        numeric_ratio = (
            numeric_version
            .notna()
            .mean()
        )


        date_ratio = 0


        if numeric_ratio < 0.8:

            try:

                date_version = (
                    pd.to_datetime(
                        series,
                        errors="coerce",
                    )
                )

                date_ratio = (
                    date_version
                    .notna()
                    .mean()
                )

            except Exception:

                date_ratio = 0


        if looks_like_id:

            kind = (
                "identifier"
            )

            profile[
                "identifiers"
            ].append(
                name
            )


        elif numeric_ratio >= 0.8:

            kind = (
                "numeric"
            )

            profile[
                "numeric"
            ].append(
                name
            )


        elif date_ratio >= 0.8:

            kind = (
                "date"
            )

            profile[
                "dates"
            ].append(
                name
            )


        else:

            kind = (
                "category"
            )

            profile[
                "categories"
            ].append(
                name
            )


        profile[
            "details"
        ][name] = {

            "type":
                kind,

            "unique":
                int(
                    series.nunique()
                ),

            "missing":
                int(
                    data[column]
                    .isna()
                    .sum()
                ),
        }


    return profile


def prepare_dataframe(
    data,
    profile,
):

    data = (
        clean_dataframe(
            data
        )
    )


    for column in profile[
        "numeric"
    ]:

        data[column] = (
            pd.to_numeric(
                data[column],
                errors="coerce",
            )
        )


    for column in profile[
        "dates"
    ]:

        try:

            data[column] = (
                pd.to_datetime(
                    data[column],
                    errors="coerce",
                )
            )

        except Exception:

            pass


    return data


# ============================================================
# BUSINESS COLUMN DETECTION
# ============================================================

def find_column(
    columns,
    words,
):

    normalized = [

        (
            column,
            str(column)
            .strip()
            .lower(),
        )

        for column in columns
    ]


    for column, lower in normalized:

        if lower in words:

            return column


    for column, lower in normalized:

        if any(
            word in lower
            for word in words
        ):

            return column


    return None


def detect_business_columns(data):

    columns = list(
        data.columns
    )


    return {

        "revenue":
            find_column(
                columns,
                [
                    "revenue",
                    "sales",
                    "income",
                    "actual",
                    "actuals",
                ],
            ),

        "budget":
            find_column(
                columns,
                [
                    "budget",
                    "plan",
                    "forecast",
                    "target",
                ],
            ),

        "costs":
            find_column(
                columns,
                [
                    "cost",
                    "costs",
                    "expense",
                    "expenses",
                ],
            ),

        "dimension":
            find_column(
                columns,
                [
                    "region",
                    "country",
                    "market",
                    "department",
                    "business unit",
                    "segment",
                ],
            ),
    }


# ============================================================
# RICH DATA SUMMARY
# ============================================================

def build_data_summary(
    data,
    profile,
    quarter,
):

    summary = {

        "quarter":
            quarter,

        "rows":
            len(data),

        "columns":
            list(
                data.columns
            ),

        "numeric_columns":
            profile[
                "numeric"
            ],

        "category_columns":
            profile[
                "categories"
            ],

        "date_columns":
            profile[
                "dates"
            ],

        "numeric_stats":
            {},

        "categories":
            {},

        "date_ranges":
            {},

        "missing_cells":
            int(
                data
                .isna()
                .sum()
                .sum()
            ),
    }


    for column in profile[
        "numeric"
    ][:20]:

        values = (
            pd.to_numeric(
                data[column],
                errors="coerce",
            )
            .dropna()
        )


        if values.empty:

            continue


        summary[
            "numeric_stats"
        ][column] = {

            "sum":
                float(
                    values.sum()
                ),

            "average":
                float(
                    values.mean()
                ),

            "minimum":
                float(
                    values.min()
                ),

            "maximum":
                float(
                    values.max()
                ),

            "count":
                int(
                    values.count()
                ),
        }


    for column in profile[
        "categories"
    ][:10]:

        values = (
            data[column]
            .dropna()
            .astype(str)
        )


        summary[
            "categories"
        ][column] = {

            "unique":
                int(
                    values.nunique()
                ),

            "top_values":
                (
                    values
                    .value_counts()
                    .head(10)
                    .to_dict()
                ),
        }


    for column in profile[
        "dates"
    ]:

        values = (
            data[column]
            .dropna()
        )


        if not values.empty:

            summary[
                "date_ranges"
            ][column] = {

                "minimum":
                    str(
                        values.min()
                    ),

                "maximum":
                    str(
                        values.max()
                    ),
            }


    return summary


# ============================================================
# COPILOT
# ============================================================

async def ask_copilot(prompt):

    token = (
        os.environ.get(
            "GITHUB_TOKEN"
        )
    )


    if not token:

        raise ValueError(
            "GITHUB_TOKEN is missing."
        )


    client = CopilotClient(

        github_token=token,

        use_logged_in_user=False,
    )


    await client.start()


    try:

        session = (
            await client.create_session(

                model="auto",

                on_permission_request=(
                    PermissionHandler
                    .approve_all
                ),
            )
        )


        response = (
            await session
            .send_and_wait(
                prompt
            )
        )


        if response:

            return (
                response
                .data
                .content
            )


        return ""


    finally:

        await client.stop()


def run_copilot(prompt):

    return asyncio.run(
        ask_copilot(
            prompt
        )
    )


# ============================================================
# JSON EXTRACTION
# ============================================================

def extract_json(text):

    start = (
        text.find(
            "{"
        )
    )

    end = (
        text.rfind(
            "}"
        )
    )


    if (
        start == -1
        or end == -1
        or end < start
    ):

        raise ValueError(
            "AI did not return "
            "a valid dashboard plan."
        )


    return json.loads(
        text[
            start:
            end + 1
        ]
    )


# ============================================================
# AI DASHBOARD PLANNER
# ============================================================

@st.cache_data(
    ttl=600,
    show_spinner=False,
)
def create_dashboard_plan(
    request,
    summary,
):

    prompt = f"""
You are designing an executive management dashboard.

USER REQUEST

{request}

DATASET

Columns:
{summary["columns"]}

Numeric columns:
{summary["numeric_columns"]}

Category columns:
{summary["category_columns"]}

Date columns:
{summary["date_columns"]}

Numeric statistics:
{summary["numeric_stats"]}

Date ranges:
{summary["date_ranges"]}

Return JSON only using this schema:

{{
  "title": "Dashboard title",
  "reason": "Short explanation",
  "kpis": [
    {{
      "column": "exact numeric column",
      "aggregation": "sum",
      "label": "KPI label"
    }}
  ],
  "charts": [
    {{
      "type": "bar",
      "title": "Chart title",
      "x": "exact column or null",
      "y": "exact numeric column or null",
      "aggregation": "sum"
    }}
  ]
}}

Allowed KPI aggregations:

sum
average
minimum
maximum
count

Allowed chart types:

bar
line
area
donut
scatter

Rules:

- Use only exact supplied column names.
- Never invent a column.
- Build between 2 and 4 useful KPIs.
- Build between 2 and 4 useful charts.
- Prefer management usefulness over visual variety.
- Use donut only for low-cardinality categories.
- Use line or area when a real date column exists.
- Use scatter only when appropriate numeric fields exist.
- Use bars for categorical comparisons.
- Do not calculate official financial values yourself.
"""


    return (
        extract_json(
            run_copilot(
                prompt
            )
        )
    )


def validate_plan(
    plan,
    profile,
):

    numeric = set(
        profile[
            "numeric"
        ]
    )

    categories = set(
        profile[
            "categories"
        ]
    )

    dates = set(
        profile[
            "dates"
        ]
    )

    all_columns = (
        numeric
        | categories
        | dates
    )


    allowed_aggs = {

        "sum",
        "average",
        "minimum",
        "maximum",
        "count",
    }


    allowed_charts = {

        "bar",
        "line",
        "area",
        "donut",
        "scatter",
    }


    clean = {

        "title":
            str(
                plan.get(
                    "title",
                    "AI Designed Dashboard",
                )
            ),

        "reason":
            str(
                plan.get(
                    "reason",
                    "",
                )
            ),

        "kpis":
            [],

        "charts":
            [],
    }


    for item in plan.get(
        "kpis",
        [],
    )[:4]:

        column = (
            item.get(
                "column"
            )
        )

        aggregation = (
            item.get(
                "aggregation",
                "sum",
            )
        )


        if column not in numeric:

            continue


        if aggregation not in allowed_aggs:

            aggregation = (
                "sum"
            )


        clean[
            "kpis"
        ].append(
            {

                "column":
                    column,

                "aggregation":
                    aggregation,

                "label":
                    str(
                        item.get(
                            "label",
                            column,
                        )
                    ),
            }
        )


    for item in plan.get(
        "charts",
        [],
    )[:4]:

        chart_type = (
            item.get(
                "type"
            )
        )

        x = (
            item.get(
                "x"
            )
        )

        y = (
            item.get(
                "y"
            )
        )

        aggregation = (
            item.get(
                "aggregation",
                "sum",
            )
        )


        if chart_type not in allowed_charts:

            continue


        if (
            x is not None
            and x not in all_columns
        ):

            continue


        if (
            y is not None
            and y not in numeric
        ):

            continue


        if aggregation not in allowed_aggs:

            aggregation = (
                "sum"
            )


        clean[
            "charts"
        ].append(
            {

                "type":
                    chart_type,

                "title":
                    str(
                        item.get(
                            "title",
                            chart_type.title(),
                        )
                    ),

                "x":
                    x,

                "y":
                    y,

                "aggregation":
                    aggregation,
            }
        )


    return clean


# ============================================================
# AI MANAGEMENT SUMMARY
# ============================================================

@st.cache_data(
    ttl=600,
    show_spinner=False,
)
def generate_management_summary(
    summary,
    plan,
):

    prompt = f"""
You are an executive management analyst.

Python has already calculated the dataset statistics.

Do not invent numbers.

Quarter:
{summary["quarter"]}

Dashboard:
{plan["title"]}

Dashboard reason:
{plan["reason"]}

Numeric statistics:
{summary["numeric_stats"]}

Categories:
{summary["categories"]}

Date ranges:
{summary["date_ranges"]}

Write a concise management summary.

Include:
- important observations
- notable high or low values
- useful comparisons
- possible risks
- 2 or 3 management questions

If the available statistics do not support
a conclusion, say so clearly.
"""


    return (
        run_copilot(
            prompt
        )
    )


# ============================================================
# CHAT
# ============================================================

def chat_answer(
    question,
    summary,
    plan,
):

    prompt = f"""
You are the conversational analyst inside
a Workiva management dashboard.

Quarter:
{summary["quarter"]}

Dashboard:
{plan["title"]}

Columns:
{summary["columns"]}

Numeric statistics:
{summary["numeric_stats"]}

Category information:
{summary["categories"]}

User question:
{question}

Never invent numbers.

Use only the supplied information.

Be concise and management-friendly.
"""


    return (
        run_copilot(
            prompt
        )
    )


# ============================================================
# KPI / CHART HELPERS
# ============================================================

def aggregate_value(
    series,
    method,
):

    values = (
        pd.to_numeric(
            series,
            errors="coerce",
        )
    )


    if method == "average":

        return (
            values.mean()
        )


    if method == "minimum":

        return (
            values.min()
        )


    if method == "maximum":

        return (
            values.max()
        )


    if method == "count":

        return (
            values.count()
        )


    return (
        values.sum()
    )


def grouped_data(
    data,
    category,
    metric,
    aggregation,
):

    working = (
        data[
            [
                category,
                metric,
            ]
        ].copy()
    )


    working[
        metric
    ] = (
        pd.to_numeric(
            working[
                metric
            ],
            errors="coerce",
        )
    )


    working = (
        working.dropna()
    )


    group = (
        working.groupby(
            category,
            dropna=False,
        )[metric]
    )


    if aggregation == "average":

        result = (
            group.mean()
        )


    elif aggregation == "minimum":

        result = (
            group.min()
        )


    elif aggregation == "maximum":

        result = (
            group.max()
        )


    elif aggregation == "count":

        result = (
            group.count()
        )


    else:

        result = (
            group.sum()
        )


    return (
        result
        .reset_index()
    )


def render_kpis(
    data,
    kpis,
):

    if not kpis:

        return


    columns = (
        st.columns(
            len(kpis)
        )
    )


    for index, item in enumerate(
        kpis
    ):

        value = (
            aggregate_value(
                data[
                    item[
                        "column"
                    ]
                ],
                item[
                    "aggregation"
                ],
            )
        )


        columns[
            index
        ].metric(
            item[
                "label"
            ],
            f"{value:,.2f}",
        )


def render_chart(
    data,
    chart,
):

    chart_type = (
        chart[
            "type"
        ]
    )

    x = (
        chart[
            "x"
        ]
    )

    y = (
        chart[
            "y"
        ]
    )

    title = (
        chart[
            "title"
        ]
    )

    aggregation = (
        chart[
            "aggregation"
        ]
    )


    if (
        x is None
        or y is None
        or x not in data.columns
        or y not in data.columns
    ):

        return


    if chart_type == "scatter":

        figure = (
            px.scatter(
                data,
                x=x,
                y=y,
                title=title,
            )
        )

        st.plotly_chart(
            figure,
            use_container_width=True,
        )

        return


    chart_data = (
        grouped_data(
            data,
            x,
            y,
            aggregation,
        )
    )


    if (
        pd.api.types
        .is_datetime64_any_dtype(
            data[x]
        )
    ):

        chart_data = (
            chart_data
            .sort_values(
                x
            )
        )


    else:

        chart_data = (
            chart_data
            .sort_values(
                y,
                ascending=False,
            )
        )


    if chart_type == "bar":

        figure = (
            px.bar(
                chart_data,
                x=x,
                y=y,
                title=title,
            )
        )


    elif chart_type == "line":

        figure = (
            px.line(
                chart_data,
                x=x,
                y=y,
                markers=True,
                title=title,
            )
        )


    elif chart_type == "area":

        figure = (
            px.area(
                chart_data,
                x=x,
                y=y,
                title=title,
            )
        )


    elif chart_type == "donut":

        figure = (
            px.pie(
                chart_data.head(
                    12
                ),
                names=x,
                values=y,
                hole=0.55,
                title=title,
            )
        )


    else:

        return


    st.plotly_chart(
        figure,
        use_container_width=True,
    )


# ============================================================
# DEFAULT MANUAL DASHBOARDS
# ============================================================

def category_is_small(
    profile,
    category,
):

    unique = (
        profile[
            "details"
        ]
        .get(
            category,
            {},
        )
        .get(
            "unique",
            999,
        )
    )


    return (
        2
        <= unique
        <= 12
    )


def manual_plan(
    mode,
    profile,
    business,
):

    numeric = (
        profile[
            "numeric"
        ]
    )

    categories = (
        profile[
            "categories"
        ]
    )

    dates = (
        profile[
            "dates"
        ]
    )


    metric = (

        business[
            "revenue"
        ]

        or business[
            "budget"
        ]

        or business[
            "costs"
        ]

        or (
            numeric[0]
            if numeric
            else None
        )
    )


    category = (

        business[
            "dimension"
        ]

        or (
            categories[0]
            if categories
            else None
        )
    )


    plan = {

        "title":
            f"{mode} Dashboard",

        "reason":
            (
                "Using the selected "
                "default dashboard."
            ),

        "kpis":
            [

                {
                    "column":
                        column,

                    "aggregation":
                        "sum",

                    "label":
                        column,
                }

                for column in numeric[
                    :4
                ]
            ],

        "charts":
            [],
    }


    if (
        mode == "Executive"
        and category
        and metric
    ):

        plan[
            "charts"
        ].append(
            {

                "type":
                    "bar",

                "title":
                    (
                        f"{metric} "
                        f"by {category}"
                    ),

                "x":
                    category,

                "y":
                    metric,

                "aggregation":
                    "sum",
            }
        )


        if category_is_small(
            profile,
            category,
        ):

            plan[
                "charts"
            ].append(
                {

                    "type":
                        "donut",

                    "title":
                        (
                            f"{metric} mix"
                        ),

                    "x":
                        category,

                    "y":
                        metric,

                    "aggregation":
                        "sum",
                }
            )


    elif (
        mode == "Trends"
        and dates
        and metric
    ):

        plan[
            "charts"
        ].append(
            {

                "type":
                    "line",

                "title":
                    (
                        f"{metric} trend"
                    ),

                "x":
                    dates[0],

                "y":
                    metric,

                "aggregation":
                    "sum",
            }
        )


    elif (
        mode == "Breakdown"
        and category
        and metric
    ):

        plan[
            "charts"
        ].append(
            {

                "type":
                    "bar",

                "title":
                    (
                        f"{metric} "
                        f"by {category}"
                    ),

                "x":
                    category,

                "y":
                    metric,

                "aggregation":
                    "sum",
            }
        )


        if category_is_small(
            profile,
            category,
        ):

            plan[
                "charts"
            ].append(
                {

                    "type":
                        "donut",

                    "title":
                        (
                            f"{metric} "
                            "distribution"
                        ),

                    "x":
                        category,

                    "y":
                        metric,

                    "aggregation":
                        "sum",
                }
            )


    return plan


# ============================================================
# WORKIVA CACHE
#
# Spreadsheet/sheet discovery is cached for 30 minutes.
# Actual sheet values are cached for 10 minutes.
# This improves speed while Refresh Workiva remains available.
# ============================================================

@st.cache_data(
    ttl=1800,
    show_spinner=False,
)
def cached_quarters():

    return (
        discover_available_quarters()
    )


@st.cache_data(
    ttl=1800,
    show_spinner=False,
)
def cached_source(quarter):

    return (
        find_quarter_source_smart(
            quarter
        )
    )


@st.cache_data(
    ttl=600,
    show_spinner=False,
)
def cached_data(
    spreadsheet_id,
    sheet_id,
):

    response = (
        get_sheet_data(
            spreadsheet_id,
            sheet_id,
        )
    )


    return (
        clean_dataframe(
            workiva_sheet_to_dataframe(
                response
            )
        )
    )


def quarter_bundle(quarter):

    source = (
        cached_source(
            quarter
        )
    )


    if not source:

        raise ValueError(
            f"No Workiva source "
            f"could be found for {quarter}."
        )


    raw = (
        cached_data(
            source[
                "spreadsheet_id"
            ],
            source[
                "sheet_id"
            ],
        )
    )


    if raw.empty:

        raise ValueError(
            f"The Workiva sheet for "
            f"{quarter} contains no usable data."
        )


    profile = (
        profile_dataframe(
            raw
        )
    )

    data = (
        prepare_dataframe(
            raw,
            profile,
        )
    )

    business = (
        detect_business_columns(
            data
        )
    )

    summary = (
        build_data_summary(
            data,
            profile,
            quarter,
        )
    )


    return {

        "quarter":
            quarter,

        "source":
            source,

        "data":
            data,

        "profile":
            profile,

        "business":
            business,

        "summary":
            summary,
    }


# ============================================================
# QUARTER COMPARISON
# ============================================================

def extract_quarters(text):

    matches = (
        re.findall(
            (
                r"\bQ([1-4])"
                r"[\s\-_\/]*"
                r"(20\d{2})\b"
            ),
            str(text),
            flags=re.IGNORECASE,
        )
    )


    quarters = []


    for q, year in matches:

        quarter = (
            f"Q{q} {year}"
        )


        if quarter not in quarters:

            quarters.append(
                quarter
            )


    return quarters


def normalized_column_map(
    columns,
):

    return {

        str(column)
        .strip()
        .lower():
            column

        for column in columns
    }


def common_columns(
    first,
    second,
):

    first_map = (
        normalized_column_map(
            first
        )
    )

    second_map = (
        normalized_column_map(
            second
        )
    )


    result = []


    for key, first_name in (
        first_map.items()
    ):

        if key in second_map:

            result.append(
                (
                    first_name,
                    second_map[
                        key
                    ],
                )
            )


    return result


def detect_comparison_focus(
    question,
    first_bundle,
    second_bundle,
    common_numeric,
):

    lower = (
        question.lower()
    )


    groups = [

        (
            [
                "budget",
                "plan",
                "forecast",
                "target",
            ],
            "budget",
        ),

        (
            [
                "cost",
                "costs",
                "expense",
                "expenses",
            ],
            "costs",
        ),

        (
            [
                "revenue",
                "sales",
                "income",
                "actual",
            ],
            "revenue",
        ),
    ]


    for words, key in groups:

        if any(
            word in lower
            for word in words
        ):

            first_metric = (
                first_bundle[
                    "business"
                ].get(
                    key
                )
            )

            second_metric = (
                second_bundle[
                    "business"
                ].get(
                    key
                )
            )


            if (
                first_metric
                and second_metric
            ):

                return (
                    first_metric,
                    second_metric,
                )


    for (
        first_name,
        second_name,
    ) in common_numeric:

        if (
            str(
                first_name
            ).lower()
            in lower
        ):

            return (
                first_name,
                second_name,
            )


    if common_numeric:

        return (
            common_numeric[0]
        )


    return (
        None,
        None,
    )


def compare_quarter_bundles(
    first_bundle,
    second_bundle,
    question="",
):

    common_numeric = (
        common_columns(
            first_bundle[
                "profile"
            ][
                "numeric"
            ],
            second_bundle[
                "profile"
            ][
                "numeric"
            ],
        )
    )


    common_categories = (
        common_columns(
            first_bundle[
                "profile"
            ][
                "categories"
            ],
            second_bundle[
                "profile"
            ][
                "categories"
            ],
        )
    )


    rows = []


    for (
        first_metric,
        second_metric,
    ) in common_numeric:

        first_total = (
            pd.to_numeric(
                first_bundle[
                    "data"
                ][
                    first_metric
                ],
                errors="coerce",
            )
            .sum()
        )


        second_total = (
            pd.to_numeric(
                second_bundle[
                    "data"
                ][
                    second_metric
                ],
                errors="coerce",
            )
            .sum()
        )


        change = (
            second_total
            - first_total
        )


        change_pct = (

            change
            / first_total
            * 100

            if first_total != 0

            else None
        )


        rows.append(
            {

                "Metric":
                    str(
                        first_metric
                    ),

                first_bundle[
                    "quarter"
                ]:
                    float(
                        first_total
                    ),

                second_bundle[
                    "quarter"
                ]:
                    float(
                        second_total
                    ),

                "Change":
                    float(
                        change
                    ),

                "Change %":
                    (
                        float(
                            change_pct
                        )

                        if change_pct
                        is not None

                        else None
                    ),
            }
        )


    metrics_df = (
        pd.DataFrame(
            rows
        )
    )


    (
        focus_first,
        focus_second,
    ) = (
        detect_comparison_focus(
            question,
            first_bundle,
            second_bundle,
            common_numeric,
        )
    )


    category_first = None

    category_second = None


    first_dimension = (
        first_bundle[
            "business"
        ].get(
            "dimension"
        )
    )

    second_dimension = (
        second_bundle[
            "business"
        ].get(
            "dimension"
        )
    )


    if (
        first_dimension
        and second_dimension
    ):

        if (
            str(
                first_dimension
            )
            .strip()
            .lower()

            ==

            str(
                second_dimension
            )
            .strip()
            .lower()
        ):

            category_first = (
                first_dimension
            )

            category_second = (
                second_dimension
            )


    if (
        category_first is None
        and common_categories
    ):

        (
            category_first,
            category_second,
        ) = (
            common_categories[0]
        )


    breakdown = None


    if (
        focus_first
        and focus_second
        and category_first
        and category_second
    ):

        first_group = (

            first_bundle[
                "data"
            ]

            .groupby(
                category_first,
                dropna=False,
            )[
                focus_first
            ]

            .sum()

            .reset_index()

            .rename(
                columns={
                    category_first:
                        "Category",

                    focus_first:
                        first_bundle[
                            "quarter"
                        ],
                }
            )
        )


        second_group = (

            second_bundle[
                "data"
            ]

            .groupby(
                category_second,
                dropna=False,
            )[
                focus_second
            ]

            .sum()

            .reset_index()

            .rename(
                columns={
                    category_second:
                        "Category",

                    focus_second:
                        second_bundle[
                            "quarter"
                        ],
                }
            )
        )


        breakdown = (
            first_group
            .merge(
                second_group,
                on="Category",
                how="outer",
            )
            .fillna(0)
        )


    return {

        "first_quarter":
            first_bundle[
                "quarter"
            ],

        "second_quarter":
            second_bundle[
                "quarter"
            ],

        "first_source":
            first_bundle[
                "source"
            ],

        "second_source":
            second_bundle[
                "source"
            ],

        "metrics":
            metrics_df,

        "focus_metric":
            (
                str(
                    focus_first
                )

                if focus_first

                else None
            ),

        "breakdown":
            breakdown,
    }


# ============================================================
# COMPARISON COMMENTARY
# ============================================================

@st.cache_data(
    ttl=600,
    show_spinner=False,
)
def comparison_commentary(
    comparison,
):

    metrics = (
        comparison[
            "metrics"
        ].to_dict(
            orient="records"
        )
    )


    breakdown = (

        comparison[
            "breakdown"
        ]
        .head(20)
        .to_dict(
            orient="records"
        )

        if comparison[
            "breakdown"
        ] is not None

        else []
    )


    prompt = f"""
You are an executive analyst comparing two Workiva quarters.

Python has already calculated all values.

Never invent values.

Quarter 1:
{comparison["first_quarter"]}

Quarter 2:
{comparison["second_quarter"]}

Metric comparison:
{metrics}

Focus metric:
{comparison["focus_metric"]}

Category comparison:
{breakdown}

Explain the most important movements.

Highlight material increases or decreases.

Give 2 or 3 management questions.

Be concise.
"""


    return (
        run_copilot(
            prompt
        )
    )


def render_comparison(
    comparison,
):

    st.subheader(
        (
            f"{comparison['first_quarter']} "
            f"vs "
            f"{comparison['second_quarter']}"
        )
    )


    st.caption(
        (
            f"Source 1: "
            f"{comparison['first_source']['spreadsheet_name']} "
            f"→ "
            f"{comparison['first_source']['sheet_name']} "
            f"| "
            f"Source 2: "
            f"{comparison['second_source']['spreadsheet_name']} "
            f"→ "
            f"{comparison['second_source']['sheet_name']}"
        )
    )


    metrics = (
        comparison[
            "metrics"
        ]
    )


    if metrics.empty:

        st.warning(
            (
                "The two quarters do not "
                "contain matching numeric "
                "columns that can be compared safely."
            )
        )

        return


    cards_data = (
        metrics.head(
            4
        )
    )


    cards = (
        st.columns(
            len(
                cards_data
            )
        )
    )


    for index, (
        _,
        row,
    ) in enumerate(
        cards_data.iterrows()
    ):

        delta = (
            row[
                "Change %"
            ]
        )


        delta_text = (

            f"{delta:+.1f}%"

            if pd.notna(
                delta
            )

            else "n/a"
        )


        cards[
            index
        ].metric(
            row[
                "Metric"
            ],
            (
                f"{row[comparison['second_quarter']]:,.2f}"
            ),
            delta_text,
        )


    long_metrics = (

        metrics[
            [
                "Metric",
                comparison[
                    "first_quarter"
                ],
                comparison[
                    "second_quarter"
                ],
            ]
        ]

        .melt(
            id_vars="Metric",
            var_name="Quarter",
            value_name="Value",
        )
    )


    figure = (
        px.bar(
            long_metrics,
            x="Metric",
            y="Value",
            color="Quarter",
            barmode="group",
            title=(
                "Quarter comparison "
                "by metric"
            ),
        )
    )


    st.plotly_chart(
        figure,
        use_container_width=True,
    )


    pct_data = (
        metrics
        .dropna(
            subset=[
                "Change %"
            ]
        )
        .sort_values(
            "Change %",
            ascending=False,
        )
    )


    if not pct_data.empty:

        figure = (
            px.bar(
                pct_data,
                x="Metric",
                y="Change %",
                title=(
                    "Percentage change"
                ),
            )
        )


        st.plotly_chart(
            figure,
            use_container_width=True,
        )


    if (
        comparison[
            "breakdown"
        ]
        is not None
    ):

        long_breakdown = (

            comparison[
                "breakdown"
            ]

            .melt(
                id_vars="Category",
                var_name="Quarter",
                value_name="Value",
            )
        )


        figure = (
            px.bar(
                long_breakdown,
                x="Category",
                y="Value",
                color="Quarter",
                barmode="group",
                title=(
                    f"{comparison['focus_metric']} "
                    "by category"
                ),
            )
        )


        st.plotly_chart(
            figure,
            use_container_width=True,
        )


    with st.spinner(
        "Generating comparison commentary..."
    ):

        commentary = (
            comparison_commentary(
                comparison
            )
        )


    st.markdown(
        commentary
    )


# ============================================================
# PDF EXPORT
# ============================================================

def build_pdf_export(
    quarter,
    source,
    summary,
    management_summary,
):

    buffer = (
        BytesIO()
    )


    document = (
        SimpleDocTemplate(

            buffer,

            pagesize=(
                landscape(
                    A4
                )
            ),

            rightMargin=(
                15 * mm
            ),

            leftMargin=(
                15 * mm
            ),

            topMargin=(
                15 * mm
            ),

            bottomMargin=(
                15 * mm
            ),
        )
    )


    styles = (
        getSampleStyleSheet()
    )


    story = [

        Paragraph(
            (
                f"Workiva AI Management Report "
                f"- {quarter}"
            ),
            styles[
                "Title"
            ],
        ),

        Spacer(
            1,
            8,
        ),

        Paragraph(
            (
                f"<b>Spreadsheet:</b> "
                f"{source['spreadsheet_name']}"
                f"<br/>"
                f"<b>Sheet:</b> "
                f"{source['sheet_name']}"
                f"<br/>"
                f"<b>Match confidence:</b> "
                f"{source.get('confidence', 'Unknown')}"
            ),
            styles[
                "BodyText"
            ],
        ),

        Spacer(
            1,
            12,
        ),

        Paragraph(
            "Key Metrics",
            styles[
                "Heading2"
            ],
        ),
    ]


    rows = [
        [
            "Metric",
            "Total",
            "Average",
            "Minimum",
            "Maximum",
        ]
    ]


    for metric, values in (
        summary[
            "numeric_stats"
        ].items()
    ):

        rows.append(
            [
                str(
                    metric
                ),

                f"{values['sum']:,.2f}",

                f"{values['average']:,.2f}",

                f"{values['minimum']:,.2f}",

                f"{values['maximum']:,.2f}",
            ]
        )


    if len(rows) == 1:

        rows.append(
            [
                "No numeric metrics detected",
                "",
                "",
                "",
                "",
            ]
        )


    table = (
        Table(
            rows,
            repeatRows=1,
        )
    )


    table.setStyle(
        TableStyle(
            [

                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.lightgrey,
                ),

                (
                    "FONTNAME",
                    (0, 0),
                    (-1, 0),
                    "Helvetica-Bold",
                ),

                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.grey,
                ),

                (
                    "ALIGN",
                    (1, 1),
                    (-1, -1),
                    "RIGHT",
                ),
            ]
        )
    )


    story.extend(
        [

            table,

            Spacer(
                1,
                14,
            ),

            Paragraph(
                "AI Management Summary",
                styles[
                    "Heading2"
                ],
            ),

            Paragraph(
                (
                    str(
                        management_summary
                    )
                    .replace(
                        "\n",
                        "<br/>",
                    )
                ),
                styles[
                    "BodyText"
                ],
            ),

            Spacer(
                1,
                14,
            ),

            Paragraph(
                "Data Profile",
                styles[
                    "Heading2"
                ],
            ),

            Paragraph(
                (
                    f"Rows: "
                    f"{summary['rows']}"
                    f"<br/>"
                    f"Columns: "
                    f"{len(summary['columns'])}"
                    f"<br/>"
                    f"Missing cells: "
                    f"{summary['missing_cells']}"
                    f"<br/>"
                    f"Numeric columns: "
                    f"{len(summary['numeric_columns'])}"
                    f"<br/>"
                    f"Category columns: "
                    f"{len(summary['category_columns'])}"
                    f"<br/>"
                    f"Date columns: "
                    f"{len(summary['date_columns'])}"
                ),
                styles[
                    "BodyText"
                ],
            ),
        ]
    )


    document.build(
        story
    )


    buffer.seek(0)


    return (
        buffer.getvalue()
    )


# ============================================================
# POWERPOINT EXPORT
# ============================================================

def build_powerpoint_export(
    quarter,
    source,
    summary,
    management_summary,
):

    presentation = (
        Presentation()
    )


    # --------------------------------------------------------
    # TITLE
    # --------------------------------------------------------

    slide = (
        presentation
        .slides
        .add_slide(
            presentation
            .slide_layouts[0]
        )
    )


    slide.shapes.title.text = (
        f"Workiva AI Report - {quarter}"
    )


    slide.placeholders[
        1
    ].text = (
        f"{source['spreadsheet_name']} "
        f"-> "
        f"{source['sheet_name']}"
    )


    # --------------------------------------------------------
    # KPIs
    # --------------------------------------------------------

    slide = (
        presentation
        .slides
        .add_slide(
            presentation
            .slide_layouts[5]
        )
    )


    slide.shapes.title.text = (
        "Key Performance Indicators"
    )


    metrics = list(
        summary[
            "numeric_stats"
        ].items()
    )[:6]


    for index, (
        metric,
        values,
    ) in enumerate(
        metrics
    ):

        row = (
            index // 2
        )

        column = (
            index % 2
        )


        box = (
            slide.shapes
            .add_textbox(

                Inches(
                    0.6
                    + column
                    * 4.6
                ),

                Inches(
                    1.5
                    + row
                    * 1.25
                ),

                Inches(
                    4.2
                ),

                Inches(
                    0.85
                ),
            )
        )


        paragraph = (
            box
            .text_frame
            .paragraphs[0]
        )


        paragraph.text = (
            f"{metric}\n"
            f"{values['sum']:,.2f}"
        )


        paragraph.font.size = (
            Pt(
                20
            )
        )


        paragraph.font.bold = (
            True
        )


    # --------------------------------------------------------
    # AI SUMMARY
    # --------------------------------------------------------

    slide = (
        presentation
        .slides
        .add_slide(
            presentation
            .slide_layouts[5]
        )
    )


    slide.shapes.title.text = (
        "AI Management Summary"
    )


    box = (
        slide.shapes
        .add_textbox(
            Inches(
                0.7
            ),
            Inches(
                1.5
            ),
            Inches(
                8.7
            ),
            Inches(
                5.2
            ),
        )
    )


    box.text_frame.word_wrap = (
        True
    )


    box.text_frame.text = (
        str(
            management_summary
        )
    )


    for paragraph in (
        box
        .text_frame
        .paragraphs
    ):

        paragraph.font.size = (
            Pt(
                16
            )
        )


    # --------------------------------------------------------
    # DATA PROFILE
    # --------------------------------------------------------

    slide = (
        presentation
        .slides
        .add_slide(
            presentation
            .slide_layouts[5]
        )
    )


    slide.shapes.title.text = (
        "Data Profile"
    )


    box = (
        slide.shapes
        .add_textbox(
            Inches(
                0.8
            ),
            Inches(
                1.6
            ),
            Inches(
                8.5
            ),
            Inches(
                4.5
            ),
        )
    )


    box.text_frame.text = (
        f"Rows: "
        f"{summary['rows']}\n"
        f"Columns: "
        f"{len(summary['columns'])}\n"
        f"Missing cells: "
        f"{summary['missing_cells']}\n"
        f"Numeric columns: "
        f"{len(summary['numeric_columns'])}\n"
        f"Category columns: "
        f"{len(summary['category_columns'])}\n"
        f"Date columns: "
        f"{len(summary['date_columns'])}\n"
        f"Source confidence: "
        f"{source.get('confidence', 'Unknown')}"
    )


    for paragraph in (
        box
        .text_frame
        .paragraphs
    ):

        paragraph.font.size = (
            Pt(
                18
            )
        )


    output = (
        BytesIO()
    )


    presentation.save(
        output
    )


    output.seek(0)


    return (
        output.getvalue()
    )


# ============================================================
# LANDING PAGE
# ============================================================

if not st.session_state[
    "app_open"
]:

    st.title(
        APP_NAME
    )


    st.subheader(
        (
            "Management reporting, "
            "analytics and AI commentary"
        )
    )


    st.write(
        (
            "Connect to Workiva, discover quarter data, "
            "build adaptive dashboards, compare reporting "
            "periods, and ask management questions "
            "in natural language."
        )
    )


    try:

        discovered = (
            cached_quarters()
        )

    except Exception:

        discovered = []


    latest = (

        discovered[0]

        if discovered

        else "Not yet discovered"
    )


    c1, c2, c3 = (
        st.columns(
            3
        )
    )


    c1.metric(
        "Latest discovered quarter",
        latest,
    )


    c2.metric(
        "Workiva access",
        "Read-only",
    )


    c3.metric(
        "AI mode",
        "Dashboard + analysis",
    )


    st.info(
        (
            "Workiva write-back is disabled. "
            "The app only reads data until explicit "
            "write permissions and confirmation "
            "controls are designed later."
        )
    )


    if st.button(
        "Open dashboard",
        type="primary",
        use_container_width=True,
    ):

        st.session_state[
            "app_open"
        ] = True


        log_event(
            "Open dashboard",
            "User entered the dashboard.",
        )


        st.rerun()


    st.stop()


# ============================================================
# MAIN PAGE
# ============================================================

st.title(
    APP_NAME
)


st.caption(
    (
        "Smart Workiva discovery, adaptive dashboards, "
        "quarter comparison and AI management analysis."
    )
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header(
        "Agent controls"
    )


    try:

        quarter_options = (
            cached_quarters()
        )

    except Exception:

        quarter_options = []


    if not quarter_options:

        quarter_options = [

            "Q1 2026",
            "Q4 2025",
            "Q3 2025",
            "Q2 2025",
            "Q1 2025",
        ]


    selected_quarter = (
        st.selectbox(
            "Quarter",
            quarter_options,
        )
    )


    board_mode = (
        st.selectbox(
            "Board mode",
            [
                "AI designed",
                "Executive",
                "Trends",
                "Breakdown",
                "Data quality",
            ],
        )
    )


    if st.button(
        "Refresh Workiva",
        use_container_width=True,
    ):

        st.cache_data.clear()


        st.session_state[
            "agent_plan"
        ] = None


        st.session_state[
            "comparison_result"
        ] = None


        st.session_state[
            "pdf_bytes"
        ] = None


        st.session_state[
            "pptx_bytes"
        ] = None


        st.session_state[
            "last_refreshed"
        ] = (
            now_text()
        )


        log_event(
            "Refresh Workiva",
            "Caches cleared.",
        )


        st.rerun()


    st.divider()


    st.success(
        "Read-only Workiva access"
    )


    if st.button(
        "Back to home",
        use_container_width=True,
    ):

        st.session_state[
            "app_open"
        ] = False


        st.rerun()


# ============================================================
# LOAD QUARTER
# ============================================================

try:

    with st.spinner(
        (
            f"Finding {selected_quarter} "
            "in Workiva..."
        )
    ):

        bundle = (
            quarter_bundle(
                selected_quarter
            )
        )


except Exception as error:

    st.error(
        friendly_error(
            error
        )
    )


    st.stop()


source = (
    bundle[
        "source"
    ]
)

data = (
    bundle[
        "data"
    ]
)

profile = (
    bundle[
        "profile"
    ]
)

business = (
    bundle[
        "business"
    ]
)

summary = (
    bundle[
        "summary"
    ]
)


source_key = (
    f"{source['spreadsheet_id']}"
    f":"
    f"{source['sheet_id']}"
)


if (
    st.session_state[
        "current_source_key"
    ]
    != source_key
):

    st.session_state[
        "current_source_key"
    ] = source_key


    st.session_state[
        "agent_plan"
    ] = None


    st.session_state[
        "chat_messages"
    ] = []


    st.session_state[
        "comparison_result"
    ] = None


    st.session_state[
        "pdf_bytes"
    ] = None


    st.session_state[
        "pptx_bytes"
    ] = None


    st.session_state[
        "last_refreshed"
    ] = (
        now_text()
    )


    log_event(
        "Load source",
        (
            f"{selected_quarter}: "
            f"{source['spreadsheet_name']} "
            f"→ "
            f"{source['sheet_name']}"
        ),
    )


# ============================================================
# SOURCE INFORMATION
# ============================================================

st.success(
    (
        f"Source: "
        f"{source['spreadsheet_name']} "
        f"→ "
        f"{source['sheet_name']}"
    )
)


st.caption(
    (
        f"Last refreshed: "
        f"{st.session_state['last_refreshed']}"
    )
)


confidence = (
    source.get(
        "confidence",
        "Unknown",
    )
)


if confidence == "Exact match":

    st.success(
        (
            "Match confidence: "
            "Exact Workiva match"
        )
    )


else:

    st.warning(
        (
            f"Match confidence: "
            f"{confidence}. "
            "Fallback matching was used."
        )
    )


st.caption(
    (
        f"{profile['rows']:,} rows • "
        f"{profile['columns']:,} columns • "
        f"{len(profile['numeric'])} numeric • "
        f"{len(profile['categories'])} categorical • "
        f"{len(profile['dates'])} date"
    )
)


# ============================================================
# AI DASHBOARD BUILDER
# ============================================================

st.subheader(
    "AI Dashboard Builder"
)


agent_request = (
    st.text_input(
        "Tell the agent what you want",
        value=(
            st.session_state[
                "agent_request"
            ]
        ),
        placeholder=(
            "Example: Build a CFO dashboard "
            "focused on budget and costs"
        ),
    )
)


if st.button(
    "Build AI dashboard",
    type="primary",
):

    if not agent_request.strip():

        st.warning(
            "Type a dashboard request first."
        )


    else:

        try:

            with st.spinner(
                "Designing dashboard..."
            ):

                raw_plan = (
                    create_dashboard_plan(
                        agent_request,
                        summary,
                    )
                )


                plan = (
                    validate_plan(
                        raw_plan,
                        profile,
                    )
                )


                st.session_state[
                    "agent_plan"
                ] = plan


            st.session_state[
                "agent_request"
            ] = agent_request


            st.session_state[
                "pdf_bytes"
            ] = None


            st.session_state[
                "pptx_bytes"
            ] = None


            log_event(
                "Build AI dashboard",
                agent_request,
            )


            st.rerun()


        except Exception as error:

            st.error(
                friendly_error(
                    error
                )
            )


# ============================================================
# ACTIVE PLAN
# ============================================================

if board_mode == "AI designed":

    if (
        st.session_state[
            "agent_plan"
        ]
        is None
    ):

        try:

            with st.spinner(
                (
                    "Creating the best "
                    "dashboard for this data..."
                )
            ):

                raw_plan = (
                    create_dashboard_plan(
                        (
                            "Build the best executive "
                            "management dashboard "
                            "for this dataset."
                        ),
                        summary,
                    )
                )


                st.session_state[
                    "agent_plan"
                ] = (
                    validate_plan(
                        raw_plan,
                        profile,
                    )
                )


        except Exception:

            st.session_state[
                "agent_plan"
            ] = (
                manual_plan(
                    "Executive",
                    profile,
                    business,
                )
            )


    active_plan = (
        st.session_state[
            "agent_plan"
        ]
    )


elif board_mode == "Data quality":

    active_plan = None


else:

    active_plan = (
        manual_plan(
            board_mode,
            profile,
            business,
        )
    )


summary_plan = (

    {
        "title":
            "Data Quality Dashboard",

        "reason":
            (
                "Reviewing dataset "
                "completeness."
            ),
    }

    if board_mode
    == "Data quality"

    else active_plan
)


# ============================================================
# TABS
# ============================================================

(
    overview_tab,
    financial_tab,
    trends_tab,
    comparison_tab,
    quality_tab,
    ask_tab,
) = st.tabs(
    [
        "Overview",
        "Financials",
        "Trends",
        "Comparison",
        "Data quality",
        "Ask AI",
    ]
)


# ============================================================
# OVERVIEW TAB
# ============================================================

with overview_tab:

    if board_mode == "Data quality":

        st.info(
            (
                "Data quality is selected. "
                "Open the Data quality tab."
            )
        )


    else:

        st.subheader(
            active_plan[
                "title"
            ]
        )


        if active_plan[
            "reason"
        ]:

            st.caption(
                active_plan[
                    "reason"
                ]
            )


        render_kpis(
            data,
            active_plan[
                "kpis"
            ],
        )


        charts = (
            active_plan[
                "charts"
            ]
        )


        for index in range(
            0,
            len(charts),
            2,
        ):

            row = (
                st.columns(
                    2
                )
            )


            for offset in range(
                2
            ):

                chart_index = (
                    index
                    + offset
                )


                if (
                    chart_index
                    >= len(charts)
                ):

                    break


                with row[
                    offset
                ]:

                    render_chart(
                        data,
                        charts[
                            chart_index
                        ],
                    )


    st.divider()


    st.subheader(
        "AI Management Summary"
    )


    try:

        with st.spinner(
            (
                "Generating management "
                "summary..."
            )
        ):

            management_summary = (
                generate_management_summary(
                    summary,
                    summary_plan,
                )
            )


        st.markdown(
            management_summary
        )


    except Exception as error:

        management_summary = (
            "AI summary unavailable: "
            + friendly_error(
                error
            )
        )


        st.warning(
            management_summary
        )


    # ========================================================
    # EXPORT
    # ========================================================

    st.divider()


    st.subheader(
        "Export"
    )


    csv_data = (
        data
        .to_csv(
            index=False
        )
        .encode(
            "utf-8"
        )
    )


    if st.download_button(
        "Download current quarter as CSV",
        data=csv_data,
        file_name=(
            selected_quarter
            .replace(
                " ",
                "_",
            )
            + "_workiva.csv"
        ),
        mime="text/csv",
    ):

        log_event(
            "Export CSV",
            selected_quarter,
        )


    # --------------------------------------------------------
    # LAZY EXPORTS
    #
    # The expensive files are only generated when requested.
    # This is intentionally faster than preparing them on
    # every Streamlit rerun.
    # --------------------------------------------------------

    prepare1, prepare2 = (
        st.columns(
            2
        )
    )


    with prepare1:

        if st.button(
            "Prepare PDF report",
            use_container_width=True,
        ):

            try:

                with st.spinner(
                    "Preparing PDF..."
                ):

                    st.session_state[
                        "pdf_bytes"
                    ] = (
                        build_pdf_export(
                            selected_quarter,
                            source,
                            summary,
                            management_summary,
                        )
                    )


                log_event(
                    "Prepare PDF",
                    selected_quarter,
                )


            except Exception as error:

                st.error(
                    friendly_error(
                        error
                    )
                )


    with prepare2:

        if st.button(
            "Prepare PowerPoint report",
            use_container_width=True,
        ):

            try:

                with st.spinner(
                    "Preparing PowerPoint..."
                ):

                    st.session_state[
                        "pptx_bytes"
                    ] = (
                        build_powerpoint_export(
                            selected_quarter,
                            source,
                            summary,
                            management_summary,
                        )
                    )


                log_event(
                    "Prepare PowerPoint",
                    selected_quarter,
                )


            except Exception as error:

                st.error(
                    friendly_error(
                        error
                    )
                )


    download1, download2 = (
        st.columns(
            2
        )
    )


    with download1:

        if (
            st.session_state[
                "pdf_bytes"
            ]
            is not None
        ):

            st.download_button(
                "Download PDF report",
                data=(
                    st.session_state[
                        "pdf_bytes"
                    ]
                ),
                file_name=(
                    selected_quarter
                    .replace(
                        " ",
                        "_",
                    )
                    + "_Workiva_AI_Report.pdf"
                ),
                mime="application/pdf",
                use_container_width=True,
            )


    with download2:

        if (
            st.session_state[
                "pptx_bytes"
            ]
            is not None
        ):

            st.download_button(
                "Download PowerPoint report",
                data=(
                    st.session_state[
                        "pptx_bytes"
                    ]
                ),
                file_name=(
                    selected_quarter
                    .replace(
                        " ",
                        "_",
                    )
                    + "_Workiva_AI_Report.pptx"
                ),
                mime=(
                    "application/"
                    "vnd.openxmlformats-officedocument."
                    "presentationml.presentation"
                ),
                use_container_width=True,
            )


# ============================================================
# FINANCIALS TAB
# ============================================================

with financial_tab:

    st.subheader(
        (
            f"{selected_quarter} "
            "Financial View"
        )
    )


    financial_metrics = []


    for column in [

        business[
            "revenue"
        ],

        business[
            "budget"
        ],

        business[
            "costs"
        ],
    ]:

        if (
            column
            and column
            not in financial_metrics
        ):

            financial_metrics.append(
                column
            )


    if not financial_metrics:

        st.info(
            (
                "No obvious revenue, budget "
                "or cost columns were detected."
            )
        )


    else:

        cards = (
            st.columns(
                len(
                    financial_metrics
                )
            )
        )


        for index, metric in enumerate(
            financial_metrics
        ):

            value = (
                pd.to_numeric(
                    data[
                        metric
                    ],
                    errors="coerce",
                )
                .sum()
            )


            cards[
                index
            ].metric(
                metric,
                f"{value:,.2f}",
            )


        if (
            business[
                "revenue"
            ]
            and business[
                "budget"
            ]
        ):

            revenue_total = (
                pd.to_numeric(
                    data[
                        business[
                            "revenue"
                        ]
                    ],
                    errors="coerce",
                )
                .sum()
            )


            budget_total = (
                pd.to_numeric(
                    data[
                        business[
                            "budget"
                        ]
                    ],
                    errors="coerce",
                )
                .sum()
            )


            variance = (
                revenue_total
                - budget_total
            )


            variance_pct = (

                variance
                / budget_total
                * 100

                if budget_total != 0

                else None
            )


            st.metric(
                "Revenue vs Budget variance",
                f"{variance:,.2f}",
                (
                    f"{variance_pct:+.1f}%"

                    if variance_pct
                    is not None

                    else "n/a"
                ),
            )


        if (
            business[
                "dimension"
            ]
            and financial_metrics
        ):

            metric = (

                business[
                    "revenue"
                ]

                or business[
                    "budget"
                ]

                or business[
                    "costs"
                ]
            )


            grouped = (
                grouped_data(
                    data,
                    business[
                        "dimension"
                    ],
                    metric,
                    "sum",
                )
            )


            figure = (
                px.bar(
                    grouped,
                    x=(
                        business[
                            "dimension"
                        ]
                    ),
                    y=metric,
                    title=(
                        f"{metric} "
                        f"by "
                        f"{business['dimension']}"
                    ),
                )
            )


            st.plotly_chart(
                figure,
                use_container_width=True,
            )


# ============================================================
# TRENDS TAB
# ============================================================

with trends_tab:

    st.subheader(
        (
            f"{selected_quarter} "
            "Trends"
        )
    )


    if not profile[
        "dates"
    ]:

        st.info(
            (
                "No reliable date column "
                "was detected in this sheet."
            )
        )


    elif not profile[
        "numeric"
    ]:

        st.info(
            (
                "No numeric measure was "
                "detected for trend analysis."
            )
        )


    else:

        date_column = (
            profile[
                "dates"
            ][0]
        )


        for metric in (
            profile[
                "numeric"
            ][:3]
        ):

            trend_data = (
                grouped_data(
                    data,
                    date_column,
                    metric,
                    "sum",
                )
                .sort_values(
                    date_column
                )
            )


            figure = (
                px.line(
                    trend_data,
                    x=date_column,
                    y=metric,
                    markers=True,
                    title=(
                        f"{metric} trend"
                    ),
                )
            )


            st.plotly_chart(
                figure,
                use_container_width=True,
            )


# ============================================================
# COMPARISON TAB
# ============================================================

with comparison_tab:

    st.subheader(
        "Quarter comparison"
    )


    quarter_list = (
        quarter_options
    )


    col1, col2 = (
        st.columns(
            2
        )
    )


    first_index = (

        quarter_list.index(
            selected_quarter
        )

        if selected_quarter
        in quarter_list

        else 0
    )


    second_default = min(
        first_index + 1,
        len(
            quarter_list
        ) - 1,
    )


    comparison_first = (
        col1.selectbox(
            "First quarter",
            quarter_list,
            index=first_index,
            key="comparison_first",
        )
    )


    comparison_second = (
        col2.selectbox(
            "Second quarter",
            quarter_list,
            index=second_default,
            key="comparison_second",
        )
    )


    comparison_focus = (
        st.text_input(
            "Optional comparison focus",
            placeholder=(
                "Example: focus on costs "
                "and regional performance"
            ),
        )
    )


    if st.button(
        "Compare quarters",
        key="compare_quarters_button",
    ):

        if (
            comparison_first
            == comparison_second
        ):

            st.warning(
                (
                    "Choose two different "
                    "quarters."
                )
            )


        else:

            try:

                with st.spinner(
                    (
                        "Loading and comparing "
                        "both Workiva quarters..."
                    )
                ):

                    first_bundle = (
                        quarter_bundle(
                            comparison_first
                        )
                    )


                    second_bundle = (
                        quarter_bundle(
                            comparison_second
                        )
                    )


                    comparison = (
                        compare_quarter_bundles(
                            first_bundle,
                            second_bundle,
                            comparison_focus,
                        )
                    )


                st.session_state[
                    "comparison_result"
                ] = comparison


                log_event(
                    "Compare quarters",
                    (
                        f"{comparison_first} "
                        f"vs "
                        f"{comparison_second}"
                    ),
                )


            except Exception as error:

                st.error(
                    friendly_error(
                        error
                    )
                )


    if (
        st.session_state[
            "comparison_result"
        ]
    ):

        render_comparison(
            st.session_state[
                "comparison_result"
            ]
        )


# ============================================================
# DATA QUALITY TAB
# ============================================================

with quality_tab:

    st.subheader(
        "Data Quality Dashboard"
    )


    c1, c2, c3 = (
        st.columns(
            3
        )
    )


    c1.metric(
        "Rows",
        len(
            data
        ),
    )


    c2.metric(
        "Columns",
        len(
            data.columns
        ),
    )


    c3.metric(
        "Missing cells",
        int(
            data
            .isna()
            .sum()
            .sum()
        ),
    )


    quality_rows = []


    for column in data.columns:

        detail = (
            profile[
                "details"
            ].get(
                column,
                {},
            )
        )


        quality_rows.append(
            {

                "Column":
                    column,

                "Detected type":
                    detail.get(
                        "type",
                        "unknown",
                    ),

                "Missing":
                    int(
                        data[
                            column
                        ]
                        .isna()
                        .sum()
                    ),

                "Unique":
                    int(
                        data[
                            column
                        ]
                        .nunique(
                            dropna=True
                        )
                    ),
            }
        )


    st.dataframe(
        pd.DataFrame(
            quality_rows
        ),
        use_container_width=True,
    )


    with st.expander(
        "What the agent detected"
    ):

        st.write(
            "Numeric:",
            profile[
                "numeric"
            ],
        )

        st.write(
            "Categories:",
            profile[
                "categories"
            ],
        )

        st.write(
            "Dates:",
            profile[
                "dates"
            ],
        )

        st.write(
            "Revenue:",
            business[
                "revenue"
            ],
        )

        st.write(
            "Budget:",
            business[
                "budget"
            ],
        )

        st.write(
            "Costs:",
            business[
                "costs"
            ],
        )

        st.write(
            "Main dimension:",
            business[
                "dimension"
            ],
        )


    with st.expander(
        "Raw data explorer"
    ):

        st.dataframe(
            data,
            use_container_width=True,
        )


# ============================================================
# ASK AI TAB
# ============================================================

with ask_tab:

    st.subheader(
        "Ask the Workiva Agent"
    )


    st.caption(
        (
            "Ask about the current quarter or type "
            "a comparison request such as "
            "'Compare Q1 2026 and Q2 2026'."
        )
    )


    for message in (
        st.session_state[
            "chat_messages"
        ]
    ):

        with st.chat_message(
            message[
                "role"
            ]
        ):

            st.markdown(
                message[
                    "content"
                ]
            )


    question = (
        st.chat_input(
            "Ask about Workiva data..."
        )
    )


    if question:

        st.session_state[
            "chat_messages"
        ].append(
            {
                "role":
                    "user",

                "content":
                    question,
            }
        )


        with st.chat_message(
            "user"
        ):

            st.markdown(
                question
            )


        with st.chat_message(
            "assistant"
        ):

            requested_quarters = (
                extract_quarters(
                    question
                )
            )


            if (
                len(
                    requested_quarters
                )
                >= 2
            ):

                first_quarter = (
                    requested_quarters[
                        0
                    ]
                )

                second_quarter = (
                    requested_quarters[
                        1
                    ]
                )


                try:

                    with st.spinner(
                        (
                            "Finding both quarters "
                            "in Workiva and comparing them..."
                        )
                    ):

                        first_bundle = (
                            quarter_bundle(
                                first_quarter
                            )
                        )


                        second_bundle = (
                            quarter_bundle(
                                second_quarter
                            )
                        )


                        comparison = (
                            compare_quarter_bundles(
                                first_bundle,
                                second_bundle,
                                question,
                            )
                        )


                    st.session_state[
                        "comparison_result"
                    ] = comparison


                    render_comparison(
                        comparison
                    )


                    answer = (
                        f"I compared "
                        f"{first_quarter} "
                        f"and "
                        f"{second_quarter} "
                        "using Workiva data."
                    )


                    log_event(
                        "AI comparison request",
                        question,
                    )


                except Exception as error:

                    answer = (
                        friendly_error(
                            error
                        )
                    )


                    st.error(
                        answer
                    )


            else:

                try:

                    with st.spinner(
                        "Analysing..."
                    ):

                        answer = (
                            chat_answer(
                                question,
                                summary,
                                summary_plan,
                            )
                        )


                    st.markdown(
                        answer
                    )


                    log_event(
                        "Ask AI",
                        question,
                    )


                except Exception as error:

                    answer = (
                        friendly_error(
                            error
                        )
                    )


                    st.error(
                        answer
                    )


        st.session_state[
            "chat_messages"
        ].append(
            {

                "role":
                    "assistant",

                "content":
                    answer,
            }
        )


# ============================================================
# AUDIT LOG
# ============================================================

st.divider()


with st.expander(
    "Session audit log"
):

    st.caption(
        (
            "Session-level activity log. "
            "Workiva write actions are disabled."
        )
    )


    if not st.session_state[
        "audit_log"
    ]:

        st.write(
            "No activity recorded yet."
        )


    else:

        st.dataframe(
            pd.DataFrame(
                st.session_state[
                    "audit_log"
                ]
            ),
            use_container_width=True,
        )