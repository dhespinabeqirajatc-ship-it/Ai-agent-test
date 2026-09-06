import asyncio
import json
import os
import re
from copy import deepcopy
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
    Image,
    PageBreak,
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
            "Try again in a moment."
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
            "Build an executive dashboard highlighting the most important "
            "KPIs, trends, risks, and business drivers in this dataset."
        ),

    "agent_request_input":
        (
            "Build an executive dashboard highlighting the most important "
            "KPIs, trends, risks, and business drivers in this dataset."
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

    "management_summary":
        None,

    "management_summary_key":
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

def _parse_quarter_value(value):
    """Convert common year/quarter labels to the first day of the quarter."""
    if value is None or pd.isna(value):
        return pd.NaT

    text = str(value).strip().upper()
    if not text:
        return pd.NaT

    patterns = (
        r"^Q([1-4])[\s\-_/]*(20\d{2})$",      # Q1 2026 / Q1-2026
        r"^(20\d{2})[\s\-_/]*Q([1-4])$",      # 2026 Q1 / 2026-Q1
        r"^([1-4])Q[\s\-_/]*(20\d{2})$",      # 1Q 2026
    )

    for index, pattern in enumerate(patterns):
        match = re.match(pattern, text)
        if not match:
            continue

        if index == 1:
            year = int(match.group(1))
            quarter = int(match.group(2))
        else:
            quarter = int(match.group(1))
            year = int(match.group(2))

        month = (quarter - 1) * 3 + 1
        return pd.Timestamp(year=year, month=month, day=1)

    return pd.NaT


def infer_date_series(series):
    """Infer dates from values rather than relying on the column heading.

    Returns a parsed datetime Series and a confidence ratio.  The parser
    recognises pandas datetime values, common quarter labels, year/month
    labels and ordinary date strings.  Pure numeric measures are deliberately
    protected from accidental timestamp conversion.
    """
    source = series.copy()
    parsed = pd.Series(pd.NaT, index=source.index, dtype="datetime64[ns]")
    non_null = source.dropna()

    if non_null.empty:
        return parsed, 0.0

    if pd.api.types.is_datetime64_any_dtype(source):
        parsed = pd.to_datetime(source, errors="coerce")
        return parsed, float(parsed.loc[non_null.index].notna().mean())

    text = source.astype("string").str.strip()
    present = source.notna() & text.ne("")

    # Quarter labels such as Q1 2026, 2026-Q1 and 1Q2026.
    quarter_values = source[present].map(_parse_quarter_value)
    quarter_mask = quarter_values.notna()
    if quarter_mask.any():
        parsed.loc[quarter_values.index[quarter_mask]] = quarter_values[quarter_mask]

    # Do not let plain numbers (amounts, IDs, counts) become nanosecond dates.
    remaining = present & parsed.isna()
    remaining_text = text[remaining]
    numeric_text = remaining_text.str.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)")
    date_candidates = remaining_text[~numeric_text.fillna(False)]

    if not date_candidates.empty:
        try:
            generic_dates = pd.to_datetime(
                date_candidates,
                errors="coerce",
                format="mixed",
            )
        except (TypeError, ValueError):
            # Compatibility fallback for older pandas versions.
            generic_dates = pd.to_datetime(
                date_candidates,
                errors="coerce",
            )
        parsed.loc[generic_dates.index] = generic_dates

    # Four-digit years are date-like when the column values consistently look
    # like plausible calendar years.  This remains value-driven and therefore
    # works even for headings such as "Period" or "Fiscal".
    remaining = present & parsed.isna()
    year_text = text[remaining]
    year_mask = year_text.str.fullmatch(r"(?:19|20|21)\d{2}")
    if year_mask.any():
        year_candidates = year_text[year_mask]
        year_ratio = len(year_candidates) / max(int(present.sum()), 1)
        if year_ratio >= 0.8:
            parsed.loc[year_candidates.index] = pd.to_datetime(
                year_candidates + "-01-01",
                errors="coerce",
            )

    confidence = float(parsed[present].notna().mean()) if present.any() else 0.0
    return parsed, confidence


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

        _, date_ratio = infer_date_series(series)

        normalized_name = re.sub(r"[\s_\-]+", " ", lower).strip()
        is_quarter_heading = normalized_name in {
            "quarter",
            "fiscal quarter",
            "financial quarter",
        }
        quarter_only_ratio = float(
            series.astype("string")
            .str.strip()
            .str.upper()
            .str.fullmatch(r"Q[1-4]")
            .fillna(False)
            .mean()
        )
        is_quarter_date = is_quarter_heading and quarter_only_ratio >= 0.8

        # A date interpretation takes precedence when the values themselves
        # strongly support it. The small heading fallback covers standalone
        # Q1/Q2/Q3/Q4 columns, which do not contain enough information to map
        # to an absolute timestamp. All other date detection is value-driven.
        if looks_like_id:

            kind = (
                "identifier"
            )

            profile[
                "identifiers"
            ].append(
                name
            )


        elif date_ratio >= 0.8 or is_quarter_date:

            kind = (
                "date"
            )

            profile[
                "dates"
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

            "date_confidence":
                round(float(max(date_ratio, quarter_only_ratio if is_quarter_date else 0.0)), 3),
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
            parsed, confidence = infer_date_series(data[column])
            if confidence >= 0.8:
                data[column] = parsed

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


def _find_financial_column(columns, aliases):
    """Prefer exact semantic matches, then safe substring matches."""
    normalized = [
        (column, str(column).strip().lower().replace("_", " ").replace("-", " "))
        for column in columns
    ]
    normalized_aliases = [
        str(alias).strip().lower().replace("_", " ").replace("-", " ")
        for alias in aliases
    ]

    for column, lower in normalized:
        if lower in normalized_aliases:
            return column

    # Avoid unsafe tiny substring aliases such as PY / LY.
    safe_aliases = [alias for alias in normalized_aliases if len(alias) >= 4]
    for column, lower in normalized:
        if any(alias in lower for alias in safe_aliases):
            return column

    return None


def detect_business_columns(data):
    """Detect a broader finance/operating vocabulary without using AI."""
    columns = list(data.columns)

    concepts = {
        "actual": ["actual", "actuals", "current actual"],
        "revenue": ["revenue", "sales", "net sales", "turnover", "income"],
        "budget": ["budget", "plan", "target", "budget amount"],
        "forecast": ["forecast", "latest estimate", "latest forecast", "outlook"],
        "costs": ["cost", "costs", "expense", "expenses", "total cost", "total costs"],
        "profit": ["profit", "gross profit", "net profit", "operating profit", "earnings"],
        "ebitda": ["ebitda", "adjusted ebitda"],
        "opex": ["opex", "operating expense", "operating expenses"],
        "capex": ["capex", "capital expenditure", "capital expenditures"],
        "headcount": ["headcount", "fte", "full time equivalent", "employees", "employee count"],
        "volume": ["volume", "units", "unit volume", "quantity", "qty"],
        "price": ["price", "unit price", "average price", "avg price", "asp"],
        "prior_year": [
            "prior year", "previous year", "last year", "prior year actual",
            "previous year actual", "prior year revenue", "ly actual", "py actual"
        ],
        "margin": ["margin", "gross margin", "operating margin", "net margin", "margin %"],
        "dimension": [
            "region", "country", "market", "department", "business unit",
            "segment", "division", "product", "product line", "customer group"
        ],
    }

    result = {
        key: _find_financial_column(columns, aliases)
        for key, aliases in concepts.items()
    }

    # Preserve the existing fallback behavior for common finance sheets.
    if result["budget"] is None and result["forecast"] is not None:
        result["budget"] = result["forecast"]
    if result["revenue"] is None and result["actual"] is not None:
        result["revenue"] = result["actual"]

    return result


def add_financial_intelligence(data, business):
    """Add deterministic derived finance measures. No Copilot call is used."""
    enriched = data.copy()
    derived = {}

    def numeric(column):
        if not column or column not in enriched.columns:
            return None
        return pd.to_numeric(enriched[column], errors="coerce")

    def add_column(name, values, formula):
        if values is None:
            return
        final_name = name
        suffix = 2
        while final_name in enriched.columns:
            final_name = f"{name} {suffix}"
            suffix += 1
        enriched[final_name] = values
        derived[final_name] = formula

    actual_col = business.get("actual") or business.get("revenue")
    budget_col = business.get("budget")
    revenue_col = business.get("revenue")
    costs_col = business.get("costs")
    profit_col = business.get("profit")
    ebitda_col = business.get("ebitda")
    opex_col = business.get("opex")
    capex_col = business.get("capex")
    prior_year_col = business.get("prior_year")

    actual = numeric(actual_col)
    budget = numeric(budget_col)
    revenue = numeric(revenue_col)
    costs = numeric(costs_col)
    profit = numeric(profit_col)
    ebitda = numeric(ebitda_col)
    opex = numeric(opex_col)
    capex = numeric(capex_col)
    prior_year = numeric(prior_year_col)

    if actual is not None and budget is not None and actual_col != budget_col:
        variance = actual - budget
        add_column(
            "Actual vs Budget Variance",
            variance,
            f"{actual_col} - {budget_col}",
        )
        denominator = budget.replace(0, pd.NA)
        add_column(
            "Actual vs Budget Variance %",
            variance / denominator * 100,
            f"({actual_col} - {budget_col}) / {budget_col} × 100",
        )

    if revenue is not None and profit is not None and revenue_col != profit_col:
        add_column(
            "Margin %",
            profit / revenue.replace(0, pd.NA) * 100,
            f"{profit_col} / {revenue_col} × 100",
        )

    if revenue is not None and costs is not None and revenue_col != costs_col:
        add_column(
            "Cost Ratio %",
            costs / revenue.replace(0, pd.NA) * 100,
            f"{costs_col} / {revenue_col} × 100",
        )

    if revenue is not None and ebitda is not None and revenue_col != ebitda_col:
        add_column(
            "EBITDA Margin %",
            ebitda / revenue.replace(0, pd.NA) * 100,
            f"{ebitda_col} / {revenue_col} × 100",
        )

    if revenue is not None and opex is not None and revenue_col != opex_col:
        add_column(
            "Opex Ratio %",
            opex / revenue.replace(0, pd.NA) * 100,
            f"{opex_col} / {revenue_col} × 100",
        )

    if revenue is not None and capex is not None and revenue_col != capex_col:
        add_column(
            "Capex Ratio %",
            capex / revenue.replace(0, pd.NA) * 100,
            f"{capex_col} / {revenue_col} × 100",
        )

    if actual is not None and prior_year is not None and actual_col != prior_year_col:
        yoy = actual - prior_year
        add_column(
            "Prior Year Variance",
            yoy,
            f"{actual_col} - {prior_year_col}",
        )
        add_column(
            "Prior Year Variance %",
            yoy / prior_year.replace(0, pd.NA) * 100,
            f"({actual_col} - {prior_year_col}) / {prior_year_col} × 100",
        )

    return enriched, derived


# ============================================================
# ANALYTICAL CONTEXT
# ============================================================

def build_analytical_context(data, profile, business):
    """Pre-calculate trustworthy facts for the AI summary and chat."""

    context = {
        "metric_totals": {},
        "top_breakdowns": {},
        "trends": {},
        "correlations": [],
        "financial_variance": {},
        "data_quality": {},
    }

    row_count = max(len(data), 1)
    missing_by_column = data.isna().sum().sort_values(ascending=False)
    context["data_quality"] = {
        "missing_cells": int(data.isna().sum().sum()),
        "missing_rate_pct": float(data.isna().sum().sum() / max(data.size, 1) * 100),
        "most_missing_columns": [
            {"column": str(col), "missing": int(value), "missing_pct": float(value / row_count * 100)}
            