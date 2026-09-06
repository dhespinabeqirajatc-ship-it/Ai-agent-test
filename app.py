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

    "forecast_result":
        None,

    "forecast_ai_commentary":
        None,

    "forecast_ai_commentary_key":
        None,

    "trends_ai_summary":
        None,

    "trends_ai_summary_key":
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

def _parse_quarter_value(value, default_year=None, allow_bare_quarter=False):
    """Convert common quarter labels to the first day of the quarter.

    ``default_year`` lets standalone values such as Q1/Q2 become real
    datetimes when the surrounding Workiva sheet already supplies the year.
    """
    if value is None or pd.isna(value):
        return pd.NaT

    text = str(value).strip().upper()
    if not text:
        return pd.NaT

    # Normalize common finance variants first.
    text = re.sub(r"\bFISCAL\s+YEAR\b", "FY", text)
    text = re.sub(r"\bFISCAL\s+QUARTER\b", "Q", text)
    text = re.sub(r"\bQUARTER\b|\bQTR\b", "Q", text)
    text = re.sub(r"\s+", " ", text).strip()

    patterns = (
        r"^Q([1-4])[\s\-_/]*(?:FY)?(20\d{2})$",   # Q1 2026 / Q1 FY2026
        r"^(?:FY)?(20\d{2})[\s\-_/]*Q([1-4])$",   # 2026 Q1 / FY2026-Q1
        r"^([1-4])Q[\s\-_/]*(?:FY)?(20\d{2})$",   # 1Q 2026
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

    # Standalone Q1/Q2/Q3/Q4 can still be made chronological when the
    # caller knows the sheet/reporting year.
    standalone = re.fullmatch(r"Q([1-4])", text)
    if standalone is None and allow_bare_quarter:
        standalone = re.fullmatch(r"([1-4])", text)
    if standalone and default_year is not None:
        quarter = int(standalone.group(1))
        month = (quarter - 1) * 3 + 1
        return pd.Timestamp(year=int(default_year), month=month, day=1)

    return pd.NaT


def infer_date_series(series, default_year=None, allow_bare_quarter=False):
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
    quarter_values = source[present].map(
        lambda value: _parse_quarter_value(
            value,
            default_year=default_year,
            allow_bare_quarter=allow_bare_quarter,
        )
    )
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


def profile_dataframe(data, default_year=None):

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

        _, date_ratio = infer_date_series(series, default_year=default_year)

        normalized_name = re.sub(r"[\s_\-]+", " ", lower).strip()
        is_quarter_heading = normalized_name in {
            "quarter",
            "fiscal quarter",
            "financial quarter",
        }
        quarter_text = (
            series.astype("string")
            .str.strip()
            .str.upper()
        )
        quarter_only_ratio = float(
            quarter_text
            .str.fullmatch(r"(?:Q|FQ|QTR\s*)[1-4]")
            .fillna(False)
            .mean()
        )
        # Quarter-shaped values are temporal even when the heading is called
        # something generic such as Period or Reporting Bucket. The heading
        # remains only a fallback, not the primary signal.
        bare_quarter_ratio = float(
            quarter_text
            .str.fullmatch(r"[1-4]")
            .fillna(False)
            .mean()
        )
        is_quarter_date = (
            quarter_only_ratio >= 0.8
            or (is_quarter_heading and bare_quarter_ratio >= 0.8)
            or (is_quarter_heading and date_ratio >= 0.5)
        )

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
    default_year=None,
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
            parsed, confidence = infer_date_series(
                data[column],
                default_year=default_year,
                allow_bare_quarter=(
                    re.sub(r"[\s_\-]+", " ", str(column).strip().lower()).strip()
                    in {"quarter", "fiscal quarter", "financial quarter"}
                ),
            )
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

    # Avoid unsafe tiny substring aliases such as PY / LY, and do not let a
    # generated measure such as "Actual vs Budget Variance" become the
    # source Actual field merely because it contains the word "actual".
    safe_aliases = [alias for alias in normalized_aliases if len(alias) >= 4]
    unsafe_generic_substrings = {"actual", "actuals", "plan", "target", "income"}
    safe_aliases = [a for a in safe_aliases if a not in unsafe_generic_substrings]
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

def _is_additive_driver_metric(metric):
    """Return True for measures where category contributions can be summed safely."""
    lower = str(metric).strip().lower()
    non_additive_tokens = (
        "%", "ratio", "margin", "rate", "average", "avg", "per ", "share",
    )
    return not any(token in lower for token in non_additive_tokens)


def build_driver_analysis(data, profile, business):
    """Deterministically explain what categories drove change between two periods.

    The analysis intentionally uses arithmetic only.  It compares the earliest and
    latest observed dates, decomposes additive metric changes by low-cardinality
    dimensions, and ranks contributors by absolute impact.
    """
    if not profile.get("dates") or not profile.get("numeric"):
        return {}

    date_col = profile["dates"][0]
    working_dates = pd.to_datetime(data[date_col], errors="coerce")
    valid_dates = working_dates.dropna().sort_values().unique()
    if len(valid_dates) < 2:
        return {}

    first_date = pd.Timestamp(valid_dates[0])
    last_date = pd.Timestamp(valid_dates[-1])

    # Put the detected management dimension first, then other usable categories.
    category_candidates = []
    preferred = business.get("dimension")
    if preferred in data.columns:
        category_candidates.append(preferred)
    for category in profile.get("categories", []):
        if category not in category_candidates and category in data.columns:
            category_candidates.append(category)

    category_candidates = [
        category for category in category_candidates
        if 2 <= data[category].nunique(dropna=True) <= 40
    ][:4]

    metric_candidates = [
        metric for metric in profile.get("numeric", [])
        if metric in data.columns and _is_additive_driver_metric(metric)
    ][:8]

    if not category_candidates or not metric_candidates:
        return {}

    result = {
        "date_column": str(date_col),
        "first_period": str(first_date),
        "last_period": str(last_date),
        "metrics": {},
    }

    for metric in metric_candidates:
        metric_values = pd.to_numeric(data[metric], errors="coerce")
        metric_result = {}

        for category in category_candidates:
            frame = pd.DataFrame({
                "__date": working_dates,
                "__category": data[category].astype("string"),
                "__value": metric_values,
            }).dropna(subset=["__date", "__category", "__value"])

            frame = frame[frame["__date"].isin([first_date, last_date])]
            if frame.empty:
                continue

            grouped = (
                frame.groupby(["__date", "__category"], dropna=False)["__value"]
                .sum()
                .unstack(fill_value=0.0)
            )
            if first_date not in grouped.index or last_date not in grouped.index:
                continue

            first_values = grouped.loc[first_date]
            last_values = grouped.loc[last_date]
            all_categories = first_values.index.union(last_values.index)
            first_values = first_values.reindex(all_categories, fill_value=0.0)
            last_values = last_values.reindex(all_categories, fill_value=0.0)
            changes = (last_values - first_values).sort_values(
                key=lambda x: x.abs(), ascending=False
            )

            total_first = float(first_values.sum())
            total_last = float(last_values.sum())
            total_change = float(total_last - total_first)
            if changes.empty or abs(total_change) < 1e-12:
                continue

            drivers = []
            for category_value, change in changes.head(8).items():
                change = float(change)
                drivers.append({
                    "category": str(category_value),
                    "first_value": float(first_values.get(category_value, 0.0)),
                    "last_value": float(last_values.get(category_value, 0.0)),
                    "change": change,
                    "contribution_pct_of_net_change": float(change / total_change * 100),
                    "direction": "increase" if change > 0 else "decrease",
                })

            metric_result[str(category)] = {
                "total_first": total_first,
                "total_last": total_last,
                "total_change": total_change,
                "total_change_pct": (
                    float(total_change / total_first * 100) if total_first else None
                ),
                "drivers": drivers,
            }

        if metric_result:
            result["metrics"][str(metric)] = metric_result

    return result if result["metrics"] else {}


def build_analytical_context(data, profile, business):
    """Pre-calculate trustworthy facts for the AI summary and chat."""

    context = {
        "metric_totals": {},
        "top_breakdowns": {},
        "trends": {},
        "correlations": [],
        "anomalies": {},
        "driver_analysis": {},
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
            for col, value in missing_by_column.head(5).items()
            if value > 0
        ],
    }

    for metric in profile["numeric"][:12]:
        values = pd.to_numeric(data[metric], errors="coerce").dropna()
        if values.empty:
            continue
        context["metric_totals"][metric] = {
            "sum": float(values.sum()),
            "average": float(values.mean()),
            "median": float(values.median()),
            "minimum": float(values.min()),
            "maximum": float(values.max()),
            "non_null_count": int(values.count()),
        }

        # Deterministic IQR outlier screening gives AI an evidence-backed
        # anomaly signal without asking the model to invent thresholds.
        if len(values) >= 4:
            q1 = float(values.quantile(0.25))
            q3 = float(values.quantile(0.75))
            iqr = q3 - q1
            if iqr > 0:
                lower_bound = q1 - 1.5 * iqr
                upper_bound = q3 + 1.5 * iqr
                outliers = values[(values < lower_bound) | (values > upper_bound)]
                if not outliers.empty:
                    context["anomalies"][metric] = {
                        "method": "IQR 1.5x",
                        "count": int(outliers.count()),
                        "lower_bound": float(lower_bound),
                        "upper_bound": float(upper_bound),
                        "minimum_outlier": float(outliers.min()),
                        "maximum_outlier": float(outliers.max()),
                    }

    for category in profile["categories"][:6]:
        unique = data[category].nunique(dropna=True)
        if unique < 2 or unique > 40:
            continue
        for metric in profile["numeric"][:6]:
            grouped = grouped_data(data, category, metric, "sum").dropna()
            if grouped.empty:
                continue
            grouped = grouped.sort_values(metric, ascending=False)
            key = f"{metric} by {category}"
            context["top_breakdowns"][key] = {
                "top": [
                    {"category": str(row[category]), "value": float(row[metric])}
                    for _, row in grouped.head(5).iterrows()
                ],
                "bottom": [
                    {"category": str(row[category]), "value": float(row[metric])}
                    for _, row in grouped.tail(3).iterrows()
                ],
            }

    if profile["dates"]:
        date_col = profile["dates"][0]
        for metric in profile["numeric"][:8]:
            working = data[[date_col, metric]].copy()
            working[date_col] = pd.to_datetime(working[date_col], errors="coerce")
            working[metric] = pd.to_numeric(working[metric], errors="coerce")
            working = working.dropna()
            if working.empty:
                continue
            grouped = working.groupby(date_col)[metric].sum().sort_index()
            if len(grouped) < 2:
                continue
            first = float(grouped.iloc[0])
            last = float(grouped.iloc[-1])
            change = last - first
            context["trends"][metric] = {
                "date_column": str(date_col),
                "first_date": str(grouped.index[0]),
                "last_date": str(grouped.index[-1]),
                "first_value": first,
                "last_value": last,
                "change": float(change),
                "change_pct": float(change / first * 100) if first else None,
                "peak_date": str(grouped.idxmax()),
                "peak_value": float(grouped.max()),
                "low_date": str(grouped.idxmin()),
                "low_value": float(grouped.min()),
            }

    numeric_frame = data[profile["numeric"][:10]].apply(pd.to_numeric, errors="coerce")
    if numeric_frame.shape[1] >= 2:
        corr = numeric_frame.corr()
        pairs = []
        columns = list(corr.columns)
        for i, first in enumerate(columns):
            for second in columns[i + 1:]:
                value = corr.loc[first, second]
                if pd.notna(value):
                    pairs.append((abs(float(value)), first, second, float(value)))
        for _, first, second, value in sorted(pairs, reverse=True)[:5]:
            context["correlations"].append({
                "metric_1": str(first),
                "metric_2": str(second),
                "correlation": value,
            })

    revenue = business.get("revenue")
    budget = business.get("budget")
    costs = business.get("costs")
    if revenue and budget:
        actual = pd.to_numeric(data[revenue], errors="coerce").sum()
        plan = pd.to_numeric(data[budget], errors="coerce").sum()
        variance = actual - plan
        context["financial_variance"]["actual_vs_budget"] = {
            "actual_metric": str(revenue),
            "budget_metric": str(budget),
            "actual": float(actual),
            "budget": float(plan),
            "variance": float(variance),
            "variance_pct": float(variance / plan * 100) if plan else None,
        }
    if revenue and costs:
        rev = pd.to_numeric(data[revenue], errors="coerce").sum()
        cost = pd.to_numeric(data[costs], errors="coerce").sum()
        context["financial_variance"]["revenue_vs_costs"] = {
            "revenue": float(rev),
            "costs": float(cost),
            "spread": float(rev - cost),
            "spread_pct_of_revenue": float((rev - cost) / rev * 100) if rev else None,
        }

    context["driver_analysis"] = build_driver_analysis(
        data,
        profile,
        business,
    )

    return context


def build_question_context(question, data, profile, business):
    """Calculate extra facts specifically relevant to the user's question."""

    lower = question.lower()
    metric_candidates = [
        c for c in profile["numeric"]
        if str(c).lower() in lower
    ]
    category_candidates = [
        c for c in profile["categories"]
        if str(c).lower() in lower
    ]

    keyword_map = {
        "revenue": business.get("revenue"),
        "sales": business.get("revenue"),
        "budget": business.get("budget"),
        "forecast": business.get("budget"),
        "cost": business.get("costs"),
        "expense": business.get("costs"),
        "profit": business.get("profit"),
        "ebitda": business.get("ebitda"),
        "opex": business.get("opex"),
        "capex": business.get("capex"),
        "headcount": business.get("headcount"),
        "fte": business.get("headcount"),
        "volume": business.get("volume"),
        "price": business.get("price"),
        "margin": business.get("margin"),
        "prior year": business.get("prior_year"),
    }
    for word, column in keyword_map.items():
        if word in lower and column and column not in metric_candidates:
            metric_candidates.append(column)

    if not metric_candidates:
        metric_candidates = profile["numeric"][:3]
    if not category_candidates and business.get("dimension"):
        category_candidates = [business["dimension"]]

    result = {"metrics": {}, "breakdowns": {}, "sample_rows": []}
    for metric in metric_candidates[:4]:
        values = pd.to_numeric(data[metric], errors="coerce").dropna()
        if not values.empty:
            result["metrics"][metric] = {
                "sum": float(values.sum()),
                "average": float(values.mean()),
                "median": float(values.median()),
                "minimum": float(values.min()),
                "maximum": float(values.max()),
            }

    for category in category_candidates[:2]:
        if category not in data.columns or data[category].nunique(dropna=True) > 50:
            continue
        for metric in metric_candidates[:3]:
            grouped = grouped_data(data, category, metric, "sum").sort_values(metric, ascending=False)
            result["breakdowns"][f"{metric} by {category}"] = grouped.head(12).to_dict(orient="records")

    matched_columns = [c for c in data.columns if str(c).lower() in lower]
    if matched_columns:
        result["sample_rows"] = data[matched_columns[:6]].head(8).astype(str).to_dict(orient="records")

    return result


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
# GROUNDED AI SAFETY
# ============================================================

def _numeric_tokens(text):
    """Extract user-visible numeric claims while ignoring list numbering."""
    text = str(text or "")
    pattern = re.compile(r"(?<![A-Za-z])[-+]?\d[\d,]*(?:\.\d+)?(?:%|x)?")
    values = []
    for match in pattern.finditer(text):
        token = match.group(0)
        start, end = match.span()
        before = text[max(0, start - 2):start]
        after = text[end:end + 1]
        # Ignore markdown/section list markers such as "1." at line starts.
        line_start = text.rfind("\n", 0, start) + 1
        prefix = text[line_start:start].strip()
        if prefix == "" and after == "." and token.lstrip("+-").isdigit():
            continue
        cleaned = token.rstrip("%x").replace(",", "")
        try:
            values.append((token, float(cleaned)))
        except ValueError:
            pass
    return values


def _allowed_evidence_numbers(evidence):
    """Build a tolerant set of numbers that are actually present in evidence."""
    serialized = json.dumps(evidence, default=str, ensure_ascii=False)
    raw = [value for _, value in _numeric_tokens(serialized)]
    allowed = set()
    for value in raw:
        # Permit normal presentation rounding, but not newly invented values.
        for decimals in range(0, 7):
            allowed.add(round(float(value), decimals))
    return allowed


def _unsupported_ai_numbers(text, evidence):
    allowed = _allowed_evidence_numbers(evidence)
    unsupported = []
    for token, value in _numeric_tokens(text):
        if not any(abs(value - candidate) <= 1e-9 for candidate in allowed):
            unsupported.append(token)
    return list(dict.fromkeys(unsupported))


@st.cache_data(ttl=900, show_spinner=False)
def grounded_ai_response(prompt, evidence, purpose="analysis"):
    """Use Copilot for interpretation while Python remains the source of truth.

    The model receives a bounded evidence package. Its numeric claims are checked
    after generation; if unsupported numbers appear, it gets one constrained
    rewrite. If the rewrite still introduces unsupported figures, the response is
    withheld rather than displaying an ungrounded numerical statement.
    """
    evidence_text = json.dumps(evidence, default=str, ensure_ascii=False)
    grounding_rules = f"""

GROUND-TRUTH EVIDENCE (calculated/retrieved by the application):
{evidence_text}

MANDATORY GROUNDING RULES:
- Treat the evidence above as the complete factual universe for this {purpose}.
- Python/Workiva values are authoritative. Never calculate or invent an official value yourself.
- Do not introduce a number unless that exact value, or an ordinary rounded display of it, is present in the evidence.
- Never invent causes, events, targets, benchmarks, seasonality, probabilities, confidence levels, or external drivers.
- You may interpret patterns that are directly supported by the evidence, but label uncertain interpretations as such.
- If the evidence cannot answer something, explicitly say the data does not establish it.
- Recommendations must be framed as management questions/actions to investigate, not as claims about facts not in evidence.
- Do not use numbered section headings; use plain markdown headings or bullets.
"""
    full_prompt = prompt + grounding_rules
    response = run_copilot(full_prompt)
    unsupported = _unsupported_ai_numbers(response, evidence)

    if unsupported:
        repair_prompt = f"""
{full_prompt}

Your previous draft contained unsupported numeric claims: {unsupported}
Previous draft:
{response}

Rewrite the answer now. Remove every unsupported number and any claim that depends on it.
Use only evidence-supported facts. Return only the corrected answer.
"""
        response = run_copilot(repair_prompt)
        unsupported = _unsupported_ai_numbers(response, evidence)

    if unsupported:
        return (
            "AI commentary was withheld because its generated text still contained "
            "numeric claims that could not be matched to the calculated dataset. "
            "The dashboard values and deterministic analysis remain available above."
        )

    return response


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
# COMPACT AI CONTEXT
# ============================================================

def compact_ai_context(summary):
    """Return a smaller, decision-useful context for Copilot prompts.

    The full analytical context remains available in Python for the UI and
    exports.  Copilot only receives the most material subset, which reduces
    prompt size and model latency without adding any extra API calls.
    """
    context = summary.get("analytical_context", {}) or {}

    compact = {
        "financial_variance": context.get("financial_variance", {}),
        "data_quality": context.get("data_quality", {}),
        "trends": {},
        "top_breakdowns": {},
        "correlations": (context.get("correlations", []) or [])[:3],
        "anomalies": dict(list((context.get("anomalies", {}) or {}).items())[:5]),
        "driver_analysis": {},
    }

    # Keep only the strongest few trend signals.
    trend_rows = []
    for metric, details in (context.get("trends", {}) or {}).items():
        pct = details.get("change_pct")
        score = abs(float(pct)) if pct is not None else 0.0
        trend_rows.append((score, metric, details))
    for _, metric, details in sorted(trend_rows, reverse=True)[:5]:
        compact["trends"][metric] = details

    # Keep driver evidence compact: strongest few metrics/dimensions and top contributors.
    drivers = context.get("driver_analysis", {}) or {}
    if drivers.get("metrics"):
        compact["driver_analysis"] = {
            "date_column": drivers.get("date_column"),
            "first_period": drivers.get("first_period"),
            "last_period": drivers.get("last_period"),
            "metrics": {},
        }
        for metric, dimensions in list(drivers.get("metrics", {}).items())[:4]:
            compact["driver_analysis"]["metrics"][metric] = {}
            for dimension, details in list(dimensions.items())[:2]:
                compact["driver_analysis"]["metrics"][metric][dimension] = {
                    "total_first": details.get("total_first"),
                    "total_last": details.get("total_last"),
                    "total_change": details.get("total_change"),
                    "total_change_pct": details.get("total_change_pct"),
                    "drivers": (details.get("drivers", []) or [])[:5],
                }

    # Keep only a few category breakdowns and fewer rows within each.
    for key, details in list((context.get("top_breakdowns", {}) or {}).items())[:6]:
        compact["top_breakdowns"][key] = {
            "top": (details.get("top", []) or [])[:3],
            "bottom": (details.get("bottom", []) or [])[:2],
        }

    return compact




def compact_prompt_summary(summary):
    """Small prompt payload for faster Copilot responses."""
    numeric_stats = summary.get("numeric_stats", {}) or {}
    categories = summary.get("categories", {}) or {}
    return {
        "numeric_stats": dict(list(numeric_stats.items())[:10]),
        "categories": dict(list(categories.items())[:6]),
        "date_ranges": summary.get("date_ranges", {}),
    }

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

    prompt_summary = compact_prompt_summary(summary)

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
{prompt_summary["numeric_stats"]}

Date ranges:
{prompt_summary["date_ranges"]}

Pre-calculated analytical context:
{compact_ai_context(summary)}


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
horizontal_bar
line
area
donut
scatter
histogram
box
treemap

Rules:

- Use only exact supplied column names.
- Never invent a column.
- Build between 3 and 6 useful KPIs.
- Build between 3 and 6 useful charts when the data supports them.
- Prefer management usefulness over visual variety.
- Use donut only for low-cardinality categories.
- Use line or area when a real date column exists.
- Use scatter only when appropriate numeric fields exist.
- Use bars for categorical comparisons.
- Use horizontal_bar for ranked categories with long labels.
- Use histogram to show a numeric distribution.
- Use box to show spread/outliers across a category.
- Use treemap for hierarchical-looking category shares when cardinality is manageable.
- Choose charts that answer distinct management questions rather than repeating the same view.
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
        "horizontal_bar",
        "line",
        "area",
        "donut",
        "scatter",
        "histogram",
        "box",
        "treemap",
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
    )[:6]:

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
    )[:6]:

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

    prompt_summary = compact_prompt_summary(summary)

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
{prompt_summary["numeric_stats"]}

Categories:
{prompt_summary["categories"]}

Date ranges:
{prompt_summary["date_ranges"]}

Pre-calculated analytical context:
{compact_ai_context(summary)}

Write an executive-grade management summary with clear sections:
1. Executive takeaway — 2 to 4 sentences on what matters most.
2. Performance signals — the strongest supported movements, concentration, trend or variance signals.
3. Risks / watch items — only risks supported by the supplied facts, with evidence.
4. Decisions / actions — 2 to 4 practical management actions or questions.

Prioritize materiality and business relevance. Quantify observations whenever a supplied value supports them.
Distinguish facts from interpretation. Never invent causes.
If the available statistics do not support a conclusion, say so clearly.
"""


    evidence = {
        "quarter": summary.get("quarter"),
        "dashboard": {
            "title": plan.get("title"),
            "reason": plan.get("reason"),
        },
        "numeric_statistics": prompt_summary.get("numeric_stats", {}),
        "categories": prompt_summary.get("categories", {}),
        "date_ranges": prompt_summary.get("date_ranges", {}),
        "analytical_context": compact_ai_context(summary),
    }

    return grounded_ai_response(
        prompt,
        evidence,
        purpose="management summary",
    )


def generate_trends_ai_summary(summary):
    """Generate a grounded executive summary specifically for the Trends board."""
    context = compact_ai_context(summary)
    evidence = {
        "quarter": summary.get("quarter"),
        "date_ranges": summary.get("date_ranges", {}),
        "trends": context.get("trends", {}),
        "anomalies": context.get("anomalies", {}),
        "correlations": context.get("correlations", []),
        "driver_analysis": context.get("driver_analysis", {}),
        "financial_variance": context.get("financial_variance", {}),
    }

    prompt = f"""
You are an executive analyst reviewing the Trends board of a Workiva dashboard.

Grounded trend evidence:
{evidence}

Rules:
- Use only the supplied evidence. Never invent dates, values, causes, seasonality, or business events.
- Describe direction, magnitude, peaks/lows, anomalies, relationships, and calculated category drivers only when directly supported.
- When driver analysis is available, distinguish mathematical contribution from causal explanation. A contributor explains where the numeric change occurred; it does not prove why it occurred.
- Treat IQR anomaly flags as screening signals, not proof of an error or root cause.
- Distinguish observed facts from interpretation.
- If there is not enough time history to support a strong trend conclusion, say so clearly.
- Keep the tone concise, executive, and decision-oriented.

Write sections:
1. Trend takeaway
2. Strongest movements
3. Key drivers of change
4. Anomalies / watch items
5. Management questions
"""
    return grounded_ai_response(
        prompt,
        evidence,
        purpose="trends summary",
    )


# ============================================================
# CHAT
# ============================================================

def chat_answer(
    question,
    summary,
    plan,
    data,
    profile,
    business,
    chat_history=None,
    forecast_result=None,
):

    question_context = build_question_context(
        question,
        data,
        profile,
        business,
    )

    # Keep enough conversational continuity without repeatedly sending a long
    # transcript to Copilot. This materially reduces prompt size/latency.
    recent_history = (chat_history or [])[-6:]

    compact_context = compact_ai_context(summary)
    compact_summary = compact_prompt_summary(summary)

    # Give Copilot the exact dashboard the user is currently looking at.
    # This is local context construction only; it does not trigger another
    # Workiva request or AI call.
    visible_kpis = [
        {
            "label": item.get("label", item.get("column")),
            "column": item.get("column"),
            "aggregation": item.get("aggregation", "sum"),
        }
        for item in (plan or {}).get("kpis", [])
    ]

    visible_chart_items = [
        item for item in (plan or {}).get("charts", [])
        if item.get("visible", True)
    ]
    visible_charts = [
        {
            "position": (
                "primary" if index == 0 else "supporting"
            ),
            "number": index + 1,
            "title": item.get("title"),
            "type": item.get("type"),
            "x": item.get("x"),
            "y": item.get("y"),
            "aggregation": item.get("aggregation", "sum"),
            "top_n": item.get("top_n", "All"),
            "sort_order": item.get("sort_order", "Descending"),
            "show_percentage": item.get("show_percentage", False),
        }
        for index, item in enumerate(visible_chart_items)
    ]

    forecast_context = None
    if forecast_result and forecast_result.get("forecasts"):
        forecast_context = {
            "historical_quarters": forecast_result.get("quarters", []),
            "forecast_horizon_quarters": forecast_result.get("horizon", 0),
            "scenario_sensitivity_pct": forecast_result.get("scenario_sensitivity_pct", 0),
            "metrics": {},
        }
        for metric, frame in forecast_result.get("forecasts", {}).items():
            backtest = (forecast_result.get("backtests", {}) or {}).get(metric, {}) or {}
            scenario_frame = (forecast_result.get("scenarios", {}) or {}).get(metric)
            forecast_context["metrics"][metric] = {
                "series": (
                    frame[["Quarter", "Value", "Series"]]
                    .to_dict(orient="records")
                ),
                "scenarios": (
                    scenario_frame[["Quarter", "Value", "Scenario"]].to_dict(orient="records")
                    if isinstance(scenario_frame, pd.DataFrame) and not scenario_frame.empty
                    else []
                ),
                "backtest": {
                    "history_points": backtest.get("history_points"),
                    "backtest_points": backtest.get("backtest_points"),
                    "mae": backtest.get("mae"),
                    "rmse": backtest.get("rmse"),
                    "mape_pct": backtest.get("mape_pct"),
                    "wape_pct": backtest.get("wape_pct"),
                    "smape_pct": backtest.get("smape_pct"),
                    "confidence": backtest.get("confidence"),
                    "confidence_reason": backtest.get("confidence_reason"),
                },
            }

    trends_context = {
        "trends": compact_context.get("trends", {}),
        "anomalies": compact_context.get("anomalies", {}),
        "correlations": compact_context.get("correlations", []),
        "driver_analysis": compact_context.get("driver_analysis", {}),
    }

    dashboard_context = {
        "title": (plan or {}).get("title"),
        "reason": (plan or {}).get("reason"),
        "layout": {
            "kpi_row": visible_kpis,
            "primary_chart": (
                visible_charts[0] if visible_charts else None
            ),
            "supporting_charts": (
                visible_charts[1:] if len(visible_charts) > 1 else []
            ),
            "management_focus_panels": [
                "Variance",
                "Risk / watch",
                "Management questions",
            ],
        },
        "charts": visible_charts,
    }

    prompt = f"""
You are a senior conversational analyst inside a Workiva management dashboard.

Quarter:
{summary["quarter"]}

Dashboard:
{plan.get("title", "Current dashboard")}

Visible dashboard configuration:
{dashboard_context}

Columns:
{summary["columns"]}

Relevant numeric statistics:
{compact_summary["numeric_stats"]}

Relevant category information:
{compact_summary["categories"]}

Pre-calculated analytical context:
{compact_context}

Trends board context:
{trends_context}

Current forecasting board context (if built):
{forecast_context}

Question-specific calculations:
{question_context}

Recent conversation:
{recent_history}

User question:
{question}

Rules:
- Never invent numbers, causes, definitions or data that are not supplied.
- Answer the question first, then give the most useful supporting evidence.
- Treat the visible dashboard configuration as authoritative for what the user can currently see.
- If the user refers to a chart by position or number (for example, "the first chart" or "the second chart"), resolve it from the visible dashboard configuration.
- If the user asks why a visible KPI or chart matters, connect that visual to the supplied analytical facts without inventing causes.
- Use exact figures when available and explain the calculation basis (sum, average, ranking, etc.).
- For ranking questions, identify leaders and laggards from the supplied breakdowns.
- For trend questions, use the Trends board context and describe direction, magnitude, peaks/lows, anomaly screening signals, correlations, and calculated drivers only when supported.
- For driver questions, explain which category values mathematically contributed to the change. Never convert contribution into an unsupported causal claim.
- For forecasting questions, use the current forecasting board context when available. Clearly distinguish historical actuals from projected values and state that projections use simple trend extrapolation.
- When forecast backtest evidence is available, use it to explain historical forecast error and the supplied confidence label. Never upgrade the supplied confidence label or imply that historical accuracy guarantees future performance.
- If the user asks about a forecast but no forecasting board has been built, say that clearly and direct them to build one rather than inventing a projection.
- If the question is ambiguous, state the most reasonable interpretation and what would change the answer.
- If the data is insufficient, say exactly what is missing and suggest the next useful question.
- Keep the tone concise, executive and analytical, but allow enough detail to explain the evidence.
"""

    evidence = {
        "quarter": summary.get("quarter"),
        "dashboard": dashboard_context,
        "columns": summary.get("columns", []),
        "numeric_statistics": compact_summary.get("numeric_stats", {}),
        "category_information": compact_summary.get("categories", {}),
        "analytical_context": compact_context,
        "trends_board_context": trends_context,
        "forecasting_board_context": forecast_context,
        "question_specific_calculations": question_context,
        "user_question": question,
    }

    return grounded_ai_response(
        prompt,
        evidence,
        purpose="dashboard question answering",
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


def build_chart_figure(data, chart):
    if not chart.get("visible", True):
        return None

    chart_type = chart.get("type", "bar")
    x = chart.get("x")
    y = chart.get("y")
    title = chart.get("title", chart_type.title())
    aggregation = chart.get("aggregation", "sum")
    sort_order = chart.get("sort_order", "Descending")
    top_n = chart.get("top_n")
    show_percentage = bool(chart.get("show_percentage", False))

    if chart_type == "histogram":
        metric = y or x
        if metric not in data.columns:
            return None
        return px.histogram(data, x=metric, title=title)

    if chart_type == "scatter":
        if x not in data.columns or y not in data.columns:
            return None
        return px.scatter(data, x=x, y=y, title=title)

    if chart_type == "box":
        if y not in data.columns:
            return None
        if x and x in data.columns:
            return px.box(data, x=x, y=y, title=title)
        return px.box(data, y=y, title=title)

    if x is None or y is None or x not in data.columns or y not in data.columns:
        return None

    chart_data = grouped_data(data, x, y, aggregation)
    if chart_data.empty:
        return None

    is_date = pd.api.types.is_datetime64_any_dtype(data[x])

    if is_date:
        chart_data = chart_data.sort_values(x)
    elif sort_order == "Ascending":
        chart_data = chart_data.sort_values(y, ascending=True)
    elif sort_order == "Descending":
        chart_data = chart_data.sort_values(y, ascending=False)

    if not is_date and top_n not in (None, "All"):
        try:
            limit = int(top_n)
            if limit > 0:
                chart_data = chart_data.head(limit)
        except (TypeError, ValueError):
            pass

    display_y = y
    if show_percentage and not is_date and chart_type in {
        "bar", "horizontal_bar", "donut", "treemap"
    }:
        total = pd.to_numeric(chart_data[y], errors="coerce").sum()
        if pd.notna(total) and total != 0:
            display_y = "Share %"
            chart_data = chart_data.copy()
            chart_data[display_y] = (
                pd.to_numeric(chart_data[y], errors="coerce") / total * 100
            )
            title = f"{title} (%)"

    if chart_type == "bar":
        figure = px.bar(chart_data, x=x, y=display_y, title=title)
    elif chart_type == "horizontal_bar":
        figure = px.bar(
            chart_data,
            x=display_y,
            y=x,
            orientation="h",
            title=title,
        )
    elif chart_type == "line":
        figure = px.line(chart_data, x=x, y=y, markers=True, title=title)
    elif chart_type == "area":
        figure = px.area(chart_data, x=x, y=y, title=title)
    elif chart_type == "donut":
        figure = px.pie(
            chart_data,
            names=x,
            values=display_y,
            hole=0.55,
            title=title,
        )
    elif chart_type == "treemap":
        figure = px.treemap(chart_data, path=[x], values=display_y, title=title)
    else:
        return None

    if display_y == "Share %" and chart_type in {"bar", "horizontal_bar"}:
        figure.update_layout(yaxis_title="Share %") if chart_type == "bar" else figure.update_layout(xaxis_title="Share %")

    return figure


def render_chart(data, chart, key):
    """Render one Plotly chart with an explicit Streamlit element key.

    Streamlit derives an internal element ID from a chart's parameters when no
    key is supplied. The same figure rendered in more than one place during a
    single rerun can therefore collide. Requiring a key here prevents that
    class of DuplicateElementId errors throughout the dashboard.

    Accessible reading also increases chart typography and interaction targets.
    This keeps the underlying data and chart choice unchanged while making the
    visual easier to read for low-vision and keyboard users.
    """
    figure = build_chart_figure(data, chart)
    if figure is not None:
        if st.session_state.get("display_mode") == "Accessible reading":
            figure.update_layout(
                font={"size": 16},
                title={"font": {"size": 22}},
                legend={"font": {"size": 15}},
                margin={"l": 70, "r": 35, "t": 75, "b": 70},
                hoverlabel={"font_size": 16},
            )
            figure.update_xaxes(
                title_font={"size": 17},
                tickfont={"size": 15},
                automargin=True,
                showgrid=True,
            )
            figure.update_yaxes(
                title_font={"size": 17},
                tickfont={"size": 15},
                automargin=True,
                showgrid=True,
            )
            figure.update_traces(
                marker_line_width=1.2,
                selector=dict(type="bar"),
            )

        st.plotly_chart(
            figure,
            use_container_width=True,
            key=key,
            config={
                "displaylogo": False,
                "responsive": True,
            },
        )

        # Accessible reading provides a concise textual equivalent of the chart
        # structure without duplicating the entire underlying table.
        if st.session_state.get("display_mode") == "Accessible reading":
            title = chart.get("title") or "Chart"
            chart_type = str(chart.get("type") or "chart").replace("_", " ")
            x = chart.get("x")
            y = chart.get("y")
            aggregation = chart.get("aggregation", "sum")
            parts = [f"{title} is a {chart_type} chart"]
            if y:
                parts.append(f"showing {aggregation} of {y}")
            if x:
                parts.append(f"by {x}")
            st.caption("Accessible chart description: " + " ".join(parts) + ".")



def customize_dashboard_plan(plan, data, profile, key_prefix="dashboard"):
    """Return a locally customized copy of the dashboard plan.

    All controls operate on in-memory data only. They do not call Workiva or Copilot.
    """
    if not plan:
        return plan

    customized = deepcopy(plan)
    numeric_options = list(profile.get("numeric", []))
    dimension_options = list(dict.fromkeys(
        profile.get("categories", []) + profile.get("dates", []) + numeric_options
    ))

    fingerprint = abs(hash(json.dumps(plan, sort_keys=True, default=str)))
    base_key = f"{key_prefix}_{fingerprint}"

    with st.expander("Customize dashboard", expanded=False):
        st.caption(
            "These controls update the visible charts locally. "
            "They do not send another request to Copilot or Workiva. "
            "Changes are applied together to avoid rerunning the app for every control."
        )

        with st.form(f"{base_key}_customizer_form", clear_on_submit=False):
            for index, chart in enumerate(customized.get("charts", [])):
                st.markdown(f"**Chart {index + 1}: {chart.get('title', 'Dashboard chart')}**")
    
                row1 = st.columns([1, 2, 2, 2])
                chart["visible"] = row1[0].checkbox(
                    "Visible",
                    value=chart.get("visible", True),
                    key=f"{base_key}_visible_{index}",
                )
    
                chart_types = [
                    "bar", "horizontal_bar", "line", "area", "donut",
                    "scatter", "histogram", "box", "treemap",
                ]
                current_type = chart.get("type", "bar")
                chart["type"] = row1[1].selectbox(
                    "Chart type",
                    chart_types,
                    index=(chart_types.index(current_type) if current_type in chart_types else 0),
                    key=f"{base_key}_type_{index}",
                )
    
                x_options = [None] + dimension_options
                current_x = chart.get("x")
                chart["x"] = row1[2].selectbox(
                    "Dimension / X",
                    x_options,
                    index=(x_options.index(current_x) if current_x in x_options else 0),
                    key=f"{base_key}_x_{index}",
                    format_func=lambda value: "None" if value is None else str(value),
                )
    
                y_options = [None] + numeric_options
                current_y = chart.get("y")
                chart["y"] = row1[3].selectbox(
                    "Metric / Y",
                    y_options,
                    index=(y_options.index(current_y) if current_y in y_options else 0),
                    key=f"{base_key}_y_{index}",
                    format_func=lambda value: "None" if value is None else str(value),
                )
    
                row2 = st.columns(4)
                aggregations = ["sum", "average", "minimum", "maximum", "count"]
                current_agg = chart.get("aggregation", "sum")
                chart["aggregation"] = row2[0].selectbox(
                    "Aggregation",
                    aggregations,
                    index=(aggregations.index(current_agg) if current_agg in aggregations else 0),
                    key=f"{base_key}_agg_{index}",
                )
    
                top_options = ["All", 5, 10, 15, 20]
                current_top = chart.get("top_n", "All")
                if current_top not in top_options:
                    current_top = "All"
                chart["top_n"] = row2[1].selectbox(
                    "Top N",
                    top_options,
                    index=top_options.index(current_top),
                    key=f"{base_key}_top_{index}",
                )
    
                sort_options = ["Descending", "Ascending", "Original"]
                current_sort = chart.get("sort_order", "Descending")
                chart["sort_order"] = row2[2].selectbox(
                    "Sort",
                    sort_options,
                    index=(sort_options.index(current_sort) if current_sort in sort_options else 0),
                    key=f"{base_key}_sort_{index}",
                )
    
                chart["show_percentage"] = row2[3].checkbox(
                    "Show %",
                    value=bool(chart.get("show_percentage", False)),
                    key=f"{base_key}_pct_{index}",
                    help="For categorical bar, horizontal bar, donut and treemap charts, show each category as a share of the displayed total.",
                )
    
                if index < len(customized.get("charts", [])) - 1:
                    st.divider()

            st.form_submit_button("Apply dashboard changes", use_container_width=True)

    return customized


# ============================================================
# SMART DASHBOARD LAYOUT HELPERS
# ============================================================

def _compact_value(value):
    """Format management values without changing the underlying calculation."""
    if value is None or pd.isna(value):
        return "n/a"
    value = float(value)
    absolute = abs(value)
    if absolute >= 1_000_000_000:
        return f"{value / 1_000_000_000:,.1f}B"
    if absolute >= 1_000_000:
        return f"{value / 1_000_000:,.1f}M"
    if absolute >= 1_000:
        return f"{value / 1_000:,.1f}K"
    return f"{value:,.2f}"


def build_key_management_signals(summary, comparison=None):
    """Return concise management signals using Python-calculated facts only."""
    context = summary.get("analytical_context", {}) or {}
    financial = context.get("financial_variance", {}) or {}
    trends = context.get("trends", {}) or {}
    breakdowns = context.get("top_breakdowns", {}) or {}
    quality = context.get("data_quality", {}) or {}
    signals = []

    actual_budget = financial.get("actual_vs_budget")
    if actual_budget:
        variance = actual_budget.get("variance")
        pct = actual_budget.get("variance_pct")
        actual_name = actual_budget.get("actual_metric", "Actual")
        budget_name = actual_budget.get("budget_metric", "Budget")
        if variance is not None:
            direction = "above" if variance >= 0 else "below"
            pct_text = f" ({abs(pct):.1f}%)" if pct is not None else ""
            signals.append(
                f"**{actual_name} vs {budget_name}:** {_compact_value(abs(variance))} "
                f"{direction} plan{pct_text}."
            )

    revenue_costs = financial.get("revenue_vs_costs")
    if revenue_costs:
        spread = revenue_costs.get("spread")
        spread_pct = revenue_costs.get("spread_pct_of_revenue")
        if spread is not None:
            signals.append(
                "**Revenue less costs:** "
                f"{_compact_value(spread)}"
                + (f" ({spread_pct:.1f}% of revenue)." if spread_pct is not None else ".")
            )

    # Largest absolute trend movement.
    trend_candidates = []
    for metric, details in trends.items():
        pct = details.get("change_pct")
        if pct is not None:
            trend_candidates.append((abs(pct), metric, pct))
    if trend_candidates:
        _, metric, pct = max(trend_candidates)
        direction = "increased" if pct >= 0 else "decreased"
        signals.append(f"**Largest trend movement:** {metric} {direction} {abs(pct):.1f}% across the observed period.")

    # Concentration signal from the strongest available category breakdown.
    for key, details in breakdowns.items():
        top = details.get("top", [])
        if len(top) >= 2:
            values = [float(item.get("value", 0) or 0) for item in top]
            positive_total = sum(v for v in values if v > 0)
            if positive_total > 0:
                top_two = sum(sorted((v for v in values if v > 0), reverse=True)[:2])
                share = top_two / positive_total * 100
                signals.append(f"**Contribution concentration:** the top two visible contributors in {key} represent {share:.1f}% of the top-five positive contribution.")
            break

    if comparison is not None:
        metrics = comparison.get("metrics")
        if metrics is not None and not metrics.empty and "Change %" in metrics.columns:
            comparable = metrics.dropna(subset=["Change %"]).copy()
            if not comparable.empty:
                idx = comparable["Change %"].abs().idxmax()
                row = comparable.loc[idx]
                signals.append(
                    f"**Quarter comparison:** {row['Metric']} moved {row['Change %']:+.1f}% "
                    f"from {comparison['first_quarter']} to {comparison['second_quarter']}."
                )

    missing_rate = quality.get("missing_rate_pct")
    if missing_rate is not None and missing_rate > 0:
        signals.append(f"**Data quality:** {missing_rate:.1f}% of cells are missing; interpret affected measures with care.")

    if not signals:
        signals.append("No material variance, trend, concentration or data-quality signal was detected from the available calculated facts.")

    return signals[:5]


def build_management_pulse(summary, forecast_result=None):
    """Return a compact, deterministic management pulse for the current state.

    No AI call is made here. The pulse only prioritizes already-calculated facts,
    so it can be rendered on every rerun without latency or hallucination risk.
    """
    context = summary.get("analytical_context", {}) or {}
    trends = context.get("trends", {}) or {}
    quality = context.get("data_quality", {}) or {}
    financial = context.get("financial_variance", {}) or {}

    trend_candidates = []
    for metric, details in trends.items():
        pct = details.get("change_pct")
        if pct is not None:
            trend_candidates.append((abs(float(pct)), metric, float(pct)))

    if trend_candidates:
        _, metric, pct = max(trend_candidates, key=lambda item: item[0])
        trend_text = f"{metric}: {pct:+.1f}% across the observed period"
        trend_question = f"Explain the {metric} trend and what management should watch next."
    else:
        trend_text = "No multi-period trend is available for this quarter."
        trend_question = "What are the most important management signals in the current data?"

    actual_budget = financial.get("actual_vs_budget") or {}
    if actual_budget and actual_budget.get("variance_pct") is not None:
        variance_text = (
            f"{actual_budget.get('actual_metric', 'Actual')} vs "
            f"{actual_budget.get('budget_metric', 'Budget')}: "
            f"{float(actual_budget['variance_pct']):+.1f}%"
        )
    else:
        missing = quality.get("missing_rate_pct")
        variance_text = (
            f"Data completeness: {100 - float(missing):.1f}%"
            if missing is not None
            else "No material variance signal calculated."
        )

    forecast_text = "Forecast board not built yet."
    forecast_question = None
    if forecast_result:
        horizon = forecast_result.get("horizon")
        quarters = forecast_result.get("quarters", []) or []
        metrics = forecast_result.get("metrics", []) or []
        confidence_labels = [
            str(details.get("confidence"))
            for details in (forecast_result.get("backtests", {}) or {}).values()
            if details and details.get("confidence")
        ]
        confidence_text = ""
        if confidence_labels:
            confidence_text = f" Confidence: {', '.join(sorted(set(confidence_labels)))}."
        forecast_text = (
            f"Forecast ready: {len(quarters)} historical quarter"
            f"{'s' if len(quarters) != 1 else ''}, {horizon} quarter"
            f"{'s' if horizon != 1 else ''} ahead, {len(metrics)} metric"
            f"{'s' if len(metrics) != 1 else ''}.{confidence_text}"
        )
        forecast_question = (
            "Summarize the forecast outlook, distinguishing historical actuals "
            "from projected values and highlighting the largest projected movement."
        )

    return {
        "trend": trend_text,
        "variance_or_quality": variance_text,
        "forecast": forecast_text,
        "trend_question": trend_question,
        "forecast_question": forecast_question,
    }


def build_contextual_ai_questions(summary, forecast_result=None, previous_quarter=None, selected_quarter=None):
    """Create high-value Ask AI shortcuts from the current calculated state."""
    pulse = build_management_pulse(summary, forecast_result)
    questions = [
        ("Explain strongest trend", pulse["trend_question"]),
        (
            "Prioritize watch items",
            "Rank the most important management watch items in the current dashboard. "
            "Use only supplied calculations and clearly separate facts from interpretation.",
        ),
    ]

    driver_analysis = (summary.get("analytical_context", {}).get("driver_analysis", {}) or {})
    driver_metrics = driver_analysis.get("metrics", {}) or {}
    if driver_metrics:
        first_metric = next(iter(driver_metrics))
        first_dimensions = driver_metrics.get(first_metric, {}) or {}
        if first_dimensions:
            first_dimension = next(iter(first_dimensions))
            questions.insert(1, (
                "Explain key drivers",
                f"Explain the calculated drivers of the change in {first_metric} by {first_dimension}. "
                "Separate mathematical contribution from causal interpretation and do not invent causes.",
            ))

    if pulse.get("forecast_question"):
        questions.insert(1, ("Explain forecast", pulse["forecast_question"]))
    elif previous_quarter and selected_quarter:
        questions.insert(1, (
            "Compare prior quarter",
            f"Compare {selected_quarter} and {previous_quarter}. Focus on the most material "
            "KPI and category movements and explain why they matter for management.",
        ))

    return questions[:3]


def build_dashboard_panels(summary):
    """Build fast, deterministic layout panels from pre-calculated facts."""
    context = summary.get("analytical_context", {}) or {}
    financial = context.get("financial_variance", {}) or {}
    trends = context.get("trends", {}) or {}
    breakdowns = context.get("top_breakdowns", {}) or {}
    quality = context.get("data_quality", {}) or {}

    variance_items = []
    risk_items = []
    questions = []

    actual_budget = financial.get("actual_vs_budget")
    if actual_budget:
        variance = actual_budget.get("variance")
        variance_pct = actual_budget.get("variance_pct")
        actual_name = actual_budget.get("actual_metric", "Actual")
        budget_name = actual_budget.get("budget_metric", "Budget")
        direction = "above" if (variance or 0) >= 0 else "below"
        pct_text = (
            f" ({abs(variance_pct):.1f}%)"
            if variance_pct is not None
            else ""
        )
        variance_items.append(
            f"**{actual_name} vs {budget_name}:** "
            f"{_compact_value(abs(variance))} {direction} plan{pct_text}."
        )
        if variance is not None and variance < 0:
            risk_items.append(
                f"{actual_name} is below {budget_name} by "
                f"{_compact_value(abs(variance))}{pct_text}."
            )
        questions.append(
            f"What are the main drivers of the {actual_name} versus {budget_name} variance?"
        )

    revenue_costs = financial.get("revenue_vs_costs")
    if revenue_costs:
        spread = revenue_costs.get("spread")
        spread_pct = revenue_costs.get("spread_pct_of_revenue")
        variance_items.append(
            "**Revenue less costs:** "
            f"{_compact_value(spread)}"
            + (
                f" ({spread_pct:.1f}% of revenue)."
                if spread_pct is not None
                else "."
            )
        )
        if spread is not None and spread < 0:
            risk_items.append(
                "Costs exceed revenue by "
                f"{_compact_value(abs(spread))}."
            )

    strongest_negative = None
    for metric, details in trends.items():
        pct = details.get("change_pct")
        if pct is None or pct >= 0:
            continue
        if strongest_negative is None or pct < strongest_negative[1]:
            strongest_negative = (metric, pct, details)

    if strongest_negative:
        metric, pct, details = strongest_negative
        risk_items.append(
            f"{metric} declined {abs(pct):.1f}% from the first to the latest observed period."
        )
        questions.append(
            f"What explains the decline in {metric}, and is it concentrated in a specific category?"
        )

    for breakdown_name, details in breakdowns.items():
        top = details.get("top", [])
        if not top:
            continue
        total = sum(abs(float(item.get("value", 0) or 0)) for item in top)
        lead_value = abs(float(top[0].get("value", 0) or 0))
        share = (lead_value / total * 100) if total else 0
        if share >= 50:
            risk_items.append(
                f"Concentration watch: {top[0].get('category')} represents about "
                f"{share:.0f}% of the top-five total for {breakdown_name}."
            )
        questions.append(
            f"Why is {top[0].get('category')} the leading contributor to {breakdown_name}?"
        )
        break

    missing_rate = quality.get("missing_rate_pct")
    if missing_rate is not None and missing_rate > 5:
        risk_items.append(
            f"Data quality watch: {missing_rate:.1f}% of cells are missing."
        )

    if not variance_items:
        variance_items.append(
            "No supported budget/variance relationship was detected in the current dataset."
        )

    if not risk_items:
        risk_items.append(
            "No material risk signal was detected by the deterministic checks currently available."
        )

    fallback_questions = [
        "Which categories are contributing most to the current result?",
        "Which movement deserves management attention first?",
        "What should management investigate before the next reporting cycle?",
    ]
    for item in fallback_questions:
        if len(questions) >= 3:
            break
        if item not in questions:
            questions.append(item)

    return {
        "variance": variance_items[:3],
        "risks": risk_items[:3],
        "questions": questions[:3],
    }


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
    ttl=1800,
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


@st.cache_data(ttl=1800, show_spinner=False)
def quarter_bundle(quarter):

    year_match = re.search(r"\b(20\d{2})\b", str(quarter))
    reporting_year = int(year_match.group(1)) if year_match else None

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
            raw,
            default_year=reporting_year,
        )
    )

    data = (
        prepare_dataframe(
            raw,
            profile,
            default_year=reporting_year,
        )
    )

    business = (
        detect_business_columns(
            data
        )
    )

    # Add deterministic finance measures once per cached quarter bundle.
    # This is vectorized pandas work and does not call Copilot.
    data, derived_measures = add_financial_intelligence(
        data,
        business,
    )

    if derived_measures:
        profile = profile_dataframe(data, default_year=reporting_year)
        data = prepare_dataframe(data, profile, default_year=reporting_year)
        business = detect_business_columns(data)

    business["derived_measures"] = derived_measures

    summary = (
        build_data_summary(
            data,
            profile,
            quarter,
        )
    )
    summary["financial_intelligence"] = {
        "detected_concepts": {
            key: value
            for key, value in business.items()
            if key != "derived_measures" and value is not None
        },
        "derived_measures": derived_measures,
    }

    summary["analytical_context"] = build_analytical_context(
        data,
        profile,
        business,
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
# MULTI-QUARTER FORECASTING
# ============================================================

def quarter_period(value):
    """Return a pandas quarterly Period for labels such as Q1 2026."""
    text = str(value).upper().strip()
    match = re.search(r"\bQ([1-4])[\s\-_\/]*(20\d{2})\b", text)
    if not match:
        match = re.search(r"\b(20\d{2})[\s\-_\/]*Q([1-4])\b", text)
        if not match:
            return None
        year = int(match.group(1))
        quarter = int(match.group(2))
    else:
        quarter = int(match.group(1))
        year = int(match.group(2))
    return pd.Period(year=year, quarter=quarter, freq="Q")


def _forecast_aggregation(metric):
    """Use averages for rates/ratios and sums for additive measures."""
    lower = str(metric).strip().lower()
    average_tokens = ("%", "ratio", "margin", "rate", "average", "avg")
    return "average" if any(token in lower for token in average_tokens) else "sum"


def build_multi_quarter_history(bundles):
    """Create one quarterly observation per common numeric metric."""
    ordered = sorted(
        bundles,
        key=lambda item: quarter_period(item["quarter"]) or pd.Period("1900Q1", freq="Q"),
    )
    if not ordered:
        return pd.DataFrame(), []

    common_numeric = set(ordered[0]["profile"]["numeric"])
    for item in ordered[1:]:
        common_numeric &= set(item["profile"]["numeric"])

    # Keep the source-column order from the newest/first bundle rather than
    # alphabetizing finance measures into an unfamiliar order.
    common_numeric = [
        column
        for column in ordered[-1]["profile"]["numeric"]
        if column in common_numeric
    ]

    rows = []
    for item in ordered:
        row = {
            "Quarter": item["quarter"],
            "Quarter Period": quarter_period(item["quarter"]),
        }
        for metric in common_numeric:
            values = pd.to_numeric(item["data"][metric], errors="coerce").dropna()
            if values.empty:
                row[metric] = pd.NA
            elif _forecast_aggregation(metric) == "average":
                row[metric] = float(values.mean())
            else:
                row[metric] = float(values.sum())
        rows.append(row)

    history = pd.DataFrame(rows)
    if not history.empty:
        history = history.sort_values("Quarter Period").reset_index(drop=True)
    return history, common_numeric


def _linear_fit(values):
    """Fit a transparent least-squares line to an ordered numeric sequence."""
    values = [float(value) for value in values]
    n = len(values)
    if not values:
        return 0.0, 0.0
    if n == 1:
        return 0.0, values[0]

    x = list(range(n))
    x_mean = sum(x) / n
    y_mean = sum(values) / n
    denominator = sum((value - x_mean) ** 2 for value in x)
    slope = (
        sum((x[i] - x_mean) * (values[i] - y_mean) for i in range(n))
        / denominator
        if denominator
        else 0.0
    )
    intercept = y_mean - slope * x_mean
    return float(slope), float(intercept)


def backtest_quarter_forecast(history, metric):
    """Walk-forward backtest using only information available before each quarter.

    Each historical quarter after the first is predicted from preceding quarters.
    With one training point the model uses a flat baseline; with two or more it uses
    the same linear fit as the live forecast. This avoids look-ahead bias.
    """
    working = history[["Quarter", "Quarter Period", metric]].dropna().copy()
    if working.empty or len(working) < 2:
        return {
            "tests": pd.DataFrame(),
            "history_points": int(len(working)),
            "backtest_points": 0,
            "mae": None,
            "rmse": None,
            "mape_pct": None,
            "wape_pct": None,
            "smape_pct": None,
            "confidence": "Limited",
            "confidence_reason": "At least two historical quarters are needed for a backtest.",
        }

    values = [float(value) for value in working[metric].tolist()]
    rows = []

    for test_index in range(1, len(values)):
        train_values = values[:test_index]
        slope, intercept = _linear_fit(train_values)
        prediction = intercept + slope * test_index
        actual = values[test_index]
        error = prediction - actual
        abs_error = abs(error)
        ape = (abs_error / abs(actual) * 100) if actual else None
        smape_denominator = abs(actual) + abs(prediction)
        smape = (
            200 * abs_error / smape_denominator
            if smape_denominator
            else 0.0
        )
        rows.append({
            "Quarter": str(working.iloc[test_index]["Quarter"]),
            "Actual": float(actual),
            "Backtest Forecast": float(prediction),
            "Error": float(error),
            "Absolute Error": float(abs_error),
            "Absolute % Error": float(ape) if ape is not None else None,
            "sMAPE %": float(smape),
            "Training Quarters": int(test_index),
        })

    tests = pd.DataFrame(rows)
    abs_errors = tests["Absolute Error"]
    mae = float(abs_errors.mean())
    rmse = float((tests["Error"].pow(2).mean()) ** 0.5)
    nonzero_ape = tests["Absolute % Error"].dropna()
    mape = float(nonzero_ape.mean()) if not nonzero_ape.empty else None
    denominator = float(tests["Actual"].abs().sum())
    wape = float(abs_errors.sum() / denominator * 100) if denominator else None
    smape = float(tests["sMAPE %"].mean())

    # Four or fewer historical quarters are intentionally never labelled High.
    # A short sample can look accurate by chance, so confidence reflects both
    # observed error and the amount of evidence available.
    backtest_points = len(tests)
    if backtest_points < 2:
        confidence = "Limited"
        confidence_reason = (
            "Only one historical holdout can be tested, so forecast reliability cannot be established."
        )
    elif wape is None:
        confidence = "Limited"
        confidence_reason = "Relative forecast error cannot be calculated reliably for this series."
    elif backtest_points >= 3 and wape <= 20:
        confidence = "Moderate"
        confidence_reason = (
            f"Walk-forward WAPE is {wape:.1f}% across {backtest_points} historical holdouts. "
            "Confidence is capped at Moderate because the selected history contains at most four quarters."
        )
    elif backtest_points >= 2 and wape <= 10:
        confidence = "Moderate"
        confidence_reason = (
            f"Walk-forward WAPE is {wape:.1f}%, but only {backtest_points} historical holdouts are available. "
            "Confidence is capped at Moderate because the sample is short."
        )
    else:
        confidence = "Limited"
        confidence_reason = (
            f"Walk-forward WAPE is {wape:.1f}% across {backtest_points} historical holdout"
            f"{'s' if backtest_points != 1 else ''}; treat the projection as directional."
        )

    return {
        "tests": tests,
        "history_points": int(len(working)),
        "backtest_points": int(backtest_points),
        "mae": mae,
        "rmse": rmse,
        "mape_pct": mape,
        "wape_pct": wape,
        "smape_pct": smape,
        "confidence": confidence,
        "confidence_reason": confidence_reason,
    }


def linear_quarter_forecast(history, metric, horizon):
    """Simple least-squares quarterly trend projection with a flat 1-point fallback."""
    working = history[["Quarter", "Quarter Period", metric]].dropna().copy()
    if working.empty:
        return pd.DataFrame()

    values = [float(value) for value in working[metric].tolist()]
    n = len(values)
    slope, intercept = _linear_fit(values)

    rows = []
    for _, row in working.iterrows():
        rows.append({
            "Quarter": str(row["Quarter"]),
            "Quarter Period": row["Quarter Period"],
            "Value": float(row[metric]),
            "Series": "Actual",
        })

    last_period = working["Quarter Period"].iloc[-1]
    for step in range(1, int(horizon) + 1):
        period = last_period + step
        estimate = intercept + slope * (n - 1 + step)
        rows.append({
            "Quarter": f"Q{period.quarter} {period.year}",
            "Quarter Period": period,
            "Value": float(estimate),
            "Series": "Forecast",
        })

    return pd.DataFrame(rows)


def forecast_trend_commentary(history, metric, forecast_frame):
    """Deterministic trend commentary so the board is useful without an AI call."""
    actual = forecast_frame[forecast_frame["Series"] == "Actual"]
    projected = forecast_frame[forecast_frame["Series"] == "Forecast"]
    if actual.empty or projected.empty:
        return "There is not enough usable data to describe this forecast."

    latest = float(actual["Value"].iloc[-1])
    future = float(projected["Value"].iloc[-1])
    change = future - latest
    change_pct = (change / abs(latest) * 100) if latest else None

    if len(actual) == 1:
        return (
            "Only one historical quarter is selected, so the projection uses the "
            "latest value as a flat baseline. Add another quarter to estimate a trend."
        )

    first = float(actual["Value"].iloc[0])
    historical_change = latest - first
    direction = "upward" if historical_change > 0 else "downward" if historical_change < 0 else "flat"

    actual_values = actual["Value"].tolist()
    deltas = [actual_values[i] - actual_values[i - 1] for i in range(1, len(actual_values))]
    reversals = sum(
        1 for i in range(1, len(deltas))
        if deltas[i] != 0 and deltas[i - 1] != 0 and (deltas[i] > 0) != (deltas[i - 1] > 0)
    )
    stability = "volatile" if reversals else "directionally consistent"

    projected_text = f"{change:+,.2f}"
    if change_pct is not None:
        projected_text += f" ({change_pct:+.1f}%)"

    return (
        f"The selected history is {direction} and {stability}. "
        f"The simple trend projection moves {metric} by {projected_text} from the latest "
        "actual quarter to the end of the forecast horizon. This is an extrapolation of "
        "the selected quarterly pattern, not a causal or scenario-based forecast."
    )


def build_forecast_scenarios(forecast_frame, sensitivity_pct):
    """Build transparent baseline/upside/downside cases from the baseline forecast.

    Scenario values are deterministic. For forecast step n, upside/downside apply a
    cumulative +/- sensitivity adjustment to the baseline value. Historical actuals
    are unchanged. "Upside" means a numerically higher value, which is not
    necessarily economically favourable for cost/expense measures.
    """
    if forecast_frame is None or forecast_frame.empty:
        return pd.DataFrame()

    sensitivity = max(float(sensitivity_pct), 0.0) / 100.0
    actual = forecast_frame[forecast_frame["Series"] == "Actual"].copy()
    projected = forecast_frame[forecast_frame["Series"] == "Forecast"].copy()

    rows = []
    for _, row in actual.iterrows():
        rows.append({
            "Quarter": row["Quarter"],
            "Quarter Period": row["Quarter Period"],
            "Value": float(row["Value"]),
            "Scenario": "Actual",
        })

    for step, (_, row) in enumerate(projected.iterrows(), start=1):
        baseline = float(row["Value"])
        rows.extend([
            {
                "Quarter": row["Quarter"],
                "Quarter Period": row["Quarter Period"],
                "Value": baseline,
                "Scenario": "Baseline",
            },
            {
                "Quarter": row["Quarter"],
                "Quarter Period": row["Quarter Period"],
                "Value": baseline * ((1.0 + sensitivity) ** step),
                "Scenario": "Upside",
            },
            {
                "Quarter": row["Quarter"],
                "Quarter Period": row["Quarter Period"],
                "Value": baseline * ((1.0 - sensitivity) ** step),
                "Scenario": "Downside",
            },
        ])

    return pd.DataFrame(rows)


def build_forecast_result(bundles, selected_metrics, horizon, scenario_sensitivity_pct=5.0):
    history, common_metrics = build_multi_quarter_history(bundles)
    metrics = [metric for metric in selected_metrics if metric in common_metrics]
    forecasts = {}
    commentary = {}
    backtests = {}
    scenarios = {}
    for metric in metrics:
        frame = linear_quarter_forecast(history, metric, horizon)
        if frame.empty:
            continue
        forecasts[metric] = frame
        commentary[metric] = forecast_trend_commentary(history, metric, frame)
        backtests[metric] = backtest_quarter_forecast(history, metric)
        scenarios[metric] = build_forecast_scenarios(frame, scenario_sensitivity_pct)
    return {
        "quarters": [item["quarter"] for item in bundles],
        "history": history,
        "common_metrics": common_metrics,
        "metrics": metrics,
        "horizon": int(horizon),
        "forecasts": forecasts,
        "commentary": commentary,
        "backtests": backtests,
        "scenarios": scenarios,
        "scenario_sensitivity_pct": float(scenario_sensitivity_pct),
    }


def generate_forecast_ai_commentary(forecast_result):
    payload = {
        "historical_quarters": forecast_result.get("quarters", []),
        "forecast_horizon_quarters": forecast_result.get("horizon", 0),
        "scenario_sensitivity_pct": forecast_result.get("scenario_sensitivity_pct", 0),
        "scenario_definition": (
            "Baseline is the trend forecast. Upside is a numerically higher case and Downside "
            "a numerically lower case using the explicit per-quarter sensitivity assumption; "
            "for costs/expenses, a numerically higher Upside is not necessarily favourable."
        ),
        "metrics": {},
    }
    for metric, frame in forecast_result.get("forecasts", {}).items():
        backtest = (forecast_result.get("backtests", {}) or {}).get(metric, {}) or {}
        tests = backtest.get("tests")
        scenario_frame = (forecast_result.get("scenarios", {}) or {}).get(metric)
        payload["metrics"][metric] = {
            "series": frame[["Quarter", "Value", "Series"]].to_dict(orient="records"),
            "scenarios": (
                scenario_frame[["Quarter", "Value", "Scenario"]].to_dict(orient="records")
                if isinstance(scenario_frame, pd.DataFrame) and not scenario_frame.empty
                else []
            ),
            "backtest": {
                "history_points": backtest.get("history_points"),
                "backtest_points": backtest.get("backtest_points"),
                "mae": backtest.get("mae"),
                "rmse": backtest.get("rmse"),
                "mape_pct": backtest.get("mape_pct"),
                "wape_pct": backtest.get("wape_pct"),
                "smape_pct": backtest.get("smape_pct"),
                "confidence": backtest.get("confidence"),
                "confidence_reason": backtest.get("confidence_reason"),
                "tests": (
                    tests.to_dict(orient="records")
                    if isinstance(tests, pd.DataFrame) and not tests.empty
                    else []
                ),
            },
        }

    prompt = f"""
You are an executive FP&A analyst reviewing a short quarterly forecasting dashboard.

Forecast dataset:
{payload}

Rules:
- Clearly separate historical actuals from projected values.
- The projection is a simple linear extrapolation from at most four historical quarters.
- Do not invent causes, seasonality, external drivers, management actions already taken, or confidence intervals.
- Point out direction, acceleration/deceleration only if the supplied sequence supports it, and material risks/opportunities.
- Use the supplied walk-forward backtest results when discussing forecast reliability.
- Never describe confidence as higher than the supplied confidence label.
- If history is too short, backtest error is high, or the supplied label is Limited, explicitly say reliability is limited.
- Backtest metrics are historical diagnostics, not guarantees of future accuracy.
- Scenario values are deterministic sensitivity cases, not separate predictive models or probabilities.
- Never describe Upside/Downside as likely outcomes or assign probabilities unless supplied.
- For costs and expenses, a numerically higher Upside case is not automatically favourable.
- Keep the response concise and management-oriented.

Write sections:
1. Forecast takeaway
2. Trend signals
3. Risks / watch items
4. Management questions
"""
    return grounded_ai_response(
        prompt,
        payload,
        purpose="forecast commentary",
    )


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


    evidence = {
        "first_quarter": comparison.get("first_quarter"),
        "second_quarter": comparison.get("second_quarter"),
        "metric_comparison": metrics,
        "focus_metric": comparison.get("focus_metric"),
        "category_comparison": breakdown,
    }

    return grounded_ai_response(
        prompt,
        evidence,
        purpose="quarter comparison commentary",
    )


def render_comparison(
    comparison,
    key_prefix="comparison",
):

    first_key = re.sub(r"[^A-Za-z0-9_-]+", "_", str(comparison["first_quarter"]))
    second_key = re.sub(r"[^A-Za-z0-9_-]+", "_", str(comparison["second_quarter"]))
    comparison_key = f"{key_prefix}_{first_key}_{second_key}"

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
        key=f"{comparison_key}_metrics",
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
            key=f"{comparison_key}_pct_change",
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
            key=f"{comparison_key}_breakdown",
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
# EXPORT HELPERS
# ============================================================

def _figure_png_bytes(figure, width=1200, height=650):
    if figure is None:
        return None
    try:
        return figure.to_image(format="png", width=width, height=height, scale=1.4)
    except Exception:
        return None


def _dashboard_kpi_rows(data, plan):
    rows = []
    for item in plan.get("kpis", [])[:6]:
        value = aggregate_value(data[item["column"]], item["aggregation"])
        rows.append([item["label"], f"{value:,.2f}"])
    return rows


# ============================================================
# PDF EXPORT
# ============================================================

def _report_sections(summary):
    """Reuse the same deterministic management-focus facts in exports."""
    panels = build_dashboard_panels(summary)
    return {
        "risks": panels.get("risks", [])[:4],
        "questions": panels.get("questions", [])[:4],
        "variance": panels.get("variance", [])[:4],
    }


def build_pdf_export(
    quarter,
    source,
    summary,
    management_summary,
    data,
    plan,
):
    """Build a board-pack style PDF from the current dashboard.

    The export is generated only when the user clicks Prepare PDF, so this
    does not slow down normal dashboard navigation.
    """
    buffer = BytesIO()
    document = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )
    styles = getSampleStyleSheet()
    sections = _report_sections(summary)

    story = [
        Paragraph(f"{plan.get('title', 'Workiva AI Dashboard')} - {quarter}", styles["Title"]),
        Spacer(1, 8),
        Paragraph(
            "Executive board-pack generated from the current dashboard and AI summary.",
            styles["BodyText"],
        ),
        Spacer(1, 8),
        Paragraph(
            f"<b>Spreadsheet:</b> {source['spreadsheet_name']}<br/>"
            f"<b>Sheet:</b> {source['sheet_name']}<br/>"
            f"<b>Match confidence:</b> {source.get('confidence', 'Unknown')}",
            styles["BodyText"],
        ),
        PageBreak(),
        Paragraph("1. Executive summary", styles["Heading1"]),
        Paragraph(str(management_summary).replace("\n", "<br/>"), styles["BodyText"]),
        PageBreak(),
        Paragraph("2. KPI overview", styles["Heading1"]),
    ]

    kpi_rows = [["KPI", "Value"]] + _dashboard_kpi_rows(data, plan)
    if len(kpi_rows) == 1:
        kpi_rows.append(["No KPI available", ""])
    kpi_table = Table(kpi_rows, colWidths=[110 * mm, 60 * mm], repeatRows=1)
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("PADDING", (0, 0), (-1, -1), 7),
    ]))
    story += [kpi_table, Spacer(1, 12)]

    if sections["variance"]:
        story.append(Paragraph("Variance / performance notes", styles["Heading2"]))
        for item in sections["variance"]:
            story.append(Paragraph(f"• {item}".replace("**", ""), styles["BodyText"]))
        story.append(Spacer(1, 8))

    chart_images = []
    for chart in plan.get("charts", [])[:6]:
        png = _figure_png_bytes(build_chart_figure(data, chart), width=1200, height=650)
        if png:
            chart_images.append((chart, png))

    story += [PageBreak(), Paragraph("3. Dashboard visuals", styles["Heading1"])]
    if not chart_images:
        story.append(Paragraph(
            "Chart images were unavailable. Install the 'kaleido' package to enable Plotly image export.",
            styles["BodyText"],
        ))
    else:
        for index, (chart, png) in enumerate(chart_images):
            label = "Primary chart" if index == 0 else "Supporting visual"
            story.append(Paragraph(f"{label}: {chart.get('title', 'Dashboard chart')}", styles["Heading2"]))
            story.append(Image(BytesIO(png), width=250 * mm, height=135 * mm))
            story.append(Spacer(1, 8))
            if index < len(chart_images) - 1:
                story.append(PageBreak())

    story += [
        PageBreak(),
        Paragraph("4. Risks and watch items", styles["Heading1"]),
    ]
    for item in sections["risks"]:
        story.append(Paragraph(f"• {item}".replace("**", ""), styles["BodyText"]))

    story += [
        Spacer(1, 14),
        Paragraph("5. Management questions / actions", styles["Heading1"]),
    ]
    for item in sections["questions"]:
        story.append(Paragraph(f"• {item}".replace("**", ""), styles["BodyText"]))

    story += [
        Spacer(1, 14),
        Paragraph("Source note", styles["Heading2"]),
        Paragraph(
            f"This report is based on {summary['rows']} rows and {len(summary['columns'])} columns from the selected Workiva sheet. "
            "Detailed data-quality information remains available in the app's Data quality tab.",
            styles["BodyText"],
        ),
    ]

    document.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
# POWERPOINT EXPORT
# ============================================================

def _add_bullets(slide, title, items, left=0.8, top=1.35, width=11.8, height=5.2):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = title
    tf.paragraphs[0].font.size = Pt(19)
    tf.paragraphs[0].font.bold = True
    for item in items:
        p = tf.add_paragraph()
        p.text = str(item).replace("**", "")
        p.level = 0
        p.font.size = Pt(15)
    return box


def build_powerpoint_export(
    quarter,
    source,
    summary,
    management_summary,
    data,
    plan,
):
    """Build a board-pack style PowerPoint from the active dashboard.

    The deck is generated lazily only when requested, keeping the app fast.
    """
    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    sections = _report_sections(summary)

    # Slide 1 — title / quarter
    slide = presentation.slides.add_slide(presentation.slide_layouts[0])
    slide.shapes.title.text = f"{plan.get('title', 'Workiva AI Dashboard')}"
    slide.placeholders[1].text = (
        f"{quarter}\n"
        f"{source['spreadsheet_name']} → {source['sheet_name']}\n"
        f"AI-generated executive board pack"
    )

    # Slide 2 — executive summary
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
    slide.shapes.title.text = "Executive summary"
    box = slide.shapes.add_textbox(Inches(0.7), Inches(1.2), Inches(12.0), Inches(5.8))
    box.text_frame.word_wrap = True
    box.text_frame.text = str(management_summary)
    for paragraph in box.text_frame.paragraphs:
        paragraph.font.size = Pt(14)

    # Slide 3 — KPI overview
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
    slide.shapes.title.text = "KPI overview"
    kpis = _dashboard_kpi_rows(data, plan)
    if not kpis:
        _add_bullets(slide, "KPI status", ["No KPI fields were selected for this dashboard."])
    else:
        for index, (label, value) in enumerate(kpis[:6]):
            row = index // 3
            col = index % 3
            box = slide.shapes.add_textbox(
                Inches(0.65 + col * 4.2),
                Inches(1.25 + row * 1.25),
                Inches(3.85),
                Inches(0.95),
            )
            tf = box.text_frame
            tf.text = f"{label}\n{value}"
            tf.paragraphs[0].font.size = Pt(16)
            tf.paragraphs[0].font.bold = True
            if len(tf.paragraphs) > 1:
                tf.paragraphs[1].font.size = Pt(22)
    if sections["variance"]:
        _add_bullets(slide, "Variance / performance notes", sections["variance"][:3], top=4.05, height=2.4)

    # Slides 4–5+ — dashboard visuals
    chart_images = []
    for chart in plan.get("charts", [])[:6]:
        png = _figure_png_bytes(build_chart_figure(data, chart), width=1200, height=650)
        if png:
            chart_images.append((chart, png))

    if chart_images:
        primary_chart, primary_png = chart_images[0]
        slide = presentation.slides.add_slide(presentation.slide_layouts[5])
        slide.shapes.title.text = "Primary dashboard visual"
        caption = slide.shapes.add_textbox(Inches(0.75), Inches(1.0), Inches(11.8), Inches(0.35))
        caption.text_frame.text = str(primary_chart.get("title", "Primary chart"))
        caption.text_frame.paragraphs[0].font.size = Pt(14)
        slide.shapes.add_picture(BytesIO(primary_png), Inches(0.75), Inches(1.45), width=Inches(11.9), height=Inches(5.45))

        supporting = chart_images[1:]
        for index in range(0, len(supporting), 2):
            slide = presentation.slides.add_slide(presentation.slide_layouts[5])
            slide.shapes.title.text = "Supporting dashboard visuals"
            batch = supporting[index:index + 2]
            for slot, (chart, png) in enumerate(batch):
                y = 1.15 + slot * 3.05
                caption = slide.shapes.add_textbox(Inches(0.65), Inches(y), Inches(12.0), Inches(0.35))
                caption.text_frame.text = str(chart.get("title", "Supporting chart"))
                caption.text_frame.paragraphs[0].font.size = Pt(14)
                slide.shapes.add_picture(BytesIO(png), Inches(0.75), Inches(y + 0.38), width=Inches(11.85), height=Inches(2.45))
    else:
        slide = presentation.slides.add_slide(presentation.slide_layouts[5])
        slide.shapes.title.text = "Dashboard visuals"
        _add_bullets(slide, "Chart export note", ["Chart images were unavailable. Install 'kaleido' to enable Plotly image export."])

    # Slide — risks / opportunities
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
    slide.shapes.title.text = "Risks and watch items"
    _add_bullets(slide, "Management watch list", sections["risks"] or ["No deterministic risk signal was detected."], top=1.2)

    # Slide — management questions / actions
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
    slide.shapes.title.text = "Management questions and actions"
    _add_bullets(slide, "Suggested discussion points", sections["questions"] or ["Review the primary chart and KPI movements with the reporting owner."], top=1.2)

    # Final provenance slide, kept concise.
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])
    slide.shapes.title.text = "Source and data note"
    _add_bullets(
        slide,
        "Report basis",
        [
            f"Quarter: {quarter}",
            f"Rows analysed: {summary['rows']}",
            f"Columns analysed: {len(summary['columns'])}",
            f"Source confidence: {source.get('confidence', 'Unknown')}",
            f"Spreadsheet: {source['spreadsheet_name']}",
            f"Sheet: {source['sheet_name']}",
            "Detailed data quality remains available inside the app.",
        ],
        top=1.2,
    )

    output = BytesIO()
    presentation.save(output)
    output.seek(0)
    return output.getvalue()


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

# Subtle application-wide accessibility and responsive refinements.
# These preserve the existing visual language instead of introducing a
# separate, visually heavier accessibility skin.
st.markdown(
    """
    <style>
    /* Respect the operating system's reduced-motion preference in every mode. */
    @media (prefers-reduced-motion: reduce) {
        *, *::before, *::after {
            scroll-behavior: auto !important;
            animation-duration: 0.001ms !important;
            animation-iteration-count: 1 !important;
            transition-duration: 0.001ms !important;
        }
    }

    /* Clear keyboard focus without changing the sophisticated neutral design. */
    button:focus-visible,
    [role="button"]:focus-visible,
    [role="tab"]:focus-visible,
    input:focus-visible,
    textarea:focus-visible,
    a:focus-visible {
        outline: 2px solid currentColor !important;
        outline-offset: 3px !important;
    }

    /* Tabs remain usable on narrower screens rather than compressing labels. */
    [data-testid="stTabs"] [role="tablist"] {
        overflow-x: auto;
        scrollbar-width: thin;
        gap: 1.15rem !important;
        align-items: center;
    }
    [data-testid="stTabs"] [role="tab"] {
        white-space: nowrap;
        min-height: 42px;
        flex: 0 0 auto !important;
        padding-left: 0.25rem !important;
        padding-right: 0.25rem !important;
    }

    /* Preserve the spacing on compact screens; horizontal scrolling is
       preferable to squeezing navigation labels together. */
    @media (max-width: 900px) {
        [data-testid="stTabs"] [role="tablist"] {
            gap: 0.9rem !important;
        }
    }

    /* Comfortable touch targets while keeping controls visually compact. */
    button, [role="button"], [data-baseweb="select"] > div {
        min-height: 40px;
    }

    /* Prevent overly wide prose while letting analytical visuals stay expansive. */
    [data-testid="stAlert"] p,
    [data-testid="stCaptionContainer"] p {
        line-height: 1.5;
    }
    </style>
    """,
    unsafe_allow_html=True,
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


    board_mode = "AI designed"


    st.divider()

    st.subheader(
        "Display & accessibility"
    )

    display_mode = st.selectbox(
        "Display mode",
        [
            "Standard",
            "Focus",
            "Accessible reading",
        ],
        index=0,
        help=(
            "Standard keeps the full dashboard. "
            "Focus reduces visual distractions and collapses secondary detail. "
            "Accessible reading uses larger text and a single-column layout."
        ),
        key="display_mode",
    )


focus_mode = (
    display_mode == "Focus"
)

accessible_mode = (
    display_mode == "Accessible reading"
)


if accessible_mode:

    st.markdown(
        """
        <style>
        /* Accessible reading: larger type, generous spacing and strong focus. */
        html {
            font-size: 18px !important;
            scroll-behavior: auto !important;
        }

        body, [data-testid="stAppViewContainer"], [data-testid="stSidebar"] {
            font-size: 1rem !important;
        }

        p, li, label, [data-testid="stCaptionContainer"] {
            line-height: 1.7 !important;
            letter-spacing: 0.01em !important;
        }

        h1 { font-size: 2.25rem !important; line-height: 1.2 !important; }
        h2 { font-size: 1.75rem !important; line-height: 1.25 !important; }
        h3 { font-size: 1.45rem !important; line-height: 1.3 !important; }

        /* Larger pointer/touch targets. */
        button,
        [role="button"],
        [role="tab"],
        [data-baseweb="select"] > div,
        input,
        textarea {
            min-height: 48px !important;
            font-size: 1rem !important;
        }

        textarea {
            line-height: 1.6 !important;
        }

        /* Make keyboard focus impossible to miss. */
        button:focus-visible,
        [role="button"]:focus-visible,
        [role="tab"]:focus-visible,
        input:focus-visible,
        textarea:focus-visible,
        [data-baseweb="select"] *:focus-visible,
        a:focus-visible {
            outline: 3px solid currentColor !important;
            outline-offset: 3px !important;
            box-shadow: 0 0 0 2px Canvas !important;
        }

        /* Tabs need a clear selected state that does not rely on colour alone. */
        [data-testid="stTabs"] [role="tablist"] {
            gap: 1rem !important;
        }

        [role="tab"] {
            padding: 0.75rem 0.35rem !important;
            font-weight: 600 !important;
            flex: 0 0 auto !important;
        }

        [role="tab"][aria-selected="true"] {
            font-weight: 800 !important;
            text-decoration: underline !important;
            text-decoration-thickness: 3px !important;
            text-underline-offset: 0.35rem !important;
        }

        /* Metrics, alerts and expanders become easier to scan. */
        [data-testid="stMetric"] {
            padding: 1rem !important;
            border: 1px solid currentColor !important;
            border-radius: 0.5rem !important;
        }

        [data-testid="stMetricLabel"] {
            font-size: 1rem !important;
            font-weight: 700 !important;
        }

        [data-testid="stMetricValue"] {
            font-size: 2.15rem !important;
            line-height: 1.2 !important;
        }

        [data-testid="stAlert"] {
            border-width: 2px !important;
        }

        [data-testid="stExpander"] summary {
            min-height: 48px !important;
            font-size: 1.05rem !important;
            font-weight: 700 !important;
        }

        /* Dataframes/tables: larger cells and visible row separation. */
        [data-testid="stDataFrame"],
        [data-testid="stTable"] {
            font-size: 1rem !important;
        }

        [data-testid="stDataFrame"] * {
            line-height: 1.45 !important;
        }

        /* Avoid motion when accessible reading is selected. */
        *, *::before, *::after {
            animation-duration: 0.001ms !important;
            animation-iteration-count: 1 !important;
            transition-duration: 0.001ms !important;
        }

        /* Keep content comfortably readable on very wide screens. */
        [data-testid="stMainBlockContainer"] {
            max-width: 1200px !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.info(
        "Accessible reading is on: larger text and controls, stronger keyboard "
        "focus, reduced motion, clearer tabs, single-column supporting charts, "
        "and enlarged chart labels are enabled."
    )


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
        "management_summary"
    ] = None

    st.session_state[
        "management_summary_key"
    ] = None


    st.session_state[
        "forecast_ai_commentary"
    ] = None


    st.session_state[
        "forecast_ai_commentary_key"
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

# Compact dynamic context: useful on every rerun, with no extra AI call.
management_pulse = build_management_pulse(
    summary,
    st.session_state.get("forecast_result"),
)
with st.container(border=True):
    pulse_a, pulse_b, pulse_c = st.columns(3)
    with pulse_a:
        st.caption("Priority signal")
        st.markdown(management_pulse["trend"])
    with pulse_b:
        st.caption("Control signal")
        st.markdown(management_pulse["variance_or_quality"])
    with pulse_c:
        st.caption("Forecast state")
        st.markdown(management_pulse["forecast"])


# ============================================================
# AI DASHBOARD BUILDER
# ============================================================

st.subheader(
    "AI Dashboard Builder"
)


st.caption(
    "Choose a suggested dashboard or describe exactly what you want."
)


SUGGESTED_DASHBOARD_PROMPTS = {

    "Executive overview": (
        "Build an executive dashboard highlighting the most important "
        "KPIs, trends, risks, and business drivers in this dataset."
    ),

    "Financial performance": (
        "Build a financial performance dashboard focused on revenue, "
        "costs, budget variance, and the most important areas requiring "
        "management attention."
    ),

    "Trends & anomalies": (
        "Build a dashboard focused on trends, unusual movements, "
        "outliers, and areas of improving or declining performance."
    ),

    "Management deep dive": (
        "Build a management dashboard showing the strongest and weakest "
        "categories, key contributors to performance, and actionable insights."
    ),
}


prompt_columns = st.columns(4)


for prompt_index, (prompt_label, prompt_text) in enumerate(
    SUGGESTED_DASHBOARD_PROMPTS.items()
):

    if prompt_columns[prompt_index].button(
        prompt_label,
        use_container_width=True,
        key=f"suggested_dashboard_prompt_{prompt_index}",
    ):

        st.session_state[
            "agent_request_input"
        ] = prompt_text

        st.session_state[
            "agent_request"
        ] = prompt_text


st.caption(
    "AI grounding: Workiva/Python calculations remain the source of truth. "
    "AI selects, explains and prioritizes evidence; unsupported numeric claims are rejected."
)

agent_request = st.text_area(
    "Tell the agent what you want",
    key="agent_request_input",
    placeholder=(
        "Example: Build a CFO dashboard focused on revenue versus budget, "
        "largest negative variances, regional performance and key risks."
    ),
    height=100,
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

                raw_plan = create_dashboard_plan(
                    agent_request,
                    summary,
                )

                plan = validate_plan(
                    raw_plan,
                    profile,
                )

                st.session_state["agent_plan"] = plan


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

    # FAST START: render a useful deterministic executive dashboard immediately.
    # Copilot is only called when the user explicitly clicks "Build AI dashboard".
    # This removes a full network/AI round trip from initial page load.
    if st.session_state["agent_plan"] is None:
        st.session_state["agent_plan"] = manual_plan(
            "Executive",
            profile,
            business,
        )

    active_plan = st.session_state["agent_plan"]


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


# The visible plan may be locally customized in the Overview tab.
# It starts as the AI/manual plan and remains available to exports and chat.
display_plan = active_plan


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
    forecast_tab,
    comparison_tab,
    quality_tab,
    ask_tab,
) = st.tabs(
    [
        "Overview",
        "Financials",
        "Trends",
        "Forecasting",
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


        # ----------------------------------------------------
        # LOCAL DASHBOARD CUSTOMIZER
        # ----------------------------------------------------
        display_plan = customize_dashboard_plan(
            active_plan,
            data,
            profile,
            key_prefix=f"{selected_quarter}_{board_mode}",
        )

        # ----------------------------------------------------
        # KPI ROW
        # ----------------------------------------------------
        # Standard and Accessible reading keep all planned KPIs.
        # Focus mode intentionally limits the first view to four KPIs.
        visible_kpis = (
            display_plan["kpis"][:4]
            if focus_mode
            else display_plan["kpis"]
        )

        render_kpis(
            data,
            visible_kpis,
        )


        charts = [
            chart
            for chart in display_plan.get("charts", [])
            if chart.get("visible", True)
        ]


        # ----------------------------------------------------
        # PRIMARY CHART
        # The AI's first chart gets visual priority.
        # ----------------------------------------------------
        if charts:

            st.markdown(
                "#### Primary view"
            )

            render_chart(
                data,
                charts[0],
                key=f"dashboard_primary_{selected_quarter}",
            )


        # ----------------------------------------------------
        # SUPPORTING BREAKDOWNS
        # Standard = existing two-column layout.
        # Focus = secondary charts collapsed.
        # Accessible reading = single-column charts.
        # ----------------------------------------------------
        supporting_charts = charts[1:]

        if supporting_charts:

            st.markdown(
                "#### Supporting breakdowns"
            )

            if focus_mode:

                with st.expander(
                    "Show supporting charts",
                    expanded=False,
                ):

                    for chart_index, chart in enumerate(
                        supporting_charts
                    ):

                        render_chart(
                            data,
                            chart,
                            key=(
                                f"dashboard_supporting_"
                                f"{selected_quarter}_"
                                f"{chart_index}"
                            ),
                        )

            elif accessible_mode:

                for chart_index, chart in enumerate(
                    supporting_charts
                ):

                    render_chart(
                        data,
                        chart,
                        key=(
                            f"dashboard_supporting_"
                            f"{selected_quarter}_"
                            f"{chart_index}"
                        ),
                    )

            else:

                for index in range(
                    0,
                    len(supporting_charts),
                    2,
                ):

                    row = st.columns(2)

                    for offset in range(2):

                        chart_index = index + offset

                        if chart_index >= len(supporting_charts):
                            break

                        with row[offset]:
                            render_chart(
                                data,
                                supporting_charts[chart_index],
                                key=(
                                    f"dashboard_supporting_"
                                    f"{selected_quarter}_"
                                    f"{chart_index}"
                                ),
                            )


        # ----------------------------------------------------
        # MANAGEMENT PANELS
        # These use already-calculated Python facts only, so
        # they add no extra Copilot or Workiva calls.
        # ----------------------------------------------------
        layout_panels = build_dashboard_panels(
            summary
        )

        st.markdown(
            "#### Management focus"
        )

        if focus_mode:

            with st.expander(
                "Show management details",
                expanded=False,
            ):

                st.markdown("**Variance panel**")
                for item in layout_panels["variance"]:
                    st.markdown(f"- {item}")

                st.markdown("**Risk / watch panel**")
                for item in layout_panels["risks"]:
                    st.markdown(f"- {item}")

                st.markdown("**Management questions**")
                for item in layout_panels["questions"]:
                    st.markdown(f"- {item}")

        elif accessible_mode:

            st.markdown("**Variance panel**")
            for item in layout_panels["variance"]:
                st.markdown(f"- {item}")

            st.markdown("**Risk / watch panel**")
            for item in layout_panels["risks"]:
                st.markdown(f"- {item}")

            with st.expander(
                "Management questions",
                expanded=False,
            ):
                for item in layout_panels["questions"]:
                    st.markdown(f"- {item}")

        else:

            variance_col, risk_col = st.columns(2)

            with variance_col:
                st.markdown("**Variance panel**")
                for item in layout_panels["variance"]:
                    st.markdown(f"- {item}")

            with risk_col:
                st.markdown("**Risk / watch panel**")
                for item in layout_panels["risks"]:
                    st.markdown(f"- {item}")

            with st.expander(
                "Management questions",
                expanded=False,
            ):
                for item in layout_panels["questions"]:
                    st.markdown(f"- {item}")


    st.divider()


    st.subheader(
        "AI Management Summary"
    )

    # FAST START: do not block every dashboard load on a second Copilot call.
    # Generate the narrative only when requested, then keep it for this source/plan.
    summary_cache_key = (
        f"{source_key}:"
        f"{abs(hash(json.dumps(summary_plan, sort_keys=True, default=str)))}"
    )

    if st.session_state.get("management_summary_key") != summary_cache_key:
        st.session_state["management_summary"] = None
        st.session_state["management_summary_key"] = summary_cache_key

    if st.session_state.get("management_summary"):
        management_summary = st.session_state["management_summary"]
        st.markdown(management_summary)
    else:
        st.caption(
            "The dashboard is ready now. Generate the narrative only when you need it."
        )
        if st.button(
            "Generate AI management summary",
            key="generate_management_summary_button",
        ):
            try:
                with st.spinner("Generating management summary..."):
                    management_summary = generate_management_summary(
                        summary,
                        summary_plan,
                    )
                st.session_state["management_summary"] = management_summary
                st.markdown(management_summary)
            except Exception as error:
                management_summary = (
                    "AI summary unavailable: " + friendly_error(error)
                )
                st.warning(management_summary)
        else:
            management_summary = ""


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
                            data,
                            display_plan,
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
                            data,
                            display_plan,
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
                key=f"finance_dimension_{selected_quarter}_{business['dimension']}_{metric}",
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


        for trend_index, metric in enumerate(
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
                key=f"trend_{selected_quarter}_{trend_index}",
            )


        # ----------------------------------------------------
        # AUTOMATIC DRIVER ANALYSIS
        # Deterministic arithmetic first; AI only interprets it.
        # ----------------------------------------------------
        driver_analysis = (
            summary.get("analytical_context", {}).get("driver_analysis", {}) or {}
        )
        driver_metrics = driver_analysis.get("metrics", {}) or {}

        if driver_metrics:
            st.markdown("#### What drove the change?")
            st.caption(
                "Calculated decomposition of the earliest-to-latest movement. "
                "Contribution shows where the change occurred; it does not assert root cause."
            )

            driver_metric_options = list(driver_metrics.keys())
            selected_driver_metric = st.selectbox(
                "Measure",
                driver_metric_options,
                key=f"driver_metric_{selected_quarter}",
            )

            driver_dimensions = driver_metrics.get(selected_driver_metric, {}) or {}
            if driver_dimensions:
                selected_driver_dimension = st.selectbox(
                    "Driver dimension",
                    list(driver_dimensions.keys()),
                    key=f"driver_dimension_{selected_quarter}_{selected_driver_metric}",
                )
                driver_details = driver_dimensions[selected_driver_dimension]

                total_change = driver_details.get("total_change")
                total_change_pct = driver_details.get("total_change_pct")
                metric_columns = st.columns(3)
                metric_columns[0].metric(
                    "Earlier period",
                    _compact_value(driver_details.get("total_first")),
                )
                metric_columns[1].metric(
                    "Latest period",
                    _compact_value(driver_details.get("total_last")),
                )
                metric_columns[2].metric(
                    "Net change",
                    _compact_value(total_change),
                    (f"{total_change_pct:+.1f}%" if total_change_pct is not None else None),
                )

                driver_rows = []
                for item in (driver_details.get("drivers", []) or [])[:8]:
                    driver_rows.append({
                        selected_driver_dimension: item.get("category"),
                        "Earlier": item.get("first_value"),
                        "Latest": item.get("last_value"),
                        "Change": item.get("change"),
                        "Contribution to net change %": item.get("contribution_pct_of_net_change"),
                    })

                if driver_rows:
                    driver_frame = pd.DataFrame(driver_rows)
                    st.dataframe(
                        driver_frame,
                        use_container_width=True,
                        hide_index=True,
                    )

                    top_driver = driver_rows[0]
                    direction = "increased" if (top_driver.get("Change") or 0) > 0 else "decreased"
                    st.caption(
                        f"Largest absolute contributor: {top_driver[selected_driver_dimension]} "
                        f"{direction} {selected_driver_metric} by "
                        f"{_compact_value(abs(top_driver.get('Change') or 0))}."
                    )

        st.divider()
        st.markdown("#### AI trend summary")
        st.caption(
            "Ask the AI analyst to synthesize the calculated trend, anomaly, and correlation "
            "signals. All numerical claims are checked against the Python-calculated evidence."
        )

        trends_summary_key = abs(hash(json.dumps({
            "quarter": selected_quarter,
            "trends": compact_ai_context(summary).get("trends", {}),
            "anomalies": compact_ai_context(summary).get("anomalies", {}),
        }, sort_keys=True, default=str)))

        if st.session_state.get("trends_ai_summary_key") != trends_summary_key:
            st.session_state["trends_ai_summary"] = None
            st.session_state["trends_ai_summary_key"] = trends_summary_key

        if st.session_state.get("trends_ai_summary"):
            st.markdown(st.session_state["trends_ai_summary"])
        elif st.button("Generate AI trend summary", key="generate_trends_ai_summary"):
            try:
                with st.spinner("Analyzing trend signals..."):
                    trends_summary = generate_trends_ai_summary(summary)
                st.session_state["trends_ai_summary"] = trends_summary
                st.markdown(trends_summary)
            except Exception as error:
                st.warning("AI trend summary unavailable: " + friendly_error(error))


# ============================================================
# FORECASTING TAB
# ============================================================

with forecast_tab:

    st.subheader("Multi-quarter forecasting")
    st.caption(
        "Combine one to four historical Workiva quarters and project the selected "
        "measures forward. Forecast values use a simple quarterly trend extrapolation, "
        "with walk-forward backtesting when enough historical observations exist."
    )

    default_forecast_quarters = quarter_options[: min(4, len(quarter_options))]

    forecast_quarters = st.multiselect(
        "Historical quarters",
        quarter_options,
        default=default_forecast_quarters,
        max_selections=4,
        help="Select between 1 and 4 quarters. More history generally makes the direction more informative.",
        key="forecast_quarters",
    )

    horizon = st.slider(
        "Forecast horizon (quarters)",
        min_value=1,
        max_value=4,
        value=4,
        step=1,
        key="forecast_horizon",
    )

    scenario_sensitivity = st.slider(
        "Scenario sensitivity per quarter (%)",
        min_value=0,
        max_value=25,
        value=5,
        step=1,
        help=(
            "Applied deterministically around the baseline forecast. Upside is the numerically "
            "higher case and Downside the numerically lower case; this is not a probability."
        ),
        key="forecast_scenario_sensitivity",
    )

    # Discover common measures from the already-loaded quarter first. The full
    # cross-quarter intersection is validated when the user builds the board.
    preferred_metrics = []
    for candidate in [
        business.get("revenue"),
        business.get("budget"),
        business.get("costs"),
        business.get("profit"),
        business.get("ebitda"),
    ]:
        if candidate and candidate in profile["numeric"] and candidate not in preferred_metrics:
            preferred_metrics.append(candidate)
    if not preferred_metrics:
        preferred_metrics = profile["numeric"][:3]

    forecast_metrics = st.multiselect(
        "Measures to forecast",
        profile["numeric"],
        default=preferred_metrics[:3],
        max_selections=4,
        help="Choose up to four numeric measures for the forecasting board.",
        key="forecast_metrics",
    )

    if st.button(
        "Build forecasting board",
        type="primary",
        key="build_forecasting_board",
    ):
        if not forecast_quarters:
            st.warning("Select at least one historical quarter.")
        elif not forecast_metrics:
            st.warning("Select at least one measure to forecast.")
        else:
            try:
                with st.spinner("Loading selected quarters and building forecast..."):
                    forecast_bundles = [quarter_bundle(item) for item in forecast_quarters]
                    result = build_forecast_result(
                        forecast_bundles,
                        forecast_metrics,
                        horizon,
                        scenario_sensitivity_pct=scenario_sensitivity,
                    )

                missing_metrics = [
                    metric for metric in forecast_metrics
                    if metric not in result["common_metrics"]
                ]
                if missing_metrics:
                    st.warning(
                        "These measures are not available as numeric fields in every selected quarter "
                        "and were excluded: " + ", ".join(map(str, missing_metrics))
                    )

                if not result["forecasts"]:
                    st.warning(
                        "No selected measure had enough compatible numeric data across the chosen quarters."
                    )
                else:
                    st.session_state["forecast_result"] = result
                    st.session_state["forecast_ai_commentary"] = None
                    st.session_state["forecast_ai_commentary_key"] = None
                    log_event(
                        "Build forecast board",
                        f"History: {', '.join(forecast_quarters)}; horizon: {horizon} quarters",
                    )
            except Exception as error:
                st.error(friendly_error(error))

    forecast_result = st.session_state.get("forecast_result")

    if forecast_result and forecast_result.get("forecasts"):
        st.divider()
        st.markdown("#### Forecast board")

        history_count = len(forecast_result.get("history", []))
        metric_count = len(forecast_result.get("forecasts", {}))
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Historical quarters", history_count)
        c2.metric("Forecast horizon", f"{forecast_result['horizon']} quarters")
        c3.metric("Measures", metric_count)
        c4.metric("Scenario sensitivity", f"±{forecast_result.get('scenario_sensitivity_pct', 0):.0f}% / qtr")

        if history_count < 2:
            st.warning(
                "Only one historical quarter is included. The board therefore uses a flat baseline, "
                "not an estimated trend."
            )
        elif history_count < 4:
            st.info(
                "The projection uses a short history. Treat it as directional planning support rather "
                "than a statistical forecast with established seasonality."
            )
        else:
            st.info(
                "Four historical quarters are included. The model still uses a simple linear quarterly "
                "trend and does not infer seasonality or external drivers."
            )

        for metric_index, metric in enumerate(forecast_result["metrics"]):
            frame = forecast_result["forecasts"].get(metric)
            if frame is None or frame.empty:
                continue

            actual = frame[frame["Series"] == "Actual"]
            projected = frame[frame["Series"] == "Forecast"]
            latest_actual = float(actual["Value"].iloc[-1])
            end_forecast = float(projected["Value"].iloc[-1])
            delta = end_forecast - latest_actual
            delta_pct = (delta / abs(latest_actual) * 100) if latest_actual else None

            st.markdown(f"##### {metric}")
            backtest = (forecast_result.get("backtests", {}) or {}).get(metric, {}) or {}
            confidence = backtest.get("confidence", "Limited")
            wape = backtest.get("wape_pct")

            card1, card2, card3, card4 = st.columns(4)
            card1.metric("Latest actual", f"{latest_actual:,.2f}")
            card2.metric("End-of-horizon forecast", f"{end_forecast:,.2f}")
            card3.metric(
                "Projected change",
                f"{delta:+,.2f}",
                f"{delta_pct:+.1f}%" if delta_pct is not None else None,
            )
            card4.metric(
                "Forecast confidence",
                confidence,
                f"WAPE {wape:.1f}%" if wape is not None else "Backtest limited",
            )

            scenario_frame = (forecast_result.get("scenarios", {}) or {}).get(metric)
            if isinstance(scenario_frame, pd.DataFrame) and not scenario_frame.empty:
                future_scenarios = scenario_frame[scenario_frame["Scenario"] != "Actual"]
                final_period = future_scenarios["Quarter Period"].max()
                final_cases = future_scenarios[future_scenarios["Quarter Period"] == final_period]
                case_values = {
                    row["Scenario"]: float(row["Value"])
                    for _, row in final_cases.iterrows()
                }
                s1, s2, s3 = st.columns(3)
                s1.metric("Downside case", f"{case_values.get('Downside', float('nan')):,.2f}")
                s2.metric("Baseline case", f"{case_values.get('Baseline', float('nan')):,.2f}")
                s3.metric("Upside case", f"{case_values.get('Upside', float('nan')):,.2f}")

                scenario_plot = scenario_frame.copy()
                scenario_plot["Quarter order"] = scenario_plot["Quarter Period"].astype(str)
                scenario_figure = px.line(
                    scenario_plot,
                    x="Quarter order",
                    y="Value",
                    color="Scenario",
                    markers=True,
                    title=f"{metric}: baseline and sensitivity scenarios",
                    labels={"Quarter order": "Quarter", "Value": metric},
                )
                st.plotly_chart(
                    scenario_figure,
                    use_container_width=True,
                    key=f"forecast_scenario_chart_{metric_index}_{metric}",
                )
                st.caption(
                    f"Scenario assumption: ±{forecast_result.get('scenario_sensitivity_pct', 0):.0f}% per forecast quarter around the baseline. "
                    "These are deterministic sensitivity cases, not probabilities. For cost/expense measures, a higher numerical case is not necessarily favourable."
                )

            chart_frame = frame.copy()
            chart_frame["Quarter order"] = chart_frame["Quarter Period"].astype(str)
            figure = px.line(
                chart_frame,
                x="Quarter order",
                y="Value",
                color="Series",
                markers=True,
                title=f"{metric}: actuals and simple trend forecast",
                labels={"Quarter order": "Quarter", "Value": metric},
            )
            st.plotly_chart(
                figure,
                use_container_width=True,
                key=f"forecast_chart_{metric_index}_{metric}",
            )

            aggregation_label = _forecast_aggregation(metric)
            st.caption(
                f"Quarterly basis: {aggregation_label}. Forecast method: straight-line trend across "
                f"{history_count} selected historical quarter(s)."
            )
            st.markdown(forecast_result["commentary"].get(metric, ""))

            confidence_reason = backtest.get("confidence_reason")
            if confidence_reason:
                st.caption(f"**Validation:** {confidence_reason}")

            tests = backtest.get("tests")
            if isinstance(tests, pd.DataFrame) and not tests.empty:
                with st.expander("Forecast validation details", expanded=False):
                    v1, v2, v3, v4 = st.columns(4)
                    v1.metric("Backtest holdouts", int(backtest.get("backtest_points") or 0))
                    v2.metric(
                        "MAE",
                        f"{float(backtest['mae']):,.2f}" if backtest.get("mae") is not None else "—",
                    )
                    v3.metric(
                        "WAPE",
                        f"{float(backtest['wape_pct']):.1f}%" if backtest.get("wape_pct") is not None else "—",
                    )
                    v4.metric(
                        "sMAPE",
                        f"{float(backtest['smape_pct']):.1f}%" if backtest.get("smape_pct") is not None else "—",
                    )

                    validation_chart = tests[["Quarter", "Actual", "Backtest Forecast"]].copy()
                    validation_long = validation_chart.melt(
                        id_vars="Quarter",
                        value_vars=["Actual", "Backtest Forecast"],
                        var_name="Series",
                        value_name="Value",
                    )
                    validation_figure = px.line(
                        validation_long,
                        x="Quarter",
                        y="Value",
                        color="Series",
                        markers=True,
                        title=f"{metric}: historical backtest",
                    )
                    st.plotly_chart(
                        validation_figure,
                        use_container_width=True,
                        key=f"backtest_chart_{metric_index}_{metric}",
                    )

                    display_columns = [
                        "Quarter",
                        "Actual",
                        "Backtest Forecast",
                        "Error",
                        "Absolute % Error",
                        "Training Quarters",
                    ]
                    st.dataframe(
                        tests[display_columns],
                        use_container_width=True,
                        hide_index=True,
                    )
                    st.caption(
                        "Walk-forward validation predicts each known quarter using only earlier quarters. "
                        "Historical backtest accuracy does not guarantee future forecast accuracy."
                    )

        st.divider()
        st.markdown("#### Executive forecast commentary")
        st.caption(
            "The metric-level commentary above is calculated locally. You can also ask the AI analyst "
            "to synthesize the board without inventing causes or external assumptions."
        )

        commentary_key = abs(hash(json.dumps({
            "quarters": forecast_result.get("quarters", []),
            "metrics": forecast_result.get("metrics", []),
            "horizon": forecast_result.get("horizon"),
            "scenario_sensitivity_pct": forecast_result.get("scenario_sensitivity_pct"),
        }, sort_keys=True, default=str)))

        if st.session_state.get("forecast_ai_commentary_key") != commentary_key:
            st.session_state["forecast_ai_commentary"] = None
            st.session_state["forecast_ai_commentary_key"] = commentary_key

        if st.session_state.get("forecast_ai_commentary"):
            st.markdown(st.session_state["forecast_ai_commentary"])
        elif st.button("Generate AI forecast commentary", key="generate_forecast_commentary"):
            try:
                with st.spinner("Analyzing forecast trends..."):
                    commentary = generate_forecast_ai_commentary(forecast_result)
                st.session_state["forecast_ai_commentary"] = commentary
                st.markdown(commentary)
            except Exception as error:
                st.warning("AI forecast commentary unavailable: " + friendly_error(error))

        export_history = forecast_result["history"].copy()
        export_history["Quarter Period"] = export_history["Quarter Period"].astype(str)
        st.download_button(
            "Download combined quarterly history as CSV",
            data=export_history.to_csv(index=False).encode("utf-8"),
            file_name="workiva_forecast_history.csv",
            mime="text/csv",
            key="download_forecast_history",
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
            ],
            key_prefix="comparison_tab",
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


        st.markdown("**Financial intelligence**")
        finance_labels = {
            "actual": "Actual",
            "forecast": "Forecast",
            "profit": "Profit",
            "ebitda": "EBITDA",
            "opex": "Opex",
            "capex": "Capex",
            "headcount": "Headcount / FTE",
            "volume": "Volume",
            "price": "Price",
            "prior_year": "Prior year",
            "margin": "Margin",
        }
        for key, label in finance_labels.items():
            if business.get(key):
                st.write(f"{label}:", business.get(key))

        derived_measures = business.get("derived_measures", {})
        if derived_measures:
            st.write("Derived measures:")
            for measure, formula in derived_measures.items():
                st.caption(f"{measure} = {formula}")
        else:
            st.caption(
                "No additional finance ratios were derived because the required source columns were not all present."
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
            "Ask about the current dashboard, trends, forecasting, or quarter comparisons. "
            "The assistant receives the active calculated evidence and distinguishes actuals "
            "from forecasts."
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


    # Context-aware shortcuts are generated from the current calculated state.
    # They remain local UI actions and do not call Copilot until clicked.
    previous_quarter = None
    try:
        current_index = quarter_options.index(selected_quarter)
        if current_index + 1 < len(quarter_options):
            previous_quarter = quarter_options[current_index + 1]
    except (ValueError, AttributeError):
        previous_quarter = None

    contextual_questions = build_contextual_ai_questions(
        summary,
        st.session_state.get("forecast_result"),
        previous_quarter=previous_quarter,
        selected_quarter=selected_quarter,
    )

    shortcut_columns = st.columns(len(contextual_questions))
    suggested_question = None
    for shortcut_index, (label, prompt) in enumerate(contextual_questions):
        if shortcut_columns[shortcut_index].button(
            label,
            use_container_width=True,
            key=f"chat_contextual_{shortcut_index}_{selected_quarter}",
        ):
            suggested_question = prompt

    typed_question = st.chat_input(
        "Ask about Workiva data..."
    )

    # A button click and a typed question share exactly the same execution
    # path, avoiding duplicate AI calls and keeping interaction fast.
    question = typed_question or suggested_question


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
                        comparison,
                        key_prefix="chat_comparison",
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
                                (display_plan or active_plan or summary_plan),
                                data,
                                profile,
                                business,
                                st.session_state["chat_messages"],
                                st.session_state.get("forecast_result"),
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


        # Refresh after storing the response so the user can immediately
        # ask another question while retaining the full conversation history.
        st.rerun()


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