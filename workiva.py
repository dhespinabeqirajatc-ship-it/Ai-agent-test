import os
import re

import pandas as pd
import requests
from dotenv import load_dotenv


load_dotenv()


BASE_URL = os.environ.get(
    "WORKIVA_BASE_URL",
    "https://api.app.wdesk.com",
)

API_VERSION = os.environ.get(
    "WORKIVA_API_VERSION",
    "2026-01-01",
)


# =================================================
# AUTHENTICATION
# =================================================

def get_workiva_token():

    client_id = os.environ[
        "WORKIVA_CLIENT_ID"
    ]

    client_secret = os.environ[
        "WORKIVA_CLIENT_SECRET"
    ]

    response = requests.post(
        f"{BASE_URL}/oauth2/token",
        headers={
            "X-Version": API_VERSION,
            "Content-Type":
                "application/x-www-form-urlencoded",
        },
        data={
            "grant_type":
                "client_credentials",
            "client_id":
                client_id,
            "client_secret":
                client_secret,
        },
        timeout=30,
    )

    response.raise_for_status()

    return response.json()[
        "access_token"
    ]


def build_headers(token):

    return {
        "Authorization":
            f"Bearer {token}",
        "X-Version":
            API_VERSION,
        "Accept":
            "application/json",
    }


# =================================================
# PAGINATION
# =================================================

def get_all_pages(
    url,
    token,
    params=None,
):

    items = []

    next_url = url
    next_params = params

    while next_url:

        response = requests.get(
            next_url,
            headers=build_headers(
                token
            ),
            params=next_params,
            timeout=30,
        )

        response.raise_for_status()

        payload = (
            response.json()
        )

        page_items = (
            payload.get(
                "data",
                [],
            )
        )

        if isinstance(
            page_items,
            list,
        ):

            items.extend(
                page_items
            )

        next_url = (
            payload.get(
                "@nextLink"
            )
            or payload.get(
                "nextLink"
            )
            or payload.get(
                "next"
            )
        )

        next_params = None

    return items


# =================================================
# SPREADSHEETS
# =================================================

def list_all_spreadsheets(
    token=None,
):

    if token is None:

        token = (
            get_workiva_token()
        )

    return get_all_pages(
        f"{BASE_URL}/spreadsheets",
        token,
        params={
            "$maxpagesize": 100,
        },
    )


def search_spreadsheets(
    search_text,
    token=None,
):

    if token is None:

        token = (
            get_workiva_token()
        )

    safe_search = (
        search_text.replace(
            "'",
            "''",
        )
    )

    return {
        "data": get_all_pages(
            f"{BASE_URL}/spreadsheets",
            token,
            params={
                "$filter":
                    f"name contains "
                    f"'{safe_search}'",
                "$maxpagesize": 100,
            },
        )
    }


# =================================================
# SHEETS
# =================================================

def list_sheets(
    spreadsheet_id,
    token=None,
):

    if token is None:

        token = (
            get_workiva_token()
        )

    sheets = get_all_pages(
        (
            f"{BASE_URL}/spreadsheets/"
            f"{spreadsheet_id}/sheets"
        ),
        token,
        params={
            "$maxpagesize": 100,
        },
    )

    return {
        "data": sheets
    }


# =================================================
# QUARTER NAME HELPERS
# =================================================

def canonical_quarter(
    value,
):

    if value is None:
        return None

    text = str(value).upper()

    match = re.search(
        r"\bQ([1-4])"
        r"[\s\-_\/]*"
        r"(20\d{2})\b",
        text,
    )

    if not match:

        return None

    quarter = match.group(1)
    year = match.group(2)

    return (
        f"Q{quarter} {year}"
    )


def quarter_sort_key(
    quarter,
):

    match = re.match(
        r"Q([1-4]) (20\d{2})",
        quarter,
    )

    if not match:

        return (0, 0)

    return (
        int(match.group(2)),
        int(match.group(1)),
    )


# =================================================
# DISCOVER AVAILABLE QUARTERS
# =================================================

def discover_available_quarters():

    token = (
        get_workiva_token()
    )

    spreadsheets = (
        list_all_spreadsheets(
            token=token
        )
    )

    quarters = set()

    for spreadsheet in spreadsheets:

        name = spreadsheet.get(
            "name",
            "",
        )

        quarter = (
            canonical_quarter(
                name
            )
        )

        if quarter:

            quarters.add(
                quarter
            )

    return sorted(
        quarters,
        key=quarter_sort_key,
        reverse=True,
    )


# =================================================
# SMART QUARTER FINDER
# =================================================

def spreadsheet_score(
    name,
    requested_quarter,
):

    name_text = str(
        name
    ).strip()

    target = (
        requested_quarter
        .strip()
        .lower()
    )

    normalized = (
        name_text.lower()
    )

    if normalized == target:
        return 100

    if (
        canonical_quarter(
            name_text
        )
        == requested_quarter
    ):
        return 90

    if target in normalized:
        return 80

    return 0


def sheet_score(
    name,
    requested_quarter,
):

    name_text = str(
        name
    ).strip()

    target = (
        requested_quarter
        .strip()
        .lower()
    )

    normalized = (
        name_text.lower()
    )

    if normalized == target:
        return 100

    if (
        canonical_quarter(
            name_text
        )
        == requested_quarter
    ):
        return 90

    if target in normalized:
        return 70

    return 0


def find_quarter_source_smart(
    quarter_name,
):

    token = (
        get_workiva_token()
    )

    requested_quarter = (
        canonical_quarter(
            quarter_name
        )
        or quarter_name.strip()
    )

    # ---------------------------------------------
    # FAST SEARCH
    # ---------------------------------------------

    search_results = (
        search_spreadsheets(
            requested_quarter,
            token=token,
        )
    )

    candidates = (
        search_results.get(
            "data",
            [],
        )
    )

    # ---------------------------------------------
    # FALLBACK:
    # get spreadsheet catalogue if exact search
    # did not return useful candidates
    # ---------------------------------------------

    if not candidates:

        candidates = (
            list_all_spreadsheets(
                token=token
            )
        )

    scored = []

    for spreadsheet in candidates:

        score = spreadsheet_score(
            spreadsheet.get(
                "name",
                "",
            ),
            requested_quarter,
        )

        if score > 0:

            scored.append(
                (
                    score,
                    spreadsheet,
                )
            )

    scored.sort(
        key=lambda item:
            item[0],
        reverse=True,
    )

    # ---------------------------------------------
    # Inspect strongest spreadsheet candidates first
    # ---------------------------------------------

    for spreadsheet_score_value, spreadsheet in scored[:10]:

        spreadsheet_id = (
            spreadsheet.get(
                "id"
            )
        )

        if not spreadsheet_id:
            continue

        try:

            result = list_sheets(
                spreadsheet_id,
                token=token,
            )

        except requests.RequestException:

            continue

        sheets = result.get(
            "data",
            [],
        )

        best_sheet = None
        best_sheet_score = 0

        for sheet in sheets:

            score = sheet_score(
                sheet.get(
                    "name",
                    "",
                ),
                requested_quarter,
            )

            if (
                score
                > best_sheet_score
            ):

                best_sheet_score = score
                best_sheet = sheet

            # Exact match:
            # stop immediately.

            if score == 100:

                return {
                    "spreadsheet_id":
                        spreadsheet_id,
                    "spreadsheet_name":
                        spreadsheet.get(
                            "name"
                        ),
                    "sheet_id":
                        sheet.get(
                            "id"
                        ),
                    "sheet_name":
                        sheet.get(
                            "name"
                        ),
                    "confidence":
                        "Exact match",
                }

        if (
            best_sheet
            and best_sheet_score >= 90
        ):

            return {
                "spreadsheet_id":
                    spreadsheet_id,
                "spreadsheet_name":
                    spreadsheet.get(
                        "name"
                    ),
                "sheet_id":
                    best_sheet.get(
                        "id"
                    ),
                "sheet_name":
                    best_sheet.get(
                        "name"
                    ),
                "confidence":
                    "Strong match",
            }

    # ---------------------------------------------
    # DEEP FALLBACK
    #
    # If the quarter sheet is hidden inside a
    # spreadsheet with an unrelated name, scan
    # spreadsheets until an exact sheet is found.
    # ---------------------------------------------

    all_spreadsheets = (
        list_all_spreadsheets(
            token=token
        )
    )

    checked_ids = {
        item[1].get("id")
        for item in scored[:10]
    }

    for spreadsheet in all_spreadsheets:

        spreadsheet_id = (
            spreadsheet.get(
                "id"
            )
        )

        if (
            not spreadsheet_id
            or spreadsheet_id
            in checked_ids
        ):

            continue

        try:

            result = list_sheets(
                spreadsheet_id,
                token=token,
            )

        except requests.RequestException:

            continue

        for sheet in result.get(
            "data",
            [],
        ):

            if (
                sheet_score(
                    sheet.get(
                        "name",
                        "",
                    ),
                    requested_quarter,
                )
                == 100
            ):

                return {
                    "spreadsheet_id":
                        spreadsheet_id,
                    "spreadsheet_name":
                        spreadsheet.get(
                            "name"
                        ),
                    "sheet_id":
                        sheet.get(
                            "id"
                        ),
                    "sheet_name":
                        sheet.get(
                            "name"
                        ),
                    "confidence":
                        "Sheet match",
                }

    return None


# =================================================
# READ SHEET DATA
# =================================================

def get_sheet_data(
    spreadsheet_id,
    sheet_id,
    cell_range=None,
):

    token = (
        get_workiva_token()
    )

    params = {}

    if cell_range:

        params[
            "$cellrange"
        ] = cell_range

    response = requests.get(
        (
            f"{BASE_URL}/spreadsheets/"
            f"{spreadsheet_id}/sheets/"
            f"{sheet_id}/sheetdata"
        ),
        headers=build_headers(
            token
        ),
        params=params,
        timeout=30,
    )

    response.raise_for_status()

    return response.json()


# =================================================
# DATAFRAME CONVERSION
# =================================================

def make_unique_headers(
    headers,
):

    cleaned = []
    seen = {}

    for index, header in enumerate(
        headers
    ):

        if header is None:
            header = ""

        header = str(
            header
        ).strip()

        if header == "":

            header = (
                f"Unnamed_"
                f"{index + 1}"
            )

        if header in seen:

            seen[header] += 1

            header = (
                f"{header}_"
                f"{seen[header]}"
            )

        else:

            seen[header] = 1

        cleaned.append(
            header
        )

    return cleaned


def workiva_sheet_to_dataframe(
    sheet_response,
):

    cells = (
        sheet_response
        .get(
            "data",
            {},
        )
        .get(
            "cells",
            [],
        )
    )

    rows = []

    for row in cells:

        row_values = []

        for cell in row:

            if cell is None:

                row_values.append(
                    None
                )

            else:

                row_values.append(
                    cell.get(
                        "calculatedValue"
                    )
                )

        rows.append(
            row_values
        )

    if not rows:

        return pd.DataFrame()

    headers = (
        make_unique_headers(
            rows[0]
        )
    )

    dataframe = pd.DataFrame(
        rows[1:],
        columns=headers,
    )

    dataframe = (
        dataframe.replace(
            "",
            pd.NA,
        )
    )

    dataframe = (
        dataframe.dropna(
            axis=1,
            how="all",
        )
    )

    dataframe = (
        dataframe.dropna(
            axis=0,
            how="all",
        )
    )

    return dataframe