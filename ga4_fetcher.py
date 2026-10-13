"""
Fetches marketing sessions from GA4 for a given date.
Metric: total Sessions where sessionSourceMedium contains 'mktng'
(matches Meta / Mktng, mktng / hotstar, igxmktng / story, etc.)
"""

import os
from google.oauth2.service_account import Credentials
from google.analytics.data_v1beta import BetaAnalyticsDataClient
from google.analytics.data_v1beta.types import (
    RunReportRequest,
    DateRange,
    Metric,
    Dimension,
    FilterExpression,
    Filter,
)

SERVICE_ACCOUNT_FILE = "service_account.json"
GA4_PROPERTY_ID = os.getenv("GA4_PROPERTY_ID", "")


def fetch_marketing_sessions(date_str: str, property_id: str = "") -> int:
    """
    Returns total Sessions for the given date where session source/medium
    contains 'mktng'. date_str format: 'YYYY-MM-DD'.
    """
    prop_id = property_id or GA4_PROPERTY_ID
    if not prop_id:
        raise ValueError("GA4_PROPERTY_ID not set. Add it as an env var or GitHub secret.")

    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=["https://www.googleapis.com/auth/analytics.readonly"],
    )
    client = BetaAnalyticsDataClient(credentials=creds)

    request = RunReportRequest(
        property=f"properties/{prop_id}",
        date_ranges=[DateRange(start_date=date_str, end_date=date_str)],
        dimensions=[Dimension(name="sessionSourceMedium")],
        metrics=[Metric(name="sessions")],
        dimension_filter=FilterExpression(
            filter=Filter(
                field_name="sessionSourceMedium",
                string_filter=Filter.StringFilter(
                    match_type=Filter.StringFilter.MatchType.CONTAINS,
                    value="mktng",
                    case_sensitive=False,
                ),
            )
        ),
    )

    response = client.run_report(request)
    total = sum(int(row.metric_values[0].value) for row in response.rows)
    print(f"  GA4 marketing sessions for {date_str}: {total}")
    return total
