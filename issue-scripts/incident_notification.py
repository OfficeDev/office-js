import os
from typing import Optional

import requests


LOGIC_APP_URL = "https://prod-07.northcentralus.logic.azure.com:443/workflows/fce230d344e3424193fa488f37a5f1cf/triggers/When_a_HTTP_request_is_received/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2FWhen_a_HTTP_request_is_received%2Frun&sv=1.0&sig=FWEPdesnuPXBarpDl_bbT6jTQLPMT2zkH8eS96ZekZw"


def send_incident_notification(
    *,
    issue_title: Optional[str],
    issue_url: Optional[str],
    regression_reason: str,
    user_login: Optional[str],
    detection_source: Optional[str] = None,
) -> bool:
    """Trigger the regression incident Logic App webhook with required fields."""

    logic_app_url = LOGIC_APP_URL
    timeout = float(os.environ.get("INCIDENT_NOTIFICATION_TIMEOUT", 10))
    payload = {
        "issue_title": issue_title,
        "issue_url": issue_url,
        "regression_reason": regression_reason,
        "user": user_login,
    }
    if detection_source:
        payload["detection_source"] = detection_source

    try:
        response = requests.post(
            logic_app_url,
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=timeout,
        )
        if 200 <= response.status_code < 300:
            print(
                "Regression incident Logic App triggered for issue "
                f"{issue_url or issue_title}"
            )
            return True

        print(
            "Failed to trigger regression incident Logic App. "
            f"Status: {response.status_code}, Response: {response.text}"
        )
        return False
    except requests.exceptions.Timeout:
        print(f"Timeout (> {timeout}s) when calling regression incident Logic App")
        return False
    except requests.exceptions.RequestException as exc:
        print(f"Request error when calling regression incident Logic App: {exc}")
        return False
    except Exception as exc:  # pragma: no cover
        print(f"Unexpected error when calling regression incident Logic App: {exc}")
        return False
