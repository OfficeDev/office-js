import os
from typing import Optional

import requests

def send_incident_notification(
    *,
    issue_title: Optional[str],
    issue_url: Optional[str],
    regression_reason: str,
    user_login: Optional[str],
    detection_source: Optional[str] = None,
) -> bool:
    """Trigger the regression incident Logic App webhook with required fields."""

    logic_app_url = os.environ.get("REGRESSION_ICM_LOGIC_APP_URL")
    if not logic_app_url:
        print(
            "Skipping regression incident notification: LOGIC_APP_URL environment "
            "variable is not set."
        )
        return False
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
