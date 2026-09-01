"""Yingdao Python module for generating random values in Excel.

The module talks to the add-in over localhost. It does not move the mouse,
activate a window, or read/write Excel through coordinates.
"""

import json
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


class YingdaoExcelError(RuntimeError):
    """Raised when the Excel add-in cannot complete an automation request."""


class ExcelCalibrationYingdao:
    def __init__(self, host="127.0.0.1", port=30771, token="", timeout=120):
        self.base_url = "http://{}:{}/api/yingdao".format(host, port).rstrip("/")
        self.token = token
        self.timeout = timeout

    def health(self):
        return self._request("GET", "/health")

    def status(self):
        return self._request("GET", "/status")

    def generate_random_numbers(self):
        """Generate values for the active workbook using the matched template."""
        return self._request("POST", "/generate", {})

    def wait_until_ready(self, retries=30, interval_seconds=1):
        """Wait until the add-in is loaded and Excel has an active workbook."""
        import time

        last_error = None
        for _ in range(retries):
            try:
                state = self.status()
                if state.get("workbookOpen") or state.get("WorkbookOpen"):
                    return state
            except YingdaoExcelError as exc:
                last_error = exc
            time.sleep(interval_seconds)

        if last_error is not None:
            raise last_error
        raise YingdaoExcelError("Excel has no open workbook.")

    def _request(self, method, path, body=None):
        data = None
        headers = {"Accept": "application/json"}
        if self.token:
            headers["X-Excel-Calibration-Token"] = self.token
        if body is not None:
            data = json.dumps(body).encode("utf-8")
            headers["Content-Type"] = "application/json"

        request = Request(self.base_url + path, data=data, headers=headers, method=method)
        try:
            with urlopen(request, timeout=self.timeout) as response:
                payload = response.read().decode("utf-8")
                result = json.loads(payload) if payload else {}
        except HTTPError as exc:
            try:
                detail = json.loads(exc.read().decode("utf-8"))
            except (ValueError, UnicodeError):
                detail = {}
            raise YingdaoExcelError(detail.get("error", "Excel add-in request failed (HTTP {}).".format(exc.code)))
        except (URLError, OSError) as exc:
            raise YingdaoExcelError(
                "Cannot connect to the Excel add-in. Confirm that Excel and the add-in are running: {}".format(exc)
            )

        if not result.get("ok", False):
            raise YingdaoExcelError(result.get("error", "The Excel add-in returned a failure."))
        return result


_client = ExcelCalibrationYingdao()


def health():
    return _client.health()


def status():
    return _client.status()


def generate_random_numbers():
    return _client.generate_random_numbers()
