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
        """Generate values for the active workbook using its matched template."""
        return self._request("POST", "/generate", {})

    def generate_after_excel_open(self, retries=60, interval_seconds=1):
        """Wait for an open workbook, then generate its matched template values."""
        self.wait_until_ready(retries=retries, interval_seconds=interval_seconds)
        return self.generate_random_numbers()

    def wait_until_ready(self, retries=30, interval_seconds=1):
        """Wait until the add-in, workbook, and matched generation template are ready."""
        import time

        try:
            attempts = int(retries)
            interval = float(interval_seconds)
        except (TypeError, ValueError):
            raise YingdaoExcelError("retries and interval_seconds must be numeric.")
        if attempts < 1:
            raise YingdaoExcelError("retries must be at least 1.")
        if interval < 0:
            raise YingdaoExcelError("interval_seconds cannot be negative.")

        last_error = None
        last_state = None
        for attempt in range(attempts):
            try:
                state = self.status()
                last_state = state
                if self._is_generation_ready(state):
                    return state
            except YingdaoExcelError as exc:
                last_error = exc
            if attempt + 1 < attempts:
                time.sleep(interval)

        if last_error is not None:
            raise last_error
        raise YingdaoExcelError(self._readiness_error(last_state))

    @staticmethod
    def _is_generation_ready(state):
        """Require a live add-in, active workbook, matched template, and rules."""
        return bool(
            (state.get("addinLoaded", state.get("AddinLoaded", True)))
            and (state.get("workbookOpen", state.get("WorkbookOpen", False)))
            and (state.get("templateMatched", state.get("TemplateMatched", False)))
            and (state.get("canGenerate", state.get("CanGenerate", False)))
        )

    @staticmethod
    def _readiness_error(state):
        if not state:
            return "Excel 加载项状态不可用。"
        message = state.get("message", state.get("Message", ""))
        if message:
            return "Excel 加载项未就绪：{}".format(message)
        if not state.get("workbookOpen", state.get("WorkbookOpen", False)):
            return "Excel 尚未打开目标工作簿。"
        return "当前工作簿未匹配到可生成模板。"

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


def generate_after_excel_open(retries=60, interval_seconds=1):
    """影刀主入口：Excel 打开后等待工作簿就绪并生成随机数。"""
    return _client.generate_after_excel_open(retries, interval_seconds)
