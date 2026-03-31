from __future__ import annotations

from datetime import date, datetime
from typing import Any
from urllib.parse import urlparse
import random
import re
import time

import requests
from requests.auth import HTTPBasicAuth

MONTH_COLUMNS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]


class WPCourseClient:
    def __init__(self, base_url: str, username: str, app_password: str):
        self.base_url = (base_url or "").strip().rstrip("/")

        clean_username = (username or "").strip()
        clean_password = (app_password or "").strip().replace(" ", "")

        self.auth = HTTPBasicAuth(clean_username, clean_password)
        self.session = requests.Session()

        # Match the successful local behavior more closely.
        self.session.headers.update({
            "User-Agent": "insomnia/11.0.2",
            "Accept": "*/*",
            "Content-Type": "application/json",
            "Origin": self.base_url,
            "Referer": f"{self.base_url}/",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
        })

        # Gentle pacing defaults to reduce pressure on Cloudflare/WP.
        self.timeout = 30
        self.min_interval_seconds = 0.85
        self.max_retries = 4
        self.base_backoff_seconds = 1.25
        self.last_request_ts = 0.0

    def _endpoint(self, path: str) -> str:
        normalized = path if path.startswith("/") else f"/{path}"
        return f"{self.base_url}{normalized}"

    def _rest_candidate_paths(self, path: str) -> list[str]:
        normalized = path if path.startswith("/") else f"/{path}"
        return [normalized]

    def _sleep_for_spacing(self) -> None:
        now = time.monotonic()
        elapsed = now - self.last_request_ts
        if elapsed < self.min_interval_seconds:
            time.sleep(self.min_interval_seconds - elapsed)

    def _mark_request_complete(self) -> None:
        self.last_request_ts = time.monotonic()

    @staticmethod
    def _is_cloudflare_challenge(response: requests.Response) -> bool:
        return (
            response.headers.get("cf-mitigated", "").lower() == "challenge"
            or "just a moment" in (response.text or "").lower()
        )

    @staticmethod
    def _retry_after_seconds(response: requests.Response) -> float | None:
        value = (response.headers.get("Retry-After") or "").strip()
        if not value:
            return None
        try:
            return max(0.0, float(value))
        except ValueError:
            return None

    def _compute_backoff(self, attempt_index: int, response: requests.Response | None = None) -> float:
        retry_after = self._retry_after_seconds(response) if response is not None else None
        if retry_after is not None:
            return retry_after + random.uniform(0.1, 0.4)

        delay = self.base_backoff_seconds * (2 ** attempt_index)
        delay = min(delay, 12.0)
        return delay + random.uniform(0.1, 0.5)

    @staticmethod
    def _raise_for_response(response: requests.Response) -> None:
        if response.ok:
            return

        server = response.headers.get("server", "")
        cf_ray = response.headers.get("cf-ray", "")
        cf_mitigated = response.headers.get("cf-mitigated", "")
        details = []
        if server:
            details.append(f"server={server}")
        if cf_ray:
            details.append(f"cf-ray={cf_ray}")
        if cf_mitigated:
            details.append(f"cf-mitigated={cf_mitigated}")
        suffix = f" ({', '.join(details)})" if details else ""

        if WPCourseClient._is_cloudflare_challenge(response):
            raise requests.HTTPError(
                f"{response.status_code} for {response.url}{suffix} - blocked by Cloudflare challenge"
            )

        raise requests.HTTPError(f"{response.status_code} for {response.url}{suffix}")

    def _request_with_retries(
        self,
        method: str,
        path: str,
        *,
        auth=None,
        retry_on_401_without_auth: bool = False,
        **kwargs,
    ) -> requests.Response:
        last_response: requests.Response | None = None
        candidate_paths = self._rest_candidate_paths(path)

        for candidate_path in candidate_paths:
            url = self._endpoint(candidate_path)

            for attempt in range(self.max_retries + 1):
                self._sleep_for_spacing()

                response = self.session.request(
                    method=method.upper(),
                    url=url,
                    auth=auth,
                    timeout=self.timeout,
                    **kwargs,
                )
                self._mark_request_complete()

                if response.ok:
                    return response

                last_response = response

                if response.status_code == 401 and retry_on_401_without_auth and auth is not None:
                    self._sleep_for_spacing()
                    fallback = self.session.request(
                        method=method.upper(),
                        url=url,
                        auth=None,
                        timeout=self.timeout,
                        **kwargs,
                    )
                    self._mark_request_complete()
                    if fallback.ok:
                        return fallback
                    last_response = fallback

                # Do not hammer the edge if challenged or throttled.
                if response.status_code in (403, 429, 500, 502, 503, 504):
                    if attempt < self.max_retries:
                        time.sleep(self._compute_backoff(attempt, response))
                        continue

                break

        if last_response is not None:
            self._raise_for_response(last_response)
        raise requests.HTTPError(f"Unable to call endpoint for path: {path}")

    def _get_with_optional_auth(self, path: str, prefer_auth: bool = True, **kwargs) -> requests.Response:
        if prefer_auth:
            return self._request_with_retries(
                "GET",
                path,
                auth=self.auth,
                retry_on_401_without_auth=True,
                **kwargs,
            )

        return self._request_with_retries(
            "GET",
            path,
            auth=None,
            retry_on_401_without_auth=False,
            **kwargs,
        )

    def get_course_by_slug(self, slug: str) -> dict[str, Any] | None:
        """Resolve WP course by slug using the canonical WP slug query endpoint."""
        if not slug:
            return None

        response = self._get_with_optional_auth(
            "/wp-json/wp/v2/cursuri",
            prefer_auth=True,
            params={"slug": slug},
        )
        data = response.json()
        if isinstance(data, list) and data:
            return data[0]
        return None

    def get_course(self, post_id: int) -> dict[str, Any]:
        """Fetch full WP post including ACF."""
        response = self._get_with_optional_auth(
            f"/wp-json/wp/v2/cursuri/{int(post_id)}",
            prefer_auth=True,
        )
        return response.json()

    def get_post_id_from_permalink(self, permalink: str) -> int | None:
        """Extract WP post id from public course HTML page."""
        url = str(permalink or "").strip()
        if not url:
            return None

        response = self._request_with_retries(
            "GET",
            url,
            auth=None,
            retry_on_401_without_auth=False,
        ) if url.startswith("http://") or url.startswith("https://") else None

        if response is None or not response.ok:
            return None

        html = response.text or ""
        patterns = [
            r'postid-(\d+)',
            r'id=["\']post-(\d+)["\']',
            r'"post_id"\s*:\s*"?(\d+)"?',
            r'"postId"\s*:\s*"?(\d+)"?',
        ]
        for pattern in patterns:
            match = re.search(pattern, html, flags=re.IGNORECASE)
            if match:
                try:
                    return int(match.group(1))
                except (TypeError, ValueError):
                    continue
        return None

    def update_course_program(self, post_id: int, final_program: list[dict], auth=None) -> dict[str, Any]:
        """POST only acf.program back to WP."""
        payload = {
            "acf": {
                "program": final_program if final_program else False,
            }
        }

        response = self._request_with_retries(
            "POST",
            f"/wp-json/wp/v2/cursuri/{int(post_id)}",
            auth=auth or self.auth,
            retry_on_401_without_auth=False,
            json=payload,
        )
        return response.json()


def extract_slug_from_permalink(url: str) -> str:
    """Return slug from course permalink."""
    parsed = urlparse((url or "").strip())
    path = (parsed.path or "").strip("/ ")
    if not path:
        return ""

    parts = [part for part in path.split("/") if part]
    if not parts:
        return ""
    return parts[-1]


def parse_single_ro_date(value: str) -> date:
    """Parse dd.mm.yyyy."""
    return datetime.strptime(value.strip(), "%d.%m.%Y").date()


def parse_effective_end_date(text: str) -> date | None:
    text = str(text or "").strip()
    if not text:
        return None

    m1 = re.fullmatch(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", text)
    if m1:
        d, m, y = map(int, m1.groups())
        try:
            return date(y, m, d)
        except ValueError:
            return None

    m2 = re.fullmatch(r"(\d{1,2})-(\d{1,2})\.(\d{1,2})\.(\d{4})", text)
    if m2:
        _start_day, end_day, month, year = map(int, m2.groups())
        try:
            return date(year, month, end_day)
        except ValueError:
            return None

    return None


def _normalize_excel_date_value(value: Any) -> str:
    if value is None:
        return ""

    if hasattr(value, "to_pydatetime"):
        value = value.to_pydatetime()

    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")

    return str(value).strip()


def expand_date_token(token: str) -> list[str]:
    """
    Expand one token:
    - 08.04.2026 -> ['08.04.2026']
    - 14-15.04.2026 -> ['14.04.2026', '15.04.2026']
    - 21-23.04.2026 -> ['21.04.2026', '22.04.2026', '23.04.2026']
    """
    token = _normalize_excel_date_value(token)
    if not token or token.lower() == "nan":
        return []

    token = token.replace("–", "-").replace("—", "-")

    if re.fullmatch(r"\d{1,2}\.\d{1,2}\.\d{4}", token):
        parsed = datetime.strptime(token, "%d.%m.%Y")
        return [parsed.strftime("%d.%m.%Y")]

    if re.fullmatch(r"\d{4}-\d{2}-\d{2}(?:\s+\d{2}:\d{2}:\d{2})?", token):
        parsed = datetime.strptime(token[:10], "%Y-%m-%d")
        return [parsed.strftime("%d.%m.%Y")]

    match = re.fullmatch(r"(\d{1,2})-(\d{1,2})\.(\d{1,2})\.(\d{4})", token)
    if not match:
        return []

    start_day = int(match.group(1))
    end_day = int(match.group(2))
    month = int(match.group(3))
    year = int(match.group(4))

    if end_day < start_day:
        return []

    out: list[str] = []
    for day in range(start_day, end_day + 1):
        normalized = f"{day:02d}.{month:02d}.{year}"
        try:
            parse_single_ro_date(normalized)
        except ValueError:
            return []
        out.append(normalized)
    return out


def split_cell_tokens(value: Any) -> list[str]:
    text = _normalize_excel_date_value(value)
    if not text or text.lower() == "nan":
        return []
    return [text]


def parse_excel_dates_from_row(row: dict) -> list[str]:
    """Read all month columns and keep original text values as provided in Excel."""
    dates: list[str] = []
    lowered_row = {str(key).strip().lower(): value for key, value in (row or {}).items()}

    for column in MONTH_COLUMNS:
        raw = lowered_row.get(column.lower())
        for token in split_cell_tokens(raw):
            normalized = str(token).strip()
            if normalized:
                dates.append(normalized)
    return dates


def _filter_existing_non_expired_rows(program: list[dict], today: date) -> list[dict[str, str]]:
    normalized: list[dict[str, str]] = []
    seen: set[str] = set()

    for row in program or []:
        raw = str((row or {}).get("data", "")).strip()
        if not raw:
            continue
        end_dt = parse_effective_end_date(raw)
        if end_dt is None:
            continue

        if end_dt >= today and raw not in seen:
            normalized.append({"data": raw})
            seen.add(raw)

    return normalized


def build_final_program(existing_program: list, excel_dates: list[str], today: date) -> list[dict]:
    """Keep non-expired existing rows unchanged, append Excel text unchanged, dedupe by exact text."""
    seen: set[str] = set()
    result: list[dict[str, str]] = []

    for row in _filter_existing_non_expired_rows(existing_program, today):
        raw = row["data"]
        if raw not in seen:
            result.append({"data": raw})
            seen.add(raw)

    for raw in excel_dates:
        normalized = str(raw).strip()
        if not normalized:
            continue
        if normalized not in seen:
            result.append({"data": normalized})
            seen.add(normalized)
    return result


def valid_existing_program(existing_program: list, today: date) -> list[dict[str, str]]:
    return _filter_existing_non_expired_rows(existing_program, today)
