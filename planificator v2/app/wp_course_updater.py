from __future__ import annotations

from datetime import date, datetime
from typing import Any
from urllib.parse import urlparse
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
        self.auth = HTTPBasicAuth((username or "").strip(), (app_password or "").strip())
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "insomnia/11.0.2",
            "Accept": "*/*",
            "Content-Type": "application/json",
        })

    def _endpoint(self, path: str) -> str:
        return f"{self.base_url}{path}"

    def _rest_candidate_paths(self, path: str) -> list[str]:
        normalized = path if path.startswith("/") else f"/{path}"
        if normalized.startswith("/wp-json/"):
            rest_route_path = normalized[len("/wp-json"):]
        else:
            rest_route_path = normalized
        rest_route = f"/?rest_route={rest_route_path}"
        return [normalized, rest_route]

    @staticmethod
    def _raise_for_response(response: requests.Response) -> None:
        if response.ok:
            return

        server = response.headers.get("server", "")
        cf_ray = response.headers.get("cf-ray", "")
        details = []
        if server:
            details.append(f"server={server}")
        if cf_ray:
            details.append(f"cf-ray={cf_ray}")
        suffix = f" ({', '.join(details)})" if details else ""
        raise requests.HTTPError(f"{response.status_code} for {response.url}{suffix}")

    def _get_with_optional_auth(self, path: str, prefer_auth: bool = True, **kwargs) -> requests.Response:
        last_response: requests.Response | None = None
        candidate_paths = self._rest_candidate_paths(path)

        auth_sequence = [self.auth] if prefer_auth else [None, self.auth]
        fallback_auth_sequence = [None] if prefer_auth else []

        for candidate_path in candidate_paths:
            for auth in auth_sequence:
                response = self.session.get(
                    self._endpoint(candidate_path),
                    auth=auth,
                    timeout=30,
                    **kwargs,
                )
                if response.ok:
                    return response
                last_response = response

                # Back off gently if edge/CDN is throttling or challenging.
                if response.status_code in (403, 429):
                    time.sleep(0.35)

                # If auth was rejected, try the same path once without auth.
                if response.status_code == 401 and fallback_auth_sequence:
                    for fallback_auth in fallback_auth_sequence:
                        fallback = self.session.get(
                            self._endpoint(candidate_path),
                            auth=fallback_auth,
                            timeout=30,
                            **kwargs,
                        )
                        if fallback.ok:
                            return fallback
                        last_response = fallback

        if last_response is not None:
            self._raise_for_response(last_response)
        raise requests.HTTPError(f"Unable to fetch endpoint for path: {path}")

    def get_course_by_slug(self, slug: str) -> dict[str, Any] | None:
        """Resolve WP course by slug."""
        if not slug:
            return None

        response = self._get_with_optional_auth('/wp-json/wp/v2/cursuri', prefer_auth=True, params={'slug': slug})
        data = response.json()
        if isinstance(data, list) and data:
            return data[0]
        return None

    def get_course(self, post_id: int) -> dict[str, Any]:
        """Fetch full WP post including ACF."""
        response = self._get_with_optional_auth(f'/wp-json/wp/v2/cursuri/{post_id}', prefer_auth=True)
        return response.json()

    def update_course_program(self, post_id: int, final_program: list[dict], auth=None) -> dict[str, Any]:
        """POST only acf.program back to WP."""
        payload = {
            'acf': {
                'program': final_program if final_program else False,
            }
        }
        last_response: requests.Response | None = None
        for candidate_path in self._rest_candidate_paths(f'/wp-json/wp/v2/cursuri/{post_id}'):
            response = self.session.post(
                self._endpoint(candidate_path),
                auth=auth or self.auth,
                json=payload,
                timeout=30,
            )
            if response.ok:
                return response.json()
            last_response = response
            if response.status_code in (403, 429):
                time.sleep(0.35)

        if last_response is not None:
            self._raise_for_response(last_response)
        raise requests.HTTPError(f"Unable to update endpoint for post_id={post_id}")


def extract_slug_from_permalink(url: str) -> str:
    """Return slug from course permalink."""
    parsed = urlparse((url or '').strip())
    path = (parsed.path or '').strip('/ ')
    if not path:
        return ''

    parts = [part for part in path.split('/') if part]
    if not parts:
        return ''
    return parts[-1]


def parse_single_ro_date(value: str) -> date:
    """Parse dd.mm.yyyy."""
    return datetime.strptime(value.strip(), "%d.%m.%Y").date()


def _normalize_excel_date_value(value: Any) -> str:
    if value is None:
        return ''

    if hasattr(value, 'to_pydatetime'):
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
    if not token or token.lower() == 'nan':
        return []

    token = token.replace('–', '-').replace('—', '-')

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
    if not text or text.lower() == 'nan':
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


def _normalize_program_rows(program: list[dict], today: date) -> list[dict[str, str]]:
    normalized: list[dict[str, str]] = []
    seen: set[str] = set()

    for row in program or []:
        raw = str((row or {}).get('data', '')).strip()
        if not raw:
            continue
        try:
            dt = parse_single_ro_date(raw)
        except ValueError:
            continue

        normalized_raw = dt.strftime("%d.%m.%Y")
        if dt >= today and normalized_raw not in seen:
            normalized.append({'data': normalized_raw})
            seen.add(normalized_raw)

    normalized.sort(key=lambda item: parse_single_ro_date(item['data']))
    return normalized


def build_final_program(existing_program: list, excel_dates: list[str], today: date) -> list[dict]:
    """Keep valid current WP dates, add Excel values as raw text, and dedupe."""
    seen: set[str] = set()
    result: list[dict[str, str]] = []

    for row in _normalize_program_rows(existing_program, today):
        raw = row['data']
        if raw not in seen:
            result.append({'data': raw})
            seen.add(raw)

    for raw in excel_dates:
        normalized = str(raw).strip()
        if not normalized:
            continue
        if normalized not in seen:
            result.append({'data': normalized})
            seen.add(normalized)
    return result


def valid_existing_program(existing_program: list, today: date) -> list[dict[str, str]]:
    return _normalize_program_rows(existing_program, today)
