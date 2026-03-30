from __future__ import annotations

from datetime import date, datetime
from typing import Any
from urllib.parse import urlparse
import re

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

    def _endpoint(self, path: str) -> str:
        return f"{self.base_url}{path}"

    def get_course_by_slug(self, slug: str) -> dict[str, Any] | None:
        """Resolve WP course by slug."""
        if not slug:
            return None

        response = requests.get(
            self._endpoint('/wp-json/wp/v2/cursuri'),
            params={'slug': slug},
            auth=self.auth,
            timeout=20,
        )
        response.raise_for_status()
        data = response.json()
        if isinstance(data, list) and data:
            return data[0]
        return None

    def get_course(self, post_id: int) -> dict[str, Any]:
        """Fetch full WP post including ACF."""
        response = requests.get(
            self._endpoint(f'/wp-json/wp/v2/cursuri/{post_id}'),
            auth=self.auth,
            timeout=20,
        )
        response.raise_for_status()
        return response.json()

    def update_course_program(self, post_id: int, final_program: list[dict], auth=None) -> dict[str, Any]:
        """POST only acf.program back to WP."""
        payload = {
            'acf': {
                'program': final_program if final_program else False,
            }
        }
        response = requests.post(
            self._endpoint(f'/wp-json/wp/v2/cursuri/{post_id}'),
            auth=auth or self.auth,
            json=payload,
            timeout=20,
        )
        response.raise_for_status()
        return response.json()


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


def expand_date_token(token: str) -> list[str]:
    """
    Expand one token:
    - 08.04.2026 -> ['08.04.2026']
    - 14-15.04.2026 -> ['14.04.2026', '15.04.2026']
    - 21-23.04.2026 -> ['21.04.2026', '22.04.2026', '23.04.2026']
    """
    token = str(token or '').strip()
    if not token:
        return []

    if re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", token):
        return [token]

    match = re.fullmatch(r"(\d{2})-(\d{2})\.(\d{2})\.(\d{4})", token)
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


def split_cell_tokens(value: str) -> list[str]:
    if value is None:
        return []
    text = str(value).strip()
    if not text or text.lower() == 'nan':
        return []
    return [part.strip() for part in re.split(r"[,\n;]+", text) if part and part.strip()]


def parse_excel_dates_from_row(row: dict) -> list[str]:
    """Read all month columns and return normalized dd.mm.yyyy strings."""
    dates: list[str] = []
    for column in MONTH_COLUMNS:
        raw = row.get(column)
        for token in split_cell_tokens(raw):
            dates.extend(expand_date_token(token))
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

        if dt >= today and raw not in seen:
            normalized.append({'data': raw})
            seen.add(raw)

    normalized.sort(key=lambda item: parse_single_ro_date(item['data']))
    return normalized


def build_final_program(existing_program: list, excel_dates: list[str], today: date) -> list[dict]:
    """Keep valid current dates, add Excel dates, dedupe, sort."""
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
        try:
            dt = parse_single_ro_date(normalized)
        except ValueError:
            continue

        if dt >= today and normalized not in seen:
            result.append({'data': normalized})
            seen.add(normalized)

    result.sort(key=lambda item: parse_single_ro_date(item['data']))
    return result


def valid_existing_program(existing_program: list, today: date) -> list[dict[str, str]]:
    return _normalize_program_rows(existing_program, today)
