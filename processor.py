from __future__ import annotations

import re
from typing import Iterable, Optional

from openpyxl.styles import PatternFill


LIGHT_GREEN_FILL = PatternFill(fill_type="solid", fgColor="C6EFCE")


def normalize_header(value: object) -> str:
    return str(value).strip().lower()


def get_headers(ws) -> dict[str, int]:
    headers: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=1, column=col).value
        if cell_value is None:
            continue
        header = normalize_header(cell_value)
        if header and header not in headers:
            headers[header] = col
    return headers


def find_column(ws, header_name: str, headers: Optional[dict[str, int]] = None) -> Optional[int]:
    local_headers = headers if headers is not None else get_headers(ws)
    return local_headers.get(normalize_header(header_name))


def find_first_column(
    ws,
    header_names: Iterable[str],
    headers: Optional[dict[str, int]] = None,
) -> Optional[int]:
    local_headers = headers if headers is not None else get_headers(ws)
    for name in header_names:
        col = local_headers.get(normalize_header(name))
        if col is not None:
            return col
    return None


def ensure_column_after(
    ws,
    anchor_header: str,
    new_header: str,
    headers: Optional[dict[str, int]] = None,
) -> int:
    local_headers = headers if headers is not None else get_headers(ws)

    existing_col = local_headers.get(normalize_header(new_header))
    if existing_col is not None:
        return existing_col

    anchor_col = local_headers.get(normalize_header(anchor_header))
    if anchor_col is None:
        raise ValueError(f'Column "{anchor_header}" not found.')

    insert_at = anchor_col + 1
    ws.insert_cols(insert_at)
    ws.cell(row=1, column=insert_at).value = new_header
    return insert_at


def fill_column_green(ws, col: int) -> None:
    max_row = ws.max_row
    for row in range(1, max_row + 1):
        ws.cell(row=row, column=col).fill = LIGHT_GREEN_FILL


def parse_yes(value: object) -> bool:
    return str(value).strip().lower() == "yes"


def clean_orientation(value: object) -> Optional[str]:
    if value is None:
        return None

    text = str(value).strip()
    if not text or ":" not in text:
        return None

    parts = text.split(":")
    if len(parts) != 2:
        return None

    hour_text = parts[0].strip()
    minute_text = parts[1].strip()

    if not hour_text.isdigit() or not minute_text.isdigit():
        return None

    hour = int(hour_text)
    minute = int(minute_text)

    if hour < 0 or hour > 12:
        return None
    if minute < 0 or minute > 59:
        return None

    if hour == 0:
        hour = 12

    return f"{hour}:{minute:02d}"


def orientation_to_degrees(orientation: object) -> Optional[int]:
    cleaned = clean_orientation(orientation)
    if not cleaned:
        return None

    hour_text, minute_text = cleaned.split(":")
    hour = int(hour_text)
    minute = int(minute_text)

    if hour == 12:
        hour = 0

    degrees = (hour * 30) + (minute * 0.5)
    return int(round(degrees))


def get_depth_percent(value: object) -> Optional[float]:
    if value is None:
        return None

    if isinstance(value, (int, float)):
        numeric = float(value)
        if numeric <= 1:
            return numeric * 100
        return numeric

    text = str(value).strip()
    if not text:
        return None

    text = text.replace("%", "").strip()

    try:
        numeric = float(text)
    except ValueError:
        return None

    return numeric


def is_girth_weld(feature_type: object) -> bool:
    return str(feature_type).strip().lower() == "girth weld"


def safe_str(value: object) -> str:
    return "" if value is None else str(value).strip()


def normalize_feature_type(value: object) -> str:
    return safe_str(value).lower()


def normalize_for_match(value: object) -> str:
    return re.sub(r"\s+", " ", safe_str(value).lower())
