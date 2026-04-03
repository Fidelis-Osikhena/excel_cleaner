from __future__ import annotations

from openpyxl import Workbook

from processor import (
    clean_orientation,
    ensure_column_after,
    fill_column_green,
    find_column,
    find_first_column,
    get_depth_percent,
    get_headers,
    is_girth_weld,
    normalize_feature_type,
    orientation_to_degrees,
    parse_yes,
)


def process_oneok_workbook(wb: Workbook) -> None:
    ws = wb.active

    process_orientation_columns(ws)
    process_delta_columns(ws)
    process_feature_type_final(ws)
    process_comments(ws)


def process_orientation_columns(ws) -> None:
    headers = get_headers(ws)

    main_col = headers.get("orientation main")
    if main_col is None:
        raise ValueError('Column "Orientation Main" not found.')

    if "orientation final" not in headers:
        ensure_column_after(ws, "Orientation Main", "Orientation Final", headers)
        headers = get_headers(ws)

    if "orientation final (degree)" not in headers:
        ensure_column_after(ws, "Orientation Final", "Orientation Final (Degree)", headers)
        headers = get_headers(ws)

    long_col = headers.get("long seam orientation")
    if long_col is not None and "long seam orientation (degree)" not in headers:
        ws.insert_cols(long_col + 1)
        ws.cell(row=1, column=long_col + 1).value = "Long Seam Orientation (Degree)"
        headers = get_headers(ws)

    main_col = headers["orientation main"]
    final_col = headers["orientation final"]
    final_deg_col = headers["orientation final (degree)"]
    long_col = headers.get("long seam orientation")
    long_deg_col = headers.get("long seam orientation (degree)")
    feature_col = headers.get("feature type")

    max_row = ws.max_row

    for row in range(2, max_row + 1):
        main_value = ws.cell(row=row, column=main_col).value
        cleaned_main = clean_orientation(main_value)

        if cleaned_main:
            ws.cell(row=row, column=final_col).value = cleaned_main
            ws.cell(row=row, column=final_deg_col).value = orientation_to_degrees(cleaned_main)
        else:
            ws.cell(row=row, column=final_col).value = ""
            ws.cell(row=row, column=final_deg_col).value = ""

        if long_col is None or long_deg_col is None:
            continue

        feature_value = ws.cell(row=row, column=feature_col).value if feature_col else None
        long_value = ws.cell(row=row, column=long_col).value

        if is_girth_weld(feature_value):
            cleaned_long = clean_orientation(long_value)

            if cleaned_long:
                long_deg = orientation_to_degrees(cleaned_long)

                ws.cell(row=row, column=long_col).value = cleaned_long
                ws.cell(row=row, column=long_deg_col).value = long_deg
                ws.cell(row=row, column=final_col).value = cleaned_long
                ws.cell(row=row, column=final_deg_col).value = long_deg
            else:
                ws.cell(row=row, column=long_col).value = ""
                ws.cell(row=row, column=long_deg_col).value = ""
                ws.cell(row=row, column=final_col).value = ""
                ws.cell(row=row, column=final_deg_col).value = ""
        else:
            ws.cell(row=row, column=long_col).value = ""
            ws.cell(row=row, column=long_deg_col).value = ""

    for row in range(2, max_row + 1):
        ws.cell(row=row, column=final_col).number_format = "@"
        ws.cell(row=row, column=final_deg_col).number_format = "0"

        if long_col is not None and long_deg_col is not None:
            ws.cell(row=row, column=long_col).number_format = "@"
            ws.cell(row=row, column=long_deg_col).number_format = "0"

    fill_column_green(ws, final_col)
    fill_column_green(ws, final_deg_col)

    if long_col is not None and long_deg_col is not None:
        fill_column_green(ws, long_col)
        fill_column_green(ws, long_deg_col)


def process_delta_columns(ws) -> None:
    headers = get_headers(ws)

    feature_col = headers.get("feature type")
    is_marker_col = headers.get("ismarker")

    weld_delta_col = _find_delta_column_from_headers(
        headers,
        [
            "US Weld Δ main (ft)",
            "US Weld ? main (ft)",
            "US Weld Delta main (ft)",
        ],
    )

    marker_delta_col = _find_delta_column_from_headers(
        headers,
        [
            "US Marker Δ main (ft)",
            "US Marker ? main (ft)",
            "US Marker Delta main (ft)",
        ],
    )

    max_row = ws.max_row

    if feature_col is not None and weld_delta_col is not None:
        for row in range(2, max_row + 1):
            if is_girth_weld(ws.cell(row=row, column=feature_col).value):
                ws.cell(row=row, column=weld_delta_col).value = 0

        fill_column_green(ws, weld_delta_col)

    if is_marker_col is not None and marker_delta_col is not None:
        for row in range(2, max_row + 1):
            if parse_yes(ws.cell(row=row, column=is_marker_col).value):
                ws.cell(row=row, column=marker_delta_col).value = 0

        fill_column_green(ws, marker_delta_col)


def process_feature_type_final(ws) -> None:
    headers = get_headers(ws)

    feature_col = headers.get("feature type")
    if feature_col is None:
        return

    if "feature type final" not in headers:
        ensure_column_after(ws, "Feature Type", "Feature Type Final", headers)
        headers = get_headers(ws)

    feature_col = headers["feature type"]
    final_col = headers["feature type final"]
    is_marker_col = headers.get("ismarker")
    depth_col = find_first_column(ws, ["Depth (%)", "Depth %", "Depth", "% Depth"], headers)

    max_row = ws.max_row

    for row in range(2, max_row + 1):
        feature_value = ws.cell(row=row, column=feature_col).value
        marker_value = ws.cell(row=row, column=is_marker_col).value if is_marker_col else None
        depth_value = ws.cell(row=row, column=depth_col).value if depth_col else None

        final_value = map_feature_type_final(
            feature_value=feature_value,
            is_marker_value=marker_value,
            depth_value=depth_value,
        )
        ws.cell(row=row, column=final_col).value = final_value

    fill_column_green(ws, final_col)


def process_comments(ws) -> None:
    headers = get_headers(ws)

    feature_col = headers.get("feature type")
    if feature_col is None:
        return

    if "comments" not in headers:
        anchor = "Feature Type Final" if "feature type final" in headers else "Feature Type"
        ensure_column_after(ws, anchor, "Comments", headers)
        headers = get_headers(ws)

    feature_col = headers["feature type"]
    comments_col = headers["comments"]

    bend_radius_col = headers.get("bend radius (xd)")
    bend_angle_col = headers.get("bend angle")
    bend_orientation_col = headers.get("bend orientation")

    max_row = ws.max_row

    for row in range(2, max_row + 1):
        feature_value = ws.cell(row=row, column=feature_col).value
        feature = normalize_feature_type(feature_value)

        comment = ""

        if feature == "bend":
            radius = ws.cell(row=row, column=bend_radius_col).value if bend_radius_col else None
            angle = ws.cell(row=row, column=bend_angle_col).value if bend_angle_col else None
            orientation = ws.cell(row=row, column=bend_orientation_col).value if bend_orientation_col else None
            comment = comment_bend_fast(radius, angle, orientation)

        ws.cell(row=row, column=comments_col).value = comment

    fill_column_green(ws, comments_col)


def comment_bend_fast(radius_value: object, angle_value: object, orientation_value: object) -> str:
    radius = format_one_decimal(radius_value)
    angle = format_one_decimal(angle_value)
    direction = bend_orientation_to_direction(orientation_value)
    return f"Bend - {radius} - {angle} - {direction}"


def format_one_decimal(value: object) -> str:
    if value is None:
        return ""

    text = str(value).strip()
    if not text:
        return ""

    try:
        return f"{float(text):.1f}"
    except (ValueError, TypeError):
        return text


def bend_orientation_to_direction(value: object) -> str:
    cleaned = clean_orientation(value)
    if not cleaned:
        return ""

    hour_text, minute_text = cleaned.split(":")
    hour = int(hour_text)
    minute = int(minute_text)

    total_minutes = bend_time_to_minutes(hour, minute)

    if total_minutes >= 630 or total_minutes <= 89:
        return "up"
    if 90 <= total_minutes <= 269:
        return "right"
    if 270 <= total_minutes <= 449:
        return "down"
    if 450 <= total_minutes <= 629:
        return "left"

    return ""


def bend_time_to_minutes(hour: int, minute: int) -> int:
    if hour == 12:
        hour = 0
    return hour * 60 + minute


def map_feature_type_final(
    *,
    feature_value: object,
    is_marker_value: object,
    depth_value: object,
) -> str:
    feature = normalize_feature_type(feature_value)

    if feature == "bend":
        return "LOCATION"

    if feature == "cp attachment":
        return "GAIN"

    if feature == "crack-like":
        return ""

    if feature in {"deformation", "deformation - ovality", "deformation w/ metal loss"}:
        depth_percent = get_depth_percent(depth_value)
        if depth_percent is not None and depth_percent >= 1.0:
            return "DENT"
        return "GMA"

    if feature in {
        "half sole repair",
        "magnet",
        "ncf-a",
        "ncf-b",
        "patch repair",
        "puddle weld repair",
        "repair marker begin",
        "repair marker end",
        "river weight",
        "weld anomaly",
        "recoat begin",
        "recoat end",
    }:
        return "MISC"

    if feature == "manufacturing anomaly":
        return "MILL ANOMALY"

    if feature in {"metal loss", "metal loss manufacturing"}:
        return "GROUP"

    if feature in {"metal object - close", "metal object - touching"}:
        return "GAIN"

    if feature in {
        "pipe support - rectangular",
        "stopple",
        "tap",
        "tee",
        "casing begin",
        "casing end",
        "flange",
        "marker band begin",
        "marker band end",
        "pipe support - circumferential",
        "sleeve begin",
        "sleeve end",
        "valve",
    }:
        if feature == "valve":
            return "AGM" if parse_yes(is_marker_value) else "LOCATION"
        return "LOCATION"

    if feature in {"sswc", "swa"}:
        return "SW GROUP"

    if feature == "swf-a":
        return "SW-A"

    if feature == "swf-b":
        return "SW-B"

    if feature == "agm":
        return "AGM"

    if feature == "girth weld":
        return "WELD"

    if feature == "coupling":
        return ""

    return ""


def _find_delta_column_from_headers(headers: dict[str, int], candidates: list[str]) -> int | None:
    for candidate in candidates:
        col = headers.get(candidate.strip().lower())
        if col is not None:
            return col
    return None
