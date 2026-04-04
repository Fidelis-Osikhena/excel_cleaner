from __future__ import annotations

import math
from openpyxl import Workbook

from processor import (
    clean_orientation,
    ensure_column_after,
    fill_column_green,
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
    process_length_width_final(ws)
    process_feature_type_final(ws)
    process_comments(ws)


# ---------------------------
# ORIENTATION
# ---------------------------

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


# ---------------------------
# DELTA
# ---------------------------

def process_delta_columns(ws) -> None:
    headers = get_headers(ws)

    feature_col = headers.get("feature type")
    is_marker_col = headers.get("ismarker")

    weld_delta_col = _find_delta_column(headers, [
        "US Weld Δ main (ft)",
        "US Weld ? main (ft)",
        "US Weld Delta main (ft)",
    ])

    marker_delta_col = _find_delta_column(headers, [
        "US Marker Δ main (ft)",
        "US Marker ? main (ft)",
        "US Marker Delta main (ft)",
    ])

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


# ---------------------------
# LENGTH / WIDTH FINAL
# ---------------------------

def process_length_width_final(ws) -> None:
    headers = get_headers(ws)

    feature_col = headers.get("feature type")
    length_col = headers.get("length (in)")
    width_col = headers.get("width (in)")

    if feature_col is None or length_col is None or width_col is None:
        return

    if "length (in) final" not in headers:
        ensure_column_after(ws, "Length (in)", "Length (in) Final", headers)
        headers = get_headers(ws)

    if "width (in) final" not in headers:
        ensure_column_after(ws, "Width (in)", "Width (in) Final", headers)
        headers = get_headers(ws)

    feature_col = headers["feature type"]
    length_col = headers["length (in)"]
    width_col = headers["width (in)"]
    length_final_col = headers["length (in) final"]
    width_final_col = headers["width (in) final"]

    target_features = {
        "cp attachment",
        "metal object - touching",
        "pipe support - rectangular",
        "river weight",
        "stopple",
        "tap",
        "tee",
    }

    for row in range(2, ws.max_row + 1):
        feature = normalize_feature_type(ws.cell(row=row, column=feature_col).value)

        length_final = ""
        width_final = ""

        if feature in target_features:
            length_val = ws.cell(row=row, column=length_col).value
            width_val = ws.cell(row=row, column=width_col).value

            rounded_length = _round_to_zero_decimal(length_val)
            rounded_width = _round_to_zero_decimal(width_val)

            length_final = "" if rounded_length is None else rounded_length
            width_final = "" if rounded_width is None else rounded_width

        ws.cell(row=row, column=length_final_col).value = length_final
        ws.cell(row=row, column=width_final_col).value = width_final

    for row in range(2, ws.max_row + 1):
        ws.cell(row=row, column=length_final_col).number_format = "0"
        ws.cell(row=row, column=width_final_col).number_format = "0"

    fill_column_green(ws, length_final_col)
    fill_column_green(ws, width_final_col)

# ---------------------------
# FEATURE TYPE FINAL
# ---------------------------

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

        ws.cell(row=row, column=final_col).value = map_feature_type_final(
            feature_value=feature_value,
            is_marker_value=marker_value,
            depth_value=depth_value,
        )

    fill_column_green(ws, final_col)


def map_feature_type_final(*, feature_value: object, is_marker_value: object, depth_value: object) -> str:
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


# ---------------------------
# COMMENTS
# ---------------------------

def process_comments(ws) -> None:
    headers = get_headers(ws)

    feature_col = headers.get("feature type")
    if feature_col is None:
        return

    if "comments" not in headers:
        anchor = "Feature Type Final" if "feature type final" in headers else "Feature Type"
        ensure_column_after(ws, anchor, "Comments", headers)
        headers = get_headers(ws)

    if "comment working 2 (gw proximity)" not in headers:
        ensure_column_after(ws, "Comments", "Comment Working 2 (GW Proximity)", headers)
        headers = get_headers(ws)

    feature_col = headers["feature type"]
    comments_col = headers["comments"]
    gw_col = headers["comment working 2 (gw proximity)"]
    feature_final_col = headers.get("feature type final")

    bend_radius_col = headers.get("bend radius (xd)")
    bend_angle_col = headers.get("bend angle")
    bend_orientation_col = headers.get("bend orientation")
    is_external_col = headers.get("is external")
    is_marker_col = headers.get("ismarker")
    length_col = headers.get("length (in)")
    width_col = headers.get("width (in)")
    orientation_final_col = headers.get("orientation final")
    orientation_main_col = headers.get("orientation main")

    us_left_col = _find_delta_column(headers, [
        "US Weld Δ from left edge (ft)",
        "US Weld ? from left edge (ft)",
        "US Weld Delta from left edge (ft)",
    ])
    ds_left_col = _find_delta_column(headers, [
        "DS Weld Δ from left edge (ft)",
        "DS Weld ? from left edge (ft)",
        "DS Weld Delta from left edge (ft)",
    ])
    us_right_col = _find_delta_column(headers, [
        "US Weld Δ from right edge (ft)",
        "US Weld ? from right edge (ft)",
        "US Weld Delta from right edge (ft)",
    ])
    ds_right_col = _find_delta_column(headers, [
        "DS Weld Δ from right edge (ft)",
        "DS Weld ? from right edge (ft)",
        "DS Weld Delta from right edge (ft)",
    ])

    max_row = ws.max_row

    for row in range(2, max_row + 1):
        feature_value = ws.cell(row=row, column=feature_col).value
        feature = normalize_feature_type(feature_value)

        existing_comment = ws.cell(row=row, column=comments_col).value
        existing_comment_text = "" if existing_comment is None else str(existing_comment).strip()

        feature_final = normalize_feature_type(
            ws.cell(row=row, column=feature_final_col).value if feature_final_col else None
        )

        is_marker_value = ws.cell(row=row, column=is_marker_col).value if is_marker_col else None

        comment = ""
        gw_comment = ""

        # Preserve existing comments
        if feature in {
            "repair marker begin",
            "agm",
            "casing begin",
            "girth weld",
            "marker band begin",
            "recoat begin",
            "sleeve begin",
        }:
            comment = existing_comment_text

        # Valve
        elif feature == "valve":
            if parse_yes(is_marker_value):
                comment = existing_comment_text
            else:
                comment = str(feature_value).strip() if feature_value else ""

        # Bend
        elif feature == "bend":
            radius = ws.cell(row=row, column=bend_radius_col).value if bend_radius_col else None
            angle = ws.cell(row=row, column=bend_angle_col).value if bend_angle_col else None
            orientation = ws.cell(row=row, column=bend_orientation_col).value if bend_orientation_col else None
            comment = comment_bend(radius, angle, orientation)

        # Simple copy rules
        elif feature in {
            "half sole repair",
            "magnet",
            "patch repair",
            "pipe support - rectangular",
            "puddle weld repair",
            "repair marker end",
            "river weight",
            "casing end",
            "flange",
            "marker band end",
            "pipe support - circumferential",
            "recoat end",
            "sleeve end",
        }:
            comment = str(feature_value).strip() if feature_value else ""

        # Deformation
        elif feature in {"deformation", "deformation - ovality", "deformation w/ metal loss"}:
            comment = "Deformation"

        # Manufacturing Anomaly
        elif feature == "manufacturing anomaly":
            comment = "Manufacturing Anomaly"

        # Metal Loss / SWA
        elif feature in {"metal loss", "metal loss manufacturing", "swa"}:
            ext_val = ws.cell(row=row, column=is_external_col).value if is_external_col else None
            side = "External" if parse_yes(ext_val) else "Internal"
            comment = f"Metal Loss - {side}"

        # NCF
        elif feature in {"ncf-a", "ncf-b"}:
            length_value = ws.cell(row=row, column=length_col).value if length_col else None
            width_value = ws.cell(row=row, column=width_col).value if width_col else None
            comment = comment_ncf(length_value, width_value)

        # SSWC
        elif feature == "sswc":
            comment = "Possible selective seam weld corrosion"

        # Weld Anomaly
        elif feature == "weld anomaly":
            comment = "Possible Girth Weld Anomaly"

        # Stopple
        elif feature == "stopple":
            orientation_value = (
                ws.cell(row=row, column=orientation_final_col).value
                if orientation_final_col else
                (ws.cell(row=row, column=orientation_main_col).value if orientation_main_col else None)
            )
            comment = comment_stopple(orientation_value)

        # Tap
        elif feature == "tap":
            orientation_value = (
                ws.cell(row=row, column=orientation_final_col).value
                if orientation_final_col else
                (ws.cell(row=row, column=orientation_main_col).value if orientation_main_col else None)
            )
            comment = comment_tap(orientation_value)

        # Tee
        elif feature == "tee":
            orientation_value = (
                ws.cell(row=row, column=orientation_final_col).value
                if orientation_final_col else
                (ws.cell(row=row, column=orientation_main_col).value if orientation_main_col else None)
            )
            comment = comment_tee(orientation_value)

        # Gain rule
        elif feature_final == "gain":
            if "attachment" in existing_comment_text.lower():
                comment = existing_comment_text
            else:
                comment = str(feature_value).strip() if feature_value else ""

        # GW proximity in separate column
        if feature in {
            "deformation",
            "deformation - ovality",
            "deformation w/ metal loss",
            "manufacturing anomaly",
            "metal loss",
            "metal loss manufacturing",
            "ncf-a",
            "ncf-b",
            "sswc",
            "stopple",
            "swa",
        }:
            if is_adjacent_to_girth_weld(
                ws, row,
                us_left_col,
                ds_left_col,
                us_right_col,
                ds_right_col,
            ):
                gw_comment = "adjacent to girth weld"

        ws.cell(row=row, column=comments_col).value = comment
        ws.cell(row=row, column=gw_col).value = gw_comment

    fill_column_green(ws, comments_col)
    fill_column_green(ws, gw_col)


# ---------------------------
# COMMENT HELPERS
# ---------------------------

def comment_bend(radius_value: object, angle_value: object, orientation_value: object) -> str:
    radius = format_one_decimal(radius_value)
    angle = format_one_decimal(angle_value)
    direction_text = bend_orientation_to_direction(orientation_value)
    return f"Bend - {radius} - {angle} - {direction_text}"


def comment_ncf(length_value: object, width_value: object) -> str:
    degrees = ncf_degrees_from_y_axis(length_value, width_value)
    if degrees is None:
        return "Possible Narrow Circumferential Feature"
    return f"Possible Narrow Circumferential Feature, {degrees} degrees from y-axis"


def comment_stopple(orientation_value: object) -> str:
    position = orientation_position(orientation_value)
    if position == "top":
        return "STOPPLE on top of pipe"
    if position == "90":
        return "STOPPLE at 90 degrees"
    if position == "bottom":
        return "STOPPLE on bottom of pipe"
    if position == "270":
        return "STOPPLE at 270 degrees"
    return "STOPPLE"


def comment_tap(orientation_value: object) -> str:
    position = orientation_position(orientation_value)
    if position == "top":
        return "Fitting on top of pipe"
    if position == "90":
        return "Tap at 90 degrees"
    if position == "bottom":
        return "Fitting on bottom of pipe"
    if position == "270":
        return "Tap at 270 degrees"
    return "Tap"


def comment_tee(orientation_value: object) -> str:
    position = orientation_position(orientation_value)
    if position == "top":
        return "TEE on top of pipe"
    if position == "90":
        return "TEE at 90 degrees"
    if position == "bottom":
        return "TEE on bottom of pipe"
    if position == "270":
        return "TEE at 270 degrees"
    return "TEE"


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

    hour, minute = map(int, cleaned.split(":"))
    total_minutes = _clock_minutes(hour, minute)

    if total_minutes >= 630 or total_minutes <= 89:
        return "up"
    if 90 <= total_minutes <= 269:
        return "right"
    if 270 <= total_minutes <= 449:
        return "down"
    if 450 <= total_minutes <= 629:
        return "left"
    return ""


def orientation_position(value: object) -> str:
    cleaned = clean_orientation(value)
    if not cleaned:
        return ""

    hour, minute = map(int, cleaned.split(":"))
    total_minutes = _clock_minutes(hour, minute)

    if total_minutes >= 630 or total_minutes <= 89:
        return "top"
    if 90 <= total_minutes <= 269:
        return "90"
    if 270 <= total_minutes <= 449:
        return "bottom"
    if 450 <= total_minutes <= 629:
        return "270"
    return ""


def _clock_minutes(hour: int, minute: int) -> int:
    if hour == 12:
        hour = 0
    return hour * 60 + minute


def ncf_degrees_from_y_axis(length_value: object, width_value: object) -> str | None:
    try:
        length_num = float(length_value)
        width_num = float(width_value)
    except (TypeError, ValueError):
        return None

    if width_num == 0:
        return None

    degrees = math.degrees(math.atan(length_num / width_num))
    return f"{degrees:.0f}"


def is_adjacent_to_girth_weld(ws, row: int, us_left_col, ds_left_col, us_right_col, ds_right_col) -> bool:
    for col in [us_left_col, ds_left_col, us_right_col, ds_right_col]:
        if not col:
            continue

        value = ws.cell(row=row, column=col).value
        num = _to_float(value)

        if num is not None and abs(num) <= 0.04:
            return True

    return False


def append_adjacent_to_girth_weld(comment: str) -> str:
    suffix = "adjacent to girth weld"
    if suffix in comment.lower():
        return comment
    return f"{comment} {suffix}"


def _to_float(value: object) -> float | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def _find_delta_column(headers: dict[str, int], candidates: list[str]) -> int | None:
    for candidate in candidates:
        col = headers.get(candidate.strip().lower())
        if col is not None:
            return col
    return None


def _round_to_zero_decimal(value: object) -> int | None:
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    try:
        return int(round(float(text)))
    except (TypeError, ValueError):
        return None
