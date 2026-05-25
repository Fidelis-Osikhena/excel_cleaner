from __future__ import annotations

import os
import csv
from openpyxl import load_workbook


JOINT_OUTPUT_COLUMNS = [
    "Upstream Girth Weld Number",
    "Previous US Girth Weld Number",
    "Odometer",
    "X_Coord",
    "Y_Coord",
    "Lat",
    "Long",
    "Height",
    "Wall Thickness",
    "Grade",
    "Diameter",
    "Seam Type",
    "Seam Position",
    "Comments",
    "MOP",
    "DPP",
    "Tool Speed",
    "Detectable Length",
]


def generate_dras_files(
    excel_path: str,
    output_folder: str,
    pipe_diameter: float,
) -> None:
    """
    Main DRAS export function.

    This is separate from ONEOK processing.
    It reads the imported Excel file and creates DRAS .txt files.
    """

    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active

    os.makedirs(output_folder, exist_ok=True)

    generate_joint_txt(
        ws=ws,
        output_folder=output_folder,
        pipe_diameter=pipe_diameter,
    )

    # will add the other 6 files here:
    # generate_file_2(...)
    # generate_file_3(...)
    # generate_file_4(...)
    # generate_file_5(...)
    # generate_file_6(...)
    # generate_file_7(...)


def generate_joint_txt(ws, output_folder: str, pipe_diameter: float) -> None:
    """
    Generates Joint.txt.

    Only imports rows where Joint = 1.
    Output is tab-delimited.
    """

    headers = get_headers(ws)

    output_path = os.path.join(output_folder, "Joint.txt")

    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=JOINT_OUTPUT_COLUMNS,
            delimiter="\t",
            lineterminator="\n",
        )

        writer.writeheader()

        for row in range(2, ws.max_row + 1):
            joint_value = get_value(ws, row, headers, ["Joint"])

            if not is_joint_row(joint_value):
                continue

            output_row = {
                "Upstream Girth Weld Number": get_value(ws, row, headers, ["Weld Id", "Weld ID"]),
                "Previous US Girth Weld Number": "",
                "Odometer": get_value(ws, row, headers, ["Odometer Main (m)", "Odometer Main (ft)"]),
                "X_Coord": get_value(ws, row, headers, ["Easting (m)", "Easting (ft)", "X Coord"]),
                "Y_Coord": get_value(ws, row, headers, ["Northing (m)", "Northing (ft)", "Y Coord"]),
                "Lat": get_value(ws, row, headers, ["Latitude", "Lat"]),
                "Long": get_value(ws, row, headers, ["Longitude", "Long"]),
                "Height": get_value(ws, row, headers, ["Height (m)", "Height (ft)"]),
                "Wall Thickness": get_value(ws, row, headers, ["Effective Depth (%)", "Wall Thickness"]),
                "Grade": get_value(ws, row, headers, ["SMYS (kPa)", "Grade"]),
                "Diameter": pipe_diameter,
                "Seam Type": get_dras_seam_type(ws, row, headers),
                "Seam Position": get_dras_seam_position(ws, row, headers),
                "Comments": get_dras_joint_comment(ws, row, headers),
                "MOP": get_mop_value(ws, row, headers),
                "DPP": get_value(ws, row, headers, ["DPP"]),
                "Tool Speed": get_value(ws, row, headers, ["Speed (m/s)", "Tool Speed"]),
                "Detectable Length": get_detectable_length(ws, row, headers),
            }

            writer.writerow(clean_output_row(output_row))


# ---------------------------
# JOINT.TXT RULE HELPERS
# ---------------------------

def get_dras_seam_type(ws, row: int, headers: dict[str, int]) -> str:
    """
    DRAS Seam Type rules:

    S = if source Seam Type is spiral
    L = if not spiral and Long Seam Orientation exists
    N = if source Seam Type is seamless
    blank = if Long Seam Orientation blank and not seamless
    """

    seam_type = normalize_text(get_value(ws, row, headers, ["Seam Type"]))
    long_seam = get_value(ws, row, headers, ["Long Seam Orientation", "Long Seam Orientation (Degree)"])

    if seam_type == "spiral":
        return "S"

    if seam_type == "seamless":
        return "N"

    if long_seam not in ("", None):
        return "L"

    return ""


def get_dras_seam_position(ws, row: int, headers: dict[str, int]) -> str:
    """
    DRAS Seam Position rule:

    Use Long Seam Orientation in degrees 0-360.
    If export already has degrees, copy it.
    If export has clock format, convert to degrees.
    """

    degree_value = get_value(ws, row, headers, ["Long Seam Orientation (Degree)"])
    if degree_value not in ("", None):
        return format_integer_if_possible(degree_value)

    clock_value = get_value(ws, row, headers, ["Long Seam Orientation"])
    degrees = clock_orientation_to_degrees(clock_value)

    return "" if degrees is None else str(degrees)


def get_dras_joint_comment(ws, row: int, headers: dict[str, int]) -> str:
    """
    Comments rule:
    concatenate Feature Type and comment.
    """

    feature = safe_str(get_value(ws, row, headers, ["Feature Type"]))
    comment = safe_str(get_value(ws, row, headers, ["Comments", "Comment Working"]))

    if feature and comment:
        return f"{feature} - {comment}"

    if feature:
        return feature

    return comment


def get_mop_value(ws, row: int, headers: dict[str, int]) -> str:
    """
    MOP rule:
    use MAOP or MOP, whichever is not blank.
    """

    maop = get_value(ws, row, headers, ["MAOP (kPa)"])
    mop = get_value(ws, row, headers, ["MOP (kPa)"])

    if maop not in ("", None):
        return maop

    if mop not in ("", None):
        return mop

    return ""


def get_detectable_length(ws, row: int, headers: dict[str, int]) -> str:
    """
    Detectable Length rule:
    If speed < 6 m/s, return 0.1.
    If speed >= 6 m/s, return blank.
    """

    speed_value = get_value(ws, row, headers, ["Speed (m/s)", "Tool Speed"])
    speed = to_float(speed_value)

    if speed is None:
        return ""

    if speed < 6:
        return "0.1"

    return ""


# ---------------------------
# GENERIC HELPERS
# ---------------------------

def get_headers(ws) -> dict[str, int]:
    """
    Returns normalized header -> column number.
    """

    headers = {}

    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=1, column=col).value
        if value is None:
            continue

        headers[normalize_header(value)] = col

    return headers


def normalize_header(value: object) -> str:
    return str(value).strip().lower()


def normalize_text(value: object) -> str:
    return safe_str(value).strip().lower()


def get_value(ws, row: int, headers: dict[str, int], possible_headers: list[str]):
    """
    Tries multiple possible header names and returns the first matching value.
    """

    for header in possible_headers:
        col = headers.get(normalize_header(header))
        if col is not None:
            value = ws.cell(row=row, column=col).value
            return "" if value is None else value

    return ""


def is_joint_row(value: object) -> bool:
    """
    Only rows where Joint = 1 are exported to Joint.txt.
    """

    if value is None:
        return False

    text = str(value).strip().lower()

    return text in {"1", "1.0", "yes", "y", "true"}


def safe_str(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def to_float(value: object) -> float | None:
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    try:
        return float(text)
    except ValueError:
        return None


def format_integer_if_possible(value: object) -> str:
    num = to_float(value)

    if num is None:
        return safe_str(value)

    if num.is_integer():
        return str(int(num))

    return str(num)


def clock_orientation_to_degrees(value: object) -> int | None:
    """
    Converts clock orientation to degrees.

    12:00 = 0
    3:00 = 90
    6:00 = 180
    9:00 = 270
    """

    if value is None:
        return None

    text = str(value).strip()
    if not text or ":" not in text:
        return None

    try:
        hour_text, minute_text = text.split(":")
        hour = int(hour_text)
        minute = int(minute_text)

        if hour == 12:
            hour = 0

        degrees = (hour * 30) + (minute * 0.5)
        return int(round(degrees))

    except ValueError:
        return None


def clean_output_row(row: dict[str, object]) -> dict[str, str]:
    """
    Converts all output values to strings for tab-delimited export.
    """

    cleaned = {}

    for key, value in row.items():
        if value is None:
            cleaned[key] = ""
        else:
            cleaned[key] = str(value).strip()

    return cleaned