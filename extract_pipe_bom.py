from pathlib import Path
import csv
import logging
import re
from typing import Iterator, List, Any

import ezdxf
from ezdxf.lldxf.const import DXFStructureError
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

INPUT_LAYER = "GT_1"
OUTPUT_FILE = "pipe_bom_export.csv"
OUTPUT_EXCEL_FILE = "pipe_bom_export.xlsx"
PIPE_PREFIX = "PIPE B36.10"
CSV_HEADERS = [
    "Drawing Number",
    "Rev.",
    "Date",
    "Reason for issue",
    "Area",
    "Pipe description",
    "material",
    "diam",
    "Sch.",
    "material code",
    "Qty.",
]

SCHEDULE_PATTERN = re.compile(r"\s*-?\s*Sch\.\s*([^\s,;]+)", re.IGNORECASE)
SCHEDULE_ONLY_PATTERN = re.compile(r"^\s*Sch\.\s*(.+?)\s*$", re.IGNORECASE)
QUANTITY_METERS_PATTERN = re.compile(r"\s*M\s*$", re.IGNORECASE)
MATERIAL_PATTERN = re.compile(r"\b(A\d{2,4})\b", re.IGNORECASE)


def extract_rev(dwg_name: str) -> str:
    """Extract revision number from drawing name, defaulting to '0'."""
    match = re.search(r"_Rev(\w+)", dwg_name, re.IGNORECASE)
    if match:
        return match.group(1)
    match = re.search(r"_R(\d+)", dwg_name, re.IGNORECASE)
    if match:
        return match.group(1)
    return "0"


def extract_area(dwg_name: str) -> str:
    """Extract 4-digit Area code from drawing name (e.g., '8906' from '...-8906-8960-...')."""
    match = re.search(r"-(\d{4})-\d{4}-", dwg_name)
    if match:
        return match.group(1)
    match = re.search(r"-(\d{4})-", dwg_name)
    if match:
        return match.group(1)
    return ""


def split_schedule(description: str) -> tuple[str, str]:
    """Return a description without its schedule and the schedule without 'Sch.'."""
    match = SCHEDULE_PATTERN.search(description)
    if not match:
        return description, ""

    schedule = match.group(1)
    description_without_schedule = (
        description[: match.start()] + description[match.end() :]
    ).strip(" -")
    return description_without_schedule, schedule


def schedule_value(text: str) -> str | None:
    """Return the value of a standalone schedule TEXT, without the 'Sch.' prefix."""
    match = SCHEDULE_ONLY_PATTERN.match(text)
    return match.group(1) if match else None


def normalize_quantity(quantity: str) -> str:
    """Remove a trailing metres unit from the exported quantity."""
    return QUANTITY_METERS_PATTERN.sub("", quantity).strip()


def extract_material(description: str) -> str:
    """Extract the first ASTM-style material code, such as A333, from a description."""
    match = MATERIAL_PATTERN.search(description)
    return match.group(1).upper() if match else ""


def extract_modelspace_text_entities(filepath: Path) -> List[tuple[str, str]]:
    """Fast, low-level line-by-line parsing of ASCII DXF file to extract TEXT entities in Model Space.
    Returns a list of (text_value, layer_name) tuples.
    """
    text_entities = []
    in_entities_section = False

    current_entity_type = None
    current_text = ""
    current_layer = ""
    is_paperspace = False
    in_model_layout = True

    try:
        with open(filepath, "r", encoding="utf-8", errors="surrogateescape") as f:
            lines = f.readlines()
    except Exception:
        with open(filepath, "r", encoding="cp1252", errors="surrogateescape") as f:
            lines = f.readlines()

    it = iter(lines)
    while True:
        try:
            line_code = next(it)
        except StopIteration:
            break
        try:
            line_val = next(it)
        except StopIteration:
            break

        code_str = line_code.strip()
        if not code_str:
            continue
        try:
            code = int(code_str)
        except ValueError:
            continue

        val = line_val.strip()

        if code == 0:
            # Save the completed TEXT entity if we were building one
            if in_entities_section and current_entity_type == "TEXT":
                if not is_paperspace and in_model_layout:
                    text_entities.append((current_text.strip(), current_layer))

            # Reset entity fields for the next entity
            current_entity_type = None
            current_text = ""
            current_layer = ""
            is_paperspace = False
            in_model_layout = True

            # Section boundaries or new entities
            if val == "SECTION":
                try:
                    next_code_line = next(it)
                    next_val_line = next(it)
                    next_code = int(next_code_line.strip())
                    next_val = next_val_line.strip()
                    if next_code == 2 and next_val == "ENTITIES":
                        in_entities_section = True
                except (StopIteration, ValueError):
                    pass
            elif val == "ENDSEC":
                in_entities_section = False
            elif in_entities_section:
                current_entity_type = val
        else:
            if in_entities_section and current_entity_type == "TEXT":
                if code == 1:
                    current_text = val
                elif code == 8:
                    current_layer = val
                elif code == 67:
                    is_paperspace = val == "1"
                elif code == 410:
                    in_model_layout = val.upper() == "MODEL"

    return text_entities


def extract_pipe_rows(dxf_path: Path) -> Iterator[List[str]]:
    """Extract pipe rows along with revision Date and Status from the DXF file."""
    all_text_entities = []
    use_fallback = False
    try:
        all_text_entities = extract_modelspace_text_entities(dxf_path)
        if not all_text_entities:
            use_fallback = True
    except Exception as exc:
        logging.debug(
            "Fast parser failed on %s: %s. Falling back to ezdxf.", dxf_path.name, exc
        )
        use_fallback = True

    if use_fallback:
        try:
            document = ezdxf.readfile(dxf_path)
            all_text_entities = [
                (entity.dxf.text, entity.dxf.layer)
                for entity in document.modelspace()
                if entity.dxftype() == "TEXT"
            ]
        except (IOError, OSError, DXFStructureError, ezdxf.DXFError) as exc:
            logging.warning("Skipping invalid DXF %s: %s", dxf_path.name, exc)
            return

    # Extract date and status from all TEXT entities in modelspace
    date_val, status_val = "", ""
    date_pattern = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
    status_keywords = {"ISSUED", "IFC", "CONSTRUCTION"}

    for i, (text, _) in enumerate(all_text_entities):
        t = text.strip()
        if date_pattern.match(t):
            date_val = t
            for j in range(i + 1, min(i + 5, len(all_text_entities))):
                nt = all_text_entities[j][0].strip()
                if any(kw in nt.upper() for kw in status_keywords):
                    status_val = nt
                    break
            break

    # Extract GT_1 TEXT values for the pipe BOM parser
    values = [
        text.strip() for text, layer in all_text_entities if layer == INPUT_LAYER
    ]

    dwg_name = dxf_path.stem
    rev_val = extract_rev(dwg_name)
    area_val = extract_area(dwg_name)

    index = 0
    while index < len(values):
        if values[index].upper().startswith(PIPE_PREFIX):
            if index + 3 >= len(values):
                logging.warning(
                    "Incomplete pipe record in %s after TEXT value %r",
                    dxf_path.name,
                    values[index],
                )
                index += 1
                continue

            description = values[index]
            standalone_schedule = schedule_value(values[index + 1])

            if standalone_schedule is not None:
                if index + 4 >= len(values):
                    logging.warning(
                        "Incomplete pipe record in %s after TEXT value %r",
                        dxf_path.name,
                        description,
                    )
                    index += 1
                    continue
                schedule = standalone_schedule
                diameter, material_code, quantity = values[index + 2 : index + 5]
                index += 5
            else:
                description, schedule = split_schedule(description)
                diameter, material_code, quantity = values[index + 1 : index + 4]
                index += 4

            yield [
                dwg_name,
                rev_val,
                date_val,
                status_val,
                area_val,
                description,
                extract_material(description),
                diameter,
                schedule,
                material_code,
                normalize_quantity(quantity),
            ]
        else:
            index += 1


def process_dxf_file(dxf_path: Path) -> List[List[str]]:
    """Worker function for multiprocessing pool to parse a single DXF file."""
    return list(extract_pipe_rows(dxf_path))


def export_to_excel(headers: List[str], all_rows: List[List[str]], excel_path: Path) -> None:
    """Export extracted pipe BOM data to formatted Excel sheet matching user template."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pipe BOM"

    # Header Row
    ws.append(headers)

    green_header_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_align = Alignment(horizontal="left", vertical="center")

    yellow_data_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    data_font = Font(name="Calibri", size=11, color="000000")

    thin_side = Side(style="thin", color="D9D9D9")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # Format Header Cells
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = green_header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = thin_border

    # Format & Insert Data Rows
    for row_idx, row_data in enumerate(all_rows, start=2):
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            col_name = headers[col_idx - 1]

            # Convert numeric strings to numerical types (float/int) for Excel
            typed_val: Any = val
            if col_name in ("Qty.", "diam"):
                clean_str = str(val).replace(",", ".").strip()
                try:
                    typed_val = float(clean_str) if "." in clean_str else int(clean_str)
                except ValueError:
                    typed_val = val
            elif col_name == "Rev.":
                try:
                    typed_val = int(val)
                except ValueError:
                    typed_val = val

            cell.value = typed_val
            cell.fill = yellow_data_fill
            cell.font = data_font
            cell.border = thin_border

            # Text alignment per column type
            if col_name in ("Qty.", "diam"):
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_name in ("Rev.", "Date", "Area", "Sch."):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Column Auto-Width
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            val_str = str(cell.value if cell.value is not None else "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Enable AutoFilter dropdowns on headers
    ws.auto_filter.ref = ws.dimensions

    wb.save(excel_path)
    logging.info("Wrote Excel file to %s", excel_path)


def main() -> None:
    folder = Path.cwd()
    output_path = folder / OUTPUT_FILE
    excel_output_path = folder / OUTPUT_EXCEL_FILE
    dxf_files = sorted(
        path
        for path in folder.rglob("*")
        if path.is_file() and path.suffix.lower() == ".dxf"
    )
    row_count = 0
    empty_dxf_files: list[Path] = []
    all_rows: list[list[str]] = []

    import multiprocessing

    with output_path.open("w", newline="", encoding="utf-8-sig") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(CSV_HEADERS)

        with multiprocessing.Pool() as pool:
            results = pool.imap(process_dxf_file, dxf_files)
            for file_number, rows in enumerate(results, start=1):
                dxf_path = dxf_files[file_number - 1]
                logging.info(
                    "[%d/%d] Processing %s",
                    file_number,
                    len(dxf_files),
                    dxf_path.relative_to(folder),
                )
                rows_from_file = 0
                for row in rows:
                    writer.writerow(row)
                    all_rows.append(row)
                    row_count += 1
                    rows_from_file += 1
                if rows_from_file == 0:
                    empty_dxf_files.append(dxf_path)

    # Automatically export Excel file formatted to match template
    export_to_excel(CSV_HEADERS, all_rows, excel_output_path)

    logging.info("Wrote %d pipe rows to %s", row_count, output_path)
    if empty_dxf_files:
        print("\nFiles that did not contribute to the CSV:")
        for path in empty_dxf_files:
            print(str(path.relative_to(folder)))


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    main()

