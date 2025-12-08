"""
Process .rnd files under Auto_*/Auto_Leq/* and write an Excel workbook
with one sheet per .rnd file.

Usage:
    - set data_folder to your main data folder (the one containing Auto_1101, Auto_1102, ...)
    - run: python main.py
"""

import os
import re
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border

# Configuration
data_folder = r"C:\\Data"
output_workbook = "Rion calculation.xlsx"
auto_prefix = "Auto_"
auto_leq_subdir = "Auto_Leq"

# Standard frequency bands
FREQ_HEADERS = [
    "16Hz",
    "31.5Hz",
    "63Hz",
    "125Hz",
    "250Hz",
    "500Hz",
    "1kHz",
    "2kHz",
    "4kHz",
    "8kHz",
    "16kHz",
]

# Weighting factors (dB)
C_WEIGHT_MAP = {
    "31.5Hz": -0.8,
    "63Hz": -0.2,
    "125Hz": 0.0,
    "250Hz": 0.0,
    "500Hz": 0.0,
    "1kHz": 0.0,
    "2kHz": -0.2,
    "4kHz": -0.8,
    "8kHz": -1.4,
    "16kHz": -1.6,
}

A_WEIGHT_MAP = {
    "31.5Hz": -39.4,
    "63Hz": -26.2,
    "125Hz": -16.1,
    "250Hz": -8.6,
    "500Hz": -3.2,
    "1kHz": 0.0,
    "2kHz": 1.2,
    "4kHz": 1.0,
    "8kHz": -1.1,
    "16kHz": -6.6,
}

MISSING_TOKENS = {"-.-", "-", "", ".-.", ".", ".,"}


# Sanitize sheet name (max 31 chars, remove invalid characters)
def make_sheet_name(name: str) -> str:
    invalid = r"[:\\/*?[\]]"
    safe = re.sub(invalid, "_", name)
    return safe[:31]


def is_number(s: str):
    try:
        float(s.replace(",", "."))
        return True
    except Exception:
        return False


def to_number(s: str):
    # Convert strings to float, handling comma decimal separators
    try:
        return float(s.replace(",", "."))
    except Exception:
        return None


def parse_measurement_time_to_seconds(s: str):
    """Parse measurement time strings like '000d 00:00:10.0' or '00:00:10.0' into seconds (float).
    Returns float seconds or None if cannot parse.
    """
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)
    s = str(s).strip()
    if s == "":
        return None
    # Try direct float conversion
    try:
        return float(s.replace(",", "."))
    except Exception:
        pass

    # Parse days + time format
    m = re.match(r"(?:(\d+)d\s*)?(\d{1,2}):(\d{2}):(\d{2}(?:\.\d+)?)", s)
    if m:
        days = int(m.group(1)) if m.group(1) else 0
        hours = int(m.group(2))
        minutes = int(m.group(3))
        seconds = float(m.group(4))
        total = days * 86400 + hours * 3600 + minutes * 60 + seconds
        return total

    # Extract first numeric value from string
    m2 = re.search(r"(\d+(?:[\.,]\d+)?)", s)
    if m2:
        try:
            return float(m2.group(1).replace(",", "."))
        except Exception:
            return None
    return None


def auto_find_rnd_files(base_folder: str):
    """Find all .rnd files in Auto_*/Auto_Leq/ subdirectories."""
    found = []
    for entry in os.listdir(base_folder):
        entry_path = os.path.join(base_folder, entry)
        if not os.path.isdir(entry_path):
            continue
        if entry.lower().startswith(auto_prefix.lower()):
            leq_dir = os.path.join(entry_path, auto_leq_subdir)
            if os.path.isdir(leq_dir):
                for f in os.listdir(leq_dir):
                    if f.lower().endswith(".rnd"):
                        found.append(os.path.join(leq_dir, f))
    return found


def parse_rnd_file(path: str, header_signature_contains="Partial Over All"):
    """Parse .rnd file into headers and address rows with measurements."""
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            content = fh.read()
    except Exception as e:
        print(f"ERROR reading file {path}: {e}")
        return None, None

    lines = [ln.rstrip("\r") for ln in content.splitlines()]
    # Find blocks (each starts with 'Address,')
    block_starts = [
        i for i, ln in enumerate(lines) if ln.strip().lower().startswith("address,")
    ]
    if not block_starts:
        block_starts = [0]

    rows = []
    headers = None

    for idx, start in enumerate(block_starts):
        # Block ends at next block start or EOF
        end = block_starts[idx + 1] if idx + 1 < len(block_starts) else len(lines)
        block = lines[start:end]

        # Extract Address, Start Time, Measurement Time
        addr = ""
        start_time = ""
        meas_time = ""

        for ln in block:
            ln_stripped = ln.strip()

            if ln_stripped.lower().startswith("address,"):
                parts = ln_stripped.split(",", 1)
                addr = parts[1].strip() if len(parts) > 1 else ""

            elif ln_stripped.lower().startswith("start time,"):
                parts = ln_stripped.split(",", 1)
                start_time = parts[1].strip() if len(parts) > 1 else ""

            elif ln_stripped.lower().startswith("measurement time,"):
                parts = ln_stripped.split(",", 1)
                meas_time = parts[1].strip() if len(parts) > 1 else ""

        # Find header line matching signature
        header_line = None
        for ln in block:
            if header_signature_contains.lower() in ln.lower():
                header_line = ln
                break

        if header_line is None:
            # Fallback: find line with frequencies
            for ln in block:
                if "16Hz" in ln or "31.5Hz" in ln:
                    header_line = ln
                    break

        if header_line is None:
            print(f"Warning: no header found in block at line {start} in {path}")
            continue

        # Parse header row (comma-separated, first token empty)
        header_parts = [h.strip() for h in header_line.split(",") if h is not None]
        if header_parts and header_parts[0] == "":
            header_parts = header_parts[1:]
        # Store headers from first block
        if headers is None:
            headers = header_parts.copy()

        # Find Leqmov line (measurement values)
        leqmov_line = None
        for ln in block:
            if ln.strip().lower().startswith("leqmov"):
                leqmov_line = ln
                break
        if leqmov_line is None:
            # Try alternate spacing
            for ln in block:
                if re.match(r"(?i)leq\s*mov\s*,", ln):
                    leqmov_line = ln
                    break

        if leqmov_line is None:
            values = [""] * len(header_parts)
        else:
            # Extract values after label
            parts = leqmov_line.split(",", 1)
            values_part = parts[1] if len(parts) > 1 else ""
            values = [v.strip() for v in values_part.split(",")]
            # Pad or truncate to match headers
            if len(values) < len(header_parts):
                values += [""] * (len(header_parts) - len(values))
            elif len(values) > len(header_parts):
                values = values[: len(header_parts)]

        # Build row with metadata and values
        row = {
            "Address": addr,
            "Start Time": start_time,
            "Measurement Time": meas_time,
        }
        for h, v in zip(header_parts, values):
            row[h] = v

        rows.append(row)

    return headers, rows


def write_to_excel(all_parsed, out_path):
    """Write parsed .rnd data to multi-table Excel workbook."""
    # Use pandas ExcelWriter
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for file_title, headers, rows in all_parsed:
            sheet_name = make_sheet_name(file_title)
            try:
                # Build raw table DataFrame
                if headers is None or rows is None:
                    print(f"Skipping {file_title} due to missing parse result.")
                    continue

                # columns: Address | Start Time | Measurement Time | <headers...>
                cols = ["Address", "Start Time", "Measurement Time"] + headers

                df_rows = []
                for r in rows:
                    rowvals = [
                        r.get("Address", ""),
                        r.get("Start Time", ""),
                        r.get("Measurement Time", ""),
                    ]
                    for h in headers:
                        val = r.get(h, "")
                        # Preserve missing tokens; convert numeric strings
                        if isinstance(val, str) and val.strip() in MISSING_TOKENS:
                            rowvals.append(val.strip())
                        elif isinstance(val, str) and is_number(val):
                            rowvals.append(to_number(val))
                        else:
                            rowvals.append(val)
                    df_rows.append(rowvals)

                raw_df = pd.DataFrame(df_rows, columns=cols)
                # Write raw table at row 4
                raw_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=3)

                # Prepare second table: C-weighting transformation
                second_startcol = len(cols) + 1
                freq_cols = [f for f in FREQ_HEADERS[1:]]

                # Standard header rows for second table
                freq_row1 = {f: f for f in freq_cols}
                c_row = {f: C_WEIGHT_MAP.get(f, "") for f in freq_cols}
                remove_row = {f: "" for f in freq_cols}
                freq_row2 = {f: f for f in freq_cols}

                # Apply C-weighting transformation (subtract C-weight values)
                transformed_data = []
                transformed_index = []
                for r in df_rows:
                    row_dict = {f: "" for f in freq_cols}
                    for i, h in enumerate(headers, start=3):
                        if h in freq_cols:
                            v = r[i]
                            if isinstance(v, str) and v.strip() in MISSING_TOKENS:
                                row_dict[h] = v.strip()
                            else:
                                try:
                                    num = float(v)
                                    cw = float(C_WEIGHT_MAP.get(h, 0.0))
                                    row_dict[h] = num - cw
                                except Exception:
                                    row_dict[h] = v
                    transformed_data.append(row_dict)
                    addr_val = r[0] if r[0] else "Unknown"
                    transformed_index.append(str(addr_val))

                # Assemble second table with header + data rows
                second_rows = []
                second_index = []
                second_rows.append(freq_row1)
                second_index.append("Frequency")
                second_rows.append(c_row)
                second_index.append("C-weight (dB)")
                remove_row = {f: "" for f in freq_cols}
                second_rows.append(remove_row)
                second_index.append("Remove C-weighting")
                second_rows.append(freq_row2)
                second_index.append("")
                # Add data rows without addresses
                for row_dict in transformed_data:
                    second_rows.append(row_dict)
                second_index.extend([""] * len(transformed_data))

                second_df = pd.DataFrame(
                    second_rows, columns=freq_cols, index=second_index
                )
                second_df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    startrow=0,
                    startcol=second_startcol,
                    index=True,
                    header=False,
                )

                # Third table: A-weighting transformation
                third_startcol = second_startcol + len(freq_cols) + 2

                # Apply A-weighting transformation (add A-weight values)
                a_transformed_data = []
                for r in df_rows:
                    row_dict = {f: "" for f in freq_cols}
                    for i, h in enumerate(headers, start=3):
                        if h in freq_cols:
                            v = r[i]
                            if isinstance(v, str) and v.strip() in MISSING_TOKENS:
                                row_dict[h] = v.strip()
                            else:
                                try:
                                    num = float(v)
                                    aw = float(A_WEIGHT_MAP.get(h, 0.0))
                                    row_dict[h] = num + aw
                                except Exception:
                                    row_dict[h] = v
                    a_transformed_data.append(row_dict)

                # Assemble third table with header + data rows
                third_rows = []
                third_index = []
                third_rows.append(freq_row1)
                third_index.append("Frequency")
                a_row = {f: A_WEIGHT_MAP.get(f, "") for f in freq_cols}
                third_rows.append(a_row)
                third_index.append("A-weight (dB)")
                apply_row = {f: "" for f in freq_cols}
                third_rows.append(apply_row)
                third_index.append("Apply A-weighting")
                third_rows.append(freq_row2)
                third_index.append("")
                for row_dict in a_transformed_data:
                    third_rows.append(row_dict)
                third_index.extend([""] * len(a_transformed_data))

                third_df = pd.DataFrame(
                    third_rows, columns=freq_cols, index=third_index
                )
                third_df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    startrow=0,
                    startcol=third_startcol,
                    index=True,
                    header=False,
                )
                
                # Add computed columns and formulas
                wb = writer.book
                ws = wb[sheet_name]                # LAeq,t_i (dBA)
                header_cell = ws.cell(row=4, column=42)
                header_cell.value = "LAeq,t_i (dBA)"
                header_cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                header_cell.font = Font(bold=True)
                ws.column_dimensions["AP"].width = 13

                # Apply formulas to column 42 (AP) for each data row
                # Data starts at row 5
                for row_idx in range(len(df_rows)):
                    excel_row = 5 + row_idx
                    target_cell = ws.cell(row=excel_row, column=42)
                    target_cell.value = f"=ROUND(10*LOG10(SUMPRODUCT(10^(AD{excel_row}:AM{excel_row}/10))), 1)"

                # t_i (seconds) - place measurement time in seconds for each address in column 44
                header_cell = ws.cell(row=4, column=44)
                header_cell.value = "t_i (s)"
                header_cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                header_cell.font = Font(bold=True)

                # Fill t_i values from the parsed measurement time (df_rows stores Measurement Time at index 2)
                for row_idx in range(len(df_rows)):
                    excel_row = 5 + row_idx
                    meas_val = df_rows[row_idx][2]
                    seconds = parse_measurement_time_to_seconds(meas_val)
                    target_cell = ws.cell(row=excel_row, column=44)
                    if seconds is None:
                        target_cell.value = None
                    else:
                        target_cell.value = float(seconds)

                # Running average: Laeq,t_t (placed in column 45)
                header_cell = ws.cell(row=3, column=45)
                header_cell.value = "Running average:"
                header_cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                header_cell.font = Font(bold=True)

                header_cell = ws.cell(row=4, column=45)
                header_cell.value = "Laeq,t_t"
                header_cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                header_cell.font = Font(bold=True)

                ws.column_dimensions["AS"].width = 15

                # Apply running-average formulas to column 45 for each data row
                # Data starts at row 5
                for row_idx in range(len(df_rows)):
                    excel_row = 5 + row_idx
                    target_cell = ws.cell(row=excel_row, column=45)
                    target_cell.value = f"=ROUND(10*LOG10(SUMPRODUCT($AR$5:AR{excel_row},10^($AP$5:AP{excel_row}/10))/SUM($AR$5:AR{excel_row})),1)"

                # --- AFTER WRITING TO THE EXCEL FILE, DO FORMATTING VIA openpyxl ---

                # Bold headers for raw table (header row is at raw_startrow + 4 because data starts at row 4 (1-based))
                header_font = Font(bold=True)
                raw_header_row = 4
                for col_idx, col_name in enumerate(cols, start=1):
                    cell = ws.cell(row=raw_header_row, column=col_idx)
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                # Also apply header formatting across the whole 4th row (including second-table and third-table columns) so the visual header style is consistent across the sheet.
                # Compute the last column index used by the third table.
                # Note: `third_startcol` is 0-based startcol; convert to 1-based and add freq cols.
                last_col_idx = (third_startcol + 1) + len(freq_cols)
                for col_idx in range(1, last_col_idx + 1):
                    cell = ws.cell(row=raw_header_row, column=col_idx)
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                # Apply header formatting to the 1st row for any cells that contain text
                # (match the style used for row 4 where applicable)
                for col_idx in range(1, last_col_idx + 1):
                    cell = ws.cell(row=1, column=col_idx)
                    if cell.value is None:
                        continue
                    # Only style text-containing cells
                    try:
                        is_text = (
                            isinstance(cell.value, str) and cell.value.strip() != ""
                        )
                    except Exception:
                        is_text = False
                    if is_text:
                        cell.font = header_font
                        cell.alignment = Alignment(
                            horizontal="center", vertical="center"
                        )

                # -------- 1. AUTO-WIDTH + ALIGNMENT FOR ALL COLUMNS (raw + second) --------
                # Use fixed width for all frequency columns in both tables for consistency
                freq_column_width = 9
                scan_rows = min(12, ws.max_row)

                # Raw table columns: first 3 are non-frequency (Address, Start Time, Measurement Time)
                # Column 4 is "Main", Column 5 is "Partial Over All" (both non-frequency, need auto-width)
                # Columns 6+ are frequency columns from FREQ_HEADERS
                for col in range(1, len(cols) + 1):
                    col_letter = get_column_letter(col)
                    # Determine if this is a frequency column
                    # Columns 1-5 are non-frequency (Address, Start Time, Measurement Time, Main, Partial Over All)
                    if col <= 5:
                        # Non-frequency columns get auto-width based on content
                        max_len = 0
                        for row_idx in range(1, scan_rows + 1):
                            cell = ws.cell(row=row_idx, column=col)
                            if cell.value is None:
                                continue
                            try:
                                text = str(cell.value)
                            except Exception:
                                text = ""
                            max_len = max(max_len, len(text))
                            if cell.row in (1, raw_header_row):
                                cell.alignment = Alignment(
                                    horizontal="center", vertical="center"
                                )
                        adjusted_width = min(max(10, max_len + 2), 20)
                        ws.column_dimensions[col_letter].width = adjusted_width
                    else:
                        # Frequency columns: use fixed width
                        ws.column_dimensions[col_letter].width = freq_column_width
                        for row_idx in range(1, scan_rows + 1):
                            cell = ws.cell(row=row_idx, column=col)
                            if cell.row in (1, raw_header_row):
                                cell.alignment = Alignment(
                                    horizontal="center", vertical="center"
                                )

                # Second table columns (index column + frequency columns)
                second_excel_start = second_startcol + 1
                for offset in range(len(freq_cols) + 1):
                    col = second_excel_start + offset
                    col_letter = get_column_letter(col)
                    if offset == 0:
                        # Index column (labels): auto-width based on content
                        max_len = 0
                        for row_idx in range(1, scan_rows + 1):
                            cell = ws.cell(row=row_idx, column=col)
                            if cell.value is None:
                                continue
                            try:
                                text = str(cell.value)
                            except Exception:
                                text = ""
                            max_len = max(max_len, len(text))
                            if cell.row in (1, raw_header_row):
                                cell.alignment = Alignment(
                                    horizontal="center", vertical="center"
                                )
                        adjusted_width = min(max(10, max_len + 2), 20)
                        ws.column_dimensions[col_letter].width = adjusted_width
                    else:
                        # Frequency columns in second table: use same fixed width
                        ws.column_dimensions[col_letter].width = freq_column_width
                        for row_idx in range(1, scan_rows + 1):
                            cell = ws.cell(row=row_idx, column=col)
                            if cell.row in (1, raw_header_row):
                                cell.alignment = Alignment(
                                    horizontal="center", vertical="center"
                                )

                # Third table columns (index column + frequency columns)
                third_excel_start = third_startcol + 1
                for offset in range(len(freq_cols) + 1):
                    col = third_excel_start + offset
                    col_letter = get_column_letter(col)
                    if offset == 0:
                        # Index column (labels): auto-width based on content
                        max_len = 0
                        for row_idx in range(1, scan_rows + 1):
                            cell = ws.cell(row=row_idx, column=col)
                            if cell.value is None:
                                continue
                            try:
                                text = str(cell.value)
                            except Exception:
                                text = ""
                            max_len = max(max_len, len(text))
                            if cell.row in (1, raw_header_row):
                                cell.alignment = Alignment(
                                    horizontal="center", vertical="center"
                                )
                        adjusted_width = min(max(10, max_len + 2), 20)
                        ws.column_dimensions[col_letter].width = adjusted_width
                    else:
                        # Frequency columns in third table: use same fixed width
                        ws.column_dimensions[col_letter].width = freq_column_width
                        for row_idx in range(1, scan_rows + 1):
                            cell = ws.cell(row=row_idx, column=col)
                            if cell.row in (1, raw_header_row):
                                cell.alignment = Alignment(
                                    horizontal="center", vertical="center"
                                )

                # Remove borders from index columns
                for col_ref in [get_column_letter(second_startcol + 1), get_column_letter(third_startcol + 1)]:
                    for cell in ws[col_ref]:
                        try:
                            cell.border = Border()
                        except Exception:
                            pass

                print(
                    f"Wrote sheet '{sheet_name}' with {raw_df.shape[0]} addresses."
                )

            except Exception as e:
                print(f"Error writing sheet for file {file_title}: {e}")

    print(f"Saved workbook: {out_path}")


def main():
    if not os.path.isdir(data_folder):
        print(f"ERROR: data_folder does not exist: {data_folder}")
        return

    rnd_files = auto_find_rnd_files(data_folder)
    if not rnd_files:
        print("No .rnd files found.")
        return

    parsed_all = []
    for p in rnd_files:
        try:
            title = os.path.splitext(os.path.basename(p))[0]
            headers, rows = parse_rnd_file(p)
            if headers and rows:
                parsed_all.append((title, headers, rows))
            else:
                print(f"Warning: no data parsed for {p}")
        except Exception as e:
            print(f"Error parsing {p}: {e}")

    if not parsed_all:
        print("No data parsed.")
        return

    try:
        write_to_excel(parsed_all, output_workbook)
    except Exception as e:
        print(f"Failed to write workbook: {e}")


if __name__ == "__main__":
    main()
