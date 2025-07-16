import os
import pandas as pd
import xlrd
from openpyxl import Workbook, load_workbook

source_folder = "source_excels"
output_excel = "main_excel_compiled.xlsx"
target_faces = ["Combustion Face", "Seat Face-1", "Seat Face-2", "Fillet-1", "Fillet-2"]

def read_sheet(file_path):
    xls = pd.ExcelFile(file_path, engine="xlrd")
    sheet_in, sheet_ex = None, None

    for sheet_name in xls.sheet_names:
        lower = sheet_name.lower()
        if "valve temperature" in lower and "in" in lower:
            sheet_in = xls.parse(sheet_name, header=None)
        elif "valve temperature" in lower and "ex" in lower:
            sheet_ex = xls.parse(sheet_name, header=None)

    return sheet_in, sheet_ex

def find_label_blocks(df):
    blocks = {}
    idx = 0
    while idx < len(df):
        cell = df.iloc[idx, 2]
        if isinstance(cell, str):
            label = cell.strip()
            if label in target_faces:
                start = idx
                end = idx
                while end + 1 < len(df) and pd.isna(df.iloc[end + 1, 2]):
                    end += 1
                blocks.setdefault(label, []).append((start, end))
                idx = end + 1
                continue
        idx += 1
    return blocks

def extract_max(df, ranges):
    values = []
    for start, end in ranges:
        for row in range(start, end + 1):
            for col in range(4, 12):
                try:
                    val = df.iloc[row, col]
                    if isinstance(val, (int, float)):
                        values.append((val, row))
                except:
                    continue
    if not values:
        for start, end in ranges:
            for row in range(start, end + 1):
                for col in range(4, 10):
                    try:
                        val = df.iloc[row, col]
                        if isinstance(val, (int, float)):
                            values.append((val, row))
                    except:
                        continue
    if values:
        max_val, max_row = max(values, key=lambda x: x[0])
        reference_d = df.iloc[max_row, 3] if max_row < len(df) else ""
        return max_val, reference_d
    else:
        return "", ""

def simplify_filename(filename):
    name = os.path.splitext(filename)[0]  # Remove .xls
    if "MSIL " in name:
        start = name.find("MSIL ") + len("MSIL ")
        # Find nearest underscore or hyphen after "MSIL "
        underscore = name.find("_", start)
        hyphen = name.find("-", start)
        end_candidates = [i for i in [underscore, hyphen] if i != -1]
        end = min(end_candidates) if end_candidates else len(name)
        return name[start:end].strip()
    return name.strip()



def get_next_column(sheet):
    row_index = 5
    col_index = 2  # Start at column B = 2
    while sheet.cell(row=row_index, column=col_index).value is not None:
        col_index += 1
    return col_index

def process_to_excel(df, target_sheet, filename):
    blocks = find_label_blocks(df)
    max_values = {}
    ref_d_value = ""

    combined_ranges = []
    for face in ["Combustion Face", "Seat Face-1", "Seat Face-2", "Fillet-1", "Fillet-2"]:
        combined_ranges.extend(blocks.get(face, []))

    overall_max, ref_d_value = extract_max(df, combined_ranges)

    if "Combustion Face" in blocks:
        val, _ = extract_max(df, blocks["Combustion Face"])
        max_values["Combustion Face"] = val
    seat_ranges = blocks.get("Seat Face-1", []) + blocks.get("Seat Face-2", [])
    if seat_ranges:
        val, _ = extract_max(df, seat_ranges)
        max_values["Seat Face"] = val
    fillet_ranges = blocks.get("Fillet-1", []) + blocks.get("Fillet-2", [])
    if fillet_ranges:
        val, _ = extract_max(df, fillet_ranges)
        max_values["Fillet"] = val

    col = get_next_column(target_sheet)
    label = simplify_filename(filename)

    target_sheet.cell(row=4, column=col, value=label)
    target_sheet.cell(row=5, column=col, value=max_values.get("Combustion Face", ""))
    target_sheet.cell(row=6, column=col, value=max_values.get("Seat Face", ""))
    target_sheet.cell(row=7, column=col, value=max_values.get("Fillet", ""))

    numeric_vals = [v for v in [max_values.get("Combustion Face"), max_values.get("Seat Face"), max_values.get("Fillet")] if isinstance(v, (int, float))]
    row8_max = max(numeric_vals) if numeric_vals else ""
    target_sheet.cell(row=8, column=col, value=row8_max)
    target_sheet.cell(row=9, column=col, value=ref_d_value)

# 📘 Open or create workbook
if os.path.exists(output_excel):
    compiled_wb = load_workbook(output_excel)
    compiled_in_sheet = compiled_wb["IN Sheet"]
    compiled_ex_sheet = compiled_wb["EX Sheet"]
    print("📂 Existing 'main_excel_compiled.xlsx' loaded.")
else:
    compiled_wb = Workbook()
    compiled_in_sheet = compiled_wb.active
    compiled_in_sheet.title = "IN Sheet"
    compiled_ex_sheet = compiled_wb.create_sheet("EX Sheet")
    print("📘 New workbook created.")

# 🔁 Process all .xls files
for file in os.listdir(source_folder):
    if file.endswith(".xls"):
        path = os.path.join(source_folder, file)
        print(f"\n🔍 Processing: {file}")
        sheet_in, sheet_ex = read_sheet(path)

        if sheet_in is not None:
            print("✏️ Writing to IN Sheet")
            process_to_excel(sheet_in, compiled_in_sheet, file)
        if sheet_ex is not None:
            print("✏️ Writing to EX Sheet")
            process_to_excel(sheet_ex, compiled_ex_sheet, file)

# 💾 Save workbook
compiled_wb.save(output_excel)
print(f"\n✅ All results saved to '{output_excel}'")
