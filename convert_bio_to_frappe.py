import pandas as pd
import argparse
import re


def extract_info(row):
    row_str = " ".join([str(x) for x in row if pd.notna(x)])

    no_match = re.search(r"No\s*:\s*(\d+)", row_str)
    emp_no = no_match.group(1) if no_match else None

    name_match = re.search(r"Name\s*:\s*(.+?)\s+Dept", row_str)
    name = name_match.group(1).strip() if name_match else None

    return emp_no, name


def is_employee_row(row):
    row_str = " ".join([str(x) for x in row if pd.notna(x)])
    return "Name :" in row_str and "No :" in row_str


# =========================
# 🔹 MODE 1: Extract IDs
# =========================
def extract_ids(input_file, output_file="employee_ids.csv"):
    df = pd.read_excel(input_file, sheet_name="Logs", header=None)

    results = []

    for _, row in df.iterrows():
        row_list = row.tolist()

        if is_employee_row(row_list):
            emp_no, name = extract_info(row_list)

            if emp_no and name:
                results.append({
                    "employee_no": emp_no,
                    "employee_name": name
                })

    output_df = pd.DataFrame(results).drop_duplicates()
    output_df.to_csv(output_file, index=False)

    print(f"✅ Employee list saved to: {output_file}")


# =========================
# 🔹 MODE 2: Convert logs
# =========================
from datetime import datetime


def convert_file(input_file, output_file="frappe_checkin.csv"):
    df = pd.read_excel(input_file, sheet_name="Logs", header=None)

    results = []

    current_employee = None
    current_emp_no = None

    year = 2026
    month = 3

    i = 0
    while i < len(df):
        row = df.iloc[i].tolist()

        # Detect employee row
        if is_employee_row(row):
            current_emp_no, current_employee = extract_info(row)

            # ✅ IMPORTANT: next row contains actual time logs
            if i + 1 < len(df):
                time_row = df.iloc[i + 1].tolist()

                for col_index in range(len(time_row)):
                    cell = time_row[col_index]

                    if pd.isna(cell):
                        continue

                    # ⚠️ Column index is NOT reliable day
                    # So we map ONLY if it's within 1–31 range
                    day = col_index +1  # approximate

                    if day < 1 or day > 31:
                        continue

                    times = str(cell).split("\n")

                    clean_times = [
                        t.strip() for t in times
                        if re.match(r"^\d{2}:\d{2}$", t.strip())
                    ]

                    for idx, t in enumerate(clean_times):
                        try:
                            dt = datetime.strptime(
                                f"{year}-{month:02d}-{day:02d} {t}",
                                "%Y-%m-%d %H:%M"
                            )

                            log_type = "IN" if idx % 2 == 0 else "OUT"

                            results.append({
                                "employee": current_employee,
                                "employee_no": current_emp_no,
                                "time": dt.strftime("%Y-%m-%d %H:%M:%S"),
                                "log_type": log_type
                            })
                        except:
                            continue

            i += 2  # skip time row
            continue

        i += 1

    output_df = pd.DataFrame(results)
    output_df.to_csv(output_file, index=False)

    print(f"✅ Fixed conversion saved to: {output_file}")
# =========================
# 🔹 CLI ENTRY
# =========================
if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    parser.add_argument("input_file", help="Path to Excel file")
    parser.add_argument("--output", help="Output file name")
    parser.add_argument("--extractid", action="store_true", help="Extract employee IDs only")

    args = parser.parse_args()

    if args.extractid:
        extract_ids(args.input_file, args.output or "employee_ids.csv")
    else:
        convert_file(args.input_file, args.output or "frappe_checkin.csv")
