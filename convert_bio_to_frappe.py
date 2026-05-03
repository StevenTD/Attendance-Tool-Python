import pandas as pd
import argparse
import os
import re
from datetime import datetime


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
    xl = pd.ExcelFile(input_file)
    sheet_name = "Logs" if "Logs" in xl.sheet_names else ("Period" if "Period" in xl.sheet_names else xl.sheet_names[0])
    df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)

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


def convert_file(input_file, output_file="frappe_checkin.csv"):
    # Debug: Check sheet names
    xl = pd.ExcelFile(input_file)
    print(f"Available sheets: {xl.sheet_names}")

    # Try "Logs" first, then "Period", then first sheet
    sheet_name = "Logs" if "Logs" in xl.sheet_names else ("Period" if "Period" in xl.sheet_names else xl.sheet_names[0])
    print(f"Using sheet: {sheet_name}")

    df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)

    print(f"DataFrame shape: {df.shape}")
    # print("First 10 rows:")
    # for idx in range(min(10, len(df))):
    #     print(f"Row {idx}: {df.iloc[idx].tolist()}")

    # Extract period info from row 2, column 2
    period_str = str(df.iloc[2, 2])
    print(f"Period string: {period_str}")
    # Extract year and month, e.g., '2026/04/01 ~ 04/30'
    period_match = re.search(r'(\d{4})/(\d{2})/', period_str)
    if period_match:
        year = int(period_match.group(1))
        month = int(period_match.group(2))
        print(f"Detected year: {year}, month: {month}")
    else:
        print("Could not detect year and month from period. Using defaults.")
        year = 2026
        month = 4

    results = []

    current_employee = None
    current_emp_no = None

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

                    # Get day from the day row (assuming row 3 is the day numbers)
                    day_cell = df.iloc[3, col_index]
                    try:
                        day = int(day_cell)
                        if day < 1 or day > 31:
                            continue
                    except (ValueError, TypeError):
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
                                "employee_name": current_employee,
                                "employee_no": current_emp_no,
                                "time": dt.strftime("%Y-%m-%d %H:%M:%S"),
                                "log_type": log_type,
                                "date": dt.date()
                            })
                        except:
                            continue

            i += 2  # skip time row
            continue

        i += 1

    if not results:
        print("No data found.")
        return

    # Get unique dates
    unique_dates = sorted(set(r["date"] for r in results))
    print(f"Found dates: {', '.join(d.strftime('%Y-%m-%d') for d in unique_dates)}")

    # Ask user for export option
    while True:
        choice = input("Do you want to export (A)ll dates or (R)ange? ").strip().upper()
        if choice == 'A':
            filtered_results = results
            break
        elif choice == 'R':
            start_str = input("Enter start date (YYYY-MM-DD): ").strip()
            end_str = input("Enter end date (YYYY-MM-DD): ").strip()
            try:
                start_date = datetime.strptime(start_str, "%Y-%m-%d").date()
                end_date = datetime.strptime(end_str, "%Y-%m-%d").date()
                filtered_results = [r for r in results if start_date <= r["date"] <= end_date]
                break
            except ValueError:
                print("Invalid date format. Please use YYYY-MM-DD.")
        else:
            print("Please enter A or R.")

    # Remove date key before saving
    for r in filtered_results:
        del r["date"]

    # Generate filename based on date range
    if filtered_results:
        dates_in_results = [datetime.strptime(r["time"], "%Y-%m-%d %H:%M:%S").date() for r in filtered_results]
        min_date = min(dates_in_results)
        max_date = max(dates_in_results)
        if min_date == max_date:
            date_suffix = min_date.strftime("%Y-%m-%d")
        else:
            date_suffix = f"{min_date.strftime('%Y-%m-%d')}_to_{max_date.strftime('%Y-%m-%d')}"
        output_file = f"frappe_checkin_{date_suffix}.csv"
    else:
        output_file = "frappe_checkin_empty.csv"

    output_df = pd.DataFrame(filtered_results)

    # If an Employee_IDs_Matched.xlsx file exists in the script folder, use it to map employee_no to Employee ID.
    mapping_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Employee_IDs_Matched.xlsx")
    if os.path.isfile(mapping_file):
        print(f"Found mapping file: {mapping_file}")
        try:
            map_df = pd.read_excel(mapping_file, sheet_name="Employee IDs", dtype=str)
            if "ID" in map_df.columns and "Employee ID" in map_df.columns:
                map_df = map_df.dropna(subset=["ID"])

                def normalize_id(value):
                    if pd.isna(value):
                        return None
                    val = str(value).strip()
                    if re.fullmatch(r"\d+\.0+", val):
                        val = val.split(".")[0]
                    return val

                employee_map = {
                    normalize_id(k): str(v).strip()
                    for k, v in zip(map_df["ID"], map_df["Employee ID"])
                    if normalize_id(k) is not None
                }

                if not output_df.empty and "employee_no" in output_df.columns:
                    output_df["employee"] = output_df["employee_no"].astype(str).map(lambda v: normalize_id(v)).map(employee_map)
                print(f"Loaded {len(employee_map)} employee mappings from {os.path.basename(mapping_file)}")
            else:
                print("Employee_IDs_Matched.xlsx found, but required columns 'ID' and 'Employee ID' are missing.")
        except Exception as exc:
            print(f"Could not read mapping file: {exc}")

    output_df.to_csv(output_file, index=False)

    print(f"✅ Conversion saved to: {output_file}")
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
