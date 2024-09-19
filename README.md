# CV-V9

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from collections import defaultdict

# Function to load workbook
def load_workbook(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
        return workbook
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return None

# Function to clear rows in a sheet
def clear_sheet(sheet):
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.value = None

# Function to apply formatting to a cell
def format_header(cell, fill_color, font_size=11, bold=True):
    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    cell.font = Font(size=font_size, bold=bold)
    cell.alignment = Alignment(horizontal="center", vertical="center")

# Get file location via dialog
def get_file_location():
    root = Tk()
    root.withdraw()  # hide root window
    file_path = askopenfilename()
    return file_path

# Main function for SR comparison
def sr_pick_list_comparison():
    # Get file locations for Sheet1 and Sheet2
    file1 = get_file_location()
    file2 = get_file_location()

    if not file1 or not file2:
        print("Please select valid files.")
        return

    # Load the workbooks
    wb1 = load_workbook(file1)
    wb2 = load_workbook(file2)

    if not wb1 or not wb2:
        return

    ws1 = wb1.active
    ws2 = wb2.active

    # Initialize dictionaries for storing SR Reason -> SR Sub Reasons
    dict1 = defaultdict(list)
    dict2 = defaultdict(list)

    # Populate dictionaries from Sheet1
    for row in ws1.iter_rows(min_row=2, values_only=True):
        sr_reason, sr_sub_reason = row[0], row[1]
        dict1[sr_reason].append(sr_sub_reason)

    # Populate dictionaries from Sheet2
    for row in ws2.iter_rows(min_row=2, values_only=True):
        sr_reason, sr_sub_reason = row[0], row[1]
        dict2[sr_reason].append(sr_sub_reason)

    # Open a new workbook to store results
    result_wb = openpyxl.Workbook()
    result_ws = result_wb.active
    result_ws.title = "Details"

    # Add headers to the result sheet
    headers = ["SR Reason", "SR Sub Reason (Sheet1)", "SR Sub Reason (Sheet2)", "Result"]
    for col_num, header in enumerate(headers, 1):
        cell = result_ws.cell(row=1, column=col_num)
        cell.value = header
        format_header(cell, fill_color="00FFFF00")  # Yellow fill for headers

    # Compare SR Reasons and Sub Reasons
    current_row = 2
    for sr_reason in set(dict1.keys()).union(dict2.keys()):
        sub_reasons1 = dict1.get(sr_reason, [])
        sub_reasons2 = dict2.get(sr_reason, [])

        if not sub_reasons2:
            for sub_reason in sub_reasons1:
                result_ws.append([sr_reason, sub_reason, "Not Available", "FAIL"])
        elif not sub_reasons1:
            for sub_reason in sub_reasons2:
                result_ws.append([sr_reason, "Not Available", sub_reason, "FAIL"])
        else:
            for sub_reason in sub_reasons1:
                if sub_reason in sub_reasons2:
                    result_ws.append([sr_reason, sub_reason, sub_reason, "PASS"])
                else:
                    result_ws.append([sr_reason, sub_reason, "Not Available", "FAIL"])
            for sub_reason in sub_reasons2:
                if sub_reason not in sub_reasons1:
                    result_ws.append([sr_reason, "Not Available", sub_reason, "FAIL"])

    # Save the result workbook
    result_wb.save("SR_Pick_List_Comparison_Result.xlsx")
    print("Comparison complete. Results saved to SR_Pick_List_Comparison_Result.xlsx")

if __name__ == "__main__":
    sr_pick_list_comparison()

    
