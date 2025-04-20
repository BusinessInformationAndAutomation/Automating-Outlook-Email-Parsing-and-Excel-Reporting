import os
import shutil
from datetime import date, timedelta
from openpyxl import load_workbook

def copy_and_replace_suffix(original_file_path):
    today = date.today()
    prev_date = (today - timedelta(days=1)).strftime("%d-%m-%Y")
    base_date = (today - timedelta(days=3)).strftime("%m-%d-%Y")

    base_name, ext = os.path.splitext(original_file_path)
    new_file_name = f"{base_name[:-9]}{base_date}{ext}"  # Trim and add date

    # Copy and update workbook
    shutil.copy(original_file_path, new_file_name)
    wb = load_workbook(new_file_name)
    source_sheet = wb["Sheet1"]

    for days_offset in [2, 3]:
        new_sheet = wb.copy_worksheet(source_sheet)
        new_sheet.title = (today - timedelta(days=days_offset)).strftime("%d-%m-%Y")

    # Rename the original sheet
    source_sheet.title = (today - timedelta(days=1)).strftime("%d-%m-%Y")

    wb.save(new_file_name)

    # Save the path to a file for main script to use
    with open("excel_path.txt", "w") as f:
        f.write(new_file_name)

    return new_file_name
