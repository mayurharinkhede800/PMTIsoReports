from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
import os
from math import ceil


def generate_tools_tackles_report(download_path, supabase):

    # -------------------------------
    # FETCH DATA
    # -------------------------------
    res = supabase.table("v_tools_tackles_report") \
        .select("*") \
        .order("taken_date") \
        .execute()

    data = res.data or []

    if not data:
        print("❌ No tools data found")
        return

    # -------------------------------
    # PATHS
    # -------------------------------
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    TEMPLATE_PATH = os.path.join(
        BASE_DIR,
        "templates",
        "tools_tackles_template.xlsx"
    )

    ITEMS_PER_SHEET = 13
    start_row = 11

    total_sheets = ceil(len(data) / ITEMS_PER_SHEET)

    # -------------------------------
    # LOAD TEMPLATE
    # -------------------------------
    final_wb = load_workbook(TEMPLATE_PATH)

    template_ws = final_wb.active

    # remove default extra sheets later
    final_wb.remove(template_ws)

    # -------------------------------
    # SAFE WRITE
    # -------------------------------
    def safe_write(ws, row, col, value, align="center"):

        cell = ws.cell(row=row, column=col)

        cell.value = "" if value is None else str(value)

        cell.font = Font(size=11)

        cell.alignment = Alignment(
            horizontal=align,
            vertical="center",
            wrap_text=True
        )

    # -------------------------------
    # CREATE SHEETS
    # -------------------------------
    for sheet_idx in range(total_sheets):

        ws = final_wb.copy_worksheet(template_ws)

        ws.title = f"Page_{sheet_idx + 1}"

        # -------------------------------
        # PAGE DATA
        # -------------------------------
        start_i = sheet_idx * ITEMS_PER_SHEET
        end_i = min(start_i + ITEMS_PER_SHEET, len(data))

        page_data = data[start_i:end_i]

        # -------------------------------
        # INSERT DATA
        # -------------------------------
        for i, item in enumerate(page_data):

            row = start_row + i

            safe_write(ws, row, 1, start_i + i + 1)                  # S.No
            safe_write(ws, row, 2, item.get("tool_name"))            # Tool
            safe_write(ws, row, 3, item.get("employee_name"))        # Employee
            safe_write(ws, row, 4, item.get("taken_date"))           # Date
            safe_write(ws, row, 5, item.get("taken_time"))           # Taken Time
            safe_write(ws, row, 6, item.get("return_time"))          # Return Time
            safe_write(ws, row, 7, item.get("status"))               # Status

    # -------------------------------
    # SAVE FILE
    # -------------------------------
    os.makedirs(download_path, exist_ok=True)

    file_path = os.path.join(
        download_path,
        "tools_tackles_report.xlsx"
    )

    final_wb.save(file_path)

    print("\n✅ Tools & Tackles Report Generated")
    print("📁 Saved at:", file_path)