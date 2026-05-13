from openpyxl import load_workbook
from openpyxl.styles import Alignment
from collections import defaultdict
import os
import subprocess
import copy


# =========================================================
# SAFE CELL WRITE
# =========================================================
def safe_write(ws, row, col, value):

    cell = ws.cell(
        row=row,
        column=col
    )

    if type(cell).__name__ == "MergedCell":
        return

    cell.value = value


# =========================================================
# COPY TEMPLATE FORMAT
# =========================================================
def copy_sheet(source, target):

    # -----------------------------------------------------
    # COPY CELL DATA + STYLE
    # -----------------------------------------------------
    for row in source.iter_rows():

        for cell in row:

            if type(cell).__name__ == "MergedCell":
                continue

            new_cell = target[cell.coordinate]

            new_cell.value = cell.value

            if cell.has_style:
                new_cell._style = copy.copy(cell._style)

            new_cell.font = copy.copy(cell.font)
            new_cell.fill = copy.copy(cell.fill)
            new_cell.border = copy.copy(cell.border)
            new_cell.alignment = copy.copy(cell.alignment)
            new_cell.number_format = copy.copy(cell.number_format)
            new_cell.protection = copy.copy(cell.protection)

    # -----------------------------------------------------
    # COLUMN WIDTH
    # -----------------------------------------------------
    for key, dim in source.column_dimensions.items():

        target.column_dimensions[key].width = dim.width

    # -----------------------------------------------------
    # ROW HEIGHT
    # -----------------------------------------------------
    for key, dim in source.row_dimensions.items():

        target.row_dimensions[key].height = dim.height

    # -----------------------------------------------------
    # MERGED CELLS
    # -----------------------------------------------------
    for merged in source.merged_cells.ranges:

        target.merge_cells(str(merged))


# =========================================================
# GENERATE REPORT
# =========================================================
def generate_preventive_maintenance_report(
    data,
    download_path,
    month
):

    # =====================================================
    # TEMPLATE
    # =====================================================
    BASE_DIR = os.path.dirname(
        os.path.dirname(
            os.path.abspath(__file__)
        )
    )

    TEMPLATE_PATH = os.path.join(
        BASE_DIR,
        "templates",
        "MT 05 Preventive Maintenance Schedule-Monthly.xlsx"
    )

    # =====================================================
    # LOAD TEMPLATE
    # =====================================================
    wb = load_workbook(TEMPLATE_PATH)

    template_ws = wb.active

    # =====================================================
    # GROUP BY MACHINE
    # =====================================================
    grouped = defaultdict(list)

    for item in data:

        machine_name = item.get(
            "machine_name",
            "Unknown Machine"
        )

        grouped[machine_name].append(item)

    # =====================================================
    # SETTINGS
    # =====================================================
    START_ROW = 8

    MAX_MACHINE_PER_SHEET = 14

    sheet_no = 1

    ws = template_ws

    current_machine = 0

    row = START_ROW

    serial_no = 1

    # =====================================================
    # MACHINE LOOP
    # =====================================================
    for machine_name, machine_rows in grouped.items():

        # -------------------------------------------------
        # NEW SHEET
        # -------------------------------------------------
        if current_machine >= MAX_MACHINE_PER_SHEET:

            sheet_no += 1

            ws = wb.create_sheet(
                title=f"Sheet{sheet_no}"
            )

            copy_sheet(
                template_ws,
                ws
            )

            row = START_ROW

            current_machine = 0

        # -------------------------------------------------
        # SR NO
        # -------------------------------------------------
        safe_write(
            ws,
            row,
            1,
            serial_no
        )

        # -------------------------------------------------
        # MACHINE NAME
        # -------------------------------------------------
        safe_write(
            ws,
            row,
            2,
            machine_name
        )

        # -------------------------------------------------
        # DONE BY
        # -------------------------------------------------
        done_by_list = []

        # -------------------------------------------------
        # REMARKS
        # -------------------------------------------------
        remarks_list = []

        # -------------------------------------------------
        # MACHINE RECORDS
        # -------------------------------------------------
        for item in machine_rows:

            # =============================================
            # PLAN DATE
            # =============================================
            if (
                item.get("plan_date")
                and item.get("is_planned") is True
            ):

                day = int(
                    str(
                        item["plan_date"]
                    ).split("-")[2]
                )

                # DATE STARTS FROM COLUMN 4
                col = day + 4

                safe_write(
                    ws,
                    row,
                    col,
                    "✔"
                )

            # =============================================
            # ACHIEVED DATE
            # =============================================
            if (
                item.get("achieved_date")
                and item.get("is_achieved") is True
            ):

                day = int(
                    str(
                        item["achieved_date"]
                    ).split("-")[2]
                )

                # SECOND ROW
                achieved_row = row + 1

                col = day + 4

                safe_write(
                    ws,
                    achieved_row,
                    col,
                    "✔"
                )

            # =============================================
            # DONE BY
            # =============================================
            done_by = item.get(
                "done_by_name"
            )

            if done_by:

                if done_by not in done_by_list:

                    done_by_list.append(
                        done_by
                    )

            # =============================================
            # REMARKS
            # =============================================
            remarks = item.get(
                "remarks"
            )

            if remarks:

                if remarks not in remarks_list:

                    remarks_list.append(
                        remarks
                    )

        # -------------------------------------------------
        # DONE BY COLUMN
        # -------------------------------------------------
        safe_write(
            ws,
            row,
            36,
            ", ".join(done_by_list)
        )

                # -------------------------------------------------
        # REMARKS COLUMN
        # -------------------------------------------------
        remarks_text = " | ".join(remarks_list)

        safe_write(
            ws,
            row,
            37,
            remarks_text
        )

        # -------------------------------------------------
        # REMARKS CELL FORMAT
        # -------------------------------------------------
        remarks_cell = ws.cell(
            row=row,
            column=37
        )

        remarks_cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
            shrink_to_fit=True
        )

        # -------------------------------------------------
        # KEEP ORIGINAL FONT
        # -------------------------------------------------
        remarks_cell.font = copy.copy(
            ws.cell(row=8, column=37).font
        )

        # -------------------------------------------------
        # DONE BY CELL FORMAT
        # -------------------------------------------------
        done_by_cell = ws.cell(
            row=row,
            column=36
        )

        done_by_cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
            shrink_to_fit=True
        )

        # -------------------------------------------------
        # DO NOT CHANGE ROW HEIGHT
        # -------------------------------------------------
        # KEEP TEMPLATE HEIGHT AS IT IS

        # -------------------------------------------------
        # DO NOT CHANGE COLUMN WIDTH
        # -------------------------------------------------
        # KEEP TEMPLATE WIDTH AS IT IS

        # -------------------------------------------------
        # ALIGNMENT
        # -------------------------------------------------
        for r in [row, row + 1]:

            for c in range(1, 38):

                cell = ws.cell(
                    row=r,
                    column=c
                )

                if type(cell).__name__ != "MergedCell":

                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center",
                        wrap_text=True
                    )

        # -------------------------------------------------
        # NEXT MACHINE
        # -------------------------------------------------
        row += 2

        current_machine += 1

        serial_no += 1

    # =====================================================
    # PAGE SETUP
    # =====================================================
    for sheet in wb.worksheets:

        sheet.page_setup.orientation = "landscape"

        sheet.page_setup.fitToWidth = 1

        sheet.page_setup.fitToHeight = 1

        sheet.sheet_properties.pageSetUpPr.fitToPage = True

        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4

        sheet.print_area = "A1:AK40"

    # =====================================================
    # CREATE DIRECTORY
    # =====================================================
    os.makedirs(
        download_path,
        exist_ok=True
    )

    # =====================================================
    # EXCEL FILE
    # =====================================================
    excel_file = os.path.join(
        download_path,
        f"MT 05 Preventive Maintenance Schedule-Monthly.xlsx"
    )

    wb.save(excel_file)

    # =====================================================
    # PDF CONVERT
    # =====================================================
    try:

        subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                excel_file,
                "--outdir",
                download_path
            ],
            check=True
        )

    except Exception as e:

        print("PDF ERROR:", e)

    # =====================================================
    # PDF PATH
    # =====================================================
    pdf_file = excel_file.replace(
        ".xlsx",
        ".pdf"
    )

    # =====================================================
    # RETURN
    # =====================================================
    return {
        "excel_path": excel_file,
        "pdf_path": pdf_file
    }