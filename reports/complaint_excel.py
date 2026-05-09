from openpyxl import load_workbook
from datetime import datetime
import os

def generate_complaint_excel(ticket, download_path):

    # -------------------------------
    # TIME FUNCTIONS
    # -------------------------------
    def parse_time(value):
        if not value:
            return None
        try:
            return datetime.fromisoformat(value)
        except:
            return None

    def f_date(dt):
        return dt.strftime("%d-%m-%Y") if dt else ""

    def f_time(dt):
        return dt.strftime("%H:%M") if dt else ""

    break_dt = parse_time(ticket.get("ticket_raised_time"))
    end_dt = parse_time(ticket.get("ticket_completion_time"))

    bd_date, bd_time = f_date(break_dt), f_time(break_dt)
    rc_date, rc_time = f_date(end_dt), f_time(end_dt)

    # -------------------------------
    # CORRECT TEMPLATE PATH
    # -------------------------------
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "complaint_register.xlsx")

    if not os.path.exists(TEMPLATE_PATH):
        print(f"❌ Template not found at: {TEMPLATE_PATH}")
        return

    # -------------------------------
    # LOAD TEMPLATE
    # -------------------------------
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # -------------------------------
    # REPLACE FUNCTION
    # -------------------------------
    def replace_all(sheet, key, value):
        value = str(value)
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and key in str(cell.value):
                    cell.value = str(cell.value).replace(key, value)

    # -------------------------------
    # REPLACE DATA
    # -------------------------------
    replace_all(ws, "{{ticket_no}}", ticket.get("ticket_no", ""))
    replace_all(ws, "{{description}}", ticket.get("title", ""))
    replace_all(ws, "{{machine_name}}", ticket.get("machine_name") or "")
    replace_all(ws, "{{area}}", ticket.get("area") or "")
    replace_all(ws, "{{raised_by}}", ticket.get("raised_by") or "")
    replace_all(ws, "{{assigned_to}}", ticket.get("assigned_to") or "")

    replace_all(ws, "{{bd_date}}", bd_date)
    replace_all(ws, "{{bd_time}}", bd_time)

    replace_all(ws, "{{rc_date}}", rc_date)
    replace_all(ws, "{{rc_time}}", rc_time)

    replace_all(ws, "{{action}}", "Maintenance Done")

    # -------------------------------
    # SAVE FILE
    # -------------------------------
    os.makedirs(download_path, exist_ok=True)
    file_path = os.path.join(download_path, f"complaint_{ticket.get('ticket_no')}.xlsx")

    wb.save(file_path)

    print("\n✅ Complaint Register Generated!")
    print("📁 Saved at:", file_path)