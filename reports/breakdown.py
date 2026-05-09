from docx import Document
from datetime import datetime
import os


def generate_breakdown_report(ticket, output_dir):

    # -------------------------------
    # TIME PARSER
    # -------------------------------
    def parse_time(value):
        if not value:
            return None
        try:
            return datetime.fromisoformat(value)
        except:
            return None

    break_dt = parse_time(ticket.get("ticket_raised_time"))
    start_dt = parse_time(ticket.get("repair_start_time"))
    end_dt = parse_time(ticket.get("ticket_completion_time"))

    def f_date(dt):
        return dt.strftime("%d-%m-%Y") if dt else ""

    def f_time(dt):
        return dt.strftime("%H:%M") if dt else ""

    bd_date, bd_time = f_date(break_dt), f_time(break_dt)
    rs_date, rs_time = f_date(start_dt), f_time(start_dt)
    rc_date, rc_time = f_date(end_dt), f_time(end_dt)

    # -------------------------------
    # DURATION
    # -------------------------------
    if start_dt and end_dt:
        mins = int((end_dt - start_dt).total_seconds() / 60)
        hrs = mins // 60
        rem = mins % 60
        rc_duration = f"{hrs} hrs {rem} min" if rem else f"{hrs} hrs"
        rc_remarks = "Completed"
    else:
        rc_duration = ""
        rc_remarks = ""

    # -------------------------------
    # LOAD TEMPLATE
    # -------------------------------
    doc = Document("templates/template.docx")

    # -------------------------------
    # BETTER REPLACE FUNCTION
    # -------------------------------
    def replace_all(doc, key, value):
        value = str(value)

        for p in doc.paragraphs:
            for run in p.runs:
                if key in run.text:
                    run.text = run.text.replace(key, value)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            if key in run.text:
                                run.text = run.text.replace(key, value)

        for section in doc.sections:
            header = section.header
            for p in header.paragraphs:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, value)

    # -------------------------------
    # SPARE FORMATTING (IMPORTANT)
    # -------------------------------
    spare_text = ticket.get("spare") or "No spares used"

    # Add spacing for clean display
    spare_text = "\n" + spare_text + "\n"

    # -------------------------------
    # REPLACE DATA
    # -------------------------------
    replace_all(doc, "{{ticket_no}}", ticket.get("ticket_no", ""))
    replace_all(doc, "{{machine_name}}", ticket.get("machine_name") or "N/A")
    replace_all(doc, "{{description}}", ticket.get("title", ""))
    replace_all(doc, "{{reason}}", ticket.get("cause_of_issue", ""))
    replace_all(doc, "{{user}}", ticket.get("raised_by") or "N/A")

    replace_all(doc, "{{bd_date}}", bd_date)
    replace_all(doc, "{{bd_time}}", bd_time)

    replace_all(doc, "{{rs_date}}", rs_date)
    replace_all(doc, "{{rs_time}}", rs_time)

    replace_all(doc, "{{rc_date}}", rc_date)
    replace_all(doc, "{{rc_time}}", rc_time)
    replace_all(doc, "{{rc_duration}}", rc_duration)
    replace_all(doc, "{{rc_remarks}}", rc_remarks)

    # ✅ SPARES
    replace_all(doc, "{{spares}}", spare_text)

    # -------------------------------
    # SAVE FILE
    # -------------------------------
    os.makedirs(output_dir, exist_ok=True)

    file_name = f"{output_dir}/report_{ticket.get('ticket_no', 'output')}.docx"
    doc.save(file_name)

    print("✅ Breakdown Report Generated:", file_name)