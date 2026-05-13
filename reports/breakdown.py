from docx import Document
from datetime import datetime
import os
import shutil
from docxcompose.composer import Composer

def generate_breakdown_report(ticket, output_dir):

    # -------------------------------
    # TIME PARSER
    # -------------------------------
    def parse_time(value):
        if not value:
            return None
        try:
            return datetime.fromisoformat(str(value))
        except:
            return None

    # Use breakdown_time if available, otherwise fallback to ticket_raised_time
    bd_str = ticket.get("breakdown_time") or ticket.get("ticket_raised_time")
    
    break_dt = parse_time(bd_str)
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
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "template.docx")
    
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template not found at: {TEMPLATE_PATH}")

    doc = Document(TEMPLATE_PATH)

    # -------------------------------
    # BETTER REPLACE FUNCTION
    # -------------------------------
    def replace_all(doc, key, value):
        value = str(value) if value is not None else ""

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
    # FORMATTING
    # -------------------------------
    spare_text = ticket.get("spare") or "No spares used"
    spare_text = "\n" + spare_text + "\n"

    # -------------------------------
    # REPLACE DATA
    # -------------------------------
    replace_all(doc, "{{ticket_no}}", ticket.get("ticket_no", ""))
    replace_all(doc, "{{equipment_code}}", ticket.get("equipment_code") or "") # NEW FIELD
    replace_all(doc, "{{machine_name}}", ticket.get("machine_name") or "N/A")
    replace_all(doc, "{{description}}", ticket.get("title", ""))
    replace_all(doc, "{{reason}}", ticket.get("cause_of_issue") or "")
    replace_all(doc, "{{action_taken}}", ticket.get("action_taken") or "")     # NEW FIELD
    replace_all(doc, "{{user}}", ticket.get("raised_by") or "N/A")

    replace_all(doc, "{{bd_date}}", bd_date)
    replace_all(doc, "{{bd_time}}", bd_time)

    replace_all(doc, "{{rs_date}}", rs_date)
    replace_all(doc, "{{rs_time}}", rs_time)

    replace_all(doc, "{{rc_date}}", rc_date)
    replace_all(doc, "{{rc_time}}", rc_time)
    replace_all(doc, "{{rc_duration}}", rc_duration)
    replace_all(doc, "{{rc_remarks}}", rc_remarks)

    # SPARES
    replace_all(doc, "{{spares}}", spare_text)

    # -------------------------------
    # SAVE FILE
    # -------------------------------
    os.makedirs(output_dir, exist_ok=True)
    file_name = os.path.join(output_dir, f"report_{ticket.get('ticket_no', 'output')}.docx")
    doc.save(file_name)

    print("✅ Breakdown Report Generated:", file_name)

def generate_monthly_breakdown_single_file(tickets, output_dir, month_abbr, year, format="docx", pdf_converter=None):
    """
    Generates multiple breakdown reports and seamlessly merges them into 
    a single multi-page Word Document (or PDF).
    """
    final_filename = f"Breakdown_Reports_{month_abbr}_{year}"
    final_docx_path = os.path.join(output_dir, f"{final_filename}.docx")
    
    # Temporary folder to hold the generated files before merging
    temp_batch_dir = os.path.join(output_dir, f"batch_{month_abbr}_{year}")
    os.makedirs(temp_batch_dir, exist_ok=True)
    
    generated_docs = []
    
    # 1. Generate all individual documents in the background
    for ticket in tickets:
        ticket_no = ticket.get("ticket_no", "Unknown")
        generate_breakdown_report(ticket, temp_batch_dir)
        
        docx_path = os.path.join(temp_batch_dir, f"report_{ticket_no.upper()}.docx")
        if os.path.exists(docx_path):
            generated_docs.append(docx_path)
            
    if not generated_docs:
        raise Exception("No documents were generated.")
        
    # 2. Merge all documents into a single master file
    master_doc = Document(generated_docs[0])
    composer = Composer(master_doc)
    
    for doc_path in generated_docs[1:]:
        # Add a clean page break before appending the next report
        master_doc.add_page_break() 
        sub_doc = Document(doc_path)
        composer.append(sub_doc)
        
    composer.save(final_docx_path)
    
    # Clean up the individual files to save server space
    shutil.rmtree(temp_batch_dir, ignore_errors=True)
    
    # 3. Handle PDF conversion if requested
    if format == "pdf" and pdf_converter:
        pdf_path = pdf_converter(final_docx_path, output_dir)
        return pdf_path
        
    return final_docx_path
    """
    Generates multiple breakdown reports, bundles them into a single .zip file,
    and returns the path to the zip file.
    """
    zip_filename = f"Breakdown_Reports_{month_abbr}_{year}.zip"
    zip_filepath = os.path.join(output_dir, zip_filename)
    
    # Temporary folder to hold the generated files before zipping
    temp_batch_dir = os.path.join(output_dir, f"batch_{month_abbr}_{year}")
    os.makedirs(temp_batch_dir, exist_ok=True)
    
    with zipfile.ZipFile(zip_filepath, 'w') as zipf:
        for ticket in tickets:
            ticket_no = ticket.get("ticket_no", "Unknown")
            
            # 1. Generate the standard .docx inside the temporary batch folder
            generate_breakdown_report(ticket, temp_batch_dir)
            docx_path = os.path.join(temp_batch_dir, f"report_{ticket_no.upper()}.docx")
            
            # 2. Add to Zip (converting to PDF first if requested)
            if format == "pdf" and pdf_converter:
                pdf_path = pdf_converter(docx_path, temp_batch_dir)
                zipf.write(pdf_path, f"Breakdown_{ticket_no.upper()}.pdf")
            else:
                zipf.write(docx_path, f"Breakdown_{ticket_no.upper()}.docx")
                
    # Clean up the unzipped files to save server space
    shutil.rmtree(temp_batch_dir, ignore_errors=True)
    
    return zip_filepath