import os
import tempfile
import subprocess
import platform
import calendar
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from supabase import create_client
from dotenv import load_dotenv
from fastapi.middleware.cors import CORSMiddleware
import os
import calendar
import tempfile

# Import your existing report functions
from reports.breakdown import generate_breakdown_report
from reports.complaint_excel import generate_complaint_excel, generate_complaint_excel_bulk
from reports.equipment_master import generate_equipment_master
from reports.tools_tackles import generate_tools_tackles_report
from reports.preventive_maintenance_Schedule import generate_preventive_maintenance_report
from reports.critical_spares import generate_critical_spare_parts_report

# Load Environment Variables
load_dotenv()
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    raise ValueError("Missing Supabase credentials in environment variables.")

# Initialize Supabase and FastAPI
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
app = FastAPI(title="Kitchen Maintenance Reports API")

# Enable CORS so the Flutter App (Phone/Web) can communicate with the backend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------------------------------------------------
# HELPER: Convert DOCX/XLSX to PDF using LibreOffice
# ---------------------------------------------------------
def convert_to_pdf(input_path: str, output_dir: str) -> str:
    """
    Uses LibreOffice headless to convert a document to PDF.
    """
    try:
        # Smart detection for Mac vs Linux/Cloud
        if platform.system() == "Darwin":  
            libreoffice_exec = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        else:
            libreoffice_exec = "libreoffice"

        command = [
            libreoffice_exec, "--headless", "--convert-to", "pdf",
            "--outdir", output_dir, input_path
        ]
        
        # Run the command
        subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        
        # The output file has the same name but a .pdf extension
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        return os.path.join(output_dir, f"{base_name}.pdf")
        
    except Exception as e:
        raise Exception(f"PDF Conversion failed. Error: {str(e)}")

# ---------------------------------------------------------
# 1A. BREAKDOWN INTIMATION REPORT (Word -> PDF)
# ---------------------------------------------------------
@app.get("/api/reports/breakdown/{ticket_no}")
async def get_breakdown_report(ticket_no: str, format: str = "docx"):
    res = supabase.table("v_breakdown_intimation__report").select("*").eq("ticket_no", ticket_no.upper()).execute()
    if not res.data:
        raise HTTPException(status_code=404, detail="Ticket not found.")
    
    temp_dir = tempfile.gettempdir()
    generate_breakdown_report(res.data[0], temp_dir)
    docx_file_path = os.path.join(temp_dir, f"report_{ticket_no.upper()}.docx")
    
    if not os.path.exists(docx_file_path):
        raise HTTPException(status_code=500, detail="Failed to generate document.")

    if format.lower() == "pdf":
        try:
            pdf_file_path = convert_to_pdf(docx_file_path, temp_dir)
            return FileResponse(path=pdf_file_path, filename=f"Breakdown_{ticket_no.upper()}.pdf", media_type="application/pdf")
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))
    
    return FileResponse(path=docx_file_path, filename=f"Breakdown_{ticket_no.upper()}.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

import calendar # <-- Add this with the other imports at the very top!

# ---------------------------------------------------------
# 1B. BREAKDOWN INTIMATION REPORT (MONTHLY MERGED)
# ---------------------------------------------------------
@app.get("/api/reports/breakdowns/monthly")
async def get_monthly_breakdowns(month: int, year: int, format: str = "docx"):
    # 1. Calculate Start and End dates for the requested month
    _, last_day = calendar.monthrange(year, month)
    start_iso = f"{year}-{month:02d}-01T00:00:00.000Z"
    end_iso = f"{year}-{month:02d}-{last_day}T23:59:59.999Z"
    
    # 2. Fetch only Breakdown tickets for that timeframe
    res = supabase.table("v_breakdown_intimation__report") \
        .select("*") \
        .gte("ticket_raised_time", start_iso) \
        .lte("ticket_raised_time", end_iso) \
        .execute()
        
    if not res.data:
        raise HTTPException(status_code=404, detail="No breakdown reports found for this month.")
        
    temp_dir = tempfile.gettempdir()
    month_abbr = calendar.month_abbr[month]
    
    # 3. Generate the Merged Single File
    from reports.breakdown import generate_monthly_breakdown_single_file
    try:
        filepath = generate_monthly_breakdown_single_file(
            tickets=res.data, 
            output_dir=temp_dir, 
            month_abbr=month_abbr, 
            year=year, 
            format=format.lower(), 
            pdf_converter=convert_to_pdf
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to merge documents: {str(e)}")
    
    # 4. Return the Single Document back to the phone
    media_type = "application/pdf" if format.lower() == "pdf" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    
    return FileResponse(
        path=filepath, 
        filename=f"Breakdown_Reports_{month_abbr}_{year}.{format.lower()}",
        media_type=media_type
    )

# ---------------------------------------------------------
# 2A. COMPLAINT REGISTER (SINGLE TICKET)
# ---------------------------------------------------------
@app.get("/api/reports/complaint/{ticket_no}")
async def get_complaint_report(ticket_no: str, format: str = "xlsx"):
    res = supabase.table("v_complaint_register").select("*").eq("ticket_no", ticket_no.upper()).execute()
    if not res.data:
        raise HTTPException(status_code=404, detail="Ticket not found.")
    
    temp_dir = tempfile.gettempdir()
    generate_complaint_excel(res.data[0], temp_dir)
    
    xlsx_file_path = os.path.join(temp_dir, f"complaint_{ticket_no.upper()}.xlsx")

    if not os.path.exists(xlsx_file_path):
        raise HTTPException(status_code=500, detail="Failed to generate Excel file.")

    if format.lower() == "pdf":
        try:
            pdf_file_path = convert_to_pdf(xlsx_file_path, temp_dir)
            return FileResponse(path=pdf_file_path, filename=f"Complaint_Register_{ticket_no.upper()}.pdf", media_type="application/pdf")
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))

    return FileResponse(path=xlsx_file_path, filename=f"Complaint_Register_{ticket_no.upper()}.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------------------------------------------
# 2B. COMPLAINT REGISTER (DATE RANGE / MULTI-SHEET)
# ---------------------------------------------------------
@app.get("/api/reports/complaints/range")
async def get_complaint_report_range(start: str, end: str, format: str = "xlsx"):
    # Ensure the end date covers the entire day
    end_date_inclusive = f"{end}T23:59:59.999Z"
    
    res = supabase.table("v_complaint_register") \
        .select("*") \
        .gte("ticket_raised_time", start) \
        .lte("ticket_raised_time", end_date_inclusive) \
        .order("ticket_raised_time") \
        .execute()
        
    if not res.data:
        raise HTTPException(status_code=404, detail="No complaints found in this date range.")
        
    temp_dir = tempfile.gettempdir()
    generate_complaint_excel_bulk(res.data, temp_dir, start, end)
    
    # Sanitize dates to prevent folder creation errors on Mac/Linux
    safe_start = str(start).replace("/", "-")
    safe_end = str(end).replace("/", "-")
    
    xlsx_file_path = os.path.join(temp_dir, f"Complaint_Register_Tabs_{safe_start}_to_{safe_end}.xlsx")

    if not os.path.exists(xlsx_file_path):
        raise HTTPException(status_code=500, detail="Failed to generate bulk Excel file.")

    if format.lower() == "pdf":
        try:
            pdf_file_path = convert_to_pdf(xlsx_file_path, temp_dir)
            return FileResponse(path=pdf_file_path, filename=f"Complaint_Register_{safe_start}_to_{safe_end}.pdf", media_type="application/pdf")
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))

    return FileResponse(path=xlsx_file_path, filename=f"Complaint_Register_Tabs_{safe_start}_to_{safe_end}.xlsx", media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------------------------------------------------------
# 3. EQUIPMENT MASTER (Excel / PDF)
# ---------------------------------------------------------
@app.get("/api/reports/equipment")
async def get_equipment_master(month: int = None, year: int = None, format: str = "xlsx"):
    query = supabase.table("v_equipment_master").select("*").order("machine_name")
    
    # Apply Monthly filters if requested
    if month and year:
        # Automatically calculate the last day of the selected month
        _, last_day = calendar.monthrange(year, month)
        
        # Format the dates for Supabase
        start_iso = f"{year}-{month:02d}-01T00:00:00.000Z"
        end_iso = f"{year}-{month:02d}-{last_day}T23:59:59.999Z"
        
        query = query.gte("created_at", start_iso).lte("created_at", end_iso)

    res = query.execute()

    if not res.data:
        raise HTTPException(status_code=404, detail="No equipment found for the selected month.")

    temp_dir = tempfile.gettempdir()
    
    # Generate a beautiful dynamic file name
    if month and year:
        month_name = calendar.month_abbr[month] # e.g., "Jan", "Feb"
        base_name = f"Equipment_Master_{month_name}_{year}"
    else:
        base_name = "Equipment_Master_All"

    xlsx_filename = f"{base_name}.xlsx"
    
    # Pass the raw data to the python generator
    generate_equipment_master(res.data, temp_dir, xlsx_filename)
    
    xlsx_file_path = os.path.join(temp_dir, xlsx_filename)

    if not os.path.exists(xlsx_file_path):
        raise HTTPException(status_code=500, detail="Failed to generate Equipment Master.")

    if format.lower() == "pdf":
        try:
            pdf_file_path = convert_to_pdf(xlsx_file_path, temp_dir)
            return FileResponse(path=pdf_file_path, filename=f"{base_name}.pdf", media_type="application/pdf")
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))

    return FileResponse(path=xlsx_file_path, filename=xlsx_filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    query = supabase.table("v_equipment_master").select("*").order("machine_name")
    
    # Apply date filters if requested
    if start and end:
        end_date_inclusive = f"{end}T23:59:59.999Z"
        query = query.gte("created_at", start).lte("created_at", end_date_inclusive)

    res = query.execute()

    if not res.data:
        raise HTTPException(status_code=404, detail="No equipment found for the selected dates.")

    temp_dir = tempfile.gettempdir()
    
    # Generate dynamic file name based on user query
    if start and end:
        safe_start = str(start).replace("/", "-")
        safe_end = str(end).replace("/", "-")
        base_name = f"Equipment_Master_{safe_start}_to_{safe_end}"
    else:
        base_name = "Equipment_Master_All"

    xlsx_filename = f"{base_name}.xlsx"
    
    # Pass the raw data to the python generator
    generate_equipment_master(res.data, temp_dir, xlsx_filename)
    
    xlsx_file_path = os.path.join(temp_dir, xlsx_filename)

    if not os.path.exists(xlsx_file_path):
        raise HTTPException(status_code=500, detail="Failed to generate Equipment Master.")

    if format.lower() == "pdf":
        try:
            pdf_file_path = convert_to_pdf(xlsx_file_path, temp_dir)
            return FileResponse(path=pdf_file_path, filename=f"{base_name}.pdf", media_type="application/pdf")
        except Exception as e:
            raise HTTPException(status_code=500, detail=str(e))

    return FileResponse(path=xlsx_file_path, filename=xlsx_filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------
# 4A. DAILY TOOLS REPORT
# ---------------------------------------------------------
@app.get("/api/reports/tools/{date}")
def get_daily_tools_report(date: str):

    try:

        # ---------------------------------------------------------
        # FETCH DATEWISE DATA
        # date format:
        # 2026-05-11
        # ---------------------------------------------------------
        res = supabase.table(
            "v_tools_tackles_report"
        ).select("*") \
         .gte("taken_time", f"{date}T00:00:00") \
         .lte("taken_time", f"{date}T23:59:59") \
         .order("taken_time") \
         .execute()

        data = res.data or []

        # ---------------------------------------------------------
        # NO DATA
        # ---------------------------------------------------------
        if not data:

            raise HTTPException(
                status_code=404,
                detail=f"No tools data found for {date}"
            )

        # ---------------------------------------------------------
        # TEMP DIR
        # ---------------------------------------------------------
        temp_dir = tempfile.gettempdir()

        # ---------------------------------------------------------
        # GENERATE REPORT
        # ---------------------------------------------------------
        file_path = generate_tools_tackles_report(
            data=data,
            download_path=temp_dir,
            filename=f"Daily_Tools_Report_{date}.docx"
        )

        # ---------------------------------------------------------
        # RETURN DOCX
        # ---------------------------------------------------------
        return FileResponse(
            path=file_path,
            filename=os.path.basename(file_path),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )


# ---------------------------------------------------------
# 4B. MONTHLY TOOLS REPORT
# ---------------------------------------------------------
@app.get("/api/reports/tools/monthly/{month}")
def get_monthly_tools_report(month: str):

    try:

        # ---------------------------------------------------------
        # FETCH MONTHWISE DATA
        # month format:
        # 2026-05
        # ---------------------------------------------------------
        res = supabase.table(
            "v_tools_tackles_report"
        ).select("*") \
         .gte("taken_time", f"{month}-01T00:00:00") \
         .lte("taken_time", f"{month}-31T23:59:59") \
         .order("taken_time") \
         .execute()

        data = res.data or []

        # ---------------------------------------------------------
        # NO DATA
        # ---------------------------------------------------------
        if not data:

            raise HTTPException(
                status_code=404,
                detail=f"No monthly data found for {month}"
            )

        # ---------------------------------------------------------
        # TEMP DIR
        # ---------------------------------------------------------
        temp_dir = tempfile.gettempdir()

        # ---------------------------------------------------------
        # GENERATE REPORT
        # ---------------------------------------------------------
        file_path = generate_tools_tackles_report(
            data=data,
            download_path=temp_dir, 
            filename=f"Monthly_Tools_Report_{month}.docx"
        )

        # ---------------------------------------------------------
        # RETURN DOCX
        # ---------------------------------------------------------
        return FileResponse(
            path=file_path,
            filename=os.path.basename(file_path),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )
    

# ---------------------------------------------------------
# 5. PREVENTIVE MAINTENANCE REPORT
# ---------------------------------------------------------
@app.get("/api/reports/preventive-maintenance/{month}")
def get_preventive_maintenance_report(
    month: str,
    format: str = "xlsx"
):

    try:

        import tempfile

        # ---------------------------------------------------------
        # VALIDATE FORMAT
        # ---------------------------------------------------------
        format = format.lower()

        if format not in ["xlsx", "pdf"]:

            raise HTTPException(
                status_code=400,
                detail="Format must be either 'xlsx' or 'pdf'"
            )

        # ---------------------------------------------------------
        # FETCH DATA FROM VIEW
        # ---------------------------------------------------------
        response = supabase.table(
            "v_preventive_maintenance_schedule"
        ).select("*") \
        .order("machine_name") \
        .execute()

        data = response.data or []

        print(data)

        # ---------------------------------------------------------
        # NO DATA
        # ---------------------------------------------------------
        if not data:

            raise HTTPException(
                status_code=404,
                detail="No preventive maintenance data found"
            )

        # ---------------------------------------------------------
        # FILTER MONTH DATA
        # ---------------------------------------------------------
        filtered_data = [
            item for item in data
            if (
                (
                    item.get("plan_date")
                    and str(item["plan_date"]).startswith(month)
                )
                or
                (
                    item.get("achieved_date")
                    and str(item["achieved_date"]).startswith(month)
                )
            )
        ]

        # ---------------------------------------------------------
        # NO FILTERED DATA
        # ---------------------------------------------------------
        if not filtered_data:

            raise HTTPException(
                status_code=404,
                detail=f"No preventive maintenance data found for {month}"
            )

        # ---------------------------------------------------------
        # TEMP DIRECTORY
        # ---------------------------------------------------------
        temp_dir = tempfile.gettempdir()

        # ---------------------------------------------------------
        # GENERATE REPORT
        # ---------------------------------------------------------
        files = generate_preventive_maintenance_report(
            data=filtered_data,
            download_path=temp_dir,
            month=month
        )

        # ---------------------------------------------------------
        # RETURN PDF
        # ---------------------------------------------------------
        if format == "pdf":

            return FileResponse(
                path=files["pdf_path"],
                filename=os.path.basename(
                    files["pdf_path"]
                ),
                media_type="application/pdf"
            )

        # ---------------------------------------------------------
        # RETURN EXCEL
        # ---------------------------------------------------------
        return FileResponse(
            path=files["excel_path"],
            filename=os.path.basename(
                files["excel_path"]
            ),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )
    
# ---------------------------------------------------------
# 6. CRITICAL SPARE PARTS REPORT
# ---------------------------------------------------------
@app.get("/api/reports/critical-spares/{month}")
def get_critical_spares_report(
    month: str,
    format: str = "docx"
):

    try:

        import tempfile

        # ---------------------------------------------------------
        # VALIDATE FORMAT
        # ---------------------------------------------------------
        format = format.lower()

        if format not in ["docx", "pdf"]:

            raise HTTPException(
                status_code=400,
                detail="Format must be docx or pdf"
            )

        # ---------------------------------------------------------
        # FETCH DATA FROM VIEW
        # ---------------------------------------------------------
        response = supabase.table(
            "v_critical_spare_parts_report"
        ).select("*") \
        .order("spare_type") \
        .execute()

        data = response.data or []

        # ---------------------------------------------------------
        # NO DATA
        # ---------------------------------------------------------
        if not data:

            raise HTTPException(
                status_code=404,
                detail="No critical spare parts found"
            )

        # ---------------------------------------------------------
        # TEMP DIR
        # ---------------------------------------------------------
        temp_dir = tempfile.gettempdir()

        # ---------------------------------------------------------
        # FILE NAME
        # ---------------------------------------------------------
        filename = (
            f"MT 15 Critical Spare parts_{month}.docx"
        )

        # ---------------------------------------------------------
        # GENERATE REPORT
        # ---------------------------------------------------------
        file_path = generate_critical_spare_parts_report(
            data=data,
            download_path=temp_dir,
            filename=filename
        )

        # ---------------------------------------------------------
        # PDF
        # ---------------------------------------------------------
        if format == "pdf":

            pdf_path = convert_to_pdf(
                file_path,
                temp_dir
            )

            return FileResponse(
                path=pdf_path,
                filename=f"MT 15 Critical Spare parts_{month}.pdf",
                media_type="application/pdf"
            )

        # ---------------------------------------------------------
        # DOCX
        # ---------------------------------------------------------
        return FileResponse(
            path=file_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:

        raise HTTPException(
            status_code=500,
            detail=str(e)
        )   