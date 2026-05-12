from supabase import create_client
from dotenv import load_dotenv
import os

from reports.breakdown import generate_breakdown_report
from reports.complaint_excel import generate_complaint_excel
from reports.equipment_master import generate_equipment_master
from reports.tools_tackles import generate_tools_tackles_report

# -------------------------------
# LOAD ENV
# -------------------------------
load_dotenv()

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# -------------------------------
# REPORT SELECTION
# -------------------------------
while True:

    print("\nSelect Report:")
    print("1 → Breakdown (Word)")
    print("2 → Complaint Register (Excel)")
    print("3 → Equipment Master (Excel)")
    print("4 → Tools & Tackles Report (Excel)")

    report_choice = input("\nEnter choice: ").strip().lower()

    if report_choice in ["1", "breakdown"]:
        report_type = "breakdown"
        break

    elif report_choice in ["2", "complaint"]:
        report_type = "complaint"
        break

    elif report_choice in ["3", "equipment"]:
        report_type = "equipment"
        break

    elif report_choice in ["4", "tools", "tools_tackles"]:
        report_type = "tools"
        break

    else:
        print("❌ Invalid choice. Try again.")

# -------------------------------
# OUTPUT FOLDER
# -------------------------------
download_path = os.path.expanduser("~/Downloads")
os.makedirs(download_path, exist_ok=True)

# -------------------------------
# EQUIPMENT REPORT
# (NO TICKET REQUIRED)
# -------------------------------
if report_type == "equipment":

    try:

        generate_equipment_master(
            download_path,
            supabase
        )

        print("\n🎉 Equipment Report generated successfully!")

    except Exception as e:
        print("❌ Error:", str(e))

    exit()

# -------------------------------
# FETCH TICKETS
# (FOR BREAKDOWN / COMPLAINT / TOOLS)
# -------------------------------
tickets = supabase.table("v_complaint_register") \
    .select("ticket_no") \
    .order("ticket_no") \
    .execute().data

if not tickets:
    print("❌ No tickets found")
    exit()

# -------------------------------
# DISPLAY TICKETS
# -------------------------------
print("\n========== AVAILABLE TICKETS ==========\n")

for i, t in enumerate(tickets, start=1):
    print(f"{i}. {t['ticket_no']}")

print("\n=======================================\n")

# -------------------------------
# GET TICKET INPUT
# -------------------------------
while True:

    choice = input("Enter Ticket Number OR Index: ").strip()

    if not choice or choice.startswith("\x1b"):
        print("❌ Invalid input. Try again.")
        continue

    # -------------------------------
    # INDEX INPUT
    # -------------------------------
    if choice.isdigit():

        index = int(choice) - 1

        if 0 <= index < len(tickets):
            ticket_no = tickets[index]["ticket_no"]
            break

        else:
            print("❌ Invalid index")

    # -------------------------------
    # MANUAL TICKET INPUT
    # -------------------------------
    else:

        ticket_no = choice.upper()

        if ticket_no.startswith("KMS"):
            break

        else:
            print("❌ Invalid Ticket Format")

print(f"\n✅ Selected Ticket: {ticket_no}")

# -------------------------------
# GENERATE REPORT
# -------------------------------
try:

    # -------------------------------
    # BREAKDOWN REPORT
    # -------------------------------
    if report_type == "breakdown":

        res = supabase.table("v_breakdown_intimation__report") \
            .select("*") \
            .eq("ticket_no", ticket_no) \
            .execute()

        if not res.data:
            print("❌ No breakdown data found")
            exit()

        generate_breakdown_report(
            res.data[0],
            download_path
        )

    # -------------------------------
    # COMPLAINT REPORT
    # -------------------------------
    elif report_type == "complaint":

        res = supabase.table("v_complaint_register") \
            .select("*") \
            .eq("ticket_no", ticket_no) \
            .execute()

        if not res.data:
            print("❌ No complaint data found")
            exit()

        generate_complaint_excel(
            res.data[0],
            download_path
        )

    # -------------------------------
    # TOOLS & TACKLES REPORT
    # -------------------------------
    elif report_type == "tools":

        res = supabase.table("v_tools_tackles_report") \
            .select("*") \
            .eq("ticket_no", ticket_no) \
            .execute()

        if not res.data:
            print("❌ No tools data found")
            exit()

        generate_tools_tackles_report(
            res.data,
            download_path
        )

    print("\n🎉 Report generated successfully!")

# -------------------------------
# ERROR
# -------------------------------
except Exception as e:

    print("❌ Error:", str(e))