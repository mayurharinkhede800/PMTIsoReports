from docx import Document
import os


def generate_critical_spare_parts_report(
    data,
    download_path,
    filename="MT 15 Critical Spare parts.docx"
):

    # ---------------------------------------------------------
    # TEMPLATE PATH
    # ---------------------------------------------------------
    BASE_DIR = os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))
    )

    TEMPLATE_PATH = os.path.join(
        BASE_DIR,
        "templates",
        "MT 15 Critical Spare parts.docx"
    )

    if not os.path.exists(TEMPLATE_PATH):

        raise FileNotFoundError(
            f"Template not found: {TEMPLATE_PATH}"
        )

    # ---------------------------------------------------------
    # LOAD DOCUMENT
    # ---------------------------------------------------------
    doc = Document(TEMPLATE_PATH)

    # ---------------------------------------------------------
    # TABLE
    # ---------------------------------------------------------
    table = doc.tables[0]

    # ---------------------------------------------------------
    # TEMPLATE DATA START ROW
    # ---------------------------------------------------------
    # Row structure:
    #
    # 0 -> Logo/Header
    # 1 -> Main Header
    # 2 -> Title
    # 3 -> Quantity Header
    # 4 -> Column Header
    # 5 onwards -> Data Rows
    #
    # Adjust if needed
    # ---------------------------------------------------------
    start_row = 2

    # ---------------------------------------------------------
    # TOTAL AVAILABLE TEMPLATE ROWS
    # ---------------------------------------------------------
    max_rows = 12

    # ---------------------------------------------------------
    # FILL TEMPLATE ROWS
    # ---------------------------------------------------------
    for index, item in enumerate(data[:max_rows]):

        current_row = start_row + index

        # SAFETY CHECK
        if current_row >= len(table.rows):
            break

        row = table.rows[current_row].cells

        # ---------------------------------------------------------
        # S.NO
        # ---------------------------------------------------------
        row[0].text = str(index + 1)

        # ---------------------------------------------------------
        # SPARE NAME
        # ---------------------------------------------------------
        row[1].text = str(
            item.get("spare_name") or ""
        )

        # ---------------------------------------------------------
        # SPARE TYPE
        # ---------------------------------------------------------
        row[2].text = str(
            item.get("spare_type") or ""
        )

        # ---------------------------------------------------------
        # ON STOCK
        # ---------------------------------------------------------
        row[3].text = str(
            item.get("on_stock") or 0
        )

        # ---------------------------------------------------------
        # MAINTAINED STOCK
        # ---------------------------------------------------------
        row[4].text = str(
            item.get("maintained_stock") or 0
        )

        # ---------------------------------------------------------
        # SIGNATURE
        # ---------------------------------------------------------
        row[5].text = ""

    # ---------------------------------------------------------
    # SAVE FILE
    # ---------------------------------------------------------
    os.makedirs(download_path, exist_ok=True)

    file_path = os.path.join(
        download_path,
        filename
    )

    doc.save(file_path)

    print(
        "✅ Critical Spare Parts Report Generated:",
        file_path
    )

    return file_path