import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side
import random

# ================= دوال استخراج البيانات =================
def extract_table_data(pdf_file):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    total_pages = len(doc)
    rows = []

    for page_num in range(total_pages):
        page = doc[page_num]
        text = page.get_text("text")
        lines = text.split('\n')

        workorder = re.search(r'WORKORDER\s*#\s*:\s*(\d+)', text)
        floor_name = None
        floor_match = re.search(r'Mataf Building Project\s*,([^,]+),([^,]+)', text)
        if floor_match:
            floor_name = floor_match.group(2).strip()

        phase = re.search(r'Phase\s*#\s*(\d+)', text)
        column = re.search(r'Column\s*([A-Z0-9]+)', text)
        axis = re.search(r'Axis\s*([A-Z0-9]+)', text)
        qty = re.search(r'Asset QTY\s*:\s*(\d+)', text)
        type_check = re.search(r'JP Code\s*:\s*([A-Z0-9\-]+)', text)
        date = re.search(r'Scheduel Start\s*:\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})', text)

        last_letter = ""
        if type_check:
            match = re.search(r'([A-Z])$', type_check.group(1))
            if match:
                last_letter = match.group(1)

        rows.append({
            "Workorder": workorder.group(1) if workorder else "",
            "Floor": floor_name if floor_name else "",
            "Phase": phase.group(1) if phase else "",
            "Column": column.group(1) if column else "",
            "Axis": axis.group(1) if axis else "",
            "Quantity": qty.group(1) if qty else "",
            "Equipment": "FHC",
            "Type of Check": last_letter,
            "Date": date.group(1) if date else ""
        })

    return rows

# ================= دالة ألوان وحدود Excel =================
def apply_colors_and_borders(excel_filename):
    wb = load_workbook(excel_filename)
    ws = wb.active

    date_col = 9  # عمود التاريخ
    max_row = ws.max_row
    dates = [ws.cell(row=r, column=date_col).value for r in range(2, max_row + 1)]
    unique_dates = list(sorted(set([d for d in dates if d is not None])))

    def light_color():
        r = random.randint(180, 255)
        g = random.randint(180, 255)
        b = random.randint(180, 255)
        return f"{r:02X}{g:02X}{b:02X}"

    color_map = {d: light_color() for d in unique_dates}

    border_style = Side(border_style="thin", color="000000")
    border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

    for r in range(2, max_row + 1):
        cell_date = ws.cell(row=r, column=date_col).value
        fill_color = color_map.get(cell_date, "FFFFFF") if cell_date else "FFFFFF"
        fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill
            cell.border = border

    wb.save(excel_filename)

# ================= دالة Streamlit =================
def main():
    st.title("PDF to Excel Converter")
    st.write("رفع ملفات PDF ومعالجتها مباشرة وتنزيل Excel")

    uploaded_file = st.file_uploader("اختر ملف PDF", type="pdf")
    if uploaded_file:
        data = extract_table_data(uploaded_file)
        if not data:
            st.warning("لم يتم العثور على بيانات صالحة في هذا الملف.")
            return

        df = pd.DataFrame(data)

        df['Floor'] = df['Floor'].str.strip().str.title()
        floor_symbol_map = {
            'Basement': 'B',
            'Bs': 'B',
            'Ground Floor': 'GF',
            'Ground Mezanin': 'GM',
            'First Floor': '1',
            'First Mezzanine Floor': '1M',
            'Second Mezzanine Floor': '2M',
            'Third Mezzanine Floor': '3M',
            'First Mezanin': '1M',
            'Second Floor': '2',
            'Second Mezanin': '2M',
            'Third Floor': '3',
            'Third Mezanin': '3M',
            'Fourth Floor': '4',
            'Fifth Floor': '5',
            'Sixth Floor': '6',
            'Seventh Floor': '7',
            'Eighth Floor': '8',
            'Ninth Floor': '9',
            'Rf': 'R',
            'Roof Floor': 'R'
        }
        df['Floor'] = df['Floor'].map(floor_symbol_map).fillna(df['Floor'])

        df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.date
        df = df.sort_values(by='Date').reset_index(drop=True)

        output_file = "output.xlsx"
        df.to_excel(output_file, index=False)
        apply_colors_and_borders(output_file)

        st.success("تم إنشاء ملف Excel!")
        st.download_button("تحميل الملف", output_file)

if __name__ == "__main__":
    main()
