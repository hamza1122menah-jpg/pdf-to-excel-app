import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import random
import io

# -----------------------------
# خريطة الأدوار والرموز
floor_symbol_map = {
    'Basement': 'B',
    'Bs': 'B',
    'Ground Floor': 'GF',
    'Ground Mezanin': 'GM',
    'First Floor': '1',
    'First Mezzanine Floor': '1M',
    'Second Floor': '2',
    'Second Mezzanine Floor': '2M',
    'Third Floor': '3',
    'Third Mezzanine Floor': '3M',
    'Fourth Floor': '4',
    'Fifth Floor': '5',
    'Sixth Floor': '6',
    'Seventh Floor': '7',
    'Eighth Floor': '8',
    'Ninth Floor': '9',
    'Rf': 'R',
    'Roof Floor': 'R'
}
# -----------------------------

# -----------------------------
# دوال المشروع 1
def extract_project1(pdf_file):
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    rows = []

    for page in doc:
        text = page.get_text("text")
        lines = text.split('\n')
        mataf_line = ""
        for line in lines:
            if "Mataf Building Project" in line and "Phase" in line:
                mataf_line = line
                break
        workorder = re.search(r'WORKORDER\s*#\s*:\s*(\d+)', text)
        floor_match = re.search(r'Mataf Building Project\s*,([^,]+),([^,]+)', mataf_line)
        phase = re.search(r'Phase\s*#\s*(\d+)', mataf_line) if mataf_line else None
        column = re.search(r'Column\s*([A-Z0-9]+)', mataf_line)
        axis = re.search(r'Axis\s*([A-Z0-9]+)', mataf_line)
        qty = re.search(r'Asset QTY\s*:\s*(\d+)', text)
        type_check = re.search(r'JP Code\s*:\s*([A-Z0-9\-]+)', text)
        date_match = re.search(r'Scheduel Start\s*:\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})', text)

        last_letter = ""
        if type_check:
            match = re.search(r'([A-Z])$', type_check.group(1))
            if match:
                last_letter = match.group(1)

        floor_name = floor_match.group(2).strip() if floor_match else ""
        rows.append({
            "workorder num": workorder.group(1) if workorder else "",
            "Floor": floor_name,
            "phase": phase.group(1) if phase else "",
            "Column": column.group(1) if column else "",
            "Axis": axis.group(1) if axis else "",
            "Quantity": qty.group(1) if qty else "",
            "Equipment": "FHC",
            "Type of check": last_letter,
            "Date": date_match.group(1) if date_match else ""
        })
    return rows

# -----------------------------
# دوال المشروع 2
def extract_project2(pdf_file):
    # هنا تضع كود المعالجة الخاص بالمشروع الثاني
    # مثال مبسط: نفس المشروع الأول لكن يمكن تغييره لاحقاً
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    rows = []
    for page in doc:
        text = page.get_text("text")
        # تعديل البحث وفق المشروع 2
        workorder = re.search(r'Order\s*#\s*:\s*(\d+)', text)
        floor_match = re.search(r'Building Project 2\s*,([^,]+),([^,]+)', text)
        phase = re.search(r'Phase\s*:\s*(\d+)', text)
        column = re.search(r'Col\s*([A-Z0-9]+)', text)
        axis = re.search(r'Axe\s*([A-Z0-9]+)', text)
        qty = re.search(r'Qty\s*:\s*(\d+)', text)
        type_check = re.search(r'Check\s*Code\s*:\s*([A-Z0-9\-]+)', text)
        date_match = re.search(r'Start Date\s*:\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})', text)

        last_letter = ""
        if type_check:
            match = re.search(r'([A-Z])$', type_check.group(1))
            if match:
                last_letter = match.group(1)

        floor_name = floor_match.group(2).strip() if floor_match else ""
        rows.append({
            "workorder num": workorder.group(1) if workorder else "",
            "Floor": floor_name,
            "phase": phase.group(1) if phase else "",
            "Column": column.group(1) if column else "",
            "Axis": axis.group(1) if axis else "",
            "Quantity": qty.group(1) if qty else "",
            "Equipment": "FHC",
            "Type of check": last_letter,
            "Date": date_match.group(1) if date_match else ""
        })
    return rows

# -----------------------------
def apply_colors_to_excel(excel_filename):
    wb = load_workbook(excel_filename)
    ws = wb.active
    date_col = 9
    max_row = ws.max_row
    dates = [ws.cell(row=row, column=date_col).value for row in range(2, max_row + 1)]
    filtered_dates = [d for d in dates if d]
    unique_dates = list(sorted(set(filtered_dates)))
    color_map = {}
    def light_random_color():
        r = random.randint(180, 255)
        g = random.randint(180, 255)
        b = random.randint(180, 255)
        return f"{r:02X}{g:02X}{b:02X}"
    for d in unique_dates:
        color_map[d] = light_random_color()
    for row in range(2, max_row + 1):
        cell_date = ws.cell(row=row, column=date_col).value
        fill_color = color_map.get(cell_date, "FFFFFF") if cell_date else "FFFFFF"
        fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).fill = fill
    wb.save(excel_filename)

# -----------------------------
# واجهة Streamlit
st.title("PDF to Excel - Multi Projects")
st.write("اختر المشروع ورفع ملف PDF:")

project_choice = st.radio(
    "اختر المشروع:",
    ("مشروع 1", "مشروع 2")
)

uploaded_file = None
if project_choice == "مشروع 1":
    uploaded_file = st.file_uploader("اختر ملف PDF للمشروع 1", type=["pdf"])
elif project_choice == "مشروع 2":
    uploaded_file = st.file_uploader("اختر ملف PDF للمشروع 2", type=["pdf"])

if uploaded_file:
    if project_choice == "مشروع 1":
        data = extract_project1(uploaded_file)
    elif project_choice == "مشروع 2":
        data = extract_project2(uploaded_file)

    if data:
        df = pd.DataFrame(data)
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.date
        df = df.sort_values(by='Date').reset_index(drop=True)
        df['Floor'] = df['Floor'].str.strip().str.title()
        df['Floor'] = df['Floor'].map(floor_symbol_map).fillna(df['Floor'])
        
        excel_filename = "extracted_workorders_colored.xlsx"
        df.to_excel(excel_filename, index=False)
        apply_colors_to_excel(excel_filename)

        with open(excel_filename, "rb") as f:
            st.download_button(
                label="⬇️ تحميل ملف Excel",
                data=f,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        st.success("✅ تم استخراج البيانات بنجاح!")
    else:
        st.warning("❌ لم يتم العثور على بيانات صالحة في الملف.")

