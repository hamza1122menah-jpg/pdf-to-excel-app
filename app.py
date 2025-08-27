import fitz  # PyMuPDF
import re
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment
import os

# خريطة الأدوار والرموز
floor_symbol_map = {
    'Basement': 'B',
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
}

# إنشاء نمط Regex لجميع الأدوار
floor_pattern = re.compile("|".join(re.escape(floor) for floor in floor_symbol_map.keys()), re.IGNORECASE)

def extract_pdf_data(pdf_path):
    doc = fitz.open(pdf_path)
    data = []

    for page in doc:
        text = page.get_text("text")
        if not text:
            continue

        # البحث عن Workorder
        wo_match = re.search(r'WORKORDER\s*#\s*:\s*(\d+)', text)
        wo_num = wo_match.group(1) if wo_match else ""

        # البحث عن JP Code لاستخراج Type of check
        jp_match = re.search(r'JP Code\s*:\s*(\S+)', text)
        check_type = jp_match.group(1).strip()[-1] if jp_match else ""

        # البحث عن Quantity
        qty_match = re.search(r'Asset QTY\s*:\s*(\d+)', text)
        qty = qty_match.group(1) if qty_match else ""

        # البحث عن Date
        date_match = re.search(r'Scheduel Start\s*:\s*([\w]+\s+\d{1,2},\s+\d{4})', text)
        if date_match:
            try:
                date_obj = datetime.strptime(date_match.group(1), '%b %d, %Y')
                formatted_date = date_obj.strftime('%d-%b-%Y')
            except:
                formatted_date = date_match.group(1)
        else:
            formatted_date = ""

        # البحث عن Phase, Column, Axis
        phase_match = re.search(r'Phase\s*#\s*(\d+)', text)
        column_match = re.search(r'Column\s*([A-Z0-9]+)', text)
        axis_match = re.search(r'Axis\s*([A-Z0-9]+)', text)

        # البحث عن Floor باستخدام Regex قوي
        floor_match = floor_pattern.search(text)
        floor_code = floor_symbol_map[floor_match.group()] if floor_match else ""

        if any([wo_num, floor_code, qty, check_type]):
            data.append({
                "Workorder num": wo_num,
                "Floor": floor_code,
                "Phase": phase_match.group(1) if phase_match else "",
                "Column": column_match.group(1) if column_match else "",
                "Axis": axis_match.group(1) if axis_match else "",
                "Quantity": qty,
                "Equipment": "FHC",
                "Type of check": check_type,
                "Date": formatted_date
            })

    return data

def save_to_excel(data, output_file="output.xlsx"):
    df = pd.DataFrame(data)
    # ترتيب حسب التاريخ
    df['Date'] = pd.to_datetime(df['Date'], format='%d-%b-%Y', errors='coerce')
    df = df.sort_values('Date')
    df['Date'] = df['Date'].dt.strftime('%d-%b-%Y')

    df.to_excel(output_file, index=False)

    # تنسيق الخلايا في Excel
    from openpyxl import load_workbook
    wb = load_workbook(output_file)
    ws = wb.active
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    wb.save(output_file)
    print(f"✅ تم حفظ الملف: {output_file}")

# ========== التشغيل ==========
pdf_file = input("ضع مسار ملف PDF هنا: ")
data = extract_pdf_data(pdf_file)
if data:
    save_to_excel(data)
else:
    print("❌ لم يتم العثور على بيانات صالحة في الملف.")
