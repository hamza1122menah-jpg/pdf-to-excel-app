import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from datetime import datetime
import io
import random
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# -----------------------
# خريطة الأدوار → رموز
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

# -----------------------
st.title("📄 PDF to Excel Converter")

uploaded_files = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)

if uploaded_files:
    all_rows = []

    for uploaded_file in uploaded_files:
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        for page in doc:
            text = page.get_text("text")
            if not text:
                continue

            # استخراج البيانات
            wo_match = re.search(r'WORKORDER\s*#\s*:\s*(\d+)', text)
            jp_match = re.search(r'JP Code\s*:\s*(\S+)', text)
            qty_match = re.search(r'Asset QTY\s*:\s*(\d+)', text)
            date_match = re.search(r'Scheduel Start\s*:\s*([A-Za-z]+\s+\d{1,2},\s+\d{4})', text)
            loc_match = re.search(r'Location Code\s*:\s*(\S+)', text)
            phase_match = re.search(r'Phase\s*#\s*(\d+)', text)
            column_match = re.search(r'Column\s*([A-Z0-9]+)', text)
            axis_match = re.search(r'Axis\s*([A-Z0-9]+)', text)

            wo_num = wo_match.group(1) if wo_match else ""
            type_check = jp_match.group(1).strip()[-1] if jp_match else ""
            qty = qty_match.group(1) if qty_match else ""
            date_str = date_match.group(1) if date_match else ""
            phase = phase_match.group(1) if phase_match else ""
            column = column_match.group(1) if column_match else ""
            axis = axis_match.group(1) if axis_match else ""

            if date_str:
                try:
                    date_obj = datetime.strptime(date_str, "%b %d, %Y")
                    formatted_date = date_obj.strftime("%d-%b-%Y")
                except:
                    formatted_date = date_str
            else:
                formatted_date = ""

            # Floor extraction and mapping
            floor = ""
            if loc_match:
                loc_code = loc_match.group(1)
                floor_code = loc_code[10:12].upper() if len(loc_code) > 10 else ""
                floor = floor_symbol_map.get(floor_code, floor_code)

            # Append row if data exists
            if all([wo_num, qty]):
                all_rows.append({
                    "workorder num": wo_num,
                    "Floor": floor,
                    "phase": phase,
                    "Column": column,
                    "Axis": axis,
                    "Quantity": qty,
                    "Equipment": "FHC",
                    "Type of check": type_check,
                    "Date": formatted_date
                })

    if all_rows:
        df = pd.DataFrame(all_rows)
        # ترتيب حسب التاريخ
        df['Date'] = pd.to_datetime(df['Date'], format='%d-%b-%Y', errors='coerce')
        df = df.sort_values('Date')
        df['Date'] = df['Date'].dt.strftime('%d-%b-%Y')

        # -----------------------
        # إنشاء ملف Excel في الذاكرة
        output = io.BytesIO()
        df.to_excel(output, index=False)
        excel_bytes = output.getvalue()

        # -----------------------
        # تلوين الصفوف بألوان فاتحة مختلفة لكل تاريخ
        wb = Workbook()
        ws = wb.active
        ws.append(list(df.columns))

        # ألوان فاتحة
        light_colors = ["FFFFCC", "CCFFCC", "CCE5FF", "FFCCCC", "FFE5CC", "E5CCFF", "FFFF99"]
        date_colors = {}

        for idx, row in df.iterrows():
            row_values = list(row)
            ws.append(row_values)
            date_val = row['Date']
            if date_val not in date_colors:
                date_colors[date_val] = random.choice(light_colors)
            fill = PatternFill(start_color=date_colors[date_val], end_color=date_colors[date_val], fill_type="solid")
            for col_idx in range(1, len(row_values)+1):
                ws.cell(row=ws.max_row, column=col_idx).fill = fill

        # حفظ الملف في الذاكرة
        output_excel = io.BytesIO()
        wb.save(output_excel)
        excel_data = output_excel.getvalue()

        # -----------------------
        # زر تحميل Excel
        st.download_button(
            label="⬇️ Download Excel",
            data=excel_data,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("✅ Excel file is ready to download!")

    else:
        st.warning("❌ No valid data found in the uploaded PDFs.")
