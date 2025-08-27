import streamlit as st
from project1 import run_project1
from project2 import run_project2

st.set_page_config(page_title="PDF to Excel App", layout="wide")

# اختيار المشروع من الشريط الجانبي
choice = st.sidebar.radio("اختر المشروع:", ["المطاف", "الشامية - الخدمات - الساحات - الأنفاق"])

# عرض اسم المشروع في أعلى الصفحة
st.title("📊 نظام استخراج البيانات من PDF إلى Excel")
st.subheader(f"المشروع الحالي: {choice}")

# تشغيل المشروع المختار
if choice == "المطاف":
    run_project1()
elif choice == "الشامية - الخدمات - الساحات - الأنفاق":
    run_project2()
