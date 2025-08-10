import streamlit as st
from docx import Document
from collections import Counter
from io import BytesIO

def extract_font_info(docx_file):
    document = Document(BytesIO(docx_file))
    font_names = []
    font_sizes = []

    for para in document.paragraphs:
        for run in para.runs:
            font = run.font
            if font.name:
                font_names.append(font.name)
            else:
                font_names.append('Default')
            
            if font.size:
                font_sizes.append(font.size.pt)  # size.pt converts to points
            else:
                font_sizes.append('Default')

    return font_names, font_sizes

def get_percentage(counter):
    total = sum(counter.values())
    return {k: (v / total) * 100 for k, v in counter.items()}

st.title("DOCX Font Analysis App")

uploaded_file = st.file_uploader("Upload a DOCX file", type=["docx"])

if uploaded_file:
    font_names, font_sizes = extract_font_info(uploaded_file.read())
    
    font_name_counts = Counter(font_names)
    font_size_counts = Counter(font_sizes)
    
    font_name_percent = get_percentage(font_name_counts)
    font_size_percent = get_percentage(font_size_counts)
    
    st.subheader("Font Name Distribution (%)")
    st.write(font_name_percent)
    
    st.subheader("Font Size Distribution (%)")
    st.write(font_size_percent)
