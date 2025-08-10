import streamlit as st
from docx import Document
from collections import Counter
from io import BytesIO

def get_font_name(run, para):
    # Try run font first
    if run.font and run.font.name:
        return run.font.name
    # Fallback to paragraph style font name
    if para.style and para.style.font and para.style.font.name:
        return para.style.font.name
    return "Default"

def get_font_size(run, para):
    if run.font and run.font.size:
        return run.font.size.pt
    if para.style and para.style.font and para.style.font.size:
        size = para.style.font.size
        if size:
            return size.pt
    return "Default"

def extract_font_info(docx_file):
    document = Document(BytesIO(docx_file))
    font_names = []
    font_sizes = []

    for para in document.paragraphs:
        for run in para.runs:
            font_names.append(get_font_name(run, para))
            font_sizes.append(get_font_size(run, para))

    return font_names, font_sizes

def get_percentage(counter):
    total = sum(counter.values())
    return {k: round((v / total) * 100, 2) for k, v in counter.items()}

st.title("Improved DOCX Font Name and Size Analysis")

uploaded_file = st.file_uploader("Upload a DOCX file", type=["docx"])

if uploaded_file is not None:
    font_names, font_sizes = extract_font_info(uploaded_file.read())
    
    font_name_counts = Counter(font_names)
    font_size_counts = Counter(font_sizes)
    
    font_name_percent = get_percentage(font_name_counts)
    font_size_percent = get_percentage(font_size_counts)
    
    st.subheader("Font Name Distribution (%)")
    for font_name, perc in font_name_percent.items():
        st.write(f"{font_name}: {perc}%")
    
    st.subheader("Font Size Distribution (%)")
    for font_size, perc in font_size_percent.items():
        st.write(f"{font_size} pt: {perc}%")
