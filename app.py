import streamlit as st
import pandas as pd
from docx import Document
from docx2pdf import convert
from PyPDF2 import PdfMerger
import io
import re
import tempfile
import os
import base64
import zipfile

# ---------------------- FUNCTIONS ----------------------
def load_excel_file(uploaded_file):
    try:
        return pd.read_excel(
            uploaded_file,
            dtype=str,          # ÉP tất cả cột về string
            keep_default_na=False  # Không biến ô trống thành NaN
        )
    except Exception as e:
        st.error(f"Lỗi khi đọc file Excel: {str(e)}")
        return None

def replace_placeholders_in_paragraph(paragraph, data_dict):
    full_text = paragraph.text
    has_placeholder = any(f"{{{{{key}}}}}" in full_text for key in data_dict)
    if has_placeholder:
        new_text = full_text
        for key, value in data_dict.items():
            new_text = new_text.replace(f"{{{{{key}}}}}", str(value))
        for run in paragraph.runs:
            run.clear()
        if paragraph.runs:
            paragraph.runs[0].text = new_text
        else:
            paragraph.add_run(new_text)

def replace_placeholders_in_table(table, data_dict):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_placeholders_in_paragraph(paragraph, data_dict)
            cell_text = cell.text
            if any(f"{{{{{key}}}}}" in cell_text for key in data_dict):
                new_text = cell_text
                for key, value in data_dict.items():
                    new_text = new_text.replace(f"{{{{{key}}}}}", str(value))
                cell.text = ""
                cell.paragraphs[0].add_run(new_text)

def process_word_template(doc_bytes, data_dict):
    try:
        doc_io = io.BytesIO(doc_bytes)
        doc = Document(doc_io)
        for paragraph in doc.paragraphs:
            replace_placeholders_in_paragraph(paragraph, data_dict)
        for table in doc.tables:
            replace_placeholders_in_table(table, data_dict)
        return doc
    except Exception as e:
        st.error(f"Lỗi khi xử lý template Word: {str(e)}")
        return None

def create_output_files(template_bytes, excel_data, selected_columns):
    output_files = []
    pdf_files = []
    temp_paths = []

    tmpdir = tempfile.mkdtemp()

    for index, row in excel_data.iterrows():
        data_dict = {col: row[col] if pd.notna(row[col]) else "" for col in selected_columns}
        doc = process_word_template(template_bytes, data_dict)
        if doc is not None:
            filename = f"output_{index + 1}.docx"
            for key in ['name', 'Name', 'ho_ten', 'ten', 'fullName', 'FullName']:
                if key in data_dict and data_dict[key]:
                    filename = f"{data_dict[key]}.docx"
                    break

            docx_path = os.path.join(tmpdir, filename)
            pdf_path = docx_path.replace(".docx", ".pdf")
            doc.save(docx_path)

            with open(docx_path, "rb") as fdocx:
                output_files.append((filename, fdocx.read()))

            temp_paths.append((docx_path, pdf_path))

    for docx_path, pdf_path in temp_paths:
        try:
            convert(docx_path, pdf_path)
            with open(pdf_path, "rb") as fpdf:
                pdf_files.append((pdf_path, fpdf.read()))
        except Exception as e:
            st.warning(f"⚠️ Không thể convert {os.path.basename(docx_path)} sang PDF: {e}")

    return output_files, pdf_files

def create_zip_file(output_files):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for filename, file_content in output_files:
            zip_file.writestr(filename, file_content)
    zip_buffer.seek(0)
    return zip_buffer.getvalue()

def merge_pdfs(pdf_contents):
    merger = PdfMerger()
    for _, pdf_data in pdf_contents:
        merger.append(io.BytesIO(pdf_data))
    output_buffer = io.BytesIO()
    merger.write(output_buffer)
    merger.close()
    output_buffer.seek(0)
    return output_buffer.getvalue()

# ---------------------- MAIN APP ----------------------
st.set_page_config(page_title="Tạo Word từ Excel", page_icon="📄", layout="wide")

st.title("📄 Tạo Word từ Excel & In hàng loạt")
st.markdown("---")

with st.sidebar:
    st.header("📁 Upload Files")
    excel_file = st.file_uploader("Chọn file Excel (.xlsx, .xls)", type=['xlsx', 'xls'])
    word_file = st.file_uploader("Chọn file Word template (.docx)", type=['docx'])

if excel_file and word_file:
    excel_data = load_excel_file(excel_file)
    template_bytes = word_file.getvalue()
    template_doc = Document(word_file)

    st.subheader("📊 Dữ liệu Excel")
    st.dataframe(excel_data.head(10), use_container_width=True)

    selected_columns = st.multiselect(
        "Chọn cột làm placeholder",
        options=excel_data.columns.tolist(),
        default=excel_data.columns.tolist()
    )

    placeholders = set()
    for paragraph in template_doc.paragraphs:
        placeholders.update(re.findall(r'\{\{([^}]+)\}\}', paragraph.text))
    for table in template_doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    placeholders.update(re.findall(r'\{\{([^}]+)\}\}', paragraph.text))

    st.subheader("🔍 Placeholder được tìm thấy:")
    for placeholder in sorted(placeholders):
        st.code(f"{{{{{placeholder}}}}}")

    if selected_columns:
        if st.button("🎯 Tạo Files", type="primary"):
            with st.spinner("Đang xử lý..."):
                output_files, pdf_files = create_output_files(template_bytes, excel_data, selected_columns)

                if output_files:
                    st.success(f"✅ Đã tạo {len(output_files)} file Word và PDF")

                    zip_content = create_zip_file(output_files)
                    st.download_button(
                        label="📦 Tải tất cả file Word (.zip)",
                        data=zip_content,
                        file_name="word_documents.zip",
                        mime="application/zip"
                    )

                    if pdf_files:
                        merged_pdf = merge_pdfs(pdf_files)
                        st.download_button(
                            label="🖨️ Tải file PDF gộp để in",
                            data=merged_pdf,
                            file_name="merged_output.pdf",
                            mime="application/pdf"
                        )

                        b64 = base64.b64encode(merged_pdf).decode()
                        st.markdown(f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="1000px"></iframe>', unsafe_allow_html=True)
                else:
                    st.warning("❌ Không tạo được file nào")
    else:
        st.warning("⚠️ Vui lòng chọn ít nhất một cột từ Excel")
else:
    st.info("👆 Vui lòng upload cả file Excel và Word để bắt đầu")

