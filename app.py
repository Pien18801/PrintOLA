import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
import io
import re
from typing import Dict, Any
import zipfile
import tempfile
import os

def load_excel_file(uploaded_file):
    """Load Excel file and return DataFrame"""
    try:
        # Đọc file Excel
        df = pd.read_excel(uploaded_file)
        return df
    except Exception as e:
        st.error(f"Lỗi khi đọc file Excel: {str(e)}")
        return None

def replace_placeholders_in_paragraph(paragraph, data_dict):
    """Thay thế các placeholder trong paragraph mà giữ nguyên định dạng"""
    full_text = paragraph.text
    has_placeholder = any(f"{{{{{key}}}}}" in full_text for key in data_dict)
    
    if has_placeholder:
        new_text = full_text
        for key, value in data_dict.items():
            placeholder = f"{{{{{key}}}}}"
            new_text = new_text.replace(placeholder, str(value))

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
            has_placeholder = any(f"{{{{{key}}}}}" in cell_text for key in data_dict)

            if has_placeholder:
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

    for index, row in excel_data.iterrows():
        data_dict = {col: row[col] if pd.notna(row[col]) else "" for col in selected_columns}
        doc = process_word_template(template_bytes, data_dict)

        if doc is not None:
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)

            filename = f"output_{index + 1}.docx"
            for key in ['name', 'Name', 'ho_ten', 'ten', 'FullName', 'fullname']:
                if key in data_dict and data_dict[key]:
                    filename = f"{data_dict[key]}.docx"
                    break

            output_files.append((filename, doc_io.getvalue()))

    return output_files

def create_zip_file(output_files):
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for filename, file_content in output_files:
            zip_file.writestr(filename, file_content)

    zip_buffer.seek(0)
    return zip_buffer.getvalue()

def main():
    st.set_page_config(
        page_title="Excel to Word Template",
        page_icon="📄",
        layout="wide"
    )

    st.title("📄 Excel to Word Template Generator")
    st.markdown("---")

    with st.sidebar:
        st.header("📁 Upload Files")

        excel_file = st.file_uploader("Chọn file Excel (.xlsx, .xls)", type=['xlsx', 'xls'])
        word_file = st.file_uploader("Chọn file Word template (.docx)", type=['docx'])

    col1, col2 = st.columns([1, 1])

    with col1:
        st.header("📊 Dữ liệu Excel")

        if excel_file is not None:
            excel_data = load_excel_file(excel_file)

            if excel_data is not None:
                st.success(f"✅ Đã load {len(excel_data)} dòng dữ liệu")
                st.subheader("Xem trước dữ liệu:")
                st.dataframe(excel_data.head(10), use_container_width=True)

                st.subheader("Chọn cột để sử dụng:")
                selected_columns = st.multiselect(
                    "Các cột được chọn sẽ làm placeholder trong template",
                    options=excel_data.columns.tolist(),
                    default=excel_data.columns.tolist()
                )

                if selected_columns:
                    st.info(f"📋 Placeholder format: {', '.join([f'{{{{{col}}}}}' for col in selected_columns[:3]])}...")
                    st.subheader("🔍 Preview data mapping:")
                    if len(excel_data) > 0:
                        first_row = excel_data.iloc[0]
                        st.write("**Dòng đầu tiên sẽ được map như sau:**")
                        for col in selected_columns:
                            value = first_row[col] if pd.notna(first_row[col]) else ""
                            st.write(f"- `{{{{{col}}}}}` → `{value}`")
        else:
            st.info("👆 Vui lòng upload file Excel")

    with col2:
        st.header("📝 Word Template")

        if word_file is not None:
            st.success("✅ Đã upload template Word")

            template_bytes = word_file.getvalue()
            template_doc = Document(word_file)

            st.subheader("Thông tin template:")
            st.write(f"- Số đoạn văn: {len(template_doc.paragraphs)}")
            st.write(f"- Số bảng: {len(template_doc.tables)}")

            placeholders = set()
            for paragraph in template_doc.paragraphs:
                placeholders.update(re.findall(r'\{\{([^}]+)\}\}', paragraph.text))
            for table in template_doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            placeholders.update(re.findall(r'\{\{([^}]+)\}\}', paragraph.text))

            if placeholders:
                st.subheader("Placeholder được tìm thấy:")
                for placeholder in sorted(placeholders):
                    st.code(f"{{{{{placeholder}}}}}")
            else:
                st.warning("⚠️ Không tìm thấy placeholder nào trong template")
        else:
            st.info("👆 Vui lòng upload file Word template")

    st.markdown("---")
    st.header("🚀 Xử lý và Tạo File")

    if excel_file and word_file and 'excel_data' in locals() and excel_data is not None and 'template_bytes' in locals():
        if selected_columns:
            col1, col2, col3 = st.columns([1, 1, 1])

            with col2:
                if st.button("🎯 Tạo Files", type="primary", use_container_width=True):
                    with st.spinner("Đang xử lý..."):
                        try:
                            output_files = create_output_files(template_bytes, excel_data, selected_columns)

                            if output_files:
                                st.success(f"✅ Đã tạo thành công {len(output_files)} file!")

                                if len(output_files) == 1:
                                    filename, file_content = output_files[0]
                                    st.download_button(
                                        label=f"📥 Tải xuống {filename}",
                                        data=file_content,
                                        file_name=filename,
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    )
                                else:
                                    zip_content = create_zip_file(output_files)
                                    st.download_button(
                                        label=f"📦 Tải xuống tất cả ({len(output_files)} files)",
                                        data=zip_content,
                                        file_name="word_documents.zip",
                                        mime="application/zip"
                                    )

                                    st.subheader("Danh sách file đã tạo:")
                                    for i, (filename, _) in enumerate(output_files, 1):
                                        st.write(f"{i}. {filename}")

                        except Exception as e:
                            st.error(f"❌ Lỗi khi tạo files: {str(e)}")
                            st.error(f"Chi tiết lỗi: {type(e).__name__}")
        else:
            st.warning("⚠️ Vui lòng chọn ít nhất một cột từ dữ liệu Excel")
    else:
        st.info("ℹ️ Vui lòng upload đầy đủ file Excel và Word template để bắt đầu")

    with st.expander("📖 Hướng dẫn sử dụng"):
        st.markdown("""
        ### Cách sử dụng:

        1. **Upload file Excel**: File chứa dữ liệu cần điền vào template
        2. **Upload file Word**: Template với các placeholder
        3. **Chọn cột**: Chọn các cột từ Excel để sử dụng
        4. **Tạo files**: Nhấn nút "Tạo Files" để xử lý

        ### Format Placeholder:
        - Sử dụng format `{{tên_cột}}` trong Word template
        - Ví dụ: `{{name}}`, `{{age}}`, `{{address}}`
        - Placeholder sẽ được thay thế bằng dữ liệu từ Excel

        ### Lưu ý:
        - Mỗi dòng trong Excel sẽ tạo ra một file Word riêng
        - Định dạng của Word template sẽ được giữ nguyên
        - Nếu có nhiều file, chúng sẽ được đóng gói trong file ZIP
        """)

if __name__ == "__main__":
    main()
