import streamlit as st
import pandas as pd
from docx import Document
from docxcompose.composer import Composer
import io
import re

def load_excel_file(uploaded_file):
    try:
        return pd.read_excel(uploaded_file)
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
            has_placeholder = any(f"{{{{{key}}}}}" in cell_text for key in data_dict)
            if has_placeholder:
                new_text = cell_text
                for key, value in data_dict.items():
                    new_text = new_text.replace(f"{{{{{key}}}}}", str(value))
                cell.text = ""
                cell.paragraphs[0].add_run(new_text)

def process_word_template(template_bytes, data_dict):
    try:
        doc = Document(io.BytesIO(template_bytes))
        for paragraph in doc.paragraphs:
            replace_placeholders_in_paragraph(paragraph, data_dict)
        for table in doc.tables:
            replace_placeholders_in_table(table, data_dict)
        return doc
    except Exception as e:
        st.error(f"Lỗi xử lý Word template: {str(e)}")
        return None

def find_placeholders_in_doc(template_bytes):
    doc = Document(io.BytesIO(template_bytes))
    placeholders = set()
    for paragraph in doc.paragraphs:
        placeholders.update(re.findall(r'\{\{([^}]+)\}\}', paragraph.text))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    placeholders.update(re.findall(r'\{\{([^}]+)\}\}', paragraph.text))
    return placeholders

def merge_word_documents(documents):
    if not documents:
        return None
    merged_doc = documents[0]
    composer = Composer(merged_doc)
    for doc in documents[1:]:
        composer.append(doc)
    output = io.BytesIO()
    composer.save(output)
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="Excel to Word Template", page_icon="📄", layout="wide")
    st.title("📄 Excel to Word Template Generator")
    st.markdown("---")

    with st.sidebar:
        st.header("📁 Upload Files")
        excel_file = st.file_uploader("Chọn file Excel (.xlsx, .xls)", type=['xlsx', 'xls'])
        word_file = st.file_uploader("Chọn file Word template (.docx)", type=['docx'])

    col1, col2 = st.columns([1, 1])

    if excel_file:
        excel_data = load_excel_file(excel_file)
    else:
        excel_data = None

    with col1:
        st.header("📊 Dữ liệu Excel")
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
                first_row = excel_data.iloc[0]
                for col in selected_columns:
                    value = first_row[col] if pd.notna(first_row[col]) else ""
                    st.write(f"- `{{{{{col}}}}}` → `{value}`")
        else:
            st.info("👆 Vui lòng upload file Excel")

    with col2:
        st.header("📝 Word Template")
        if word_file:
            template_bytes = word_file.getvalue()
            st.success("✅ Đã upload template Word")

            placeholders = find_placeholders_in_doc(template_bytes)

            if placeholders:
                st.subheader("Placeholder được tìm thấy:")
                for placeholder in sorted(placeholders):
                    st.code(f"{{{{{placeholder}}}}}")
            else:
                st.warning("⚠️ Không tìm thấy placeholder nào trong template")
        else:
            template_bytes = None
            st.info("👆 Vui lòng upload file Word template")

    st.markdown("---")
    st.header("🚀 Xử lý và Tạo File Word Gộp")

    if excel_data is not None and template_bytes is not None:
        if selected_columns:
            placeholders = find_placeholders_in_doc(template_bytes)
            missing = placeholders - set(selected_columns)
            if missing:
                st.warning(f"⚠️ Các placeholder không có trong dữ liệu Excel: {', '.join(sorted(missing))}")

            if st.button("📘 Tạo File Word Gộp", type="primary", use_container_width=True):
                with st.spinner("Đang xử lý..."):
                    documents = []
                    for _, row in excel_data.iterrows():
                        data_dict = {col: row[col] if pd.notna(row[col]) else "" for col in selected_columns}
                        doc = process_word_template(template_bytes, data_dict)
                        if doc:
                            documents.append(doc)

                    if documents:
                        merged_output = merge_word_documents(documents)
                        st.success("✅ Đã tạo file Word gộp thành công!")
                        st.download_button(
                            label="📥 Tải xuống file Word gộp",
                            data=merged_output.getvalue(),
                            file_name="word_merged_output.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    else:
                        st.error("Không có tài liệu nào được tạo.")
        else:
            st.warning("⚠️ Vui lòng chọn ít nhất một cột từ dữ liệu Excel")
    else:
        st.info("ℹ️ Vui lòng upload đầy đủ file Excel và Word template để bắt đầu")

    with st.expander("📖 Hướng dẫn sử dụng"):
        st.markdown("""
        ### Cách sử dụng:

        1. **Upload file Excel**: File chứa dữ liệu cần điền vào template
        2. **Upload file Word**: Template với các placeholder
        3. **Chọn cột**: Chọn các cột từ Excel để sử dụng làm `{{placeholder}}`
        4. **Tạo file Word gộp**: Nhấn nút để tạo một file Word duy nhất

        ### Format Placeholder:
        - Sử dụng định dạng `{{tên_cột}}` trong Word
        - Ví dụ: `{{name}}`, `{{age}}`, `{{address}}`

        ### Lưu ý:
        - Mỗi dòng trong Excel sẽ tạo thành một trang trong file Word gộp
        - Giữ nguyên định dạng file Word gốc
        - Kiểm tra placeholder không khớp để cảnh báo
        """)

if __name__ == "__main__":
    main()
