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
from openpyxl import load_workbook
import tempfile

def load_excel_file(uploaded_file):
    try:
        # Lưu file tạm
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        wb = load_workbook(tmp_path, data_only=True)
        ws = wb.active

        headers = [cell.value for cell in ws[1]]

        data = []
        for row in ws.iter_rows(min_row=2):
            row_data = {}
            for header, cell in zip(headers, row):
                if cell.value is None:
                    row_data[header] = ""
                else:
                    # 👉 LẤY GIÁ TRỊ HIỂN THỊ (KHÔNG PHẢI RAW)
                    if cell.is_date:
                        row_data[header] = cell.value.strftime("%d/%m/%Y")
                    else:
                        row_data[header] = str(cell.value)
            data.append(row_data)

        return pd.DataFrame(data)

    except Exception as e:
        st.error(f"Lỗi khi đọc file Excel: {str(e)}")
        return None

def replace_placeholders_in_paragraph(paragraph, data_dict):
    """
    Thay thế placeholder trong paragraph và giữ nguyên định dạng chữ
    Xử lý trường hợp placeholder bị tách thành nhiều runs
    """
    # Ghép toàn bộ text của paragraph
    full_text = ''.join(run.text for run in paragraph.runs)
    
    # Kiểm tra xem có placeholder nào không
    has_changes = False
    for key in data_dict.keys():
        if f"{{{{{key}}}}}" in full_text:
            has_changes = True
            break
    
    if not has_changes:
        return
    
    # Tạo map vị trí của từng run
    run_positions = []
    pos = 0
    for run in paragraph.runs:
        run_length = len(run.text)
        run_positions.append({
            'run': run,
            'start': pos,
            'end': pos + run_length,
            'text': run.text,
            'font_name': run.font.name,
            'font_size': run.font.size,
            'bold': run.font.bold,
            'italic': run.font.italic,
            'underline': run.font.underline,
            'color': run.font.color.rgb if run.font.color.rgb else None,
            'highlight': run.font.highlight_color
        })
        pos += run_length
    
    # Thay thế placeholder trong text
    new_text = full_text
    replacements = []
    for key, value in data_dict.items():
        placeholder = f"{{{{{key}}}}}"
        if placeholder in new_text:
            # Tìm vị trí của placeholder
            start_idx = new_text.find(placeholder)
            end_idx = start_idx + len(placeholder)
            
            # Tìm run chứa điểm bắt đầu của placeholder
            format_to_use = None
            for rp in run_positions:
                if rp['start'] <= start_idx < rp['end']:
                    format_to_use = rp
                    break
            
            # Thay thế
            new_text = new_text.replace(placeholder, str(value), 1)
            
            # Lưu thông tin replacement
            replacements.append({
                'old_start': start_idx,
                'old_end': end_idx,
                'new_length': len(str(value)),
                'format': format_to_use
            })
    
    # Xóa tất cả runs cũ
    for run in paragraph.runs:
        run.text = ''
    
    # Tạo run mới với text đã thay thế
    if paragraph.runs:
        new_run = paragraph.runs[0]
    else:
        new_run = paragraph.add_run()
    
    new_run.text = new_text
    
    # Áp dụng định dạng từ run gốc chứa placeholder
    if replacements and replacements[0]['format']:
        fmt = replacements[0]['format']
        if fmt['font_name']:
            new_run.font.name = fmt['font_name']
        if fmt['font_size']:
            new_run.font.size = fmt['font_size']
        if fmt['bold'] is not None:
            new_run.font.bold = fmt['bold']
        if fmt['italic'] is not None:
            new_run.font.italic = fmt['italic']
        if fmt['underline'] is not None:
            new_run.font.underline = fmt['underline']
        if fmt['color']:
            new_run.font.color.rgb = fmt['color']
        if fmt['highlight']:
            new_run.font.highlight_color = fmt['highlight']

def replace_placeholders_in_table(table, data_dict):
    """
    Thay thế placeholder trong bảng và giữ nguyên định dạng
    Xử lý cả paragraph và cell text
    """
    for row in table.rows:
        for cell in row.cells:
            # Xử lý từng paragraph trong cell
            for paragraph in cell.paragraphs:
                replace_placeholders_in_paragraph(paragraph, data_dict)
            
            # Xử lý trường hợp placeholder nằm trong cell.text
            # (một số template có placeholder trực tiếp trong cell)
            cell_text = cell.text
            has_placeholder = any(f"{{{{{key}}}}}" in cell_text for key in data_dict.keys())
            
            if has_placeholder and len(cell.paragraphs) > 0:
                # Lấy định dạng từ run đầu tiên của paragraph đầu tiên
                first_para = cell.paragraphs[0]
                if first_para.runs:
                    first_run = first_para.runs[0]
                    
                    # Thay thế text
                    new_text = cell_text
                    for key, value in data_dict.items():
                        placeholder = f"{{{{{key}}}}}"
                        new_text = new_text.replace(placeholder, str(value))
                    
                    # Xóa tất cả nội dung cũ trong cell
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.text = ''
                    
                    # Tạo run mới với định dạng gốc
                    new_run = first_para.runs[0] if first_para.runs else first_para.add_run()
                    new_run.text = new_text
                    
                    # Giữ nguyên định dạng
                    if first_run.font.name:
                        new_run.font.name = first_run.font.name
                    if first_run.font.size:
                        new_run.font.size = first_run.font.size
                    if first_run.font.bold is not None:
                        new_run.font.bold = first_run.font.bold
                    if first_run.font.italic is not None:
                        new_run.font.italic = first_run.font.italic
                    if first_run.font.underline is not None:
                        new_run.font.underline = first_run.font.underline
                    if first_run.font.color.rgb:
                        new_run.font.color.rgb = first_run.font.color.rgb
                    if first_run.font.highlight_color:
                        new_run.font.highlight_color = first_run.font.highlight_color

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
            for key in ['name', 'Name', 'ho_ten', 'ten', 'fullName', 'FullName', 'StudentName']:
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






