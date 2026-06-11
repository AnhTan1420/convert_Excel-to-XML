import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import time
import io
import zipfile

# App Workspace Configuration
st.set_page_config(page_title="MOE Jordan", page_icon="⚙️", layout="centered")

st.title("MOE convert XLSX to XML then compress to ZIP")

# Initialize Session States to prevent download button data-loss on click
if "download_queue" not in st.session_state:
    st.session_state.download_queue = []
if "run_summary" not in st.session_state:
    st.session_state.run_summary = {"success_count": 0, "total_records": 0}

# --- CHỨC NĂNG TẠO TEMPLATE EXCEL TRONG BỘ NHỚ ---
def generate_template(headers):
    buffer = io.BytesIO()
    # Tạo DataFrame trống chỉ có các cột tiêu đề
    df_template = pd.DataFrame(columns=headers)
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False)
    return buffer.getvalue()

# Định nghĩa danh sách cột chính xác từ file mẫu của bạn
TEMPLATE_COLUMNS = {
    "MAPPING": [
        "UNIQUE_ID", "STUDENT_UIN_FIN_NO", "PARENT_UIN_FIN_NO", 
        "STUDENT_UINFIN_TYPE_ICODE", "PREV_NRIC_UIN_FIN_NO"
    ],
    "PERSONAL": [
        "RECORD_ID", "UNIQUE_ID", "STUDENT_NAME", "HANYU_PINYIN_NAME", "BIRTH_DATE", 
        "CITIZENSHIP_CODE", "CITIZENSHIP_SGDRM_CODE", "RACE_CODE", "RELIGION_CODE", 
        "RELIGION_SGDRM_CODE", "SEX_CODE", "EMAIL_ADDRESS", "CITIZENSHIP_EFFECTIVE_DATE", 
        "CONTACT_SAMEAS_OFFICIAL_IND", "CONTACTADD_BLK_HSE_NO", "CONTACTADD_STREET_NAME", 
        "CONTACTADD_FLOOR_NO", "CONTACTADD_UNIT_NO", "CONTACTADD_BLDG_NAME", 
        "CONTACTADD_POSTAL_ECODE", "TELEPHONE_NO", "HANDPHONE_NO", "OTHER_CONTACT_NO", 
        "RES_TYPE_CODE", "OFFICIALADD_BLK_HSE_NO", "OFFICIALADD_STREET_NAME", 
        "OFFICIALADD_FLOOR_NO", "OFFICIALADD_UNIT_NO", "OFFICIALADD_BLDG_NAME", 
        "OFFICIALADD_POSTAL_ECODE", "FOREIGNADD_LINE1_DESC", "FOREIGNADD_LINE2_DESC", 
        "FOREIGNADD_POSTAL_ECODE", "FOREIGNADD_COUNTRY_CODE", "FOREIGNADD_CONTACTCODE_NO_OLD", 
        "FOREIGNADD_CONTACT_NO_OLD", "FOREIGNADD_CONTACTCODE_NO", "FOREIGNADD_CONTACT_NO", 
        "FOREIGNADD_COUNTRY_SGDRM_CODE", "ADDRESS_IND", "CONTACTADD_STREET_CODE", 
        "OFFICIALADD_STREET_CODE", "GUARDIAN_TYPE_ICODE", "PASS_TYPE_CODE", 
        "PASS_ISSUE_DATE", "PASS_EXPIRY_DATE", "RACE_REQUEST_DATE", "PR_TYPE"
    ],
    "SCHOOL": [
        "UNIQUE_ID", "SCHOOL_CODE", "SCHOOL_NAME", "ADMISSION_DATE", 
        "STREAM_CODE", "CLASS_CODE", "STATUS_CODE"
    ] # Thêm các cột School thực tế của dự án bạn vào đây nếu cần chỉnh sửa
}

# Hộp tiện ích tải File Mẫu (Templates)
with st.expander("📥 Download Excel Sample Templates for QA Testing", expanded=True):
    st.markdown("Chọn loại template cần thiết bên dưới để tải về file Excel có sẵn cấu trúc cột chuẩn:")
    col_t1, col_t2, col_t3 = st.columns(3)
    
    with col_t1:
        mapping_data = generate_template(TEMPLATE_COLUMNS["MAPPING"])
        st.download_button(
            label="📁 ID_MAPPING Template",
            data=mapping_data,
            file_name="Template_ID_MAPPING.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_t2:
        personal_data = generate_template(TEMPLATE_COLUMNS["PERSONAL"])
        st.download_button(
            label="📁 BASIC_PERSONAL Template",
            data=personal_data,
            file_name="Template_BASIC_PERSONAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_t3:
        school_data = generate_template(TEMPLATE_COLUMNS["SCHOOL"])
        st.download_button(
            label="📁 BASIC_SCHOOL Template",
            data=school_data,
            file_name="Template_BASIC_SCHOOL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# Collapsible Documentation Matrix
with st.expander("📘 System Guidelines & Target Interface Identifiers", expanded=False):
    st.markdown("""
    The processing engine matches files based on case-insensitive keyword tokens within the filename. 
    Each valid file type will generate its own dedicated, isolated ZIP archive.
    
    | Target System Interface | Accepted File Keyword | Generated Payload Prefix | XML Row Structure |
    | :--- | :--- | :--- | :--- |
    | **ID Mapping Profile** | `Mapping` | `FULL_SFS_ID_MAPPING_MK_*` | `<ID_Mapping UNIQUE_ID="...">` |
    | **Basic Personal** | `Personal` | `FULL_SFS_STUDENT_BASIC_PERSONAL_MK_*` | `<STUDENT_BASIC_PERSONAL>` + Attr |
    | **Basic School** | `School` | `FULL_SFF_STUDENT_BASIC_SCHOOL_MK_*` | `<STUDENT_BASIC_SCHOOL>` + Attr |
    """)

st.markdown("### 📤 Source File Upload")
uploaded_files = st.file_uploader(
    "Drag and drop or browse Excel workbooks (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

# Rule Mapping Engine Configuration
MAPPING_RULES = {
    "MAPPING": {"prefix": "FULL_SFS_ID_MAPPING_MK", "interface": "STUDENT_ID_Mapping_INFO"},
    "PERSONAL": {"prefix": "FULL_SFS_STUDENT_BASIC_PERSONAL_MK", "interface": "STUDENT_Personal_INFO"},
    "PERSONLA": {"prefix": "FULL_SFS_STUDENT_BASIC_PERSONAL_MK", "interface": "STUDENT_Personal_INFO"}, # Legacy typo fallback
    "SCHOOL": {"prefix": "FULL_SFF_STUDENT_BASIC_SCHOOL_MK", "interface": "STUDENT_School_INFO"}
}

if uploaded_files:
    st.info(f"📋 **Stage Queue:** {len(uploaded_files)} file(s) loaded into session memory.")
    
    if st.button("🚀 Execute Split Transformation", type="primary", use_container_width=True):
        # Reset local cache queue for the new execution run
        st.session_state.download_queue = []
        success_count = 0
        total_records_processed = 0
        current_time = time.strftime('%Y%m%d%H%M%S')
        
        for uploaded_file in uploaded_files:
            file_name = uploaded_file.name
            matched_rule = None
            
            # Match token against criteria matrix
            for key, rule in MAPPING_RULES.items():
                if key.upper() in file_name.upper():
                    matched_rule = rule
                    break
            
            if matched_rule:
                prefix = matched_rule["prefix"]
                interface = matched_rule["interface"]
                xml_filename = f"{prefix}_{current_time}.xml"
                zip_filename = f"{prefix}_{current_time}.zip"
                
                # Identify specific interface rules based on prefix
                is_personal = (prefix == "FULL_SFS_STUDENT_BASIC_PERSONAL_MK")
                is_school = (prefix == "FULL_SFF_STUDENT_BASIC_SCHOOL_MK")
                
                try:
                    # Đọc file Excel vào DataFrame
                    df = pd.read_excel(uploaded_file, engine='openpyxl')
                    
                    # Xử lý các ô trống (NaN) thành chuỗi rỗng "" trước, rồi mới ép kiểu string
                    df = df.fillna("")
                    df = df.astype(str)
                    
                    # Loại bỏ luôn dấu chấm ".0" phát sinh ở các cột số do Pandas tự dịch sang kiểu float
                    for col in df.columns:
                        df[col] = df[col].apply(
                            lambda x: x.split(".")[0] if x.endswith(".0") else x
                        )
                    
                    # Initialize XML Element Tree Structure
                    NS = 'http://www.w3.org/2001/XMLSchema-instance'
                    ET.register_namespace('xs', NS)
                    
                    root = ET.Element('INTERFACE', {
                        'INTERFACE_NAME': interface,
                        'FILE_CREATED_TIME': str(int(time.time() * 1000)),
                        'FILE_NAME': xml_filename, 
                        'NO_RECORD': str(len(df)) 
                    })
                    
                    # Process rows dynamically based on the file type rule
                    for index, row in df.iterrows():
                        if 'UNIQUE_ID' not in df.columns or row['UNIQUE_ID'] == "":
                            continue
                        
                        unique_id_val = str(row['UNIQUE_ID'])
                        
                        # Apply conditional branching for specialized XML tags
                        if is_personal:
                            row_element = ET.SubElement(root, 'STUDENT_BASIC_PERSONAL')
                            uid_child = ET.SubElement(row_element, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'})
                            uid_child.text = unique_id_val
                        elif is_school:
                            row_element = ET.SubElement(root, 'STUDENT_BASIC_SCHOOL')
                            uid_child = ET.SubElement(row_element, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'})
                            uid_child.text = unique_id_val
                        else:
                            row_element = ET.SubElement(root, 'ID_Mapping', {'UNIQUE_ID': unique_id_val})
                            
                        # Build remaining column nodes
                        for col_name in df.columns:
                            if col_name == 'UNIQUE_ID':
                                continue
                                
                            child = ET.SubElement(row_element, col_name)
                            val = row[col_name]
                            
                            # Kiểm tra nếu giá trị là chuỗi rỗng thì gán thuộc tính nil="true"
                            if val == "":
                                child.set(f"{{{NS}}}nil", "true")
                            else:
                                child.text = str(val)
                                
                    tree = ET.ElementTree(root)
                    try:
                        ET.indent(tree, space="  ", level=0)
                    except AttributeError:
                        pass
                    
                    # Direct render payload to streaming buffer
                    xml_buffer = io.BytesIO()
                    tree.write(xml_buffer, encoding="UTF-8", xml_declaration=True)
                    
                    # Create an isolated zip archive package for this specific target prefix
                    individual_zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(individual_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        zipf.writestr(xml_filename, xml_buffer.getvalue())
                    
                    # Store generated zip metadata into session state queue
                    st.session_state.download_queue.append({
                        "zip_name": zip_filename,
                        "zip_data": individual_zip_buffer.getvalue(),
                        "source_name": file_name,
                        "records": len(df)
                    })
                    
                    success_count += 1
                    total_records_processed += len(df)
                    
                except Exception as e:
                    st.error(f"❌ **Pipeline Failure** on `{file_name}`: {e}")
            else:
                st.warning(f"⏭️ **Ignored:** `{file_name}` does not match any known target system criteria.")
        
        # Cache summary results for display rendering
        st.session_state.run_summary = {
            "success_count": success_count,
            "total_records": total_records_processed
        }

    # Persistent UI Rendering Section
    if st.session_state.download_queue:
        st.markdown("---")
        st.markdown("### Execution Run Summary")
        col1, col2 = st.columns(2)
        with col1:
            st.metric(label="Generated Payloads", value=f"{st.session_state.run_summary['success_count']} File(s)")
        with col2:
            st.metric(label="Total Dataset Records", value=f"{st.session_state.run_summary['total_records']} Rows")
        
        st.write("The processing has generated zip files:")
        
        for item in st.session_state.download_queue:
            with st.container(border=True):
                st.markdown(f"🔹 **Source Workbook:** `{item['source_name']}` | **Payload Size:** {item['records']} elements")
                st.download_button(
                    label=f"📦 Download {item['zip_name']}",
                    data=item['zip_data'],
                    file_name=item['zip_name'],
                    mime="application/zip",
                    key=f"btn_{item['zip_name']}", 
                    use_container_width=True
                )
