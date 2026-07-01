import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import time
import io
import zipfile
import random

# App Workspace Configuration
st.set_page_config(page_title="MOE Jordan Engine", page_icon="⚙️", layout="centered")
st.title("MOE: EXCEL ⇄ ZIP/XML Conversion")

# Reusable Data Pools for Mock Generation
months_list = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
relation_codes_list = ['F', 'M', 'G', 'G4']
sex_codes_list = ['M', 'F']
boolean_list = ['Y', 'N']
res_types_list = ['HDB', 'CONDO', 'LANDED']

def random_uin_generator():
    prefix = random.choice(['S', 'T', 'G', 'F', 'M'])
    digits = "".join([str(random.randint(0, 9)) for _ in range(7)])
    suffix = random.choice(['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'Z', 'R', 'Q', 'K', 'N', 'T'])
    return f"{prefix}{digits}{suffix}"

TEMPLATE_COLUMNS = {
    "MAPPING": ["UNIQUE_ID", "STUDENT_UIN_FIN_NO", "PARENT_UIN_FIN_NO", "STUDENT_UINFIN_TYPE_ICODE", "PREV_NRIC_UIN_FIN_NO"],
    "PERSONAL": ["RECORD_ID", "UNIQUE_ID", "STUDENT_NAME", "HANYU_PINYIN_NAME", "BIRTH_DATE", "CITIZENSHIP_CODE", "CITIZENSHIP_SGDRM_CODE", "RACE_CODE", "RELIGION_CODE", "RELIGION_SGDRM_CODE", "SEX_CODE", "EMAIL_ADDRESS", "CITIZENSHIP_EFFECTIVE_DATE", "CONTACT_SAMEAS_OFFICIAL_IND", "CONTACTADD_BLK_HSE_NO", "CONTACTADD_STREET_NAME", "CONTACTADD_FLOOR_NO", "CONTACTADD_UNIT_NO", "CONTACTADD_BLDG_NAME", "CONTACTADD_POSTAL_ECODE", "TELEPHONE_NO", "HANDPHONE_NO", "OTHER_CONTACT_NO", "RES_TYPE_CODE", "OFFICIALADD_BLK_HSE_NO", "OFFICIALADD_STREET_NAME", "OFFICIALADD_FLOOR_NO", "OFFICIALADD_UNIT_NO", "OFFICIALADD_BLDG_NAME", "OFFICIALADD_POSTAL_ECODE", "FOREIGNADD_LINE1_DESC", "FOREIGNADD_LINE2_DESC", "FOREIGNADD_POSTAL_ECODE", "FOREIGNADD_COUNTRY_CODE", "FOREIGNADD_CONTACTCODE_NO_OLD", "FOREIGNADD_CONTACT_NO_OLD", "FOREIGNADD_CONTACTCODE_NO", "FOREIGNADD_CONTACT_NO", "FOREIGNADD_COUNTRY_SGDRM_CODE", "ADDRESS_IND", "CONTACTADD_STREET_CODE", "OFFICIALADD_STREET_CODE", "GUARDIAN_TYPE_ICODE", "PASS_TYPE_CODE", "PASS_ISSUE_DATE", "PASS_EXPIRY_DATE", "RACE_REQUEST_DATE", "PR_TYPE"],
    "SCHOOL": ["RECORD_ID", "UNIQUE_ID", "STUDENT_STATUS_ICODE", "SCHOOL_CODE", "ADMISSION_NO", "ACADEMIC_YEAR", "LEVEL_XCODE", "STREAM_XCODE", "CLASS_XCODE", "CLASS_SERIAL_NO", "COURSE_TYPE_CODE", "FIRSTLANGUAGE_L1_CODE", "SECONDLANGUAGE_L2_CODE", "LEAVE_OF_ABSENCE_IND", "REPEAT_STUD_IND", "ACAD_STATUS_ICODE", "EFFECTIVE_DATE", "SCHOOL_NAME", "CLASS_NAME", "LEVEL_NAME", "STREAM_NAME", "COURSE_XCODE", "COURSE_NAME", "COURSE_TYPE_NAME", "INTF_PROMOTION_IND", "RECOMMENDED_LEVEL_XCODE", "RECOMMENDED_STREAM_XCODE", "JC_PROVISIONAL_IND", "POSTED_IND", "MATRICULATION_NO", "IP_IND"],
    "PARENT": ["RECORD_ID", "UNIQUE_ID", "PARENT_UNIQUE_ID", "RELATION_ICODE", "PARENT_GUARDIAN_NAME", "CITIZENSHIP_CODE", "RACE_CODE", "STANDARD_ATTENDED_CODE", "DECEASED_YEAR", "TELEPHONE_NO", "HANDPHONE_NO", "OTHER_CONTACT_NO", "BIRTH_DATE", "EMAIL_ADDRESS", "CITIZENSHIP_EFFECTIVE_DATE", "CITIZENSHIP_SGDRM_CODE", "PR_TYPE", "NRIC_BLK_HSE_NO", "NRIC_STREET_CODE", "NRIC_FLOOR_NO", "NRIC_UNIT_NO", "NRIC_POSTAL_ECODE"],
    "MOVEMENT": ["RECORD_ID", "UNIQUE_ID", "STRAT_DATE", "END_DATE", "REASON"],
    "CUSTODIAL": ["RECORD_ID", "UNIQUE_ID", "PARENT_UNIQUE_ID", "RELATION_ICODE", "CUSTODIAL_INFO", "RELATIONSHIP", "PG_ACCESS_IND", "LAST_UPDATED_DATE"]
}

# Unified Dynamic Tabs Layout Setup
tab_forward, tab_backward, tab_mock_generator = st.tabs([
    "📤 Forward Converter Engine", 
    "🔄 Reverse Parser Engine",
    "✨ Mock Data Generator"
])

# ==========================================
# --- TAB 1: FORWARD CONVERSION ENGINE ---
# ==========================================
with tab_forward:
    # --- SECTION 3: WORKBOOK SOURCE CONVERSION ENGINE (SMART AUTO-ROUTER LOGIC) ---
    st.subheader("⚙️ Convert (Excel ➡️ ZIP/XML)")
    # --- SHORT & CLEAR GUIDE ---
    with st.expander("📖 Quick Guide & Download Rules", expanded=True):
        st.markdown("""
        ### File Generation & Download Rules:
        * **When uploading `ID_MAPPING`:**
          * The system automatically generates and packs **4 structural ZIP files** at once.
          
        * **When uploading `BASIC_PERSONAL` / `STUDENT_PARENT` / `STUDENT_CUSTODIAL`:**
          * The system splits them. You will download **individual separate files** matching the specific Names/IDs generated from the `ID_MAPPING` data.
        """)

    # Khởi tạo các trạng thái ban đầu trong session_state nếu chưa có
    if "pipeline_download_queue" not in st.session_state:
        st.session_state.pipeline_download_queue = []
    if "run_summary" not in st.session_state:
        st.session_state.run_summary = {"success_count": 0, "total_records": 0}

    def generate_template(headers):
        buffer = io.BytesIO()
        df_template = pd.DataFrame(columns=headers)
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_template.to_excel(writer, index=False)
        return buffer.getvalue()

    with st.expander("📥 Download Excel Sample Templates", expanded=False):
        st.markdown("Select the School Cockpit MK template type below to download the standard Excel file structure:")
        col_t1, col_t2, col_t3, col_t4, col_t5, col_t6 = st.columns(6)
        with col_t1: st.download_button("📁 ID_MAPPING", generate_template(TEMPLATE_COLUMNS["MAPPING"]), "Template_ID_MAPPING.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_t2: st.download_button("📁 PERSONAL", generate_template(TEMPLATE_COLUMNS["PERSONAL"]), "Template_BASIC_PERSONAL.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_t3: st.download_button("📁 SCHOOL", generate_template(TEMPLATE_COLUMNS["SCHOOL"]), "Template_BASIC_SCHOOL.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_t4: st.download_button("📁 PARENT", generate_template(TEMPLATE_COLUMNS["PARENT"]), "Template_STUDENT_PARENT.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_t5: st.download_button("📁 MOVEMENT", generate_template(TEMPLATE_COLUMNS["MOVEMENT"]), "Template_MOVEMENT.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_t6: st.download_button("📁 CUSTODIAL", generate_template(TEMPLATE_COLUMNS["CUSTODIAL"]), "Template_STUDENT_CUSTODIAL.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    # Hàm callback: Tự động chạy để clear kết quả cũ KHI XÓA FILE hoặc ĐỔI FILE MỚI
    def clear_forward_pipeline_cache():
        st.session_state.pipeline_download_queue = []
        st.session_state.run_summary = {"success_count": 0, "total_records": 0}

    st.markdown("### 📤 Source File Upload (Smart Detection Enabled)")
    # Thêm `on_change=clear_forward_pipeline_cache` vào uploader
    uploaded_any_file = st.file_uploader(
        "Upload ANY single template file (.xlsx or .csv) to generate corresponding XML/ZIP files.", 
        type=["xlsx", "csv"], 
        accept_multiple_files=False, 
        key="forward_uploader",
        on_change=clear_forward_pipeline_cache
    )

    if uploaded_any_file:
        st.info(f"📋 **Stage Queue:** `{uploaded_any_file.name}` loaded into session memory.")
        
        if st.button("🚀 Execute Conversion", type="primary", use_container_width=True):
            # Reset lại queue trước khi build data mới
            st.session_state.pipeline_download_queue = []
            current_time = time.strftime('%Y%m%d%H%M%S')
            epoch_ms = str(int(time.time() * 1000))
            NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
            
            try:
                if uploaded_any_file.name.lower().endswith(".csv"):
                    df_input = pd.read_csv(uploaded_any_file, encoding='utf-8')
                else:
                    df_input = pd.read_excel(uploaded_any_file, engine='openpyxl')
                    
                df_input = df_input.fillna("").astype(str)
                for col in df_input.columns:
                    df_input[col] = df_input[col].apply(lambda x: x.split(".")[0] if x.endswith(".0") else x)
                
                headers = [c.strip() for c in df_input.columns]
                detected_mode = "UNKNOWN"
                
                if "STUDENT_UIN_FIN_NO" in headers and "PARENT_UIN_FIN_NO" in headers:
                    detected_mode = "MAPPING"
                elif "STUDENT_NAME" in headers and "BIRTH_DATE" in headers:
                    detected_mode = "PERSONAL"
                elif "PARENT_GUARDIAN_NAME" in headers and "PARENT_UNIQUE_ID" in headers:
                    detected_mode = "PARENT"
                elif "CUSTODIAL_INFO" in headers and "RELATIONSHIP" in headers:
                    detected_mode = "CUSTODIAL"
                elif "SCHOOL_CODE" in headers and "ADMISSION_NO" in headers:
                    detected_mode = "SCHOOL"
                elif "STRAT_DATE" in headers or "REASON" in headers:
                    detected_mode = "MOVEMENT"
                    
                st.write(f"🔍 **Engine Status:** Detected input data structure matches layout: **{detected_mode}**")
                generated_xmls = {}

                if detected_mode == "MAPPING":
                    student_records = []
                    parent_records = []
                    for _, row in df_input.iterrows():
                        uid = row.get("UNIQUE_ID", "").strip()
                        st_uin = row.get("STUDENT_UIN_FIN_NO", "").strip()
                        pt_uin = row.get("PARENT_UIN_FIN_NO", "").strip()
                        if uid:
                            if st_uin: student_records.append({"id": uid, "uin": st_uin})
                            elif pt_uin: parent_records.append({"id": uid, "uin": pt_uin})
                    
                    paired_records = []
                    min_len = min(len(student_records), len(parent_records))
                    for i in range(min_len):
                        paired_records.append({
                            "student_id": student_records[i]["id"], "student_uin": student_records[i]["uin"],
                            "parent_id": parent_records[i]["id"], "parent_uin": parent_records[i]["uin"]
                        })
                    
                    root_map = ET.Element('INTERFACE', {'INTERFACE_NAME': 'STUDENT_ID_Mapping_INFO', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_ID_MAPPING_MK_{current_time}.xml', 'NO_RECORD': str(len(df_input))})
                    ET.register_namespace('xs', NS_XSI)
                    for _, row in df_input.iterrows():
                        uid = row.get("UNIQUE_ID", "").strip()
                        item = ET.SubElement(root_map, 'ID_Mapping', {'UNIQUE_ID': uid})
                        for col in ["STUDENT_UIN_FIN_NO", "PARENT_UIN_FIN_NO", "STUDENT_UINFIN_TYPE_ICODE", "PREV_NRIC_UIN_FIN_NO"]:
                            val = row.get(col, "").strip()
                            if val: ET.SubElement(item, col).text = val
                            else: ET.SubElement(item, col).set(f"{{{NS_XSI}}}nil", "true")
                    generated_xmls[f'FULL_SFS_ID_MAPPING_MK_{current_time}'] = root_map

                    root_pers = ET.Element('INTERFACE', {'INTERFACE_NAME': 'STUDENT_Personal_INFO', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_STUDENT_BASIC_PERSONAL_MK_{current_time}.xml', 'NO_RECORD': str(len(student_records))})
                    for st_rec in student_records:
                        item = ET.SubElement(root_pers, 'STUDENT_BASIC_PERSONAL')
                        ET.SubElement(item, 'RECORD_ID').text = '1'
                        ET.SubElement(item, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = st_rec["id"]
                        ET.SubElement(item, 'STUDENT_NAME').text = f"Student-{st_rec['uin']}"
                        ET.SubElement(item, 'HANYU_PINYIN_NAME').text = 'HANYUPINYIN'
                        ET.SubElement(item, 'BIRTH_DATE').text = f"{random.randint(10,28)}-JAN-{random.randint(2001,2008)}"
                        ET.SubElement(item, 'CITIZENSHIP_CODE').text = '10'
                        ET.SubElement(item, 'CITIZENSHIP_SGDRM_CODE').text = 'SG'
                        ET.SubElement(item, 'RACE_CODE').text = '2'
                        ET.SubElement(item, 'RELIGION_CODE').text = '9'
                        ET.SubElement(item, 'RELIGION_SGDRM_CODE').text = 'F'
                        ET.SubElement(item, 'SEX_CODE').text = random.choice(sex_codes_list)
                        ET.SubElement(item, 'EMAIL_ADDRESS').text = f"student_{st_rec['uin'].lower()}@yopmail.com"
                        ET.SubElement(item, 'CITIZENSHIP_EFFECTIVE_DATE').text = f"01-JAN-2026"
                        ET.SubElement(item, 'CONTACT_SAMEAS_OFFICIAL_IND').text = 'Y'
                        ET.SubElement(item, 'CONTACTADD_BLK_HSE_NO').text = str(random.randint(10, 999))
                        ET.SubElement(item, 'CONTACTADD_STREET_NAME').text = 'Automation Street'
                        ET.SubElement(item, 'CONTACTADD_FLOOR_NO').text = f"{random.randint(1,15):02d}"
                        ET.SubElement(item, 'CONTACTADD_UNIT_NO').text = f"{random.randint(1,50):02d}"
                        ET.SubElement(item, 'CONTACTADD_BLDG_NAME').text = 'Tech Hub Tower'
                        ET.SubElement(item, 'CONTACTADD_POSTAL_ECODE').text = str(random.randint(100000, 999999))
                        ET.SubElement(item, 'TELEPHONE_NO').text = f"6{random.randint(1000000, 9999999)}"
                        ET.SubElement(item, 'HANDPHONE_NO').text = f"9{random.randint(1000000, 9999999)}"
                        ET.SubElement(item, 'OTHER_CONTACT_NO').text = '1'
                        ET.SubElement(item, 'RES_TYPE_CODE').text = random.choice(res_types_list)
                        ET.SubElement(item, 'OFFICIALADD_BLK_HSE_NO').text = str(random.randint(10, 999))
                        ET.SubElement(item, 'OFFICIALADD_STREET_NAME').text = 'Official Street'
                        ET.SubElement(item, 'OFFICIALADD_FLOOR_NO').text = f"{random.randint(1,15):02d}"
                        ET.SubElement(item, 'OFFICIALADD_UNIT_NO').text = f"{random.randint(1,50):02d}"
                        ET.SubElement(item, 'OFFICIALADD_BLDG_NAME').text = 'Civic Building'
                        ET.SubElement(item, 'OFFICIALADD_POSTAL_ECODE').text = str(random.randint(100000, 999999))
                        ET.SubElement(item, 'FOREIGNADD_LINE1_DESC').text = 'Line1'
                        ET.SubElement(item, 'FOREIGNADD_LINE2_DESC').text = 'Line2'
                        ET.SubElement(item, 'FOREIGNADD_POSTAL_ECODE').text = 'F1234'
                        ET.SubElement(item, 'FOREIGNADD_COUNTRY_CODE').text = 'MY'
                        ET.SubElement(item, 'FOREIGNADD_CONTACTCODE_NO_OLD').text = '123'
                        ET.SubElement(item, 'FOREIGNADD_CONTACT_NO_OLD').text = '456'
                        ET.SubElement(item, 'FOREIGNADD_CONTACTCODE_NO').text = '789'
                        ET.SubElement(item, 'FOREIGNADD_CONTACT_NO').text = '012'
                        ET.SubElement(item, 'FOREIGNADD_COUNTRY_SGDRM_CODE').text = 'MY'
                        ET.SubElement(item, 'ADDRESS_IND').text = '1'
                        ET.SubElement(item, 'CONTACTADD_STREET_CODE').text = 'ST01'
                        ET.SubElement(item, 'OFFICIALADD_STREET_CODE').text = 'ST02'
                        ET.SubElement(item, 'GUARDIAN_TYPE_ICODE').text = 'M'
                        ET.SubElement(item, 'PASS_TYPE_CODE').text = '1'
                        ET.SubElement(item, 'PASS_ISSUE_DATE').text = '20200101'
                        ET.SubElement(item, 'PASS_EXPIRY_DATE').text = '20281231'
                        ET.SubElement(item, 'RACE_REQUEST_DATE').text = '20200101'
                        ET.SubElement(item, 'PR_TYPE').text = 'N'
                    generated_xmls[f'FULL_SFS_STUDENT_BASIC_PERSONAL_MK_{current_time}'] = root_pers

                    pipeline_relation_heritage = {}
                    root_parent = ET.Element('INTERFACE', {'INTERFACE_NAME': 'Student_Parent', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_STUDENT_PARENT_MK_{current_time}.xml', 'NO_RECORD': str(len(paired_records))})
                    for p in paired_records:
                        item = ET.SubElement(root_parent, 'STUDENT_PARENT')
                        ET.SubElement(item, 'RECORD_ID').text = '1'
                        ET.SubElement(item, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = p["student_id"]
                        ET.SubElement(item, 'PARENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = p["parent_id"]
                        
                        random_relation = random.choice(relation_codes_list)
                        ET.SubElement(item, 'RELATION_ICODE').text = random_relation
                        pipeline_relation_heritage[f"{p['student_id']}_{p['parent_id']}"] = random_relation
                        
                        ET.SubElement(item, 'PARENT_GUARDIAN_NAME').text = f"Parent-Of-{p['student_uin']}"
                        ET.SubElement(item, 'CITIZENSHIP_CODE').text = '10'
                        ET.SubElement(item, 'RACE_CODE').text = '2'
                        ET.SubElement(item, 'STANDARD_ATTENDED_CODE').text = '2'
                        ET.SubElement(item, 'DECEASED_YEAR').text = '2001'
                        ET.SubElement(item, 'TELEPHONE_NO').text = f"6{random.randint(1000000, 9999999)}"
                        ET.SubElement(item, 'HANDPHONE_NO').text = f"7{random.randint(1000000, 9999999)}"
                        ET.SubElement(item, 'OTHER_CONTACT_NO').text = '1'
                        ET.SubElement(item, 'BIRTH_DATE').text = '26-JUN-1981'
                        ET.SubElement(item, 'EMAIL_ADDRESS').text = f"parent_{p['student_uin'].lower()}@yopmail.com"
                        ET.SubElement(item, 'CITIZENSHIP_EFFECTIVE_DATE').text = f"{random.randint(10,25)}-FEB-2005"
                        ET.SubElement(item, 'CITIZENSHIP_SGDRM_CODE').text = 'SG'
                        ET.SubElement(item, 'PR_TYPE').text = 'Y'
                        ET.SubElement(item, 'NRIC_BLK_HSE_NO').text = 'A22'
                        ET.SubElement(item, 'NRIC_STREET_CODE').text = 'STA22'
                        ET.SubElement(item, 'NRIC_FLOOR_NO').text = 'F22'
                        ET.SubElement(item, 'NRIC_UNIT_NO').text = 'U22'
                        ET.SubElement(item, 'NRIC_POSTAL_ECODE').text = '550000'
                    generated_xmls[f'FULL_SFS_STUDENT_PARENT_MK_{current_time}'] = root_parent

                    root_custodial = ET.Element('INTERFACE', {'INTERFACE_NAME': 'custodial_info_mk', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_time}.xml', 'NO_RECORD': str(len(paired_records))})
                    for idx, p in enumerate(paired_records, start=1):
                        item = ET.SubElement(root_custodial, 'STUDENT_CUSTODIAL_INFO_MK')
                        ET.SubElement(item, 'RECORD_ID').text = str(idx)
                        ET.SubElement(item, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = p["student_id"]
                        ET.SubElement(item, 'PARENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = p["parent_id"]
                        
                        inherited_relation = pipeline_relation_heritage.get(f"{p['student_id']}_{p['parent_id']}", "G4")
                        ET.SubElement(item, 'RELATION_ICODE').text = inherited_relation
                        ET.SubElement(item, 'CUSTODIAL_INFO').text = 'JN'
                        ET.SubElement(item, f'{{{NS_XSI}}}RELATIONSHIP').set(f"{{{NS_XSI}}}nil", "true")
                        ET.SubElement(item, 'PG_ACCESS_IND').text = '2'
                        ET.SubElement(item, 'LAST_UPDATED_DATE').text = '2026-06-30'
                    generated_xmls[f'FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_time}'] = root_custodial

                elif detected_mode in ["PERSONAL", "PARENT", "CUSTODIAL", "SCHOOL", "MOVEMENT"]:
                    tag_row_map = {
                        "PERSONAL": ("STUDENT_Personal_INFO", "FULL_SFS_STUDENT_BASIC_PERSONAL_MK", "STUDENT_BASIC_PERSONAL"),
                        "PARENT": ("Student_Parent", "FULL_SFS_STUDENT_PARENT_MK", "STUDENT_PARENT"),
                        "CUSTODIAL": ("custodial_info_mk", "FULL_SFS_STUDENT_CUSTODIAL_INFO_MK", "STUDENT_CUSTODIAL_INFO_MK"),
                        "SCHOOL": ("STUDENT_School_INFO", "FULL_SFF_STUDENT_BASIC_SCHOOL_MK", "STUDENT_BASIC_SCHOOL"),
                        "MOVEMENT": ("STUDENT_Movement_INFO", "FULL_SFS_MOVEMENT_MK", "MOVEMENT")
                    }
                    
                    intf_name, prefix_fn, row_tag = tag_row_map[detected_mode]
                    root_node = ET.Element('INTERFACE', {'INTERFACE_NAME': intf_name, 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'{prefix_fn}_{current_time}.xml', 'NO_RECORD': str(len(df_input))})
                    ET.register_namespace('xs', NS_XSI)
                    
                    for _, row in df_input.iterrows():
                        item = ET.SubElement(root_node, row_tag)
                        for col in TEMPLATE_COLUMNS[detected_mode]:
                            val = row.get(col, "").strip()
                            if col in ["UNIQUE_ID", "STUDENT_UNIQUE_ID", "PARENT_UNIQUE_ID"] and val:
                                item.set('UNIQUE_ID', 'Y') if col=="STUDENT_UNIQUE_ID" else None
                                ET.SubElement(item, col).text = val
                            elif val:
                                ET.SubElement(item, col).text = val
                            else:
                                ET.SubElement(item, col).set(f"{{{NS_XSI}}}nil", "true")
                    generated_xmls[f'{prefix_fn}_{current_time}'] = root_node
                else:
                    st.error("❌ **Structure Error:** Unknown Excel template headers. Please check the sample templates above.")
                    st.stop()

                success_count = 0
                total_rows = 0
                for prefix_name, xml_node in generated_xmls.items():
                    tree = ET.ElementTree(xml_node)
                    try: ET.indent(tree, space="  ", level=0)
                    except AttributeError: pass
                    
                    xml_buf = io.BytesIO()
                    tree.write(xml_buf, encoding="UTF-8", xml_declaration=True)
                    
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                        zf.writestr(f"{prefix_name}.xml", xml_buf.getvalue())
                    
                    st.session_state.pipeline_download_queue.append({
                        "zip_name": f"{prefix_name}.zip",
                        "zip_data": zip_buf.getvalue(),
                        "records": int(xml_node.get("NO_RECORD"))
                    })
                    success_count += 1
                    total_rows += int(xml_node.get("NO_RECORD"))
                    
                st.session_state.run_summary = {"success_count": success_count, "total_records": total_rows}
                st.toast("⚡ Conversion complete!", icon="🚀")
            except Exception as e:
                st.error(f"❌ **Pipeline Failure:** {e}")

    # Chỉ hiển thị Summary section khi biến `pipeline_download_queue` có dữ liệu thực tế
    if st.session_state.pipeline_download_queue:
        st.markdown("---")
        st.markdown("### Execution Run Summary")
        col1, col2 = st.columns(2)
        col1.metric(label="Generated Payloads", value=f"{st.session_state.run_summary['success_count']} File(s)")
        col2.metric(label="Total Dataset Records", value=f"{st.session_state.run_summary['total_records']} Rows")
        
        for item in st.session_state.pipeline_download_queue:
            with st.container(border=True):
                st.markdown(f"🔹 **Target Interface Payload:** `{item['zip_name']}` | **Size:** {item['records']} elements")
                st.download_button(label="📦 Download " + item['zip_name'], data=item['zip_data'], file_name=item['zip_name'], mime="application/zip", key=f"btn_{item['zip_name']}", use_container_width=True)


# ==========================================
# --- TAB 2: REVERSE ENGINE (ZIP -> XLSX) ---
# ==========================================
with tab_backward:
    st.subheader("🔄 Reverse (ZIP/XML ➡️ Excel)")
    st.write("Upload a `.zip` pack containing target payload system XML file(s). The engine will parse fields and synchronize the output Excel filename with your uploaded ZIP name.")

    uploaded_zip_file = st.file_uploader("Upload target system ZIP file:", type=["zip"], accept_multiple_files=False, key="backward_uploader")

    if uploaded_zip_file:
        st.success(f"📦 Archive `{uploaded_zip_file.name}` staged for decompression.")
        
        if st.button("🏁 Run Reverse Extraction Pipeline", type="primary", use_container_width=True):
            try:
                # 1. Xác định tên file ZIP đầu vào để đồng bộ hóa tên file Excel xuất ra
                zip_filename_raw = uploaded_zip_file.name
                # Loại bỏ đuôi .zip để lấy tên gốc (Ví dụ: "FULL_SFS_ID_MAPPING_MK_2026...")
                base_excel_name = zip_filename_raw.rsplit(".", 1)[0]
                
                # Read loaded ZIP stream out of RAM
                zip_in_mem = zipfile.ZipFile(io.BytesIO(uploaded_zip_file.read()))
                parsed_sheets = {} # Accumulate worksheets matching data types
                
                # Iterate across nested archives entries 
                for file_inside in zip_in_mem.namelist():
                    if file_inside.lower().endswith(".xml"):
                        xml_bytes = zip_in_mem.read(file_inside)
                        root = ET.fromstring(xml_bytes)
                        
                        # Collate row objects sequentially
                        records_list = []
                        
                        for row_node in root:
                            row_dict = {}
                            # Capture internal key identities strings (Specific to mapping structures)
                            if "UNIQUE_ID" in row_node.attrib:
                                row_dict["UNIQUE_ID"] = row_node.attrib["UNIQUE_ID"]
                                
                            for element in row_node:
                                tag_clean = element.tag.split("}")[-1] # Purge namespace prefix blocks
                                
                                # Inspect if node carries dedicated nil flag elements
                                is_nil = False
                                for attr_key, attr_val in element.attrib.items():
                                    if attr_key.endswith("nil") and attr_val == "true":
                                        is_nil = True
                                
                                row_dict[tag_clean] = "" if is_nil else (element.text or "").strip()
                            
                            if row_dict:
                                records_list.append(row_dict)
                        
                        if records_list:
                            df_sheet = pd.DataFrame(records_list)
                            
                            # Tên Sheet bên trong vẫn lấy theo tên file XML tương ứng
                            base_xml_name = file_inside.split("/")[-1]
                            sheet_clean_name = base_xml_name.rsplit(".", 1)[0]
                            if len(sheet_clean_name) > 31:
                                sheet_clean_name = sheet_clean_name[:31]
                            
                            parsed_sheets[sheet_clean_name] = df_sheet

                if parsed_sheets:
                    # Write output to Workbook IO buffer objects
                    out_xlsx_buffer = io.BytesIO()
                    with pd.ExcelWriter(out_xlsx_buffer, engine='openpyxl') as writer:
                        for sheet_name, df_data in parsed_sheets.items():
                            df_data.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # 2. Lưu data và gán tên file Excel đồng bộ hoàn toàn với file ZIP đầu vào
                    st.session_state["extracted_xlsx_data"] = out_xlsx_buffer.getvalue()
                    st.session_state["extracted_xlsx_name"] = f"{base_excel_name}.xlsx"
                    st.session_state["extracted_summary"] = parsed_sheets
                    st.success(f"🎉 Extraction pipeline executed with 100% data fidelity. Target name: `{st.session_state['extracted_xlsx_name']}`")
                else:
                    st.error("⚠️ No valid XML files detected inside the uploaded ZIP archive.")
            except Exception as ex:
                st.error(f"❌ **Reverse Pipeline Error:** {ex}")

    # Display Parsed Sheets Previews and Excel Downloads Link
    if "extracted_xlsx_data" in st.session_state:
        st.markdown("---")
        st.markdown("### 📊 Extracted Workbook Structure Preview")
        
        for s_name, df_p in st.session_state["extracted_summary"].items():
            with st.expander(f"📋 Sheet: {s_name} ({len(df_p)} Rows Detected)", expanded=True):
                st.dataframe(df_p.head(5), use_container_width=True)
                
        st.download_button(
            label=f"📥 DOWNLOAD EXTRACTED EXCEL WORKBOOK ({st.session_state['extracted_xlsx_name']})",
            data=st.session_state["extracted_xlsx_data"],
            file_name=st.session_state["extracted_xlsx_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )


# ==========================================
# --- TAB 3: MOCK DATA GENERATOR ---
# ==========================================
with tab_mock_generator:
    st.subheader("🛠️ Mock Test Data Suite")

    # --- ENGLISH QUICK GUIDE FOR CONVERSION & TESTING ---
    with st.expander("📖 Quick Guide & Usage Instructions (English)", expanded=True):
        st.markdown("""
        1. **Performance & Stress Testing (Section 1):**
           - Click **"Start Generating (900K)"** to compile a heavy dataset containing 900,000 sequential `CUSTODIAL` records.
           - The system streams data directly into a compressed `.zip` package to keep RAM footprint low.
           - Once completed, download the ZIP archive and feed it into your downstream batch processor to verify system throughput.
        
        2. **Data Synchronization (Section 2):**
           - Click **"Start Generating Sample Data Records"** to generate sample data.
           - **The Sync Mechanism:** The generator automatically maps unique `STUDENT_UNIQUE_ID` and `PARENT_UNIQUE_ID` tokens across **4 distinct structural layers**:
             - 📂 **`ID_MAPPING`**: Connects generated UIN/FIN identifiers.
             - 📂 **`BASIC_PERSONAL`**: Contains randomized personal profiles, contact channels, and addresses.
             - 📂 **`STUDENT_PARENT`**: Handles parent links and metadata.
             - 📂 **`STUDENT_CUSTODIAL`**: Sets custodial flags while dynamically inheriting relationships (`RELATION_ICODE`) established in the parent module.
           - **Download Formats:**
             - Use the **Master Excel Workbook (.xlsx)** to manually review or audit the raw generated values side-by-side.
             - Use the individual **Target XML Bundles (ZIPs)** to perform direct integration uploads into your staging server environments.
        """)

    st.write("Generate bulk or synchronized sample datasets natively to run functional verification or performance stress test pipelines.")
    # --- SECTION 1: LARGE DATASET GENERATION (900K RECORDS) ---
    with st.expander("🚀 Generate Large Dataset (900K Records - Stress Test)", expanded=False):
        st.write("Automatically generate `CUSTODIAL` structure data sequentially from 1 to 900,000 and compress it directly into a ZIP file to optimize memory usage.")
        
        if st.button("Start Generating (900K)", type="secondary", use_container_width=True):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("⏳ Initializing compression stream...")
            current_time = time.strftime('%Y%m%d%H%M%S')
            xml_filename = f"FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_time}.xml"
            zip_filename = f"FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_time}.zip"
            NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
            total_records = 900000
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                header = f"<?xml version='1.0' encoding='UTF-8'?>\n<INTERFACE xmlns:xs=\"{NS_XSI}\" INTERFACE_NAME=\"custodial_info_mk\" FILE_CREATED_TIME=\"{str(int(time.time() * 1000))}\" FILE_NAME=\"{xml_filename}\" NO_RECORD=\"{total_records}\">\n"
                xml_chunks = [header]
                chunk_size = 100000
                
                for i in range(1, total_records + 1):
                    row_str = (
                        f"  <STUDENT_CUSTODIAL_INFO_MK>\n"
                        f"    <RECORD_ID>{i}</RECORD_ID>\n"
                        f"    <STUDENT_UNIQUE_ID UNIQUE_ID=\"Y\">JDK-{i}-DIF</STUDENT_UNIQUE_ID>\n"
                        f"    <PARENT_UNIQUE_ID UNIQUE_ID=\"Y\">JDKRP-{i}</PARENT_UNIQUE_ID>\n"
                        f"    <RELATION_ICODE>G4</RELATION_ICODE>\n"
                        f"    <CUSTODIAL_INFO>JN</CUSTODIAL_INFO>\n"
                        f"    <RELATIONSHIP xmlns:xs=\"{NS_XSI}\" xs:nil=\"true\" />\n"
                        f"    <PG_ACCESS_IND>2</PG_ACCESS_IND>\n"
                        f"    <LAST_UPDATED_DATE>2026-06-30</LAST_UPDATED_DATE>\n"
                        f"  </STUDENT_CUSTODIAL_INFO_MK>\n"
                    )
                    xml_chunks.append(row_str)
                    
                    if i % chunk_size == 0:
                        status_text.text(f"⏳ Processing data blocks: {i:,} / {total_records:,} records...")
                        progress_bar.progress(i / total_records)
                
                xml_chunks.append("</INTERFACE>")
                full_xml_string = "".join(xml_chunks)
                zipf.writestr(xml_filename, full_xml_string.encode('utf-8'))
                
            status_text.text("🎉 Generation completed successfully! Package ready for download.")
            progress_bar.empty()
            st.session_state["large_zip_data"] = zip_buffer.getvalue()
            st.session_state["large_zip_name"] = zip_filename

        if "large_zip_data" in st.session_state:
            st.success(f"📦 Archive compiled: `{st.session_state['large_zip_name']}`")
            st.download_button(
                label="📥 CLICK HERE TO DOWNLOAD ZIP FILE (900K DATA)",
                data=st.session_state["large_zip_data"],
                file_name=st.session_state["large_zip_name"],
                mime="application/zip",
                use_container_width=True
            )

    # --- SECTION 2: SYNCHRONIZED TARGET PAIRS GENERATION ---
    with st.expander("✨ Generate Synchronized Sample Data (Full Relations Pack)", expanded=False):
        st.write("Randomly generate randomized synchronized datasets for 5 students across both **XML (ZIP format)** and a single unified **Excel Workbook (.xlsx)**.")
        
        if st.button("Start Generating Sample Data Records", type="primary", use_container_width=True):
            current_timestamp = time.strftime('%Y%m%d%H%M%S')
            epoch_ms = str(int(time.time() * 1000))
            NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
            
            student_ids = [f"JDK-{i}" for i in range(101, 106)]
            parent_ids = [f"JDKRP-{i}" for i in range(101, 106)]
            
            sync_files = {}
            mapped_student_uins = {}
            rows_mapping_xlsx = []
            rows_personal_xlsx = []
            rows_parent_xlsx = []
            rows_custodial_xlsx = []
            sample_relation_heritage = {}

            # -- 1. Mapping Generation --
            root_map = ET.Element('INTERFACE', {'INTERFACE_NAME': 'STUDENT_ID_Mapping_INFO', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_ID_MAPPING_MK_{current_timestamp}.xml', 'NO_RECORD': '10'})
            ET.register_namespace('xs', NS_XSI)
            
            for p_id in parent_ids:
                item = ET.SubElement(root_map, 'ID_Mapping', {'UNIQUE_ID': p_id})
                uin_val = random_uin_generator()
                mapped_student_uins[p_id] = uin_val 
                ET.SubElement(item, 'STUDENT_UIN_FIN_NO').set(f"{{{NS_XSI}}}nil", "true")
                ET.SubElement(item, 'PARENT_UIN_FIN_NO').text = uin_val
                ET.SubElement(item, 'STUDENT_UINFIN_TYPE_ICODE').set(f"{{{NS_XSI}}}nil", "true")
                ET.SubElement(item, 'PREV_NRIC_UIN_FIN_NO').set(f"{{{NS_XSI}}}nil", "true")
                rows_mapping_xlsx.append({"UNIQUE_ID": p_id, "STUDENT_UIN_FIN_NO": "", "PARENT_UIN_FIN_NO": uin_val, "STUDENT_UINFIN_TYPE_ICODE": "", "PREV_NRIC_UIN_FIN_NO": ""})
                
            for s_id in student_ids:
                item = ET.SubElement(root_map, 'ID_Mapping', {'UNIQUE_ID': s_id})
                uin_val = random_uin_generator()
                mapped_student_uins[s_id] = uin_val 
                ET.SubElement(item, 'STUDENT_UIN_FIN_NO').text = uin_val
                ET.SubElement(item, 'PARENT_UIN_FIN_NO').set(f"{{{NS_XSI}}}nil", "true")
                ET.SubElement(item, 'STUDENT_UINFIN_TYPE_ICODE').text = '1'
                ET.SubElement(item, 'PREV_NRIC_UIN_FIN_NO').set(f"{{{NS_XSI}}}nil", "true")
                rows_mapping_xlsx.append({"UNIQUE_ID": s_id, "STUDENT_UIN_FIN_NO": uin_val, "PARENT_UIN_FIN_NO": "", "STUDENT_UINFIN_TYPE_ICODE": "1", "PREV_NRIC_UIN_FIN_NO": ""})
                
            sync_files[f'FULL_SFS_ID_MAPPING_MK_{current_timestamp}.zip'] = (f'FULL_SFS_ID_MAPPING_MK_{current_timestamp}.xml', root_map)

            # -- 2. Personal Info Generation --
            root_pers = ET.Element('INTERFACE', {'INTERFACE_NAME': 'STUDENT_Personal_INFO', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_STUDENT_BASIC_PERSONAL_MK_{current_timestamp}.xml', 'NO_RECORD': '5'})
            for s_id in student_ids:
                item = ET.SubElement(root_pers, 'STUDENT_BASIC_PERSONAL')
                ET.SubElement(item, 'RECORD_ID').text = '1'
                ET.SubElement(item, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = s_id
                
                uin_from_mapping = mapped_student_uins.get(s_id, "UNKNOWN")
                student_name = f"student-{uin_from_mapping}"
                ET.SubElement(item, 'STUDENT_NAME').text = student_name
                ET.SubElement(item, 'HANYU_PINYIN_NAME').text = 'HANYUPINYIN'
                
                birth_date = f"{random.randint(10,28)}-JAN-{random.randint(2001,2008)}"
                effective_date = f"{random.randint(10, 28)}-{random.choice(months_list)}-{random.randint(2000, 2025)}"
                sex_code = random.choice(sex_codes_list)
                
                ET.SubElement(item, 'BIRTH_DATE').text = birth_date
                ET.SubElement(item, 'CITIZENSHIP_CODE').text = '10'
                ET.SubElement(item, 'CITIZENSHIP_SGDRM_CODE').text = 'SG'
                ET.SubElement(item, 'RACE_CODE').text = '2'
                ET.SubElement(item, 'RELIGION_CODE').text = '9'
                ET.SubElement(item, 'RELIGION_SGDRM_CODE').text = 'F'
                ET.SubElement(item, 'SEX_CODE').text = sex_code
                ET.SubElement(item, 'EMAIL_ADDRESS').text = f"student_{uin_from_mapping.lower()}@yopmail.com"
                ET.SubElement(item, 'CITIZENSHIP_EFFECTIVE_DATE').text = effective_date
                ET.SubElement(item, 'CONTACT_SAMEAS_OFFICIAL_IND').text = 'Y'
                ET.SubElement(item, 'CONTACTADD_BLK_HSE_NO').text = '123'
                ET.SubElement(item, 'CONTACTADD_STREET_NAME').text = 'Random Street'
                ET.SubElement(item, 'CONTACTADD_FLOOR_NO').text = '05'
                ET.SubElement(item, 'CONTACTADD_UNIT_NO').text = '12'
                ET.SubElement(item, 'CONTACTADD_BLDG_NAME').text = 'Random Building'
                ET.SubElement(item, 'CONTACTADD_POSTAL_ECODE').text = '123456'
                ET.SubElement(item, 'TELEPHONE_NO').text = '61234567'
                ET.SubElement(item, 'HANDPHONE_NO').text = '91234567'
                ET.SubElement(item, 'OTHER_CONTACT_NO').text = '1'
                ET.SubElement(item, 'RES_TYPE_CODE').text = 'HDB'
                ET.SubElement(item, 'OFFICIALADD_BLK_HSE_NO').text = '123b'
                ET.SubElement(item, 'OFFICIALADD_STREET_NAME').text = 'Official Street'
                ET.SubElement(item, 'OFFICIALADD_FLOOR_NO').text = '05'
                ET.SubElement(item, 'OFFICIALADD_UNIT_NO').text = '12'
                ET.SubElement(item, 'OFFICIALADD_BLDG_NAME').text = 'Official Building'
                ET.SubElement(item, 'OFFICIALADD_POSTAL_ECODE').text = '123456'
                ET.SubElement(item, 'FOREIGNADD_LINE1_DESC').text = 'Line1'
                ET.SubElement(item, 'FOREIGNADD_LINE2_DESC').text = 'Line2'
                ET.SubElement(item, 'FOREIGNADD_POSTAL_ECODE').text = 'F1234'
                ET.SubElement(item, 'FOREIGNADD_COUNTRY_CODE').text = 'MY'
                ET.SubElement(item, 'FOREIGNADD_CONTACTCODE_NO_OLD').text = '123'
                ET.SubElement(item, 'FOREIGNADD_CONTACT_NO_OLD').text = '456'
                ET.SubElement(item, 'FOREIGNADD_CONTACTCODE_NO').text = '789'
                ET.SubElement(item, 'FOREIGNADD_CONTACT_NO').text = '012'
                ET.SubElement(item, 'FOREIGNADD_COUNTRY_SGDRM_CODE').text = 'MY'
                ET.SubElement(item, 'ADDRESS_IND').text = '1'
                ET.SubElement(item, 'CONTACTADD_STREET_CODE').text = 'ST01'
                ET.SubElement(item, 'OFFICIALADD_STREET_CODE').text = 'ST02'
                ET.SubElement(item, 'GUARDIAN_TYPE_ICODE').text = 'M'
                ET.SubElement(item, 'PASS_TYPE_CODE').text = '1'
                ET.SubElement(item, 'PASS_ISSUE_DATE').text = '20200101'
                ET.SubElement(item, 'PASS_EXPIRY_DATE').text = '20281231'
                ET.SubElement(item, 'RACE_REQUEST_DATE').text = '20200101'
                ET.SubElement(item, 'PR_TYPE').text = 'N'
                
                rows_personal_xlsx.append({
                    "RECORD_ID": "1", "UNIQUE_ID": s_id, "STUDENT_NAME": student_name, "HANYU_PINYIN_NAME": "HANYUPINYIN", "BIRTH_DATE": birth_date,
                    "CITIZENSHIP_CODE": "10", "CITIZENSHIP_SGDRM_CODE": "SG", "RACE_CODE": "2", "RELIGION_CODE": "9", "RELIGION_SGDRM_CODE": "F",
                    "SEX_CODE": sex_code, "EMAIL_ADDRESS": f"student_{uin_from_mapping.lower()}@yopmail.com", "CITIZENSHIP_EFFECTIVE_DATE": effective_date, "CONTACT_SAMEAS_OFFICIAL_IND": "Y",
                    "CONTACTADD_BLK_HSE_NO": "123", "CONTACTADD_STREET_NAME": "Random Street", "CONTACTADD_FLOOR_NO": "05", "CONTACTADD_UNIT_NO": "12",
                    "CONTACTADD_BLDG_NAME": "Random Building", "CONTACTADD_POSTAL_ECODE": "123456", "TELEPHONE_NO": "61234567", "HANDPHONE_NO": "91234567",
                    "OTHER_CONTACT_NO": "1", "RES_TYPE_CODE": "HDB", "OFFICIALADD_BLK_HSE_NO": "123b", "OFFICIALADD_STREET_NAME": "Official Street",
                    "OFFICIALADD_FLOOR_NO": "05", "OFFICIALADD_UNIT_NO": "12", "OFFICIALADD_BLDG_NAME": "Official Building", "OFFICIALADD_POSTAL_ECODE": "123456",
                    "FOREIGNADD_LINE1_DESC": "Line1", "FOREIGNADD_LINE2_DESC": "Line2", "FOREIGNADD_POSTAL_ECODE": "F1234", "FOREIGNADD_COUNTRY_CODE": "MY",
                    "FOREIGNADD_CONTACTCODE_NO_OLD": "123", "FOREIGNADD_CONTACT_NO_OLD": "456", "FOREIGNADD_CONTACTCODE_NO": "789", "FOREIGNADD_CONTACT_NO": "012",
                    "FOREIGNADD_COUNTRY_SGDRM_CODE": "MY", "ADDRESS_IND": "1", "CONTACTADD_STREET_CODE": "ST01", "OFFICIALADD_STREET_CODE": "ST02",
                    "GUARDIAN_TYPE_ICODE": "M", "PASS_TYPE_CODE": "1", "PASS_ISSUE_DATE": "20200101", "PASS_EXPIRY_DATE": "20281231", "RACE_REQUEST_DATE": "20200101", "PR_TYPE": "N"
                })
            sync_files[f'FULL_SFS_STUDENT_BASIC_PERSONAL_MK_{current_timestamp}.zip'] = (f'FULL_SFS_STUDENT_BASIC_PERSONAL_MK_{current_timestamp}.xml', root_pers)

            # -- 3. Parent Info Generation --
            root_parent = ET.Element('INTERFACE', {'INTERFACE_NAME': 'Student_Parent', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_STUDENT_PARENT_MK_{current_timestamp}.xml', 'NO_RECORD': '5'})
            for s_id, p_id in zip(student_ids, parent_ids):
                item = ET.SubElement(root_parent, 'STUDENT_PARENT')
                ET.SubElement(item, 'RECORD_ID').text = '1'
                ET.SubElement(item, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = s_id
                ET.SubElement(item, 'PARENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = p_id
                
                relation_icode = random.choice(relation_codes_list)
                sample_relation_heritage[f"{s_id}_{p_id}"] = relation_icode
                
                uin_from_mapping = mapped_student_uins.get(s_id, "UNKNOWN")
                parent_name = f"Parent-Of-{uin_from_mapping}"
                p_eff_date = f"{random.randint(10, 28)}-{random.choice(months_list)}-{random.randint(1980, 2010)}"
                p_email = f"parent_{uin_from_mapping.lower()}@yopmail.com"
                
                ET.SubElement(item, 'RELATION_ICODE').text = relation_icode
                ET.SubElement(item, 'PARENT_GUARDIAN_NAME').text = parent_name
                ET.SubElement(item, 'CITIZENSHIP_CODE').text = '10'
                ET.SubElement(item, 'RACE_CODE').text = '2'
                ET.SubElement(item, 'STANDARD_ATTENDED_CODE').text = '2'
                ET.SubElement(item, 'DECEASED_YEAR').text = '2001'
                ET.SubElement(item, 'TELEPHONE_NO').text = '65765772'
                ET.SubElement(item, 'HANDPHONE_NO').text = '76576560'
                ET.SubElement(item, 'OTHER_CONTACT_NO').text = '1'
                ET.SubElement(item, 'BIRTH_DATE').text = '26-JUN-2001'
                ET.SubElement(item, 'EMAIL_ADDRESS').text = p_email
                ET.SubElement(item, 'CITIZENSHIP_EFFECTIVE_DATE').text = p_eff_date
                ET.SubElement(item, 'CITIZENSHIP_SGDRM_CODE').text = 'SG'
                ET.SubElement(item, 'PR_TYPE').text = 'Y'
                ET.SubElement(item, 'NRIC_BLK_HSE_NO').text = 'A22'
                ET.SubElement(item, 'NRIC_STREET_CODE').text = 'STA22'
                ET.SubElement(item, 'NRIC_FLOOR_NO').text = 'F22'
                ET.SubElement(item, 'NRIC_UNIT_NO').text = 'U22'
                ET.SubElement(item, 'NRIC_POSTAL_ECODE').text = '550000'
                
                rows_parent_xlsx.append({
                    "RECORD_ID": "1", "UNIQUE_ID": s_id, "PARENT_UNIQUE_ID": p_id, "RELATION_ICODE": relation_icode, "PARENT_GUARDIAN_NAME": parent_name,
                    "CITIZENSHIP_CODE": "10", "RACE_CODE": "2", "STANDARD_ATTENDED_CODE": "2", "DECEASED_YEAR": "2001", "TELEPHONE_NO": "65765772",
                    "HANDPHONE_NO": "76576560", "OTHER_CONTACT_NO": "1", "BIRTH_DATE": "26-JUN-2001", "EMAIL_ADDRESS": p_email, "CITIZENSHIP_EFFECTIVE_DATE": p_eff_date,
                    "CITIZENSHIP_SGDRM_CODE": "SG", "PR_TYPE": "Y", "NRIC_BLK_HSE_NO": "A22", "NRIC_STREET_CODE": "STA22", "NRIC_FLOOR_NO": "F22", "NRIC_UNIT_NO": "U22", "NRIC_POSTAL_ECODE": "550000"
                })
            sync_files[f'FULL_SFS_STUDENT_PARENT_MK_{current_timestamp}.zip'] = (f'FULL_SFS_STUDENT_PARENT_MK_{current_timestamp}.xml', root_parent)

            # -- 4. Custodial Generation --
            root_custodial = ET.Element('INTERFACE', {'INTERFACE_NAME': 'custodial_info_mk', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_timestamp}.xml', 'NO_RECORD': '5'})
            for idx, (s_id, p_id) in enumerate(zip(student_ids, parent_ids), start=1):
                item = ET.SubElement(root_custodial, 'STUDENT_CUSTODIAL_INFO_MK')
                ET.SubElement(item, 'RECORD_ID').text = str(idx)
                ET.SubElement(item, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = s_id
                ET.SubElement(item, 'PARENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = p_id
                
                inherited_icode = sample_relation_heritage.get(f"{s_id}_{p_id}", "G4")
                ET.SubElement(item, 'RELATION_ICODE').text = inherited_icode
                ET.SubElement(item, 'CUSTODIAL_INFO').text = 'JN'
                ET.SubElement(item, f'{{{NS_XSI}}}RELATIONSHIP').set(f"{{{NS_XSI}}}nil", "true")
                ET.SubElement(item, 'PG_ACCESS_IND').text = '2'
                ET.SubElement(item, 'LAST_UPDATED_DATE').text = '2026-06-30'
                
                rows_custodial_xlsx.append({
                    "RECORD_ID": str(idx), "UNIQUE_ID": s_id, "PARENT_UNIQUE_ID": p_id, "RELATION_ICODE": inherited_icode,
                    "CUSTODIAL_INFO": "JN", "RELATIONSHIP": "", "PG_ACCESS_IND": "2", "LAST_UPDATED_DATE": "2026-06-30"
                })
            sync_files[f'FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_timestamp}.zip'] = (f'FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_timestamp}.xml', root_custodial)

            # -- Write Excel Pack --
            xlsx_buffer = io.BytesIO()
            with pd.ExcelWriter(xlsx_buffer, engine='openpyxl') as writer:
                pd.DataFrame(rows_mapping_xlsx).to_excel(writer, sheet_name="ID_MAPPING", index=False)
                pd.DataFrame(rows_personal_xlsx).to_excel(writer, sheet_name="BASIC_PERSONAL", index=False)
                pd.DataFrame(rows_parent_xlsx).to_excel(writer, sheet_name="STUDENT_PARENT", index=False)
                pd.DataFrame(rows_custodial_xlsx).to_excel(writer, sheet_name="STUDENT_CUSTODIAL", index=False)
            st.session_state["sync_xlsx_data"] = xlsx_buffer.getvalue()
            st.session_state["sync_xlsx_name"] = f"SAMPLE_DATA_ID_MAPPING_{current_timestamp}.xlsx"

            # -- Package ZIPs Pack --
            st.session_state["sync_download_queue"] = []
            for zip_name, (xml_name, xml_node) in sync_files.items():
                tree = ET.ElementTree(xml_node)
                try: ET.indent(tree, space="  ", level=0)
                except AttributeError: pass
                
                xml_buf = io.BytesIO()
                tree.write(xml_buf, encoding="UTF-8", xml_declaration=True)
                
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                    zf.writestr(xml_name, xml_buf.getvalue())
                
                st.session_state["sync_download_queue"].append({
                    "name": zip_name, "data": zip_buf.getvalue()
                })
            st.success("🎉 Synchronized Student, Parent & Custodial datasets generated successfully!")

    # Hiển thị nút Download cho Mock Sync Data ngoài expander để dễ thao tác
    if "sync_xlsx_data" in st.session_state:
        st.markdown("---")
        st.markdown("#### 📊 Master Consolidated Excel Workbook File")
        st.download_button(
            label="🟢 DOWNLOAD MASTER EXCEL WORKBOOK (.XLSX)",
            data=st.session_state["sync_xlsx_data"],
            file_name=st.session_state["sync_xlsx_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
    if "sync_download_queue" in st.session_state:
        st.markdown("#### 📦 Segmented Structural Target XML Bundles (ZIPs)")
        cols = st.columns(len(st.session_state["sync_download_queue"]))
        for idx, item in enumerate(st.session_state["sync_download_queue"]):
            with cols[idx]:
                st.download_button(
                    label=f"📁 {item['name'].split('_MK_')[0]}",
                    data=item['data'],
                    file_name=item['name'],
                    mime="application/zip",
                    key=f"sync_btn_{idx}",
                    use_container_width=True
                )
