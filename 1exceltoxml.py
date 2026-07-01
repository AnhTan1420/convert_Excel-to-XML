import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import time
import io
import zipfile
import random

# App Workspace Configuration
st.set_page_config(page_title="MOE Jordan", page_icon="⚙️", layout="centered")
st.title("MOE: Convert XLSX/CSV to XML and compress to ZIP")

# --- SECTION 1: LARGE DATASET GENERATION (900K RECORDS) ---
st.markdown("---")
with st.expander("Generate 900K Records", expanded=False):
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
st.markdown("---")
with st.expander("✨ Generate Random 5 Sync Records (Student, Parent & Custodial)", expanded=False):
    st.write("Randomly generate randomized synchronized datasets for 5 students across both **XML (ZIP format)** and a single unified **Excel Workbook (.xlsx)**.")
    
    if st.button("Start Generating 5 Sample Data Records", type="primary", use_container_width=True):
        current_timestamp = time.strftime('%Y%m%d%H%M%S')
        epoch_ms = str(int(time.time() * 1000))
        NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
        
        student_ids = [f"JDK-{i}" for i in range(101, 106)]
        parent_ids = [f"JDKRP-{i}" for i in range(101, 106)]
        
        def random_uin():
            prefix = random.choice(['S', 'T', 'G', 'F', 'M'])
            digits = "".join([str(random.randint(0, 9)) for _ in range(7)])
            suffix = random.choice(['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'Z', 'R', 'Q', 'K', 'N', 'T'])
            return f"{prefix}{digits}{suffix}"
            
        sync_files = {}
        mapped_student_uins = {}
        months_list = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
        relation_codes_list = ['F', 'M', 'G', 'G4']

        rows_mapping_xlsx = []
        rows_personal_xlsx = []
        rows_parent_xlsx = []
        rows_custodial_xlsx = []

        root_map = ET.Element('INTERFACE', {'INTERFACE_NAME': 'STUDENT_ID_Mapping_INFO', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_ID_MAPPING_MK_{current_timestamp}.xml', 'NO_RECORD': '10'})
        ET.register_namespace('xs', NS_XSI)
        
        for p_id in parent_ids:
            item = ET.SubElement(root_map, 'ID_Mapping', {'UNIQUE_ID': p_id})
            uin_val = random_uin()
            mapped_student_uins[p_id] = uin_val 
            ET.SubElement(item, 'STUDENT_UIN_FIN_NO').set(f"{{{NS_XSI}}}nil", "true")
            ET.SubElement(item, 'PARENT_UIN_FIN_NO').text = uin_val
            ET.SubElement(item, f'{{{NS_XSI}}}STUDENT_UINFIN_TYPE_ICODE').set(f"{{{NS_XSI}}}nil", "true")
            ET.SubElement(item, f'{{{NS_XSI}}}PREV_NRIC_UIN_FIN_NO').set(f"{{{NS_XSI}}}nil", "true")
            rows_mapping_xlsx.append({"UNIQUE_ID": p_id, "STUDENT_UIN_FIN_NO": "", "PARENT_UIN_FIN_NO": uin_val, "STUDENT_UINFIN_TYPE_ICODE": "", "PREV_NRIC_UIN_FIN_NO": ""})
            
        for s_id in student_ids:
            item = ET.SubElement(root_map, 'ID_Mapping', {'UNIQUE_ID': s_id})
            uin_val = random_uin()
            mapped_student_uins[s_id] = uin_val 
            ET.SubElement(item, 'STUDENT_UIN_FIN_NO').text = uin_val
            ET.SubElement(item, f'{{{NS_XSI}}}PARENT_UIN_FIN_NO').set(f"{{{NS_XSI}}}nil", "true")
            ET.SubElement(item, 'STUDENT_UINFIN_TYPE_ICODE').text = '1'
            ET.SubElement(item, f'{{{NS_XSI}}}PREV_NRIC_UIN_FIN_NO').set(f"{{{NS_XSI}}}nil", "true")
            rows_mapping_xlsx.append({"UNIQUE_ID": s_id, "STUDENT_UIN_FIN_NO": uin_val, "PARENT_UIN_FIN_NO": "", "STUDENT_UINFIN_TYPE_ICODE": "1", "PREV_NRIC_UIN_FIN_NO": ""})
            
        sync_files[f'FULL_SFS_ID_MAPPING_MK_{current_timestamp}.zip'] = (f'FULL_SFS_ID_MAPPING_MK_{current_timestamp}.xml', root_map)

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
            sex_code = random.choice(['M', 'F'])
            
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

        root_parent = ET.Element('INTERFACE', {'INTERFACE_NAME': 'Student_Parent', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_STUDENT_PARENT_MK_{current_timestamp}.xml', 'NO_RECORD': '5'})
        for s_id, p_id in zip(student_ids, parent_ids):
            item = ET.SubElement(root_parent, 'STUDENT_PARENT')
            ET.SubElement(item, 'RECORD_ID').text = '1'
            ET.SubElement(item, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = s_id
            ET.SubElement(item, 'PARENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = p_id
            
            relation_icode = random.choice(relation_codes_list)
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

        root_custodial = ET.Element('INTERFACE', {'INTERFACE_NAME': 'custodial_info_mk', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_timestamp}.xml', 'NO_RECORD': '5'})
        for idx, (s_id, p_id) in enumerate(zip(student_ids, parent_ids), start=1):
            item = ET.SubElement(root_custodial, 'STUDENT_CUSTODIAL_INFO_MK')
            ET.SubElement(item, 'RECORD_ID').text = str(idx)
            ET.SubElement(item, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = s_id
            ET.SubElement(item, 'PARENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = p_id
            ET.SubElement(item, 'RELATION_ICODE').text = 'G4'
            ET.SubElement(item, 'CUSTODIAL_INFO').text = 'JN'
            ET.SubElement(item, f'{{{NS_XSI}}}RELATIONSHIP').set(f"{{{NS_XSI}}}nil", "true")
            ET.SubElement(item, 'PG_ACCESS_IND').text = '2'
            ET.SubElement(item, 'LAST_UPDATED_DATE').text = '2026-06-30'
            
            rows_custodial_xlsx.append({
                "RECORD_ID": str(idx), "UNIQUE_ID": s_id, "PARENT_UNIQUE_ID": p_id, "RELATION_ICODE": "G4",
                "CUSTODIAL_INFO": "JN", "RELATIONSHIP": "", "PG_ACCESS_IND": "2", "LAST_UPDATED_DATE": "2026-06-30"
            })
        sync_files[f'FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_timestamp}.zip'] = (f'FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_timestamp}.xml', root_custodial)

        xlsx_buffer = io.BytesIO()
        with pd.ExcelWriter(xlsx_buffer, engine='openpyxl') as writer:
            pd.DataFrame(rows_mapping_xlsx).to_excel(writer, sheet_name="ID_MAPPING", index=False)
            pd.DataFrame(rows_personal_xlsx).to_excel(writer, sheet_name="BASIC_PERSONAL", index=False)
            pd.DataFrame(rows_parent_xlsx).to_excel(writer, sheet_name="STUDENT_PARENT", index=False)
            pd.DataFrame(rows_custodial_xlsx).to_excel(writer, sheet_name="STUDENT_CUSTODIAL", index=False)
        st.session_state["sync_xlsx_data"] = xlsx_buffer.getvalue()
        st.session_state["sync_xlsx_name"] = f"SYNC_SYSTEM_SAMPLE_DATA_{current_timestamp}.xlsx"

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
                "name": zip_name,
                "data": zip_buf.getvalue()
            })
        st.success("🎉 Synchronized Student, Parent & Custodial datasets generated successfully!")

    if "sync_xlsx_data" in st.session_state:
        st.markdown("#### 📊 Master Consolidated Excel Workbook File")
        st.download_button(
            label=f"🟢 DOWNLOAD MASTER EXCEL WORKBOOK (.XLSX)",
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

# --- SECTION 3: AUTOMATED FULL PIPELINE GENERATOR FROM 1 MAPPING FILE ---
st.markdown("---")
st.subheader("⚙️ Automated System-wide Generation from ID Mapping")
st.write("Upload only **one single ID Mapping file**, the engine will automatically split and construct all related target schemas (`ID_MAPPING`, `BASIC_PERSONAL`, `STUDENT_PARENT`, `STUDENT_CUSTODIAL`) in ZIP formats.")

if "pipeline_download_queue" not in st.session_state:
    st.session_state.pipeline_download_queue = []

# Đơn giản hoá giao diện, chỉ nhận đúng 1 file Mapping duy nhất
uploaded_mapping_file = st.file_uploader("Upload ID Mapping Excel (.xlsx) or CSV (.csv)", type=["xlsx", "csv"], accept_multiple_files=False)

if uploaded_mapping_file:
    if st.button("🚀 Process & Generate All 4 XML Files", type="primary", use_container_width=True):
        st.session_state.pipeline_download_queue = []
        current_time = time.strftime('%Y%m%d%H%M%S')
        epoch_ms = str(int(time.time() * 1000))
        NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
        
        try:
            if uploaded_mapping_file.name.lower().endswith(".csv"):
                df_map = pd.read_csv(uploaded_mapping_file, encoding='utf-8')
            else:
                df_map = pd.read_excel(uploaded_mapping_file, engine='openpyxl')
                
            df_map = df_map.fillna("").astype(str)
            for col in df_map.columns:
                df_map[col] = df_map[col].apply(lambda x: x.split(".")[0] if x.endswith(".0") else x)
            
            # Phân tách dữ liệu
            student_records = []
            parent_records = []
            
            for _, row in df_map.iterrows():
                uid = row.get("UNIQUE_ID", "").strip()
                st_uin = row.get("STUDENT_UIN_FIN_NO", "").strip()
                pt_uin = row.get("PARENT_UIN_FIN_NO", "").strip()
                st_type = row.get("STUDENT_UINFIN_TYPE_ICODE", "").strip()
                prev_nric = row.get("PREV_NRIC_UIN_FIN_NO", "").strip()
                
                if uid:
                    if st_uin != "":
                        student_records.append({"id": uid, "uin": st_uin})
                    elif pt_uin != "":
                        parent_records.append({"id": uid, "uin": pt_uin})

            # Ghép cặp student và parent tương ứng dựa trên thứ tự xuất hiện hoặc index (Logic ghép cặp Sync 1-1 giống Sec 2)
            paired_records = []
            min_len = min(len(student_records), len(parent_records))
            for i in range(min_len):
                paired_records.append({
                    "student_id": student_records[i]["id"],
                    "student_uin": student_records[i]["uin"],
                    "parent_id": parent_records[i]["id"],
                    "parent_uin": parent_records[i]["uin"]
                })
            
            months_list = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
            relation_codes_list = ['F', 'M', 'G', 'G4']
            generated_xmls = {}

            # --- 1. FILE ID MAPPING XML ---
            root_map = ET.Element('INTERFACE', {'INTERFACE_NAME': 'STUDENT_ID_Mapping_INFO', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_ID_MAPPING_MK_{current_time}.xml', 'NO_RECORD': str(len(df_map))})
            ET.register_namespace('xs', NS_XSI)
            for _, row in df_map.iterrows():
                uid = row.get("UNIQUE_ID", "").strip()
                item = ET.SubElement(root_map, 'ID_Mapping', {'UNIQUE_ID': uid})
                
                st_uin = row.get("STUDENT_UIN_FIN_NO", "").strip()
                pt_uin = row.get("PARENT_UIN_FIN_NO", "").strip()
                st_type = row.get("STUDENT_UINFIN_TYPE_ICODE", "").strip()
                prev_nric = row.get("PREV_NRIC_UIN_FIN_NO", "").strip()
                
                if st_uin: ET.SubElement(item, 'STUDENT_UIN_FIN_NO').text = st_uin
                else: ET.SubElement(item, 'STUDENT_UIN_FIN_NO').set(f"{{{NS_XSI}}}nil", "true")
                    
                if pt_uin: ET.SubElement(item, 'PARENT_UIN_FIN_NO').text = pt_uin
                else: ET.SubElement(item, 'PARENT_UIN_FIN_NO').set(f"{{{NS_XSI}}}nil", "true")
                    
                if st_type: ET.SubElement(item, 'STUDENT_UINFIN_TYPE_ICODE').text = st_type
                else: ET.SubElement(item, 'STUDENT_UINFIN_TYPE_ICODE').set(f"{{{NS_XSI}}}nil", "true")
                    
                if prev_nric: ET.SubElement(item, 'PREV_NRIC_UIN_FIN_NO').text = prev_nric
                else: ET.SubElement(item, 'PREV_NRIC_UIN_FIN_NO').set(f"{{{NS_XSI}}}nil", "true")
                
            generated_xmls[f'FULL_SFS_ID_MAPPING_MK_{current_time}'] = root_map

            # --- 2. FILE BASIC PERSONAL XML (Dựa trên danh sách Students) ---
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
                ET.SubElement(item, 'SEX_CODE').text = random.choice(['M', 'F'])
                ET.SubElement(item, 'EMAIL_ADDRESS').text = f"student_{st_rec['uin'].lower()}@yopmail.com"
                ET.SubElement(item, 'CITIZENSHIP_EFFECTIVE_DATE').text = f"01-JAN-2020"
                ET.SubElement(item, 'CONTACT_SAMEAS_OFFICIAL_IND').text = 'Y'
                ET.SubElement(item, 'CONTACTADD_BLK_HSE_NO').text = '100'
                ET.SubElement(item, 'CONTACTADD_STREET_NAME').text = 'Main Street'
                ET.SubElement(item, 'CONTACTADD_POSTAL_ECODE').text = '654321'
                # Điền thẻ trống có nil=true cho cấu trúc hoàn chỉnh
                for blank_tag in ['CONTACTADD_FLOOR_NO','CONTACTADD_UNIT_NO','CONTACTADD_BLDG_NAME','TELEPHONE_NO','HANDPHONE_NO','OTHER_CONTACT_NO','RES_TYPE_CODE','OFFICIALADD_BLK_HSE_NO','OFFICIALADD_STREET_NAME','OFFICIALADD_FLOOR_NO','OFFICIALADD_UNIT_NO','OFFICIALADD_BLDG_NAME','OFFICIALADD_POSTAL_ECODE','FOREIGNADD_LINE1_DESC','FOREIGNADD_LINE2_DESC','FOREIGNADD_POSTAL_ECODE','FOREIGNADD_COUNTRY_CODE','FOREIGNADD_CONTACTCODE_NO_OLD','FOREIGNADD_CONTACT_NO_OLD','FOREIGNADD_CONTACTCODE_NO','FOREIGNADD_CONTACT_NO','FOREIGNADD_COUNTRY_SGDRM_CODE','ADDRESS_IND','CONTACTADD_STREET_CODE','OFFICIALADD_STREET_CODE','GUARDIAN_TYPE_ICODE','PASS_TYPE_CODE','PASS_ISSUE_DATE','PASS_EXPIRY_DATE','RACE_REQUEST_DATE','PR_TYPE']:
                    ET.SubElement(item, f'{{{NS_XSI}}}{blank_tag}').set(f"{{{NS_XSI}}}nil", "true")
            generated_xmls[f'FULL_SFS_STUDENT_BASIC_PERSONAL_MK_{current_time}'] = root_pers

            # --- 3. FILE STUDENT PARENT XML (Dựa trên danh sách Cặp ghép đôi) ---
            root_parent = ET.Element('INTERFACE', {'INTERFACE_NAME': 'Student_Parent', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_STUDENT_PARENT_MK_{current_time}.xml', 'NO_RECORD': str(len(paired_records))})
            for pair in paired_records:
                item = ET.SubElement(root_parent, 'STUDENT_PARENT')
                ET.SubElement(item, 'RECORD_ID').text = '1'
                ET.SubElement(item, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = pair["student_id"]
                ET.SubElement(item, 'PARENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = pair["parent_id"]
                ET.SubElement(item, 'RELATION_ICODE').text = random.choice(relation_codes_list)
                ET.SubElement(item, 'PARENT_GUARDIAN_NAME').text = f"Parent-Of-{pair['student_uin']}"
                ET.SubElement(item, 'CITIZENSHIP_CODE').text = '10'
                ET.SubElement(item, 'RACE_CODE').text = '2'
                ET.SubElement(item, 'EMAIL_ADDRESS').text = f"parent_{pair['parent_id'].lower()}@yopmail.com"
                for blank_tag in ['STANDARD_ATTENDED_CODE','DECEASED_YEAR','TELEPHONE_NO','HANDPHONE_NO','OTHER_CONTACT_NO','BIRTH_DATE','CITIZENSHIP_EFFECTIVE_DATE','CITIZENSHIP_SGDRM_CODE','PR_TYPE','NRIC_BLK_HSE_NO','NRIC_STREET_CODE','NRIC_FLOOR_NO','NRIC_UNIT_NO','NRIC_POSTAL_ECODE']:
                    ET.SubElement(item, f'{{{NS_XSI}}}{blank_tag}').set(f"{{{NS_XSI}}}nil", "true")
            generated_xmls[f'FULL_SFS_STUDENT_PARENT_MK_{current_time}'] = root_parent

            # --- 4. FILE STUDENT CUSTODIAL XML (Dựa trên danh sách Cặp ghép đôi) ---
            root_custodial = ET.Element('INTERFACE', {'INTERFACE_NAME': 'custodial_info_mk', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_time}.xml', 'NO_RECORD': str(len(paired_records))})
            for idx, pair in enumerate(paired_records, start=1):
                item = ET.SubElement(root_custodial, 'STUDENT_CUSTODIAL_INFO_MK')
                ET.SubElement(item, 'RECORD_ID').text = str(idx)
                ET.SubElement(item, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = pair["student_id"]
                ET.SubElement(item, 'PARENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = pair["parent_id"]
                ET.SubElement(item, 'RELATION_ICODE').text = 'G4'
                ET.SubElement(item, 'CUSTODIAL_INFO').text = 'JN'
                ET.SubElement(item, f'{{{NS_XSI}}}RELATIONSHIP').set(f"{{{NS_XSI}}}nil", "true")
                ET.SubElement(item, 'PG_ACCESS_IND').text = '2'
                ET.SubElement(item, 'LAST_UPDATED_DATE').text = '2026-06-30'
            generated_xmls[f'FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_time}'] = root_custodial

            # Đóng gói ZIP riêng biệt từng file thành hàng đợi
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
                    "total_records": xml_node.get("NO_RECORD")
                })
                
            st.success("🎉 Processed! 4 synchronized target XML Packages generated from your ID Mapping file.")
        except Exception as e:
            st.error(f"❌ Failed to execute pipeline conversion: {e}")

# Hiển thị kết quả tải xuống cho quy trình tự động hóa Section 3
if st.session_state.pipeline_download_queue:
    st.markdown("#### 📥 Download Generated Interface Packages")
    grid_cols = st.columns(2)
    for idx, item in enumerate(st.session_state.pipeline_download_queue):
        col_target = grid_cols[idx % 2]
        with col_target:
            with st.container(border=True):
                st.markdown(f"📦 **{item['zip_name']}**")
                st.caption(f"Payload Rows: `{item['total_records']}` rows parsed.")
                st.download_button(
                    label=f"Download ZIP Package",
                    data=item['zip_data'],
                    file_name=item['zip_name'],
                    mime="application/zip",
                    key=f"pipeline_btn_{idx}",
                    use_container_width=True
                )
