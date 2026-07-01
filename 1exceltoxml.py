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
    
    # State control for heavy generation
    if st.button("Start Generating (900K)", type="secondary", use_container_width=True):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        status_text.text("⏳ Initializing compression stream...")
        current_time = time.strftime('%Y%m%d%H%M%S')
        xml_filename = f"FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_time}.xml"
        zip_filename = f"FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_time}.zip"
        
        # Fixed Schema Definitions
        NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
        total_records = 900000
        
        # Output buffer to process zip directly in RAM
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Initialize XML content stream using string builder for maximum processing speed
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
                
                # Stream out memory blocks periodically to update UI and keep execution footprint low
                if i % chunk_size == 0:
                    status_text.text(f"⏳ Processing data blocks: {i:,} / {total_records:,} records...")
                    progress_bar.progress(i / total_records)
            
            xml_chunks.append("</INTERFACE>")
            
            # Combine all chunks and inject into the archive file
            full_xml_string = "".join(xml_chunks)
            zipf.writestr(xml_filename, full_xml_string.encode('utf-8'))
            
        status_text.text("🎉 Generation completed successfully! Package ready for download.")
        progress_bar.empty()
        
        # Save payload to session state to maintain download stability
        st.session_state["large_zip_data"] = zip_buffer.getvalue()
        st.session_state["large_zip_name"] = zip_filename

    # Render download button persistent layout
    if "large_zip_data" in st.session_state:
        st.success(f"📦 Archive compiled: `{st.session_state['large_zip_name']}`")
        st.download_button(
            label="📥 CLICK HERE TO DOWNLOAD ZIP FILE (900K DATA)",
            data=st.session_state["large_zip_data"],
            file_name=st.session_state["large_zip_name"],
            mime="application/zip",
            use_container_width=True
        )

# --- SECTION 2: SYNCHRONIZED TARGET PAIRS GENERATION (3 TABLES) ---
st.markdown("---")
with st.expander("✨ Generate Random 5 Sync Records (Student & Parent)", expanded=False):
    st.write("Randomly generate synchronized datasets for 5 students across 3 structural target files: ID Mapping, Basic Personal, and Student Parent.")
    
    if st.button("Start Generating 5 Sample Data Records", type="primary", use_container_width=True):
        current_timestamp = time.strftime('%Y%m%d%H%M%S')
        epoch_ms = str(int(time.time() * 1000))
        NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
        
        # Initialize 5 unified pairs
        student_ids = [f"JDK-{i}" for i in range(101, 106)]
        parent_ids = [f"JDKRP-{i}" for i in range(101, 106)]
        
        def random_uin():
            prefix = random.choice(['S', 'T', 'G', 'F', 'M'])
            digits = "".join([str(random.randint(0, 9)) for _ in range(7)])
            suffix = random.choice(['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'Z', 'R', 'Q', 'K', 'N', 'T'])
            return f"{prefix}{digits}{suffix}"
            
        sync_files = {}
        mapped_student_uins = {}
        mapped_parent_uins = {}
        months_list = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
        relation_codes_list = ['F', 'M', 'G', 'G4']

        # --- 2.1 ID MAPPING (5 Parents + 5 Students = 10 Records Total) ---
        root_map = ET.Element('INTERFACE', {'INTERFACE_NAME': 'STUDENT_ID_Mapping_INFO', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_ID_MAPPING_MK_{current_timestamp}.xml', 'NO_RECORD': '10'})
        ET.register_namespace('xs', NS_XSI)
        
        # Map parents to unique validation elements
        for p_id in parent_ids:
            item = ET.SubElement(root_map, 'ID_Mapping', {'UNIQUE_ID': p_id})
            uin_val = random_uin()
            mapped_parent_uins[p_id] = uin_val 
            
            ET.SubElement(item, f'{{{NS_XSI}}}STUDENT_UIN_FIN_N').set(f"{{{NS_XSI}}}nil", "true")
            ET.SubElement(item, 'PARENT_UIN_FIN_NO').text = uin_val
            ET.SubElement(item, f'{{{NS_XSI}}}STUDENT_UINFIN_TYPE_ICODE').set(f"{{{NS_XSI}}}nil", "true")
            ET.SubElement(item, f'{{{NS_XSI}}}PREV_NRIC_UIN_FIN_NO').set(f"{{{NS_XSI}}}nil", "true")
            
        # Map primary students to unique validation elements
        for s_id in student_ids:
            item = ET.SubElement(root_map, 'ID_Mapping', {'UNIQUE_ID': s_id})
            uin_val = random_uin()
            mapped_student_uins[s_id] = uin_val 
            
            ET.SubElement(item, 'STUDENT_UIN_FIN_NO').text = uin_val
            ET.SubElement(item, f'{{{NS_XSI}}}PARENT_UIN_FIN_NO').set(f"{{{NS_XSI}}}nil", "true")
            ET.SubElement(item, 'STUDENT_UINFIN_TYPE_ICODE').text = '1'
            ET.SubElement(item, f'{{{NS_XSI}}}PREV_NRIC_UIN_FIN_NO').set(f"{{{NS_XSI}}}nil", "true")
            
        sync_files[f'FULL_SFS_ID_MAPPING_MK_{current_timestamp}.zip'] = (f'FULL_SFS_ID_MAPPING_MK_{current_timestamp}.xml', root_map)

        # --- 2.2 BASIC PERSONAL - STUDENT ATTRIBUTES (5 Records) ---
        root_pers = ET.Element('INTERFACE', {'INTERFACE_NAME': 'STUDENT_Personal_INFO', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_STUDENT_BASIC_PERSONAL_MK_{current_timestamp}.xml', 'NO_RECORD': '5'})
        for s_id in student_ids:
            item = ET.SubElement(root_pers, 'STUDENT_BASIC_PERSONAL')
            ET.SubElement(item, 'RECORD_ID').text = '1'
            ET.SubElement(item, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = s_id
            
            # Synchronization mapping binding
            uin_from_mapping = mapped_student_uins.get(s_id, "UNKNOWN")
            ET.SubElement(item, 'STUDENT_NAME').text = f"RANDOM-NAME-{uin_from_mapping}"
            
            ET.SubElement(item, 'HANYU_PINYIN_NAME').text = 'HANYUPINYIN'
            ET.SubElement(item, 'BIRTH_DATE').text = f"{random.randint(10,28)}-JAN-{random.randint(2001,2008)}"
            ET.SubElement(item, 'CITIZENSHIP_CODE').text = '10'
            ET.SubElement(item, 'CITIZENSHIP_SGDRM_CODE').text = 'SG'
            ET.SubElement(item, 'RACE_CODE').text = '2'
            ET.SubElement(item, 'RELIGION_CODE').text = '9'
            ET.SubElement(item, 'RELIGION_SGDRM_CODE').text = 'F'
            ET.SubElement(item, 'SEX_CODE').text = random.choice(['M', 'F'])
            ET.SubElement(item, f'{{{NS_XSI}}}EMAIL_ADDRESS').set(f"{{{NS_XSI}}}nil", "true")
            
            ET.SubElement(item, 'CITIZENSHIP_EFFECTIVE_DATE').text = f"{random.randint(10, 28)}-{random.choice(months_list)}-{random.randint(2000, 2025)}"
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
        sync_files[f'FULL_SFS_STUDENT_BASIC_PERSONAL_MK_{current_timestamp}.zip'] = (f'FULL_SFS_STUDENT_BASIC_PERSONAL_MK_{current_timestamp}.xml', root_pers)

        # --- 2.3 STUDENT PARENT - PARENT ATTRIBUTES (5 Records) ---
        root_parent = ET.Element('INTERFACE', {'INTERFACE_NAME': 'Student_Parent', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_STUDENT_PARENT_MK_{current_timestamp}.xml', 'NO_RECORD': '5'})
        for s_id, p_id in zip(student_ids, parent_ids):
            item = ET.SubElement(root_parent, 'STUDENT_PARENT')
            ET.SubElement(item, 'RECORD_ID').text = '1'
            ET.SubElement(item, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = s_id
            ET.SubElement(item, 'PARENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = p_id
            
            ET.SubElement(item, 'RELATION_ICODE').text = random.choice(relation_codes_list)
            
            # Cross-referenced child identification linking
            uin_from_mapping = mapped_student_uins.get(s_id, "UNKNOWN")
            ET.SubElement(item, 'PARENT_GUARDIAN_NAME').text = f"Parent-Of-{uin_from_mapping}"
            
            ET.SubElement(item, 'CITIZENSHIP_CODE').text = '10'
            ET.SubElement(item, 'RACE_CODE').text = '2'
            ET.SubElement(item, 'STANDARD_ATTENDED_CODE').text = '2'
            ET.SubElement(item, 'DECEASED_YEAR').text = '2001'
            ET.SubElement(item, 'TELEPHONE_NO').text = '65765772'
            ET.SubElement(item, 'HANDPHONE_NO').text = '76576560'
            ET.SubElement(item, 'OTHER_CONTACT_NO').text = '1'
            ET.SubElement(item, 'BIRTH_DATE').text = '26-JUN-2001'
            ET.SubElement(item, 'EMAIL_ADDRESS').text = f"parent_{uin_from_mapping.lower()}@yopmail.com"
            
            ET.SubElement(item, 'CITIZENSHIP_EFFECTIVE_DATE').text = f"{random.randint(10, 28)}-{random.choice(months_list)}-{random.randint(1980, 2010)}"
            ET.SubElement(item, 'CITIZENSHIP_SGDRM_CODE').text = 'SG'
            ET.SubElement(item, 'PR_TYPE').text = 'Y'
            ET.SubElement(item, 'NRIC_BLK_HSE_NO').text = 'A22'
            ET.SubElement(item, 'NRIC_STREET_CODE').text = 'STA22'
            ET.SubElement(item, 'NRIC_FLOOR_NO').text = 'F22'
            ET.SubElement(item, 'NRIC_UNIT_NO').text = 'U22'
            ET.SubElement(item, 'NRIC_POSTAL_ECODE').text = '550000'
        sync_files[f'FULL_SFS_STUDENT_PARENT_MK_{current_timestamp}.zip'] = (f'FULL_SFS_STUDENT_PARENT_MK_{current_timestamp}.xml', root_parent)

        # --- 2.4 STUDENT CUSTODIAL INFO - CUSTODIAL TABLE (5 Records) ---
        root_custodial = ET.Element('INTERFACE', {'INTERFACE_NAME': 'custodial_info_mk', 'FILE_CREATED_TIME': epoch_ms, 'FILE_NAME': f'FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_timestamp}.xml', 'NO_RECORD': '5'})
        for idx, (d_id, p_id) in enumerate(zip(student_ids, parent_ids), start=1):
            item = ET.SubElement(root_custodial, 'STUDENT_CUSTODIAL_INFO_MK')
            ET.SubElement(item, 'RECORD_ID').text = str(idx)
            ET.SubElement(item, 'STUDENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = d_id
            ET.SubElement(item, 'PARENT_UNIQUE_ID', {'UNIQUE_ID': 'Y'}).text = p_id
            ET.SubElement(item, 'RELATION_ICODE').text = 'G4'
            ET.SubElement(item, 'CUSTODIAL_INFO').text = 'JN'
            ET.SubElement(item, f'{{{NS_XSI}}}RELATIONSHIP').set(f"{{{NS_XSI}}}nil", "true")
            ET.SubElement(item, 'PG_ACCESS_IND').text = '2'
            ET.SubElement(item, 'LAST_UPDATED_DATE').text = '2026-06-30'
        sync_files[f'FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_timestamp}.zip'] = (f'FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_{current_timestamp}.xml', root_custodial)
        
        # Assemble and structure binary packages
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
        st.success("🎉 Synchronized Student & Parent datasets generated successfully!")

    # Render localized generation package nodes
    if "sync_download_queue" in st.session_state:
        st.write("📥 **Download Synchronized Files:**")
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

# --- SECTION 3: WORKBOOK SOURCE CONVERSION ENGINE ---
st.markdown("---")

# Initialize Session States to prevent download button data-loss on click
if "download_queue" not in st.session_state:
    st.session_state.download_queue = []
if "run_summary" not in st.session_state:
    st.session_state.run_summary = {"success_count": 0, "total_records": 0}

# Create excel template in memory ram
def generate_template(headers):
    buffer = io.BytesIO()
    df_template = pd.DataFrame(columns=headers)
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False)
    return buffer.getvalue()

# Schema Structures definition
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
        "RECORD_ID", "UNIQUE_ID", "STUDENT_STATUS_ICODE", "SCHOOL_CODE", "ADMISSION_NO", 
        "ACADEMIC_YEAR", "LEVEL_XCODE", "STREAM_XCODE", "CLASS_XCODE", "CLASS_SERIAL_NO", 
        "COURSE_TYPE_CODE", "FIRSTLANGUAGE_L1_CODE", "SECONDLANGUAGE_L2_CODE", 
        "LEAVE_OF_ABSENCE_IND", "REPEAT_STUD_IND", "ACAD_STATUS_ICODE", "EFFECTIVE_DATE", 
        "SCHOOL_NAME", "CLASS_NAME", "LEVEL_NAME", "STREAM_NAME", "COURSE_XCODE", 
        "COURSE_NAME", "COURSE_TYPE_NAME", "INTF_PROMOTION_IND", "RECOMMENDED_LEVEL_XCODE", 
        "RECOMMENDED_STREAM_XCODE", "JC_PROVISIONAL_IND", "POSTED_IND", "MATRICULATION_NO", "IP_IND"
    ],
    "PARENT": [
        "RECORD_ID", "UNIQUE_ID", "PARENT_UNIQUE_ID", "RELATION_ICODE", "PARENT_GUARDIAN_NAME", 
        "CITIZENSHIP_CODE", "RACE_CODE", "STANDARD_ATTENDED_CODE", "DECEASED_YEAR", 
        "TELEPHONE_NO", "HANDPHONE_NO", "OTHER_CONTACT_NO", "BIRTH_DATE", "EMAIL_ADDRESS", 
        "CITIZENSHIP_EFFECTIVE_DATE", "CITIZENSHIP_SGDRM_CODE", "PR_TYPE", "NRIC_BLK_HSE_NO", 
        "NRIC_STREET_CODE", "NRIC_FLOOR_NO", "NRIC_UNIT_NO", "NRIC_POSTAL_ECODE"
    ],
    "MOVEMENT": [
        "RECORD_ID", "UNIQUE_ID", "PARENT_UNIQUE_ID", "RELATION_ICODE", "PARENT_GUARDIAN_NAME", 
        "CITIZENSHIP_CODE", "RACE_CODE", "STANDARD_ATTENDED_CODE", "DECEASED_YEAR", 
        "TELEPHONE_NO", "HANDPHONE_NO", "OTHER_CONTACT_NO", "BIRTH_DATE", "EMAIL_ADDRESS", 
        "CITIZENSHIP_EFFECTIVE_DATE", "CITIZENSHIP_SGDRM_CODE", "PR_TYPE", "NRIC_BLK_HSE_NO", 
        "NRIC_STREET_CODE", "NRIC_FLOOR_NO", "NRIC_UNIT_NO", "NRIC_POSTAL_ECODE"
    ],
     "CUSTODIAL": [
        "RECORD_ID", "UNIQUE_ID", "PARENT_UNIQUE_ID", "RELATION_ICODE", "CUSTODIAL_INFO",
        "RELATIONSHIP", "PG_ACCESS_IND", "LAST_UPDATED_DATE"
    ]
}

# Sample file download box (Templates)
with st.expander("📥 Download Excel Sample Templates", expanded=True):
    st.markdown("Select the School Cockpit MK template type below to download the standard Excel file structure:")
    col_t1, col_t2, col_t3, col_t4, col_t5, col_t6 = st.columns(6)
    
    with col_t1:
        st.download_button(
            label="📁 ID_MAPPING",
            data=generate_template(TEMPLATE_COLUMNS["MAPPING"]),
            file_name="Template_ID_MAPPING.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_t2:
        st.download_button(
            label="📁 PERSONAL",
            data=generate_template(TEMPLATE_COLUMNS["PERSONAL"]),
            file_name="Template_BASIC_PERSONAL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_t3:
        st.download_button(
            label="📁 SCHOOL",
            data=generate_template(TEMPLATE_COLUMNS["SCHOOL"]),
            file_name="Template_BASIC_SCHOOL.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_t4:
        st.download_button(
            label="📁 PARENT",
            data=generate_template(TEMPLATE_COLUMNS["PARENT"]),
            file_name="Template_STUDENT_PARENT.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_t5:
        st.download_button(
            label="📁 MOVEMENT",
            data=generate_template(TEMPLATE_COLUMNS["MOVEMENT"]),
            file_name="Template_MOVEMENT.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_t6:
        st.download_button(
            label="📁 STUDENT_CUSTODIAL",
            data=generate_template(TEMPLATE_COLUMNS["CUSTODIAL"]),
            file_name="Template_STUDENT_CUSTODIAL.xlsx",
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
    | **ID Mapping** | `Mapping` | `FULL_SFS_ID_MAPPING_MK_*` | `<ID_Mapping UNIQUE_ID="...">` |
    | **Basic Personal** | `Personal` | `FULL_SFS_STUDENT_BASIC_PERSONAL_MK_*` | `<STUDENT_BASIC_PERSONAL>` |
    | **Basic School** | `School` | `FULL_SFF_STUDENT_BASIC_SCHOOL_MK_*` | `<STUDENT_BASIC_SCHOOL>` |
    | **Student Parent** | `Parent` | `FULL_SFS_STUDENT_PARENT_MK_*` | `<STUDENT_PARENT>` |
    | **Student Custodial** | `Custodial` | `FULL_SFS_STUDENT_CUSTODIAL_INFO_MK_*` | `<STUDENT_CUSTODIAL_INFO_MK>` |
    | **Movement** | `Movement` | `FULL_SFS_MOVEMENT_MK_*` | `<MOVEMENT>` |
    """)

st.markdown("### 📤 Source File Upload")
uploaded_files = st.file_uploader(
    "Drag and drop or browse Excel workbooks (.xlsx) or CSV files (.csv)", 
    type=["xlsx", "csv"], 
    accept_multiple_files=True
)

# Rule Mapping Engine Configuration
MAPPING_RULES = {
    "MAPPING": {"prefix": "FULL_SFS_ID_MAPPING_MK", "interface": "STUDENT_ID_Mapping_INFO"},
    "PERSONAL": {"prefix": "FULL_SFS_STUDENT_BASIC_PERSONAL_MK", "interface": "STUDENT_Personal_INFO"},
    "SCHOOL": {"prefix": "FULL_SFF_STUDENT_BASIC_SCHOOL_MK", "interface": "STUDENT_BASIC_SCHOOL"},
    "PARENT": {"prefix": "FULL_SFS_STUDENT_PARENT_MK", "interface": "Student_Parent"},
    "MOVEMENT": {"prefix": "FULL_SFS_MOVEMENT_MK", "interface": "MOVEMENT"},
    "CUSTODIAL": {"prefix": "FULL_SFS_STUDENT_CUSTODIAL_INFO_MK", "interface": "custodial_info_mk"}
}

if uploaded_files:
    st.info(f"📋 **Stage Queue:** {len(uploaded_files)} file(s) loaded into session memory.")
    
    if st.button("🚀 Execute Split Transformation", type="primary", use_container_width=True):
        st.session_state.download_queue = []
        success_count = 0
        total_records_processed = 0
        current_time = time.strftime('%Y%m%d%H%M%S')
        
        for uploaded_file in uploaded_files:
            file_name = uploaded_file.name
            matched_rule = None
            
            for key, rule in MAPPING_RULES.items():
                if key.upper() in file_name.upper():
                    matched_rule = rule
                    break
            
            if matched_rule:
                base_prefix = matched_rule["prefix"]
                interface = matched_rule["interface"]
                
                if file_name.upper().startswith("DELTA_"):
                    prefix = base_prefix.replace("FULL_", "DELTA_")
                else:
                    prefix = base_prefix
                    
                xml_filename = f"{prefix}_{current_time}.xml"
                zip_filename = f"{prefix}_{current_time}.zip"
                
                is_personal = (base_prefix == "FULL_SFS_STUDENT_BASIC_PERSONAL_MK")
                is_school = (base_prefix == "FULL_SFF_STUDENT_BASIC_SCHOOL_MK")
                is_parent = (base_prefix == "FULL_SFS_STUDENT_PARENT_MK")
                is_movement = (base_prefix == "FULL_SFS_MOVEMENT_MK")
                is_custodial = (base_prefix == "FULL_SFS_STUDENT_CUSTODIAL_INFO_MK")
                
                try:
                    if file_name.lower().endswith(".csv"):
                        df = pd.read_csv(uploaded_file, encoding='utf-8')
                    else:
                        df = pd.read_excel(uploaded_file, engine='openpyxl')
                    
                    df = df.fillna("")
                    df = df.astype(str)
                    
                    for col in df.columns:
                        df[col] = df[col].apply(
                            lambda x: x.split(".")[0] if x.endswith(".0") else x
                        )
                    
                    NS = 'http://www.w3.org/2001/XMLSchema-instance'
                    ET.register_namespace('xs', NS)
                    
                    root = ET.Element('INTERFACE', {
                        'INTERFACE_NAME': interface,
                        'FILE_CREATED_TIME': str(int(time.time() * 1000)),
                        'FILE_NAME': xml_filename, 
                        'NO_RECORD': str(len(df)) 
                    })
                    
                    for index, row in df.iterrows():
                        id_col = 'UNIQUE_ID' if 'UNIQUE_ID' in df.columns else ('STUDENT_UNIQUE_ID' if 'STUDENT_UNIQUE_ID' in df.columns else None)
                        unique_id_val = str(row[id_col]) if id_col and row[id_col] != "" else ""
                        
                        if is_personal:
                            row_element = ET.SubElement(root, 'STUDENT_BASIC_PERSONAL')
                        elif is_school:
                            row_element = ET.SubElement(root, 'STUDENT_BASIC_SCHOOL')
                        elif is_parent:
                            row_element = ET.SubElement(root, 'STUDENT_PARENT')
                        elif is_movement:
                            row_element = ET.SubElement(root, 'MOVEMENT')
                        elif is_custodial:
                            row_element = ET.SubElement(root, 'STUDENT_CUSTODIAL_INFO_MK')
                        else:
                            row_element = ET.SubElement(root, 'ID_Mapping', {'UNIQUE_ID': unique_id_val})
                            
                        for col_name in df.columns:
                            if not (is_personal or is_school or is_parent or is_movement or is_custodial) and col_name == 'UNIQUE_ID':
                                continue
                                
                            xml_tag = col_name
                            if col_name == 'UNIQUE_ID' and (is_personal or is_school or is_parent or is_movement or is_custodial):
                                xml_tag = 'STUDENT_UNIQUE_ID'
                                
                            child = ET.SubElement(row_element, xml_tag)
                            val = row[col_name]
                            
                            if xml_tag in ['STUDENT_UNIQUE_ID', 'PARENT_UNIQUE_ID']:
                                child.set('UNIQUE_ID', 'Y')
                                
                            if val == "":
                                child.set(f"{{{NS}}}nil", "true")
                            else:
                                child.text = str(val)
                                
                    tree = ET.ElementTree(root)
                    try:
                        ET.indent(tree, space="  ", level=0)
                    except AttributeError:
                        pass
                    
                    xml_buffer = io.BytesIO()
                    tree.write(xml_buffer, encoding="UTF-8", xml_declaration=True)
                    
                    individual_zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(individual_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        zipf.writestr(xml_filename, xml_buffer.getvalue())
                    
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
                    label="📦 Download " + item['zip_name'],
                    data=item['zip_data'],
                    file_name=item['zip_name'],
                    mime="application/zip",
                    key=f"btn_{item['zip_name']}", 
                    use_container_width=True
                )
