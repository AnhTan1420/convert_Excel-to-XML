import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import time
import io
import zipfile

# App Workspace Configuration
st.set_page_config(page_title="Test Data Engine", page_icon="⚙️", layout="centered")

st.title("MOE convert from XLSS to XML(ZIP)")
st.markdown("Streamline UAT setups by transforming staging Excel spreadsheets into fully schema-compliant interface XML payloads.")

# Collapsible Documentation Matrix to clean up the main view
with st.expander("📘 System Guidelines & Target Interface Identifiers", expanded=False):
    st.markdown("""
    The processing engine matches files based on case-insensitive keyword tokens within the filename. 
    Ensure your files adhere to the naming criteria below:
    
    | Target System Interface | Accepted File Keyword | Generated Payload Prefix |
    | :--- | :--- | :--- |
    | **ID Mapping** | `Mapping` | `FULL_SFS_ID_MAPPING_MK_*` |
    | **Basic Personal** | `Personal` | `FULL_SFS_BASIC_PERSONLA_MK_*` |
    | **Basic School** | `School` | `FULL_SFF_BASIC_SCHOOL_MK_*` |
    
    *Note: The system ignores Excel temporary ownership files (`~$...`). Only columns containing data rows bound to a valid `UNIQUE_ID` are parsed.*
    """)

st.markdown("### 📤 Source File Upload")
uploaded_files = st.file_uploader(
    "Drag and drop or browse Excel workbooks (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

# Rule Mapping Engine Configuration (Preserving backward compatibility tags)
MAPPING_RULES = {
    "MAPPING": {"prefix": "FULL_SFS_ID_MAPPING_MK", "interface": "STUDENT_ID_Mapping_INFO"},
    "PERSONAL": {"prefix": "FULL_SFS_BASIC_PERSONLA_MK", "interface": "STUDENT_Personal_INFO"},
    "PERSONLA": {"prefix": "FULL_SFS_BASIC_PERSONLA_MK", "interface": "STUDENT_Personal_INFO"}, # Legacy typo support fallback
    "SCHOOL": {"prefix": "FULL_SFF_BASIC_SCHOOL_MK", "interface": "STUDENT_School_INFO"}
}

if uploaded_files:
    st.info(f"📋 **Stage Queue:** {len(uploaded_files)} file(s) loaded into session memory.")
    
    if st.button("🚀 Execute Transformation", type="primary", use_container_width=True):
        zip_buffer = io.BytesIO()
        current_time = time.strftime('%Y%m%d%H%M%S')
        success_count = 0
        total_records_processed = 0
        
        # Open in-memory compilation pipeline
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
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
                    
                    try:
                        # Extract data elements
                        df = pd.read_excel(uploaded_file, dtype=str, engine='openpyxl')
                        df = df.where(pd.notnull(df), None)
                        
                        # Initialize XML Element Tree Structure
                        NS = 'http://www.w3.org/2001/XMLSchema-instance'
                        ET.register_namespace('xs', NS)
                        
                        root = ET.Element('INTERFACE', {
                            'INTERFACE_NAME': interface,
                            'FILE_CREATED_TIME': str(int(time.time() * 1000)),
                            'FILE_NAME': xml_filename, 
                            'NO_RECORD': str(len(df)) 
                        })
                        
                        for index, row in df.iterrows():
                            if 'UNIQUE_ID' not in df.columns or row['UNIQUE_ID'] is None:
                                continue
                                
                            mapping = ET.SubElement(root, 'ID_Mapping', {'UNIQUE_ID': str(row['UNIQUE_ID'])})
                            for col_name in df.columns:
                                if col_name == 'UNIQUE_ID':
                                    continue
                                child = ET.SubElement(mapping, col_name)
                                val = row[col_name]
                                if val is None:
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
                        
                        # Add compiled asset into archive matrix
                        zipf.writestr(xml_filename, xml_buffer.getvalue())
                        st.success(f"✔️ **Processed:** `{file_name}` ➔ `{xml_filename}` ({len(df)} records)")
                        
                        success_count += 1
                        total_records_processed += len(df)
                        
                    except Exception as e:
                        st.error(f"❌ **Pipeline Failure** on `{file_name}`: {e}")
                else:
                    st.warning(f"⏭️ **Ignored:** `{file_name}` does not match any known target rules.")
        
        # Render a clean executive dashboard post-run
        if success_count > 0:
            st.markdown("---")
            st.markdown("### 📊 Execution Run Summary")
            col1, col2 = st.columns(2)
            with col1:
                st.metric(label="Generated Payloads", value=f"{success_count} / {len(uploaded_files)}")
            with col2:
                st.metric(label="Total Dataset Records", value=f"{total_records_processed} Rows")
            
            st.balloons()
            
            # Master file bundle extraction terminal
            st.download_button(
                label="📥 DOWNLOAD PACKAGED PAYLOADS (.ZIP)",
                data=zip_buffer.getvalue(),
                file_name=f"{prefix}_{current_time}.zip",
                mime="application/zip",
                use_container_width=True
            )
