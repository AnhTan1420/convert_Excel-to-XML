import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import time
import io
import zipfile

# App Workspace Configuration
st.set_page_config(page_title="MOE Jordan", page_icon="⚙️", layout="centered")

st.title("MOE CONVERT EXCEL to XML(ZIP)")

# Initialize Session States to prevent download button data-loss on click
if "download_queue" not in st.session_state:
    st.session_state.download_queue = []
if "run_summary" not in st.session_state:
    st.session_state.run_summary = {"success_count": 0, "total_records": 0}

# Collapsible Documentation Matrix
with st.expander("System Guidelines & Target Interface Identifiers", expanded=False):
  st.markdown("""
    The processing engine matches files based on case-insensitive keyword tokens within the filename. 
    Each valid file type will generate its own dedicated, isolated ZIP archive.
    
    | Target System Interface | Accepted File Keyword | Generated Payload Prefix |
    | :--- | :--- | :--- |
    | **ID Mapping Profile** | `Mapping` | `FULL_SFS_ID_MAPPING_MK_*` |
    | **Basic Personal** | `Personal` | `FULL_SFS_BASIC_PERSONLA_MK_*` |
    | **Basic School** | `School` | `FULL_SFF_BASIC_SCHOOL_MK_*` |
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
    "PERSONAL": {"prefix": "FULL_SFS_BASIC_PERSONAL_MK", "interface": "STUDENT_Personal_INFO"},
    "PERSONLA": {"prefix": "FULL_SFS_BASIC_PERSONAL_MK", "interface": "STUDENT_Personal_INFO"},
    "SCHOOL": {"prefix": "FULL_SFF_BASIC_SCHOOL_MK", "interface": "STUDENT_School_INFO"}
}

if uploaded_files:
    st.info(f"**Stage Queue:** {len(uploaded_files)} file(s) loaded into session memory.")
    
    if st.button("Execute Split Transformation", type="primary", use_container_width=True):
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
                    
                    # CREATE ISOLATED ZIP ARCHIVE FOR THIS INDIVIDUAL FILE TYPE
                    individual_zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(individual_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        zipf.writestr(xml_filename, xml_buffer.getvalue())
                    
                    # Store generated data into persistent download queue
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
                st.warning(f"⏭️ **Ignored:** `{file_name}` does not match any known target rules.")
        
        # Save structural metadata for rendering
        st.session_state.run_summary = {
            "success_count": success_count,
            "total_records": total_records_processed
        }

    # Persistent UI Rendering Section (Executes seamlessly outside/after the button click process)
    if st.session_state.download_queue:
        st.markdown("---")
        st.markdown("### Execution Run Summary")
        col1, col2 = st.columns(2)
        with col1:
            st.metric(label="Generated Payloads", value=f"{st.session_state.run_summary['success_count']} File(s)")
        with col2:
            st.metric(label="Total Dataset Records", value=f"{st.session_state.run_summary['total_records']} Rows")
        
        
        st.markdown("### 📥 Payload Downloads")
        st.write("Download the respective ZIP configurations below:")
        
        # Display separate download block layout for each successfully generated prefix
        for item in st.session_state.download_queue:
            with st.container(border=True):
                st.markdown(f"🔹 **Source:** `{item['source_name']}` | **Total:** {item['records']} records")
                st.download_button(
                    label=f"📦 Download {item['zip_name']}",
                    data=item['zip_data'],
                    file_name=item['zip_name'],
                    mime="application/zip",
                    key=f"btn_{item['zip_name']}", # Prevent duplicate key widget collisions
                    use_container_width=True
                )
