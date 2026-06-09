import xml.etree.ElementTree as ET
import pandas as pd
import time
import os
import zipfile

def compress_file_to_zip(file_path, zip_path):
    print(f"4. Compressing {file_path} into {zip_path}...")
    if not os.path.exists(file_path):
        print(f"Error: Source file {file_path} not found for compression!")
        return False
    
    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # os.path.basename prevents storing the absolute folder structure inside the zip
            zipf.write(file_path, arcname=os.path.basename(file_path))
        print(f"5. Zip archive created successfully: {zip_path}")
        return True
    except Exception as e:
        print(f"Error during compression: {e}")
        return False


def generate_xml_from_excel(xlsx_path, xml_output_path):
    # 1. Define the Namespace identical to the sample file
    NS = 'http://www.w3.org/2001/XMLSchema-instance'
    ET.register_namespace('xs', NS)
    
    print(f"1. Reading data from Excel file: {xlsx_path}...")
    if not os.path.exists(xlsx_path):
        print(f"Error: File {xlsx_path} not found!")
        return
        
    # Read Excel file, casting all columns to string to avoid Excel numeric formatting issues
    df = pd.read_excel(xlsx_path, dtype=str)
    
    # Replace empty cells (NaN in Pandas) with Python's None type
    df = df.where(pd.notnull(df), None)
    
    # Extract only the actual file name from the full output path
    actual_file_name = os.path.basename(xml_output_path)
    
    # 2. Initialize the root <INTERFACE> element with dynamic attributes
    root = ET.Element('INTERFACE', {
        'INTERFACE_NAME': 'STUDENT_ID_Mapping_INFO',
        'FILE_CREATED_TIME': str(int(time.time() * 1000)), # Current Epoch timestamp in milliseconds
        'FILE_NAME': actual_file_name, 
        'NO_RECORD': str(len(df)) 
    })
    
    # 3. Iterate through each row in Excel to construct child tags
    for index, row in df.iterrows():
        if row['UNIQUE_ID'] is None:
            print(f"Warning: Row {index + 2} is missing UNIQUE_ID, skipping this row.")
            continue
            
        # Create <ID_Mapping UNIQUE_ID="..."> tag
        mapping = ET.SubElement(root, 'ID_Mapping', {'UNIQUE_ID': str(row['UNIQUE_ID'])})
        
        # Automatically generate child tags based on the remaining Excel columns
        for col_name in df.columns:
            if col_name == 'UNIQUE_ID':
                continue # Already handled above
                
            child = ET.SubElement(mapping, col_name)
            val = row[col_name]
            
            if val is None:
                # If Excel cell is empty -> generate xs:nil="true" attribute
                child.set(f"{{{NS}}}nil", "true")
            else:
                child.text = str(val)
                
    # 4. Write data to XML file and indent for better readability
    tree = ET.ElementTree(root)
    try:
        ET.indent(tree, space="  ", level=0) # Requires Python >= 3.9
    except AttributeError:
        pass
        
    tree.write(xml_output_path, encoding="UTF-8", xml_declaration=True)
    print(f"2. XML file exported successfully: {xml_output_path}")
    print(f"3. Total processed records (NO_RECORD): {len(df)}")

if __name__ == "__main__":
    # Configure input Excel file path
    excel_input = "ID_Mapping_TestData.xlsx"
    
    # 1. Get current real-time timestamp (YYYYMMDDHHMMSS)
    current_time = time.strftime('%Y%m%d%H%M%S')
    
    # 2. Generate dynamic real-time filenames for both XML and ZIP
    xml_output = f"FULL_SFS_ID_MAPPING_MK_{current_time}.xml"
    zip_output = f"FULL_SFS_ID_MAPPING_MK_{current_time}.zip"
    
    # Step 1: Generate the XML file from Excel data
    generate_xml_from_excel(excel_input, xml_output)
    
    # Step 2: Compress the newly generated XML file into a ZIP archive
    if os.path.exists(xml_output):
        compress_file_to_zip(xml_output, zip_output)