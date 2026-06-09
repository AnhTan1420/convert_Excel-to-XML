import xml.etree.ElementTree as ET
import pandas as pd
import time

# Declare the namespace used in the XML file
NS = 'http://www.w3.org/2001/XMLSchema-instance'
ET.register_namespace('xs', NS)

def xml_to_xlsx(xml_file_path, xlsx_file_path):
    """
    Reads an XML file and exports the data into an Excel file (.xlsx)
    """
    print(f"🔄 Converting {xml_file_path} to {xlsx_file_path}...")
    tree = ET.parse(xml_file_path)
    root = tree.getroot()
    
    data = []
    
    # Loop through all ID_Mapping tags
    for mapping in root.findall('ID_Mapping'):
        row = {'UNIQUE_ID': mapping.get('UNIQUE_ID')} # Get the UNIQUE_ID attribute
        
        # Iterate through the child tags inside (STUDENT_UIN, PARENT_UIN, etc.)
        for child in mapping:
            # Check if this tag has the xs:nil="true" attribute
            is_nil = child.get(f"{{{NS}}}nil")
            
            if is_nil == 'true':
                row[child.tag] = None
            else:
                row[child.tag] = child.text
                
        data.append(row)
        
    # Load data into a Pandas DataFrame and export to Excel
    df = pd.DataFrame(data)
    df.to_excel(xlsx_file_path, index=False)
    print("✅ Export successful!")


def xlsx_to_xml(xlsx_file_path, xml_file_path):
    """
    Reads an Excel file and regenerates the XML file while maintaining the standard structure
    """
    print(f"🔄 Converting {xlsx_file_path} to {xml_file_path}...")
    df = pd.read_excel(xlsx_file_path)
    
    # Convert empty cells (NaN in Pandas) to Python's None type
    df = df.where(pd.notnull(df), None)
    
    # Create the root <INTERFACE> element with its mandatory attributes
    root = ET.Element('INTERFACE', {
        'INTERFACE_NAME': 'STUDENT_ID_Mapping_INFO',
        'FILE_CREATED_TIME': str(int(time.time() * 1000)), # Generate the current timestamp
        'FILE_NAME': xml_file_path.split('/')[-1], # Automatically extract the filename
        'NO_RECORD': str(len(df)) # Count the actual total number of records
    })
    
    # Create ID_Mapping tags from the Excel data
    for index, row in df.iterrows():
        mapping = ET.SubElement(root, 'ID_Mapping', {'UNIQUE_ID': str(row['UNIQUE_ID'])})
        
        # Automatically generate child tags based on the Excel column names
        for col_name in df.columns:
            if col_name == 'UNIQUE_ID':
                continue # Skip this column since it was already used as an attribute
                
            child = ET.SubElement(mapping, col_name)
            val = row[col_name]
            
            if val is None:
                # If there is no data, set the attribute xs:nil = "true"
                child.set(f"{{{NS}}}nil", "true")
            else:
                # Cast float values back to integers if affected by Excel's decimal format
                if isinstance(val, float) and val.is_integer():
                    val = int(val)
                child.text = str(val)
                
    # Write to the XML file
    tree = ET.ElementTree(root)
    
    # Format (Indent) the XML file for readability (Requires Python >= 3.9)
    try:
        ET.indent(tree, space="  ", level=0)
    except AttributeError:
        pass # Skip the indent step if running an older Python version
        
    tree.write(xml_file_path, encoding="UTF-8", xml_declaration=True)
    print("✅ XML generation successful!")

# ==================================
# EXECUTION ZONE (RUNNING CODE)
# ==================================
if __name__ == "__main__":
    xml_input = "FULL_SFS_ID_MAPPING_20260115193401.xml"
    excel_output = "ID_Mapping_TestData.xlsx"
    xml_generated = "GENERATED_ID_MAPPING.xml"
    
    # 1. Run test to convert from XML to Excel
    xml_to_xlsx(xml_input, excel_output)
    
    # 2. Read the newly created Excel file and convert it back to XML
    xlsx_to_xml(excel_output, xml_generated)