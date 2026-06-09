import xml.etree.ElementTree as ET
import pandas as pd
import time
import os
import zipfile
import sys
import glob

def compress_file_to_zip(file_path, zip_path):
    print(f"📦 Compressing {file_path} into {zip_path}...")
    if not os.path.exists(file_path):
        print(f"❌ Error: Source file {file_path} not found for compression!")
        return False
    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.write(file_path, arcname=os.path.basename(file_path))
        print(f"✅ Zip archive created successfully: {zip_path}\n")
        return True
    except Exception as e:
        print(f"❌ Error during compression: {e}\n")
        return False


def generate_xml_from_excel(xlsx_path, xml_output_path, interface_name):
    NS = 'http://www.w3.org/2001/XMLSchema-instance'
    ET.register_namespace('xs', NS)
    
    print(f"🔄 Reading data from Excel: {xlsx_path}...")
    df = pd.read_excel(xlsx_path, dtype=str, engine='openpyxl')
    df = df.where(pd.notnull(df), None)
    
    actual_file_name = os.path.basename(xml_output_path)
    
    # Cấu hình root element dynamic theo loại INTERFACE_NAME
    root = ET.Element('INTERFACE', {
        'INTERFACE_NAME': interface_name,
        'FILE_CREATED_TIME': str(int(time.time() * 1000)),
        'FILE_NAME': actual_file_name, 
        'NO_RECORD': str(len(df)) 
    })
    
    for index, row in df.iterrows():
        if 'UNIQUE_ID' not in df.columns or row['UNIQUE_ID'] is None:
            print(f"⚠️ Warning: Row {index + 2} is missing UNIQUE_ID. Skipping.")
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
        
    tree.write(xml_output_path, encoding="UTF-8", xml_declaration=True)
    print(f"🎉 XML exported: {xml_output_path}")
    print(f"📊 Total records (NO_RECORD): {len(df)}")


if __name__ == "__main__":
    try:
        print("🔍 Scanning directory for Excel test data files...")
        # Lấy tất cả các file .xlsx trong thư mục (bỏ qua file tạm của Excel dạng ~$...)
        all_files = glob.glob("*.xlsx")
        excel_files = [f for f in all_files if not f.startswith("~$")]
        
        if not excel_files:
            print("❌ No Excel (.xlsx) files found in the current directory!")
            sys.exit(1)
            
        # Định nghĩa luật ánh xạ: Từ khóa trong tên file -> (Prefix Đầu ra, Tên Cổng Giao Tiếp)
        MAPPING_RULES = {
            "MAPPING": {"prefix": "FULL_SFS_ID_MAPPING_MK", "interface": "STUDENT_ID_Mapping_INFO"},
            "PERSONAL": {"prefix": "FULL_SFS_BASIC_PERSONAL_MK", "interface": "STUDENT_Personal_INFO"},
            "SCHOOL": {"prefix": "FULL_SFF_BASIC_SCHOOL_MK", "interface": "STUDENT_School_INFO"}
        }
        
        processed_count = 0
        current_time = time.strftime('%Y%m%d%H%M%S')
        
        for file_name in excel_files:
            matched_rule = None
            # Kiểm tra xem tên file chứa từ khóa nào để áp luật
            for key, rule in MAPPING_RULES.items():
                if key.upper() in file_name.upper():
                    matched_rule = rule
                    break
            
            if matched_rule:
                print(f"\n🎯 Found matching file: '{file_name}'")
                prefix = matched_rule["prefix"]
                interface = matched_rule["interface"]
                
                # Khởi tạo tên file XML và ZIP động theo chuẩn thời gian real-time
                xml_output = f"{prefix}_{current_time}.xml"
                zip_output = f"{prefix}_{current_time}.zip"
                
                # Tiến hành chuyển đổi và nén
                generate_xml_from_excel(file_name, xml_output, interface)
                compress_file_to_zip(xml_output, zip_output)
                processed_count += 1
            else:
                print(f"⏭️ Skipped: '{file_name}' (Does not match any defined rule)")
                
        print(f"\n==========================================")
        print(f"✅ DONE! Successfully processed {processed_count} file(s).")
        
    except Exception as e:
        print(f"❌ CRITICAL RUNTIME ERROR: {e}")
        sys.exit(1)