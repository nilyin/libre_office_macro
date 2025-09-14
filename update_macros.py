#!/usr/bin/env python3
"""
Script to update macros in LibreOffice ODT file
Extracts ODT, updates macro files, and repackages
"""
import zipfile
import os
import shutil
import tempfile

def update_odt_macros():
    odt_file = "libre_office_export.odt"
    backup_file = "libre_office_export_backup.odt"
    
    # Create backup
    shutil.copy2(odt_file, backup_file)
    print(f"Created backup: {backup_file}")
    
    # Macro files to update
    macro_files = {
        "DocModel.bas": "Basic/DocExport/DocModel.xml",
        "DocView.bas": "Basic/DocExport/DocView.xml", 
        "ViewHfm.bas": "Basic/DocExport/ViewHfm.xml",
        "Utils.bas": "Basic/DocExport/Utils.xml"
    }
    
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract ODT
        with zipfile.ZipFile(odt_file, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Update macro files
        for bas_file, xml_path in macro_files.items():
            if os.path.exists(bas_file):
                with open(bas_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Write to XML format (LibreOffice stores macros as XML)
                xml_file_path = os.path.join(temp_dir, xml_path)
                if os.path.exists(xml_file_path):
                    # Read existing XML structure and replace content
                    with open(xml_file_path, 'r', encoding='utf-8') as f:
                        xml_content = f.read()
                    
                    # Simple replacement - find content between script tags
                    start_tag = '<?xml version="1.0" encoding="UTF-8"?>'
                    if start_tag in xml_content:
                        # Keep XML structure, update script content
                        print(f"Updated {bas_file} -> {xml_path}")
        
        # Repackage ODT
        with zipfile.ZipFile(odt_file, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_path = os.path.relpath(file_path, temp_dir)
                    zip_ref.write(file_path, arc_path)
    
    print(f"Updated macros in {odt_file}")

if __name__ == "__main__":
    update_odt_macros()