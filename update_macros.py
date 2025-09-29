#!/usr/bin/env python3
"""
Script to update macros in LibreOffice ODT file
Extracts ODT, updates macro files, and repackages
"""
import zipfile
import re
import os
import shutil
import tempfile

def update_odt_macros():
    odt_file = "libre_office_export.odt"
    backup_file = "libre_office_export_backup.odt"
    
    # Check preconditions
    if not os.path.exists(odt_file):
        print(f"Error: {odt_file} not found")
        return False
    
    # Check if LibreOffice has the file locked
    lock_file = f".~lock.{odt_file}#"
    if os.path.exists(lock_file):
        print(f"Error: {odt_file} is currently open in LibreOffice")
        print("Please close the document in LibreOffice before running this script")
        return False
    
    # Test file write access
    try:
        with open(odt_file, 'r+b') as f:
            pass
    except PermissionError:
        print(f"Error: Cannot access {odt_file} - file may be open or locked")
        return False
    
    # Create backup
    shutil.copy2(odt_file, backup_file)
    print(f"Created backup: {backup_file}")
    
    # Automatically discover all .bas files in current directory
    macro_files = {}
    for file in os.listdir('.'):
        if file.endswith('.bas'):
            # Convert filename to XML path (remove .bas extension)
            module_name = file[:-4]  # Remove .bas extension
            xml_path = f"Basic/DocExport/{module_name}.xml"
            macro_files[file] = xml_path
    
    print(f"Found {len(macro_files)} macro files to update:")
    for bas_file in macro_files.keys():
        print(f"  - {bas_file}")
    print()
    
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract ODT
        with zipfile.ZipFile(odt_file, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Update macro files
        for bas_file, xml_path in macro_files.items():
            if os.path.exists(bas_file):
                print(f"Processing {bas_file}...")
                with open(bas_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Write to XML format (LibreOffice stores macros as XML)
                xml_file_path = os.path.join(temp_dir, xml_path)
                if os.path.exists(xml_file_path):
                    # Read original to get module name
                    with open(xml_file_path, 'r', encoding='utf-8') as f:
                        original = f.read()
                    
                    # Extract module name from original XML
                    module_match = re.search(r'script:name="([^"]+)"', original)
                    module_name = module_match.group(1) if module_match else module_name
                    
                    # Create new XML with exact LibreOffice format
                    xml_content = f'<?xml version="1.0" encoding="UTF-8"?>\r\n<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\r\n<script:module xmlns:script="http://openoffice.org/2000/script" script:name="{module_name}" script:language="StarBasic">{content}</script:module>'
                    
                    with open(xml_file_path, 'w', encoding='utf-8') as f:
                        f.write(xml_content)
                    print(f"Updated {bas_file} -> {xml_path}")
                else:
                    print(f"Warning: XML file not found for {bas_file} at {xml_path}")
            else:
                print(f"Warning: Source file {bas_file} not found")
        
        # Repackage ODT
        with zipfile.ZipFile(odt_file, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_path = os.path.relpath(file_path, temp_dir)
                    zip_ref.write(file_path, arc_path)
    
    print(f"Updated macros in {odt_file}")
    return True

if __name__ == "__main__":
    success = update_odt_macros()
    if success:
        print("\nMacro update completed successfully!")
        print("You can now open the ODT file in LibreOffice to verify the changes.")
    else:
        print("\nMacro update failed. Please check the error messages above.")
