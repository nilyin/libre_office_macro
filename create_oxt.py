#!/usr/bin/env python3
"""
Script to package LibreOffice macros into an .oxt extension file.
"""
import os
import shutil
import tempfile
import zipfile
import re

# --- Configuration ---
OXT_FILENAME = "DocExport.oxt"
LIBRARY_NAME = "DocExport"
EXTENSION_IDENTIFIER = "com.github.bda-user.docexport.ext"
EXTENSION_VERSION = "1.0.0"
EXTENSION_DISPLAY_NAME = "Document Export Macros"

# List of .bas files to include in the extension
BAS_FILES = [
    "DocModel.bas",
    "DocView.bas",
    "ViewHfm.bas",
    "ViewHtml.bas", # Assuming this should be included
    "Utils.bas",
    "mMath.bas",
    "vLatex.bas"
]

def create_oxt():
    """Builds and packages the .oxt extension file."""
    with tempfile.TemporaryDirectory() as temp_dir:
        print(f"Using temporary directory: {temp_dir}")

        # 1. Create directory structure
        basic_lib_path = os.path.join(temp_dir, "Basic", LIBRARY_NAME)
        meta_inf_path = os.path.join(temp_dir, "META-INF")
        os.makedirs(basic_lib_path)
        os.makedirs(meta_inf_path)

        # 2. Generate description.xml
        description_xml_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<description xmlns="http://openoffice.org/extensions/description/2006"
    xmlns:d="http://openoffice.org/extensions/description/2006"
    xmlns:xlink="http://www.w3.org/1999/xlink">
    <identifier value="{EXTENSION_IDENTIFIER}" />
    <version value="{EXTENSION_VERSION}" />
    <display-name><name lang="en">{EXTENSION_DISPLAY_NAME}</name></display-name>
</description>
""".replace('\n', '\r\n')
        with open(os.path.join(temp_dir, "description.xml"), "w", encoding="utf-8") as f:
            f.write(description_xml_content)

        # 3. Prepare manifest entries and create macro XML files
        manifest_file_entries = ""
        script_xlb_modules = ""

        for bas_file in BAS_FILES:
            module_name = os.path.splitext(bas_file)[0]
            script_xlb_modules += f' <library:module library:name="{module_name}"/>\n'

            if os.path.exists(bas_file):
                with open(bas_file, 'r', encoding='utf-8') as f:
                    bas_content = f.read()
            else:
                print(f"Error: Required source file '{bas_file}' not found. Aborting.")
                return

            xml_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="{module_name}" script:language="StarBasic">
{bas_content}</script:module>
""".replace('\n', '\r\n')
            # The path for the individual module xml
            module_xml_path = os.path.join(basic_lib_path, f"{module_name}.xml")
            with open(module_xml_path, "w", encoding="utf-8") as f:
                f.write(xml_content)
            print(f"Packaged {bas_file} into Basic/{LIBRARY_NAME}/{module_name}.xml")

        # 4. Generate META-INF/manifest.xml using the collected entries
        manifest_xml_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE manifest:manifest PUBLIC "-//OpenOffice.org//DTD Manifest 1.0//EN" "Manifest.dtd">
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:media-type="application/vnd.sun.xml.uno-description;type=OpenOffice-Extension"
                       manifest:full-path="description.xml"/>
  <manifest:file-entry manifest:media-type="application/vnd.sun.xml.script;type=StarBasic" manifest:full-path="Basic/{LIBRARY_NAME}/script.xlb"/>
  <manifest:file-entry manifest:media-type="application/vnd.sun.xml.script;type=StarBasic" manifest:full-path="Basic/{LIBRARY_NAME}/dialog.xlb"/>
</manifest:manifest>
""".replace('\n', '\r\n')
        with open(os.path.join(meta_inf_path, "manifest.xml"), "w", encoding="utf-8") as f:
            f.write(manifest_xml_content)

        # 5. Create library files (script.xlb and dialog.xlb)
        script_xlb_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="{LIBRARY_NAME}" library:readonly="false" library:passwordprotected="false">
{script_xlb_modules}
</library:library>
""".replace('\n', '\r\n')
        with open(os.path.join(basic_lib_path, "script.xlb"), "w", encoding="utf-8") as f:
            f.write(script_xlb_content)

        dialog_xlb_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="{LIBRARY_NAME}" library:readonly="false" library:passwordprotected="false"/>
"""
        dialog_xlb_content = dialog_xlb_content.replace('\n', '\r\n')
        with open(os.path.join(basic_lib_path, "dialog.xlb"), "w", encoding="utf-8") as f:
            f.write(dialog_xlb_content)

        # 6. Create the OXT (ZIP) file
        if os.path.exists(OXT_FILENAME):
            os.remove(OXT_FILENAME)
            
        with zipfile.ZipFile(OXT_FILENAME, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_path = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, arc_path)
        
        print(f"\nSuccessfully created extension: {OXT_FILENAME}")

if __name__ == "__main__":
    create_oxt()