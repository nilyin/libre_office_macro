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
"""
        with open(os.path.join(temp_dir, "description.xml"), "w", encoding="utf-8") as f:
            f.write(description_xml_content)

        # 3. Generate META-INF/manifest.xml
        manifest_file_entries = '<manifest:file-entry manifest:media-type="application/vnd.sun.xml.uno-description;type=OpenOffice-Extension" manifest:full-path="description.xml"/>\n'
        manifest_file_entries += f'<manifest:file-entry manifest:media-type="application/vnd.sun.xml.script;type=StarBasic" manifest:full-path="Basic/{LIBRARY_NAME}/script.xlb"/>\n'
        manifest_file_entries += f'<manifest:file-entry manifest:media-type="application/vnd.sun.xml.script;type=StarBasic" manifest:full-path="Basic/{LIBRARY_NAME}/dialog.xlb"/>\n'

        # 4. Create macro XML files and add to manifest
        script_xlb_modules = ""
        for bas_file in BAS_FILES:
            module_name = os.path.splitext(bas_file)[0]
            xml_path = f"Basic/{LIBRARY_NAME}/{module_name}.xml"
            manifest_file_entries += f'<manifest:file-entry manifest:media-type="application/vnd.sun.xml.script;type=StarBasic" manifest:full-path="{xml_path}"/>\n'
            script_xlb_modules += f' <library:module library:name="{module_name}"/>\n'

            if os.path.exists(bas_file):
                with open(bas_file, 'r', encoding='utf-8') as f:
                    bas_content = f.read()
                
                # Escape CDATA closing sequence if present in the code
                bas_content = bas_content.replace("]]>", "]]&gt;")

                xml_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="{module_name}" script:language="StarBasic">
<script:source-code><![CDATA[{bas_content}]]></script:source-code>
</script:module>
"""
                with open(os.path.join(temp_dir, xml_path), "w", encoding="utf-8") as f:
                    f.write(xml_content)
                print(f"Packaged {bas_file} into {xml_path}")

        manifest_xml_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
{manifest_file_entries}
</manifest:manifest>
"""
        with open(os.path.join(meta_inf_path, "manifest.xml"), "w", encoding="utf-8") as f:
            f.write(manifest_xml_content)

        # 5. Create library files (script.xlb and dialog.xlb)
        script_xlb_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="{LIBRARY_NAME}" library:readonly="true" library:passwordprotected="false">
{script_xlb_modules}
</library:library>
"""
        with open(os.path.join(basic_lib_path, "script.xlb"), "w", encoding="utf-8") as f:
            f.write(script_xlb_content)

        dialog_xlb_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="{LIBRARY_NAME}" library:readonly="true" library:passwordprotected="false"/>
"""
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