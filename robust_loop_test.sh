#!/bin/bash
echo "Starting robust file loop conversion..."

# Clean up any existing processes and files
pkill -f soffice 2>/dev/null
sleep 2
rm -f /tmp/*.md /tmp/macro_debug.log /tmp/libreoffice_debug.log
rm -rf /tmp/img_*

cd /tmp

# Process each file individually with proper cleanup
for odt_file in doc2.odt doc3.odt; do
    if [ -f "$odt_file" ]; then
        echo "Processing: $odt_file"
        
        # Kill any existing processes
        pkill -f soffice 2>/dev/null
        sleep 1
        
        # Run conversion
        soffice --headless --invisible --nologo --norestore "$odt_file" 'macro:///DocExport.DocModel.MakeDocHfmView'
        
        # Wait for completion
        sleep 5
        
        # Check results
        base_name=$(basename "$odt_file" .odt)
        if [ -f "${base_name}.md" ]; then
            echo "SUCCESS: ${base_name}.md created"
            ls -la "${base_name}.md"
        else
            echo "FAILED: ${base_name}.md not created"
        fi
        
        if [ -d "img_${base_name}" ]; then
            echo "SUCCESS: img_${base_name} folder created"
            ls -la "img_${base_name}/"
        else
            echo "FAILED: img_${base_name} folder not created"
        fi
        
        echo "---"
    fi
done

echo "Robust file loop conversion finished."