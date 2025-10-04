#!/bin/bash

# LibreOffice ODT to Markdown Batch Converter
# Usage: ./convert_batch.sh [input_dir] [output_dir]

INPUT_DIR="${1:-.}"
OUTPUT_DIR="${2:-converted_docs}"
CONVERTED=0
FAILED=0

echo "=== LibreOffice Batch Conversion ==="
echo "Input: $INPUT_DIR"
echo "Output: $OUTPUT_DIR"

# Create output directory
mkdir -p "$OUTPUT_DIR"

# Process each ODT file
for odt_file in "$INPUT_DIR"/*.odt; do
    [ ! -f "$odt_file" ] && continue
    
    filename=$(basename "$odt_file" .odt)
    echo "Converting: $filename"
    
    # Kill any existing LibreOffice processes
    pkill -f soffice 2>/dev/null || true
    sleep 1
    
    # Convert single file
    if timeout 60 soffice --headless --invisible --nologo --norestore "$odt_file" 'macro:///DocExport.DocModel.MakeDocHfmView'; then
        # Move generated files
        if [ -f "${odt_file%.*}.md" ]; then
            mv "${odt_file%.*}.md" "$OUTPUT_DIR/$filename.md"
            echo "✓ $filename.md"
            ((CONVERTED++))
            
            # Move image folder if exists
            img_dir="${INPUT_DIR}/img_$filename"
            if [ -d "$img_dir" ]; then
                mv "$img_dir" "$OUTPUT_DIR/"
                echo "✓ img_$filename/"
            fi
        else
            echo "✗ Failed: $filename"
            ((FAILED++))
        fi
    else
        echo "✗ Timeout: $filename"
        ((FAILED++))
    fi
done

echo "=== Summary ==="
echo "Converted: $CONVERTED"
echo "Failed: $FAILED"
echo "Output: $OUTPUT_DIR"