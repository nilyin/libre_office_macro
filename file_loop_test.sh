#!/bin/bash
echo "Starting file loop conversion..."
for odt_file in *.odt; do
    if [ -f "$odt_file" ]; then
        echo "Processing: $odt_file"
        soffice --headless --invisible --nologo --norestore "$odt_file" 'macro:///DocExport.DocModel.MakeDocHfmView'
        sleep 2
        echo "Completed: $odt_file"
    fi
done
echo "File loop conversion finished."