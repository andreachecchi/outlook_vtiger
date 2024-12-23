#!/bin/bash

file="../manifest.xml"
sed -i 's|https://andreachecchi.github.io/outlook_vtiger/dist/|https://localhost:3000/|g' "$file"
echo "Sostituzione completata nel file $file"