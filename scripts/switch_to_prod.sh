#!/bin/bash

file="../manifest.xml"
sed -i 's|https://localhost:3000/|https://andreachecchi.github.io/outlook_vtiger/dist/|g' "$file"
echo "Sostituzione completata nel file $file"