#!/bin/bash
filePath=./feature.ts
echo "Update feature settings in $filePath if alpha or preview release"
varNames=$(grep -o "const [a-zA-Z0-9_]\+" $filePath | grep -v "releaseControlledFeatureSettings" | sed 's/const //')

# Process each variable
for varName in $varNames; do
  # Check if current value is false
  if grep -q "$varName = false" $filePath; then
    echo "Setting $varName to true"
    sed -i -e "s@const $varName = false@const $varName = true@g" $filePath
  else
    echo "$varName is already set to true or has another value"
  fi
done

echo "Prerelease feature setting update done."
