$ErrorActionPreference = 'Continue'

# Read manifest ID
$manifest = Get-Content 'appPackage/manifest.json' -Raw | ConvertFrom-Json
$manifestId = $manifest.id

Write-Host "Uninstalling add-in with Manifest ID: $manifestId"
Write-Host ""

# Uninstall via atk
npx -p @microsoft/m365agentstoolkit-cli atk uninstall --mode manifest-id --manifest-id $manifestId --interactive false --options m365-app

Write-Host ""

# Stop debugging
office-addin-debugging stop appPackage/manifest.json 2>$null

# Clear Office cache
$wefPath = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
if (Test-Path $wefPath) {
    Remove-Item -Recurse -Force $wefPath
    Write-Host "Office cache cleared"
}

Write-Host ""
Write-Host "Add-in uninstalled successfully!"
