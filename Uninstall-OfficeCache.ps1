<#
.SYNOPSIS
    Removes the Office Cache Cleaner detection marker.

.DESCRIPTION
    Deletes the HKCU registry key so the app can be redeployed via Intune.

.NOTES
    EIFER IT-Tools
#>

$DetectionKey = "HKCU:\SOFTWARE\EIFER\OfficeCacheCleaner"

if (Test-Path $DetectionKey) {
    Remove-Item -Path $DetectionKey -Force
}

# Clean up parent key if empty
$ParentKey = "HKCU:\SOFTWARE\EIFER"
if ((Test-Path $ParentKey) -and -not (Get-ChildItem -Path $ParentKey)) {
    Remove-Item -Path $ParentKey -Force
}

exit 0
