<#
.SYNOPSIS
    Entfernt den Detection-Marker des Office Cache Cleaners.

.DESCRIPTION
    Loescht den HKCU Registry-Key, damit die App erneut ueber Intune deployed werden kann.

.NOTES
    EIFER IT-Tools
#>

$DetectionKey = "HKCU:\SOFTWARE\EIFER\OfficeCacheCleaner"

if (Test-Path $DetectionKey) {
    Remove-Item -Path $DetectionKey -Force
}

# Eltern-Key aufraeumen falls leer
$ParentKey = "HKCU:\SOFTWARE\EIFER"
if ((Test-Path $ParentKey) -and -not (Get-ChildItem -Path $ParentKey)) {
    Remove-Item -Path $ParentKey -Force
}

exit 0
