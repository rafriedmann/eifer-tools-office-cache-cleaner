<#
.SYNOPSIS
    Clears the Microsoft 365 Office cache to fix issues with pinned documents.

.DESCRIPTION
    Deletes Office MRU cache, Jump Lists, and relevant registry entries.
    Designed for Intune Win32 app deployment in user context.

.NOTES
    EIFER IT-Tools
    Requires: PowerShell 5.1+, user context execution
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Continue"

# --- Toast Notification ---
function Show-Toast {
    param(
        [string]$Title,
        [string]$Message
    )
    try {
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
        [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null

        $Template = @"
<toast>
    <visual>
        <binding template="ToastGeneric">
            <text>$Title</text>
            <text>$Message</text>
        </binding>
    </visual>
</toast>
"@
        $Xml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $Xml.LoadXml($Template)
        $Toast = [Windows.UI.Notifications.ToastNotification]::new($Xml)
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("EIFER IT-Tools").Show($Toast)
    } catch {
        Write-Log "Failed to show toast notification: $_"
    }
}

# --- Logging ---
$LogDir = "$env:LOCALAPPDATA\EIFER\Logs"
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
$LogFile = Join-Path $LogDir "Clear-OfficeCache_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$Message)
    $Entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    $Entry | Out-File -Append -FilePath $LogFile -Encoding utf8
    Write-Host $Entry
}

Write-Log "=== Office Cache Cleaner started ==="
Write-Log "User: $env:USERNAME | Computer: $env:COMPUTERNAME"

# --- Close Office processes ---
$OfficeProcesses = @("WINWORD", "EXCEL", "POWERPNT", "OUTLOOK", "ONENOTE", "MSACCESS", "MSPUB")
$RunningOffice = Get-Process -Name $OfficeProcesses -ErrorAction SilentlyContinue

if ($RunningOffice) {
    Write-Log "Running Office apps found: $($RunningOffice.Name -join ', '). Attempting to close..."
    foreach ($proc in $RunningOffice) {
        try {
            $proc.CloseMainWindow() | Out-Null
        } catch {
            Write-Log "WARNING: Could not close $($proc.Name) via MainWindow."
        }
    }
    Start-Sleep -Seconds 5

    # Force-kill remaining processes
    $StillRunning = Get-Process -Name $OfficeProcesses -ErrorAction SilentlyContinue
    if ($StillRunning) {
        Write-Log "Force-closing: $($StillRunning.Name -join ', ')"
        $StillRunning | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
    Write-Log "Office apps closed."
} else {
    Write-Log "No Office apps running."
}

# --- 1. Office MRU / Recent Files Cache ---
$OfficePaths = @(
    "$env:APPDATA\Microsoft\Office\Recent",
    "$env:APPDATA\Microsoft\Office\Recent\AutoRecoverSave"
)

foreach ($path in $OfficePaths) {
    if (Test-Path $path) {
        $count = (Get-ChildItem -Path $path -File -ErrorAction SilentlyContinue).Count
        Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Cleaned: $path ($count files)"
    } else {
        Write-Log "Not found (skipped): $path"
    }
}

# --- 2. Office File Cache (local cache for cloud documents) ---
$OfficeFileCache = "$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache"
if (Test-Path $OfficeFileCache) {
    $count = (Get-ChildItem -Path $OfficeFileCache -File -Recurse -ErrorAction SilentlyContinue).Count
    Remove-Item -Path "$OfficeFileCache\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Log "Cleaned: $OfficeFileCache ($count files)"
} else {
    Write-Log "Not found (skipped): $OfficeFileCache"
}

# --- 3. Office Upload Center Cache ---
$UploadCenter = "$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeUploadCenter"
if (Test-Path $UploadCenter) {
    Remove-Item -Path "$UploadCenter\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Log "Cleaned: $UploadCenter"
}

# --- 4. Windows Jump Lists (Pinned + Recent in Taskbar/Explorer) ---
$JumpListPaths = @(
    "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations",
    "$env:APPDATA\Microsoft\Windows\Recent\CustomDestinations"
)

foreach ($path in $JumpListPaths) {
    if (Test-Path $path) {
        $count = (Get-ChildItem -Path $path -File -ErrorAction SilentlyContinue).Count
        Remove-Item -Path "$path\*" -Force -ErrorAction SilentlyContinue
        Write-Log "Cleaned: $path ($count files)"
    }
}

# --- 5. Registry: Office MRU entries ---
$OfficeApps = @("Word", "Excel", "PowerPoint", "Access", "Publisher")
$MRUBasePath = "HKCU:\Software\Microsoft\Office\16.0"

foreach ($app in $OfficeApps) {
    $mruPath = "$MRUBasePath\$app\File MRU"
    if (Test-Path $mruPath) {
        Remove-Item -Path $mruPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Registry cleaned: $mruPath"
    }

    $placeMru = "$MRUBasePath\$app\Place MRU"
    if (Test-Path $placeMru) {
        Remove-Item -Path $placeMru -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Registry cleaned: $placeMru"
    }

    $userMru = "$MRUBasePath\$app\User MRU"
    if (Test-Path $userMru) {
        Remove-Item -Path $userMru -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Registry cleaned: $userMru"
    }
}

# --- 6. Office Roaming data for MRU ---
$RoamingPath = "$MRUBasePath\Common\Roaming"
if (Test-Path $RoamingPath) {
    Remove-Item -Path $RoamingPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Log "Registry cleaned: $RoamingPath"
}

# --- Detection marker (for Intune) ---
$DetectionKey = "HKCU:\SOFTWARE\EIFER\OfficeCacheCleaner"
if (-not (Test-Path $DetectionKey)) {
    New-Item -Path $DetectionKey -Force | Out-Null
}
Set-ItemProperty -Path $DetectionKey -Name "LastRun" -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Set-ItemProperty -Path $DetectionKey -Name "Version" -Value "1.0"
Write-Log "Detection marker set: $DetectionKey"

Write-Log "=== Office Cache Cleaner completed ==="

Show-Toast -Title "EIFER Office Cache Cleaner" -Message "Office cache has been cleared successfully. Please restart your Office applications."

exit 0
