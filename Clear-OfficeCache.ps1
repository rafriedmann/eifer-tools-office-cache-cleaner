<#
.SYNOPSIS
    Bereinigt den Microsoft 365 Office-Cache um Probleme mit angepinnten Dokumenten zu beheben.

.DESCRIPTION
    Loescht Office MRU-Cache, Jump Lists und relevante Registry-Eintraege.
    Gedacht fuer Intune Win32-App Deployment im User-Kontext.

.NOTES
    EIFER IT-Tools
    Erfordert: PowerShell 5.1+, Ausfuehrung im User-Kontext
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Continue"

# --- Toast-Notification Funktion ---
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
        Write-Log "Toast-Notification konnte nicht angezeigt werden: $_"
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

Write-Log "=== Office Cache Cleaner gestartet ==="
Write-Log "Benutzer: $env:USERNAME | Computer: $env:COMPUTERNAME"

# --- Office-Prozesse pruefen und beenden ---
$OfficeProcesses = @("WINWORD", "EXCEL", "POWERPNT", "OUTLOOK", "ONENOTE", "MSACCESS", "MSPUB")
$RunningOffice = Get-Process -Name $OfficeProcesses -ErrorAction SilentlyContinue

if ($RunningOffice) {
    Write-Log "Offene Office-Apps gefunden: $($RunningOffice.Name -join ', '). Versuche zu schliessen..."
    foreach ($proc in $RunningOffice) {
        try {
            $proc.CloseMainWindow() | Out-Null
        } catch {
            Write-Log "WARNUNG: Konnte $($proc.Name) nicht ueber MainWindow schliessen."
        }
    }
    Start-Sleep -Seconds 5

    # Noch laufende Prozesse hart beenden
    $StillRunning = Get-Process -Name $OfficeProcesses -ErrorAction SilentlyContinue
    if ($StillRunning) {
        Write-Log "Erzwinge Beendigung von: $($StillRunning.Name -join ', ')"
        $StillRunning | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
    Write-Log "Office-Apps geschlossen."
} else {
    Write-Log "Keine Office-Apps geoeffnet."
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
        Write-Log "Bereinigt: $path ($count Dateien)"
    } else {
        Write-Log "Nicht vorhanden (uebersprungen): $path"
    }
}

# --- 2. Office File Cache (lokaler Cache fuer Cloud-Dokumente) ---
$OfficeFileCache = "$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache"
if (Test-Path $OfficeFileCache) {
    $count = (Get-ChildItem -Path $OfficeFileCache -File -Recurse -ErrorAction SilentlyContinue).Count
    Remove-Item -Path "$OfficeFileCache\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Log "Bereinigt: $OfficeFileCache ($count Dateien)"
} else {
    Write-Log "Nicht vorhanden (uebersprungen): $OfficeFileCache"
}

# --- 3. Office Upload Center Cache ---
$UploadCenter = "$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeUploadCenter"
if (Test-Path $UploadCenter) {
    Remove-Item -Path "$UploadCenter\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Log "Bereinigt: $UploadCenter"
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
        Write-Log "Bereinigt: $path ($count Dateien)"
    }
}

# --- 5. Registry: Office MRU Eintraege ---
$OfficeApps = @("Word", "Excel", "PowerPoint", "Access", "Publisher")
$MRUBasePath = "HKCU:\Software\Microsoft\Office\16.0"

foreach ($app in $OfficeApps) {
    # File MRU
    $mruPath = "$MRUBasePath\$app\File MRU"
    if (Test-Path $mruPath) {
        Remove-Item -Path $mruPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Registry bereinigt: $mruPath"
    }

    # Place MRU (Ordner-Favoriten)
    $placeMru = "$MRUBasePath\$app\Place MRU"
    if (Test-Path $placeMru) {
        Remove-Item -Path $placeMru -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Registry bereinigt: $placeMru"
    }

    # User MRU (Cloud-basierte MRU)
    $userMru = "$MRUBasePath\$app\User MRU"
    if (Test-Path $userMru) {
        Remove-Item -Path $userMru -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Registry bereinigt: $userMru"
    }
}

# --- 6. Office Roaming-Daten fuer MRU ---
$RoamingPath = "$MRUBasePath\Common\Roaming"
if (Test-Path $RoamingPath) {
    Remove-Item -Path $RoamingPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Log "Registry bereinigt: $RoamingPath"
}

# --- Detection-Marker setzen (fuer Intune) ---
$DetectionKey = "HKCU:\SOFTWARE\EIFER\OfficeCacheCleaner"
if (-not (Test-Path $DetectionKey)) {
    New-Item -Path $DetectionKey -Force | Out-Null
}
Set-ItemProperty -Path $DetectionKey -Name "LastRun" -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Set-ItemProperty -Path $DetectionKey -Name "Version" -Value "1.0"
Write-Log "Detection-Marker gesetzt: $DetectionKey"

Write-Log "=== Office Cache Cleaner abgeschlossen ==="

Show-Toast -Title "EIFER Office Cache Cleaner" -Message "Office-Cache wurde erfolgreich bereinigt. Bitte starten Sie Ihre Office-Anwendungen neu."

exit 0
