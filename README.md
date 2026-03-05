<p align="center">
  <img src="icon.svg" alt="Office Cache Cleaner" width="128">
</p>

# Office Cache Cleaner

PowerShell script to clear the Microsoft 365 Office cache. Fixes issues with pinning documents in Word, Excel, Explorer, etc.

## What gets cleaned?

| Area | Path / Location |
|------|----------------|
| Office Recent Files | `%APPDATA%\Microsoft\Office\Recent\` |
| Office File Cache | `%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache\` |
| Office Upload Center | `%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeUploadCenter\` |
| Windows Jump Lists | `%APPDATA%\Microsoft\Windows\Recent\*Destinations\` |
| Registry File MRU | `HKCU:\Software\Microsoft\Office\16.0\<App>\File MRU` |
| Registry Place MRU | `HKCU:\Software\Microsoft\Office\16.0\<App>\Place MRU` |
| Registry User MRU | `HKCU:\Software\Microsoft\Office\16.0\<App>\User MRU` |
| Registry Roaming | `HKCU:\Software\Microsoft\Office\16.0\Common\Roaming` |

**Note:** The script automatically closes running Office apps before cleaning.

## Manual execution

```powershell
powershell.exe -ExecutionPolicy Bypass -File Clear-OfficeCache.ps1
```

## Intune Deployment (Win32 App)

### 1. Create .intunewin package

Use the [Microsoft Win32 Content Prep Tool](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool):

```cmd
IntuneWinAppUtil.exe -c .\  -s Clear-OfficeCache.ps1 -o .\output
```

### 2. Configure in Intune

| Setting | Value |
|---------|-------|
| **Install command** | `powershell.exe -ExecutionPolicy Bypass -File Clear-OfficeCache.ps1` |
| **Uninstall command** | `powershell.exe -ExecutionPolicy Bypass -File Uninstall-OfficeCache.ps1` |
| **Install behavior** | User |
| **Detection rule** | Registry: `HKCU\SOFTWARE\EIFER\OfficeCacheCleaner`, Value: `Version`, String equals `1.0` |

### 3. Assignment

- Assign as **Required** to affected users or device groups
- Or publish as **Available** in the Company Portal for self-service

## Logs

Log files are written to:

```
%LOCALAPPDATA%\EIFER\Logs\Clear-OfficeCache_<timestamp>.log
```
