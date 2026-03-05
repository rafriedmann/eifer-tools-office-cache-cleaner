<p align="center">
  <img src="icon.svg" alt="Office Cache Cleaner" width="128">
</p>

# Office Cache Cleaner

PowerShell script to clear the Microsoft 365 Office cache. Fixes issues with pinning documents in Word, Excel, Explorer, etc.

## What does this tool do?

This tool resets the Office "memory" of recently used and pinned files. After running it, Office apps behave as if opened for the first time regarding file history.

**What the user will notice:**
- The **"Recent Documents"** list in Word, Excel, PowerPoint etc. will be empty
- **Pinned documents** in Office apps will be removed (can be re-pinned afterwards)
- **Jump Lists** (right-click on taskbar icons) will be cleared
- The **"Recent files"** section in File Explorer will be reset
- **No documents, files, or personal data are deleted** - only the shortcuts/references to them

**What is NOT affected:**
- All actual files and documents remain untouched
- OneDrive / SharePoint sync is not affected
- Outlook emails, calendar, and contacts are not affected
- Office settings, templates, and add-ins remain unchanged

**Important:** The script will automatically close all running Office apps (Word, Excel, PowerPoint, Outlook, OneNote, Access, Publisher) before cleaning. Users should save their work beforehand.

## Technical details

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
