# Office Cache Cleaner

PowerShell-Script zur Bereinigung des Microsoft 365 Office-Caches. Loest Probleme mit angepinnten Dokumenten in Word, Excel, Explorer etc.

## Was wird bereinigt?

| Bereich | Pfad / Ort |
|---------|-----------|
| Office Recent Files | `%APPDATA%\Microsoft\Office\Recent\` |
| Office File Cache | `%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache\` |
| Office Upload Center | `%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeUploadCenter\` |
| Windows Jump Lists | `%APPDATA%\Microsoft\Windows\Recent\*Destinations\` |
| Registry File MRU | `HKCU:\Software\Microsoft\Office\16.0\<App>\File MRU` |
| Registry Place MRU | `HKCU:\Software\Microsoft\Office\16.0\<App>\Place MRU` |
| Registry User MRU | `HKCU:\Software\Microsoft\Office\16.0\<App>\User MRU` |
| Registry Roaming | `HKCU:\Software\Microsoft\Office\16.0\Common\Roaming` |

**Hinweis:** Das Script schliesst laufende Office-Apps automatisch vor der Bereinigung.

## Manuell ausfuehren

```powershell
powershell.exe -ExecutionPolicy Bypass -File Clear-OfficeCache.ps1
```

## Intune Deployment (Win32-App)

### 1. .intunewin Paket erstellen

Das [Microsoft Win32 Content Prep Tool](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool) verwenden:

```cmd
IntuneWinAppUtil.exe -c .\  -s Clear-OfficeCache.ps1 -o .\output
```

### 2. In Intune konfigurieren

| Einstellung | Wert |
|------------|------|
| **Install command** | `powershell.exe -ExecutionPolicy Bypass -File Clear-OfficeCache.ps1` |
| **Uninstall command** | `powershell.exe -ExecutionPolicy Bypass -File Uninstall-OfficeCache.ps1` |
| **Install behavior** | User |
| **Detection rule** | Registry: `HKCU\SOFTWARE\EIFER\OfficeCacheCleaner`, Value: `Version`, String equals `1.0` |

### 3. Zuweisung

- Als **Required** dem betroffenen Benutzer oder einer Geraetegruppe zuweisen
- Oder als **Available** im Company Portal bereitstellen (Mitarbeiter kann es bei Bedarf selbst ausfuehren)

## Logs

Logdateien werden geschrieben nach:

```
%LOCALAPPDATA%\EIFER\Logs\Clear-OfficeCache_<timestamp>.log
```
