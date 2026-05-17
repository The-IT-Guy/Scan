# Scan

> **Windows System Health & Diagnostic Scripts** by The IT Guy

A collection of PowerShell scripts for diagnosing, auditing, repairing, and optimizing Windows 11 systems. Includes both standalone one-shot scripts and a full interactive menu-driven utility (`MasterStack`). All output is rendered as styled HTML reports that open automatically in your browser.

Licensed under the MIT License.

---

## Table of Contents

- [Scripts Overview](#scripts-overview)
- [MasterStack — Interactive Utility](#masterstack--interactive-utility)
- [Standalone Scripts](#standalone-scripts)
- [Debloat & Revert](#debloat--revert)
- [Requirements](#requirements)
- [Usage](#usage)
- [Output](#output)
- [Security Note](#security-note)

---

## Scripts Overview

| Script / File                  | Type          | What It Does                                          |
|-------------------------------|---------------|-------------------------------------------------------|
| `MasterStack`                 | Interactive   | Full 14-tool diagnostic and repair console menu       |
| `PCHealthScan.ps1`            | Standalone    | Comprehensive scan: OS, disks, RAM, CHKDSK, events    |
| `SoftwareScan.ps1`            | Standalone    | Installed software inventory + winget update check    |
| `HardwareScan.ps1`            | Standalone    | Physical disk health and RAM slot details             |
| `BatteryHealth.ps1`           | Standalone    | Battery wear report via `powercfg /batteryreport`     |
| `OSErrors.ps1`                | Standalone    | SFC verify + System event log critical errors         |
| `MASTER_DEBLOAT_IT_GUY.ps1`  | System mod    | 6-step Windows 11 debloat and privacy hardening       |
| `REVERT_IT_GUY_CHANGES.ps1`  | System mod    | Reverses every change made by the debloat script      |
| `MasterAudit`                 | Draft/excerpt | `Invoke-FullAuditReport` function (no entry point)    |

---

## MasterStack — Interactive Utility

`MasterStack` is the flagship tool. Run it and you get a numbered console menu with 14 options across three categories.

### Diagnostics & Audits

| # | Option                        | What It Does                                                                 |
|---|-------------------------------|------------------------------------------------------------------------------|
| 1 | System Health Check           | OS info, disk usage, live CPU load, RAM usage                                |
| 2 | Run Ultimate System Audit     | All-in-one report: OS, hardware, storage, battery, network, crashes, AV, updates |
| 3 | Crash Diagnostics             | System event log — Critical and Error events from the last 30 days           |
| 4 | Startup Programs Analyzer     | Lists all programs in `HKLM` and `HKCU` Run registry keys                   |
| 5 | Hardware & Serial Number      | CPU, BIOS serial, motherboard, GPU, RAM slots                                |
| 6 | Generate Battery Report       | `powercfg /batteryreport` → saved to Desktop and opened                     |
| 7 | Extract Wi-Fi Passwords       | Lists all saved SSIDs and reveals stored plaintext passwords                 |

### Maintenance & Repair

| # | Option                        | What It Does                                                                 |
|---|-------------------------------|------------------------------------------------------------------------------|
| 8  | Create System Restore Point  | Enables System Restore on C: and creates a labeled checkpoint                |
| 9  | Repair Windows (CHKDSK/DISM/SFC) | Runs `chkdsk /scan`, DISM `RestoreHealth`, then `sfc /scannow` in sequence |
| 10 | Cleanup Temp Files & Caches  | Deletes `%TEMP%`, empties Recycle Bin, clears Windows Error Reporting archive|
| 11 | Network Troubleshoot & Reset | `ipconfig` release/renew/flushdns, Winsock reset, connectivity ping test     |
| 12 | Fix Printing & Spooler Issues | Stops Spooler, clears print queue, restarts Spooler, lists installed printers|

### Security & Updates

| # | Option                        | What It Does                                                                 |
|---|-------------------------------|------------------------------------------------------------------------------|
| 13 | Check Security, AV & BitLocker | BitLocker status, Windows Firewall state, antivirus product and signature age |
| 14 | Check Outdated Software       | Runs `winget upgrade` and displays available updates                         |

---

## Standalone Scripts

Run any of these individually for a focused report.

### PCHealthScan.ps1
The most comprehensive single-file scan. Covers:
- OS name, version, total and free RAM, uptime
- Physical disk health (MediaType, OperationalStatus, HealthStatus)
- RAM slots (capacity in GB, speed in MHz)
- CHKDSK scan on all fixed drives
- Last 10 Error-level System event log entries
- Processes using >50% CPU or >500MB RAM

Output: `PCHealthReport.html` on the Desktop.

### SoftwareScan.ps1
- Full installed software inventory from the Windows registry (faster than `Win32_Product`)
- Flags known high-risk or deprecated software (Flash, Silverlight, QuickTime, Java 6/7, Ask Toolbar, CCleaner, uTorrent) in red
- Runs `winget upgrade` to show available updates

Output: `SoftwareRiskReport.html` on the Desktop.

### HardwareScan.ps1
- Physical disk inventory: device ID, media type, operational status, health status
- RAM inventory: slot location, capacity (GB), speed (MHz)

Output: `HardwareScanReport.html` on the Desktop.

### BatteryHealth.ps1
- Runs `powercfg /batteryreport`
- Wraps the output in a clean styled HTML shell
- Opens automatically in your default browser

Output: `BatteryHealthReport.html` on the Desktop.

### OSErrors.ps1
- Runs `sfc /verifyonly` and captures output
- Queries the System event log for Critical events (Level 1) in the last 7 days

Output: `OSErrorCheckReport.html` on the Desktop.

---

## Debloat & Revert

### MASTER_DEBLOAT_IT_GUY.ps1

A 6-step Windows 11 hardening and debloat script:

| Step | Action |
|------|--------|
| 1 | Disable Microsoft Recall and AI Manager (DISM + registry) |
| 2 | Remove consumer bloatware: Clipchamp, LinkedIn, Xbox apps, Widgets, Copilot, Phone Link, Bing Search, Feedback Hub |
| 3 | Purge Xbox/Gaming services and disable GameMode |
| 4 | Disable telemetry scheduled tasks and set `AllowTelemetry = 0` |
| 5 | Remove Start Menu ads and disable Bing search integration |
| 6 | Set DNS to Cloudflare (1.1.1.1 / 1.0.0.1), defer quality updates 7 days, feature updates 180 days |

A reboot is required after running.

### REVERT_IT_GUY_CHANGES.ps1

Reverses every change made by the debloat script:
- Re-enables Recall and AI features
- Re-enables telemetry tasks and resets `AllowTelemetry = 1`
- Restores Xbox/Gaming services to Manual startup
- Resets DNS to DHCP auto-configuration
- Re-enables Bing search and removes update deferral policies

---

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 (built into Windows — no install needed)
- Administrator privileges (right-click → "Run as Administrator")
- `winget` (Windows Package Manager) — required only for `SoftwareScan.ps1` and option 14 in MasterStack; pre-installed on Windows 11

---

## Usage

### Run MasterStack (recommended starting point)

```powershell
# Open PowerShell as Administrator, then:
Set-ExecutionPolicy Bypass -Scope Process -Force
.\MasterStack
```

### Run a standalone script

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\PCHealthScan.ps1
```

### Run the debloat script

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\MASTER_DEBLOAT_IT_GUY.ps1
```

---

## Output

All scripts generate HTML reports that open automatically. Reports are saved to `%USERPROFILE%\Desktop\` or `%TEMP%\` depending on the script.

| Script            | Output file                    | Location    |
|-------------------|-------------------------------|-------------|
| MasterStack       | Various per option            | `%TEMP%`    |
| PCHealthScan.ps1  | `PCHealthReport.html`         | Desktop     |
| SoftwareScan.ps1  | `SoftwareRiskReport.html`     | Desktop     |
| HardwareScan.ps1  | `HardwareScanReport.html`     | Desktop     |
| BatteryHealth.ps1 | `BatteryHealthReport.html`    | Desktop     |
| OSErrors.ps1      | `OSErrorCheckReport.html`     | Desktop     |

---

## Security Note

**Option 7 in MasterStack (Wi-Fi Password Extractor)** displays saved Wi-Fi passwords in plaintext using `netsh wlan show profile key=clear`. This requires Administrator access. Use only on machines you own or are authorized to manage.

The debloat script makes significant system changes. Review it before running and use `REVERT_IT_GUY_CHANGES.ps1` if you need to restore defaults. Creating a System Restore Point (option 8 in MasterStack) before running the debloat is strongly recommended.
