# Windows System Health & Diagnostic Scripts

A collection of PowerShell scripts designed to monitor, diagnose, and report on various aspects of Windows system health. Each script generates a clean, easy-to-read HTML report saved directly to your desktop.

## 🚀 Scripts Included

| Script | Description |
| :--- | :--- |
| `BatteryHealth.ps1` | Generates a detailed report of the system's battery health and usage history. |
| `HardwareScan.ps1` | Scans physical disk health, memory status, and general system information. |
| `OSErrors.ps1` | Checks for disk errors, system file corruption (SFC), image health (DISM), and critical event logs. |
| `PCHealthScan.ps1` | A comprehensive scan combining system summary, hardware health, and high-resource process monitoring. |
| `SoftwareScan.ps1` | Lists all installed software and checks for potential version compatibility issues. |

## 🛠️ Requirements

* **Operating System**: Windows 10 or Windows 11.
* **Permissions**: All scripts must be run in **PowerShell as Administrator** to access system-level diagnostics like DISM, SFC, and CHKDSK.
* **Environment**: PowerShell 5.1 or higher.

## 📋 Usage

1. Download or clone this repository.
2. Right-click on the script you wish to run (e.g., `PCHealthScan.ps1`).
3. Select **Run with PowerShell**. 
   * *Note: If prompted, ensure you are running as an Administrator.*
4. Once the scan is complete, an HTML report will automatically open in your default browser and be saved to your **Desktop**.

## ⚠️ Disclaimer
These scripts are intended for diagnostic purposes. While they include standard Windows repair tools (like SFC and DISM), always ensure you have backups of important data before running system-level repairs.
