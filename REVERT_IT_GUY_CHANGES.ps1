<#
.SYNOPSIS
    The IT Guy - Windows 11 Restore Script
.DESCRIPTION
    Reverses optimizations, re-enables AI/Recall, and restores default Windows settings.
.COPYRIGHT
    Copyright (c) 2026 Jeff Lemons. Licensed under the MIT License.
#>

Write-Host "--- THE IT GUY: REVERTING SYSTEM CHANGES ---" -ForegroundColor Yellow

# 1. RE-ENABLE AI & RECALL
Write-Host "[1/5] Restoring Microsoft Recall and AI components..." -ForegroundColor Cyan
Dism /online /Enable-Feature /FeatureName:Recall /NoRestart /Quiet
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI" -Name "DisableAIDataAnalysis" -Value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\AI" -Name "AIPermitted" -Value 1

# 2. RESTORE TELEMETRY & SCHEDULED TASKS
Write-Host "[2/5] Re-enabling background telemetry tasks..." -ForegroundColor White
$tasks = @(
    "\Microsoft\Windows\WindowsAI\SnapshotRetention",
    "\Microsoft\Windows\WindowsAI\SnapshotCleanup",
    "\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser",
    "\Microsoft\Windows\Customer Experience Improvement Program\Consolidator",
    "\Microsoft\Windows\DeliveryOptimization\UsageStatsLog"
)
foreach ($task in $tasks) {
    Enable-ScheduledTask -TaskName $task -ErrorAction SilentlyContinue
}
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 1

# 3. RESTORE GAMING & XBOX SERVICES
Write-Host "[3/5] Restoring Gaming and Xbox services..." -ForegroundColor Green
Get-Service -Name "XblAuthManager", "XblGameSave", "XboxNetApiSvc" | Set-Service -StartupType Manual
Set-ItemProperty -Path "HKCU:\Software\Microsoft\GameBar" -Name "AllowAutoGameMode" -Value 1

# 4. RESET NETWORK & UI ADS
Write-Host "[4/5] Resetting DNS and UI 'Suggested' content..." -ForegroundColor Gray
$adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
foreach ($a in $adapters) {
    Set-DnsClientServerAddress -InterfaceAlias $a.Name -ResetServerAddresses
}

# Re-enable Start Menu Web Search (Bing)
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "BingSearchEnabled" -Value 1
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\SearchSettings" -Name "IsDynamicSearchBoxEnabled" -Value 1

# 5. RESTORE DEFAULT UPDATE POLICY
Write-Host "[5/5] Resetting Windows Update to default settings..." -ForegroundColor Magenta
$UpdatePath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
Remove-ItemProperty -Path $UpdatePath -Name "DeferQualityUpdatesPeriodInDays" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $UpdatePath -Name "DeferFeatureUpdatesPeriodInDays" -ErrorAction SilentlyContinue

# Set Windows Update to Auto Download/Install (Default)
$AUPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
if (Test-Path $AUPath) { Set-ItemProperty -Path $AUPath -Name "AUOptions" -Value 4 }

Write-Host "--- REVERT COMPLETE. A REBOOT IS REQUIRED ---" -ForegroundColor Green
