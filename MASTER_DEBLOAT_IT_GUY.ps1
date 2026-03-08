<#
.SYNOPSIS
    The IT Guy - Windows 11 Master Debloat & Optimization
.DESCRIPTION
    Comprehensive script to remove Recall, AI, Ads, and Bloatware. 
    Optimized for high-performance business workstations.
.COPYRIGHT
    Copyright (c) 2026 Jeff Lemons. Licensed under the MIT License.
#>

Write-Host "--- THE IT GUY: STARTING SYSTEM OPTIMIZATION ---" -ForegroundColor Green

# 1. REMOVE MICROSOFT RECALL & AI COMPONENTS
Write-Host "[1/6] Disabling Microsoft Recall and AI Manager..." -ForegroundColor Red
Dism /online /Disable-Feature /FeatureName:Recall /NoRestart /Quiet
$recallManager = Get-AppxPackage -Name "*aimgr*" -AllUsers
if ($recallManager) { $recallManager | Remove-AppxPackage -ErrorAction SilentlyContinue }
$RecallPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI"
if (!(Test-Path $RecallPolicyPath)) { New-Item $RecallPolicyPath -Force }
Set-ItemProperty -Path $RecallPolicyPath -Name "DisableAIDataAnalysis" -Value 1

# 2. SURGICAL BLOATWARE REMOVAL
Write-Host "[2/6] Removing consumer bloatware..." -ForegroundColor Cyan
$bloatApps = @(
    "Clipchamp.Clipchamp", "7EE7776C.LinkedInforWindows", "Microsoft.ZuneMusic",
    "Microsoft.ZuneVideo", "Microsoft.YourPhone", "Microsoft.BingSearch",
    "Microsoft.Copilot", "Microsoft.WindowsFeedbackHub", "Microsoft.GetHelp",
    "MicrosoftWindows.Client.WebExperience"
)
foreach ($app in $bloatApps) {
    Get-AppxPackage -Name $app -AllUsers | Remove-AppxPackage -ErrorAction SilentlyContinue
}

# 3. XBOX & GAMING PURGE
Write-Host "[3/6] Purging Xbox and Gaming services..." -ForegroundColor Yellow
$gamingApps = @("Microsoft.XboxGamingOverlay", "Microsoft.XboxApp", "Microsoft.GamingApp", "Microsoft.SolitaireCollection")
foreach ($gApp in $gamingApps) { winget uninstall --id "$gApp" --silent --accept-source-agreements -e }
Get-Service -Name "XblAuthManager", "XblGameSave", "XboxNetApiSvc" | Set-Service -StartupType Disabled -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\GameBar" -Name "AllowAutoGameMode" -Value 0

# 4. TELEMETRY & AI BACKGROUND TASKS
Write-Host "[4/6] Killing background telemetry tasks..." -ForegroundColor Magenta
$tasks = @(
    "\Microsoft\Windows\WindowsAI\SnapshotRetention",
    "\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser",
    "\Microsoft\Windows\Customer Experience Improvement Program\Consolidator",
    "\Microsoft\Windows\DeliveryOptimization\UsageStatsLog"
)
foreach ($task in $tasks) { Disable-ScheduledTask -TaskName $task -ErrorAction SilentlyContinue }
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0

# 5. VISUAL DE-CLUTTER & AD REMOVAL
Write-Host "[5/6] Banning ads and 'Suggested' content..." -ForegroundColor Cyan
$ContentPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
Set-ItemProperty -Path $ContentPath -Name "SubscribedContent-338393Enabled" -Value 0
Set-ItemProperty -Path $ContentPath -Name "SubscribedContent-353694Enabled" -Value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "BingSearchEnabled" -Value 0
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\SearchSettings" -Name "IsDynamicSearchBoxEnabled" -Value 0

# 6. NETWORK & UPDATE POLICY
Write-Host "[6/6] Configuring Cloudflare DNS & Update Policies..." -ForegroundColor Green
$adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
foreach ($a in $adapters) { Set-DnsClientServerAddress -InterfaceAlias $a.Name -ServerAddresses ("1.1.1.1", "1.0.0.1") }
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "DeferQualityUpdatesPeriodInDays" -Value 7
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "DeferFeatureUpdatesPeriodInDays" -Value 180

Write-Host "--- OPTIMIZATION COMPLETE. PLEASE REBOOT ---" -ForegroundColor Green
