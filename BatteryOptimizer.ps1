<#
.SYNOPSIS
Optimizes Windows 11 battery settings with WiFi control options.
.DESCRIPTION
- Optimize-Power: Basic battery savings (WiFi disabled by default)
- Optimize-Ultra: Aggressive battery savings (WiFi always disabled)
- New -EnableWiFi parameter for Power mode
- Strict mode transition enforcement
- Hidden windows during operations
#>

#Requires -RunAsAdministrator

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Optimize-Power", "Optimize-Ultra", "Revert")]
    [string]$Mode,
    [int]$RevertAfterMinutes = 0,
    [switch]$EnableWiFi  # New parameter for WiFi control in Power mode
)

# Create Logs directory if it doesn't exist
$logDir = "$PSScriptRoot\Logs"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

# Log file with timestamp
$logFile = "$logDir\BatteryOptimizer_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$message, [string]$level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$level] $message"
    Add-Content -Path $logFile -Value $logEntry
    
    if ($level -eq "WARNING") {
        Write-Host $logEntry -ForegroundColor Yellow
    } elseif ($level -eq "ERROR") {
        Write-Host $logEntry -ForegroundColor Red
    } else {
        Write-Host $logEntry
    }
}

# State tracking file
$stateFile = "$env:APPDATA\BatteryOptimizerState.txt"

function Get-CurrentMode {
    if (Test-Path $stateFile) {
        return Get-Content $stateFile
    }
    return $null
}

function Set-CurrentMode {
    param([string]$mode)
    $mode | Out-File $stateFile -Force
}

# Check for valid mode transitions
$currentMode = Get-CurrentMode
if ($currentMode) {
    switch ($Mode) {
        "Optimize-Power" {
            if ($currentMode -eq "Optimize-Ultra") {
                Write-Log "Cannot switch from Ultra to Power mode directly. Please revert first." -level "ERROR"
                exit
            }
        }
        "Optimize-Ultra" {
            if ($currentMode -eq "Optimize-Power") {
                Write-Log "Cannot switch from Power to Ultra mode directly. Please revert first." -level "ERROR"
                exit
            }
        }
        "Revert" {
            if ($currentMode -eq "Revert") {
                Write-Log "Already in Revert mode. No action needed." -level "WARNING"
                exit
            }
        }
    }
}

# Temporarily enable script execution
$originalExecutionPolicy = Get-ExecutionPolicy
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
Write-Log "Execution policy set to Bypass for this session"

# Saved Settings File
$settingsFile = "$env:APPDATA\BatterySettingsBackup.json"

# Settings Configuration
$settings = @{
    OriginalPowerPlan = $null
    BatteryPowerPlan = "SCHEME_BATTERY"
    OriginalBrightness = $null
    PowerBrightness = 50
    UltraBrightness = 30
    WiFiAdapterName = $null
    WiFiEnabled = $false  # Track WiFi state
    OriginalServices = @()
    OriginalStartupItems = @()
}

function Schedule-Revert {
    param([int]$minutes)
    Write-Log "Scheduling revert in $minutes minutes..."
    
    $scriptPath = (Resolve-Path $PSCommandPath).Path
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument `
        "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command `"& {Start-Process powershell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \`"$scriptPath\`" -Mode Revert' -Verb RunAs}`""
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes($minutes)
    $principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries
    
    try {
        Unregister-ScheduledTask -TaskName "BatteryOptimizerRevert" -Confirm:$false -ErrorAction SilentlyContinue
        Register-ScheduledTask -TaskName "BatteryOptimizerRevert" `
            -Action $action -Trigger $trigger `
            -Principal $principal -Settings $settings `
            -Force -ErrorAction Stop | Out-Null
        Write-Log "Successfully scheduled revert task"
        return $true
    } catch {
        Write-Log "Failed to schedule revert task: $_" -level "ERROR"
        return $false
    }
}

function Save-OriginalSettings {
    Write-Log "Saving original system settings..."
    
    $brightness = 70
    try {
        $brightness = (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness).CurrentBrightness
        Write-Log "Current brightness: $brightness%"
    } catch {
        Write-Log "Could not get current brightness, using default 70" -level "WARNING"
    }

    $wifiAdapter = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match "Wi-Fi|Wireless" } | Select-Object -First 1
    $settings.WiFiAdapterName = $wifiAdapter.Name
    $settings.WiFiEnabled = ($wifiAdapter.Status -eq "Up")
    Write-Log "Detected Wi-Fi adapter: $($settings.WiFiAdapterName) (Status: $($wifiAdapter.Status))"

    $startupItems = @()
    $servicesToTrack = @("BluetoothUserService", "BthAvctpSvc", "DiagTrack", "SysMain", "WSearch")
    
    try {
        $registryPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
        foreach ($path in $registryPaths) {
            if (Test-Path $path) {
                $props = Get-ItemProperty -Path $path
                $props.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" -and $_.Name -ne "(default)" } | ForEach-Object {
                    $startupItems += @{
                        Type = "Registry"
                        Path = $path
                        Name = $_.Name
                        Value = $_.Value
                    }
                }
            }
        }

        $startupFolder = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
        if (Test-Path $startupFolder) {
            Get-ChildItem $startupFolder | ForEach-Object {
                $shortcut = (New-Object -ComObject WScript.Shell).CreateShortcut($_.FullName)
                $startupItems += @{
                    Type = "Shortcut"
                    Path = $_.FullName
                    Target = $shortcut.TargetPath
                    Arguments = $shortcut.Arguments
                    WorkingDirectory = $shortcut.WorkingDirectory
                }
            }
        }

        $services = @()
        foreach ($service in $servicesToTrack) {
            $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
            if ($svc) {
                $services += @{
                    Name = $svc.Name
                    StartupType = $svc.StartType
                    Status = $svc.Status
                }
            }
        }
    } catch {
        Write-Log "Error while gathering system settings: $_" -level "ERROR"
    }

    $backup = @{
        OriginalPowerPlan = (powercfg /getactivescheme | Select-String -Pattern 'GUID: (.+?) \(').Matches.Groups[1].Value
        OriginalBrightness = $brightness
        OriginalExecutionPolicy = $originalExecutionPolicy
        OriginalStartupItems = $startupItems
        OriginalServices = $services
        WiFiAdapterName = $settings.WiFiAdapterName
        WiFiEnabled = $settings.WiFiEnabled
        BackupTime = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    }

    try {
        $backup | ConvertTo-Json -Depth 5 | Out-File $settingsFile -Force
        Write-Log "Original settings saved to $settingsFile"
    } catch {
        Write-Log "Failed to save settings backup: $_" -level "ERROR"
    }
}

function Enable-Bluetooth {
    Write-Log "Attempting to enable Bluetooth services..."
    $success = $false
    $maxRetries = 5
    $retryDelay = 5
    
    $btServices = @("BluetoothUserService", "BTAGService", "BthAvctpSvc", "bthserv")

    for ($i = 1; $i -le $maxRetries; $i++) {
        try {
            foreach ($service in $btServices) {
                $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
                if ($svc) {
                    # Run service commands with hidden window
                    $cmd = "Set-Service -Name $($svc.Name) -StartupType Automatic; Start-Service -Name $($svc.Name)"
                    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command $cmd" -Wait
                }
            }
            
            Start-Process "pnputil" -ArgumentList "/enable-device", "BTHENUM\*" -WindowStyle Hidden -Wait
            Start-Sleep -Seconds 2
            $success = $true
            break
        } catch {
            Write-Log "Attempt $i failed: $_" -level "WARNING"
            if ($i -lt $maxRetries) { Start-Sleep -Seconds $retryDelay }
        }
    }
    
    if ($success) {
        Write-Log "Bluetooth services enabled successfully"
    } else {
        Write-Log "Failed to enable Bluetooth after $maxRetries attempts" -level "ERROR"
    }
    return $success
}

function Restore-OriginalSettings {
    if (-not (Test-Path $settingsFile)) {
        Write-Log "No backup found. Run optimization first." -level "WARNING"
        return $false
    }

    Write-Log "Beginning system restoration from backup..."
    $restoreSuccess = $true

    try {
        $original = Get-Content $settingsFile | ConvertFrom-Json
        Write-Log "Restoring settings from backup made at $($original.BackupTime)"

        try {
            powercfg /setactive $original.OriginalPowerPlan 2>&1 | Out-Null
            Write-Log "Restored power plan"
        } catch {
            Write-Log "Failed to restore power plan: $_" -level "ERROR"
            $restoreSuccess = $false
        }

        try {
            (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1, $original.OriginalBrightness)
            Write-Log "Restored brightness"
        } catch {
            Write-Log "Could not restore brightness" -level "WARNING"
        }

        if ($original.WiFiAdapterName) {
            try {
                if ($original.WiFiEnabled) {
                    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command `"Enable-NetAdapter -Name '$($original.WiFiAdapterName)' -Confirm:`$false`"" -Wait
                    Write-Log "Enabled Wi-Fi adapter (restored original state)"
                } else {
                    Write-Log "Wi-Fi was originally disabled, leaving disabled"
                }
            } catch {
                Write-Log "Could not restore Wi-Fi adapter state" -level "ERROR"
                $restoreSuccess = $false
            }
        }

        $btResult = Enable-Bluetooth
        if (-not $btResult) { $restoreSuccess = $false }

        foreach ($service in $original.OriginalServices) {
            try {
                $cmd = "Set-Service -Name '$($service.Name)' -StartupType '$($service.StartupType)'; "
                if ($service.Status -eq "Running") { $cmd += "Start-Service -Name '$($service.Name)'" }
                Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command $cmd" -Wait
                Write-Log "Restored service $($service.Name)"
            } catch {
                Write-Log "Could not restore service $($service.Name)" -level "WARNING"
            }
        }

        try {
            foreach ($item in $original.OriginalStartupItems) {
                if ($item.Type -eq "Registry") {
                    $command = "Set-ItemProperty -Path '$($item.Path)' -Name '$($item.Name)' -Value '$($item.Value)'"
                    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command $command" -Wait
                } elseif ($item.Type -eq "Shortcut") {
                    $ws = New-Object -ComObject WScript.Shell
                    $shortcut = $ws.CreateShortcut($item.Path)
                    $shortcut.TargetPath = $item.Target
                    $shortcut.Arguments = $item.Arguments
                    $shortcut.WorkingDirectory = $item.WorkingDirectory
                    $shortcut.Save()
                }
            }
            Write-Log "Restored startup items"
        } catch {
            Write-Log "Error restoring startup items: $_" -level "ERROR"
            $restoreSuccess = $false
        }

        try {
            Remove-Item $settingsFile -Force
            Write-Log "Removed settings backup file"
        } catch {
            Write-Log "Could not remove settings backup file" -level "WARNING"
        }

        try {
            Unregister-ScheduledTask -TaskName "BatteryOptimizerRevert" -Confirm:$false -ErrorAction Stop
            Write-Log "Removed scheduled revert task"
        } catch {
            Write-Log "No scheduled task to remove or removal failed" -level "WARNING"
        }

        if ($restoreSuccess) {
            Write-Log "System restoration completed successfully!"
        } else {
            Write-Log "System restoration completed with some errors" -level "WARNING"
        }
        return $restoreSuccess
    } catch {
        Write-Log "Fatal error during restoration: $_" -level "ERROR"
        return $false
    }
}

function Optimize-Power {
    Write-Log "Applying Power optimization (high-impact settings)..."
    
    # Apply power plan
    powercfg /setactive $settings.BatteryPowerPlan 2>&1 | Out-Null
    Write-Log "Set power plan to Battery Saver"

    # Set brightness
    try {
        (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1, $settings.PowerBrightness)
        Write-Log "Set brightness to $($settings.PowerBrightness)%"
    } catch {
        Write-Log "Could not adjust brightness" -level "WARNING"
    }

    # Handle WiFi based on parameter
    if ($EnableWiFi) {
        try {
            if ($settings.WiFiAdapterName) {
                Enable-NetAdapter -Name $settings.WiFiAdapterName -Confirm:$false -ErrorAction Stop
                Write-Log "WiFi explicitly enabled (via -EnableWiFi parameter)"
                $settings.WiFiEnabled = $true
            }
        } catch {
            Write-Log "Could not enable WiFi adapter" -level "WARNING"
        }
    } else {
        try {
            if ($settings.WiFiAdapterName) {
                Disable-NetAdapter -Name $settings.WiFiAdapterName -Confirm:$false -ErrorAction Stop
                Write-Log "Disabled WiFi adapter (default for Power mode)"
                $settings.WiFiEnabled = $false
            }
        } catch {
            Write-Log "Could not disable WiFi adapter" -level "WARNING"
        }
    }

    # Disable Bluetooth
    try {
        Stop-Service "BluetoothUserService", "BthAvctpSvc" -ErrorAction Stop
        Set-Service "BluetoothUserService", "BthAvctpSvc" -StartupType Disabled -ErrorAction Stop
        Write-Log "Disabled Bluetooth services"
    } catch {
        Write-Log "Could not disable Bluetooth" -level "WARNING"
    }

    # Set battery saver threshold
    powercfg /setdcvalueindex SCHEME_CURRENT SUB_ENERGYSAVER ESBATTTHRESHOLD 20 2>&1 | Out-Null
    Write-Log "Set Battery Saver threshold to 20%"

    # Set sleep timer
    powercfg /change monitor-timeout-dc 5 2>&1 | Out-Null
    Write-Log "Set display sleep timer to 5 minutes"

    # Disable background apps
    try {
        Get-ChildItem "HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications" | ForEach-Object {
            Set-ItemProperty -Path $_.PSPath -Name "Disabled" -Value 1 -ErrorAction Stop
        }
        Write-Log "Disabled background apps"
    } catch {
        Write-Log "Could not disable background apps" -level "WARNING"
    }

    Write-Log "Power optimization complete!"
    Set-CurrentMode "Optimize-Power"
}

function Optimize-Ultra {
    Write-Log "Applying Ultra optimization (high + medium impact settings)..."

    # Apply all Power optimizations first
    powercfg /setactive $settings.BatteryPowerPlan 2>&1 | Out-Null
    Write-Log "Set power plan to Battery Saver"

    try {
        (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1, $settings.UltraBrightness)
        Write-Log "Set aggressive brightness to $($settings.UltraBrightness)%"
    } catch {
        Write-Log "Could not adjust brightness" -level "WARNING"
    }

    # Always disable WiFi in Ultra mode
    try {
        if ($settings.WiFiAdapterName) {
            Disable-NetAdapter -Name $settings.WiFiAdapterName -Confirm:$false -ErrorAction Stop
            Write-Log "Disabled WiFi adapter (always disabled in Ultra mode)"
            $settings.WiFiEnabled = $false
        }
    } catch {
        Write-Log "Could not disable WiFi adapter" -level "WARNING"
    }

    # Disable Bluetooth
    try {
        Stop-Service "BluetoothUserService", "BthAvctpSvc" -ErrorAction Stop
        Set-Service "BluetoothUserService", "BthAvctpSvc" -StartupType Disabled -ErrorAction Stop
        Write-Log "Disabled Bluetooth services"
    } catch {
        Write-Log "Could not disable Bluetooth" -level "WARNING"
    }

    # Set battery saver threshold
    powercfg /setdcvalueindex SCHEME_CURRENT SUB_ENERGYSAVER ESBATTTHRESHOLD 20 2>&1 | Out-Null
    Write-Log "Set Battery Saver threshold to 20%"

    # Set more aggressive sleep timer
    powercfg /change monitor-timeout-dc 3 2>&1 | Out-Null
    Write-Log "Set display sleep timer to 3 minutes"

    # Disable background apps
    try {
        Get-ChildItem "HKCU:\Software\Microsoft\Windows\CurrentVersion\BackgroundAccessApplications" | ForEach-Object {
            Set-ItemProperty -Path $_.PSPath -Name "Disabled" -Value 1 -ErrorAction Stop
        }
        Write-Log "Disabled background apps"
    } catch {
        Write-Log "Could not disable background apps" -level "WARNING"
    }

    # Ultra-specific optimizations
    try {
        Stop-Service "WSearch" -ErrorAction Stop
        Set-Service "WSearch" -StartupType Disabled -ErrorAction Stop
        Write-Log "Disabled search indexing service"
    } catch {
        Write-Log "Could not disable search indexing" -level "WARNING"
    }

    try {
        Stop-Service "SysMain" -ErrorAction Stop
        Set-Service "SysMain" -StartupType Disabled -ErrorAction Stop
        Write-Log "Disabled SysMain service"
    } catch {
        Write-Log "Could not disable SysMain" -level "WARNING"
    }

    try {
        Stop-Service "DiagTrack" -ErrorAction Stop
        Set-Service "DiagTrack" -StartupType Disabled -ErrorAction Stop
        Write-Log "Disabled diagnostics tracking"
    } catch {
        Write-Log "Could not disable diagnostics tracking" -level "WARNING"
    }

    try {
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SystemPaneSuggestionsEnabled" -Value 0 -ErrorAction Stop
        Write-Log "Disabled live tiles"
    } catch {
        Write-Log "Could not disable live tiles" -level "WARNING"
    }

    try {
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\GameDVR" -Name "AppCaptureEnabled" -Value 0 -ErrorAction Stop
        Write-Log "Disabled Game Bar"
    } catch {
        Write-Log "Could not disable Game Bar" -level "WARNING"
    }

    try {
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" -Name "VisualFXSetting" -Value 2 -ErrorAction Stop
        Write-Log "Disabled visual effects"
    } catch {
        Write-Log "Could not disable visual effects" -level "WARNING"
    }

    Write-Log "Ultra optimization complete!"
    Set-CurrentMode "Optimize-Ultra"
}

# Main execution
try {
    switch ($Mode) {
        "Optimize-Power" {
            if ((Get-CurrentMode) -eq "Optimize-Ultra") {
                Write-Log "Cannot switch from Ultra to Power mode directly. Please revert first." -level "ERROR"
                break
            }
            Save-OriginalSettings
            Optimize-Power
            if ($RevertAfterMinutes -gt 0) {
                if (-not (Schedule-Revert $RevertAfterMinutes)) {
                    Write-Log "Failed to schedule automatic revert - will need manual revert" -level "ERROR"
                }
            }
        }
        "Optimize-Ultra" {
            if ((Get-CurrentMode) -eq "Optimize-Power") {
                Write-Log "Cannot switch from Power to Ultra mode directly. Please revert first." -level "ERROR"
                break
            }
            Save-OriginalSettings
            Optimize-Ultra
            if ($RevertAfterMinutes -gt 0) {
                if (-not (Schedule-Revert $RevertAfterMinutes)) {
                    Write-Log "Failed to schedule automatic revert - will need manual revert" -level "ERROR"
                }
            }
        }
        "Revert" {
            if ((Get-CurrentMode) -eq "Revert") {
                Write-Log "Already in Revert mode. No action needed." -level "WARNING"
                break
            }
            $restoreResult = Restore-OriginalSettings
            if ($restoreResult) {
                Set-CurrentMode "Revert"
                if (-not $restoreResult) {
                    Write-Log "Restoration encountered errors - some settings may not have been restored" -level "ERROR"
                }
            }
        }
        default {
            Write-Log "Usage: .\BatteryOptimizer.ps1 -Mode Optimize-Power [-EnableWiFi]|Optimize-Ultra [-RevertAfterMinutes <minutes>]"
            Write-Log "       .\BatteryOptimizer.ps1 -Mode Revert"
            Write-Log "Note: -EnableWiFi parameter keeps WiFi enabled in Power mode"
        }
    }
} catch {
    Write-Log "Script encountered fatal error: $_" -level "ERROR"
} finally {
    Set-ExecutionPolicy -ExecutionPolicy $originalExecutionPolicy -Scope Process -Force -ErrorAction SilentlyContinue
    Write-Log "Execution policy restored to $originalExecutionPolicy"
    Write-Log "Script execution complete. Log saved to $logFile"
}