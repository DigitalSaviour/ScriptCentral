# Remediation Script for Windows Hello Face Sign-in Issue
# Designed for Intune Proactive Remediation
# Returns exit code 0 (success) or 1 (failure)

# Configuration
$LogPath = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\WinHelloFix.log"
$LogDir = Split-Path $LogPath -Parent
$RestartFlagPath = "HKLM:\SOFTWARE\Custom\WinHelloFix"
$ErrorOccurred = $false

# Function to log messages
function Write-Log {
    param ($Message)
    $logMessage = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Write-Output $logMessage
    Add-Content -Path $LogPath -Value $logMessage -ErrorAction SilentlyContinue
}

# Ensure log directory exists
if (-not (Test-Path $LogDir)) {
    New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
}

Write-Log "Starting Windows Hello Face remediation script"

try {
    # 1. Check Windows Version
    $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    Write-Log "OS: $($osInfo.Caption), Version: $($osInfo.Version), Build: $($osInfo.BuildNumber)"

    # 2. Enable Sign-in Options
    $signInOptions = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Settings\AllowSignInOptions" -Name "value" -ErrorAction SilentlyContinue
    if ($signInOptions.value -eq 0 -or -not $signInOptions) {
        Write-Log "Enabling all sign-in options..."
        New-Item -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Settings" -Name "AllowSignInOptions" -Force -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Settings\AllowSignInOptions" -Name "value" -Value 1 -Force
        Write-Log "Sign-in options enabled. Restart required."
        $restartRequired = $true
    }

    # 3. Update Biometric and Camera Drivers
    Write-Log "Checking and updating biometric devices..."
    $bioDevices = Get-PnpDevice -Class Biometric
    foreach ($device in $bioDevices) {
        Write-Log "Biometric Device: $($device.Name), Status: $($device.Status)"
        if ($device.Status -ne "OK") {
            try {
                Update-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction Stop
                Write-Log "Updated driver for $($device.Name)"
            } catch {
                Write-Log "Failed to update driver for $($device.Name): $_"
                $ErrorOccurred = $true
            }
            $restartRequired = $true
        }
    }

    Write-Log "Checking and updating camera devices..."
    $camDevices = Get-CimInstance -ClassName Win32_PnPEntity | Where-Object { $_.PNPClass -eq "Camera" }
    if ($camDevices) {
        pnputil /scan-devices
        Write-Log "Rescanned devices to reinstall camera drivers."
        $restartRequired = $true
    }

    # 4. Run System File Checker (SFC)
    Write-Log "Running System File Checker..."
    $sfcResult = sfc /scannow
    Write-Log "SFC Output: $sfcResult"
    if ($sfcResult -match "found corrupt files") {
        Write-Log "Corrupt files repaired. Restart required."
        $restartRequired = $true
    }

    # 5. Install Windows Updates
    Write-Log "Checking for Windows Updates..."
    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $searcher = $updateSession.CreateUpdateSearcher()
    $updates = $searcher.Search("IsInstalled=0").Updates
    if ($updates.Count -gt 0) {
        Write-Log "Installing $($updates.Count) updates..."
        $downloader = $updateSession.CreateUpdateDownloader()
        $downloader.Updates = $updates
        $downloader.Download()
        $installer = $updateSession.CreateUpdateInstaller()
        $installer.Updates = $updates
        $installer.Install() | Out-Null
        Write-Log "Updates installed. Restart required."
        $restartRequired = $true
    } else {
        Write-Log "No pending updates found."
    }

    # 6. Reset Windows Hello Biometric Data
    Write-Log "Resetting Windows Hello biometric data..."
    Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WinBio\Database\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Log "Biometric data cleared. User must re-enroll face locally via Settings."

    # 7. Check for Third-Party Software Conflicts
    Write-Log "Checking for conflicting software..."
    $software = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*Password*" -or $_.Name -like "*1Password*" -or $_.Name -like "*LastPass*" }
    if ($software) {
        Write-Log "Potential conflicting software found: $($software.Name -join ', ')"
        # Note: Uncomment to uninstall conflicting software if confirmed
        # foreach ($app in $software) {
        #     $app.Uninstall()
        #     Write-Log "Uninstalled $($app.Name)"
        #     $restartRequired = $true
        # }
    } else {
        Write-Log "No known conflicting software detected."
    }

    # 8. Set Restart Flag
    if ($restartRequired) {
        Write-Log "Setting restart flag in registry..."
        New-Item -Path "HKLM:\SOFTWARE\Custom" -Name "WinHelloFix" -Force -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path $RestartFlagPath -Name "RestartRequired" -Value 1 -Force
        Write-Log "Restart required. Notify user or schedule via Intune."
    }

    # Exit with success if no errors
    Write-Log "Remediation completed successfully."
    exit 0
}
catch {
    Write-Log "Error during remediation: $_"
    $ErrorOccurred = $true
    exit 1
}
finally {
    if ($ErrorOccurred) {
        Write-Log "Remediation completed with errors. Check log for details."
        exit 1
    }
}