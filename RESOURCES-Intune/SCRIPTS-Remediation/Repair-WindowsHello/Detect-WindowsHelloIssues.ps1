# Detection Script for Windows Hello Face Sign-in Issue
# Designed for Intune Proactive Remediation
# Returns exit code 0 (compliant) or 1 (non-compliant)

# Configuration
$LogPath = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\WinHelloDetect.log"
$LogDir = Split-Path $LogPath -Parent

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

Write-Log "Starting Windows Hello Face detection script"

# 1. Check Windows Version
$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
Write-Log "OS: $($osInfo.Caption), Version: $($osInfo.Version), Build: $($osInfo.BuildNumber)"
if ($osInfo.Caption -like "*Windows 10*") {
    Write-Log "Note: Windows 10 support ends October 14, 2025."
}

# 2. Check for IR Camera (Windows Hello Face requires an IR camera)
$irCamera = Get-CimInstance -ClassName Win32_PnPEntity | Where-Object { $_.PNPClass -eq "Camera" -and $_.Name -like "*IR*" }
if (-not $irCamera) {
    Write-Log "No IR camera detected. Windows Hello Face not supported on this device."
    exit 0  # Compliant: Device doesn't support Windows Hello Face
}

# 3. Check Biometric Device Status
$bioDevices = Get-PnpDevice -Class Biometric | Where-Object { $_.Name -like "*Windows Hello Face*" }
if (-not $bioDevices) {
    Write-Log "No Windows Hello Face device found despite IR camera presence."
    exit 1  # Non-compliant: Should have Windows Hello Face
}

foreach ($device in $bioDevices) {
    Write-Log "Biometric Device: $($device.Name), Status: $($device.Status), Problem: $($device.Problem)"
    if ($device.Status -ne "OK" -or $device.Problem) {
        Write-Log "Biometric device has issues."
        exit 1  # Non-compliant: Device issue detected
    }
}

# 4. Check Windows Hello Configuration
$winBio = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WinBio" -ErrorAction SilentlyContinue
if (-not $winBio -or $winBio.BiometricEnabled -eq 0) {
    Write-Log "Windows Hello biometric settings missing or disabled."
    exit 1  # Non-compliant: Configuration issue
}

# 5. Check Sign-in Options
$signInOptions = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\Settings\AllowSignInOptions" -Name "value" -ErrorAction SilentlyContinue
if ($signInOptions.value -eq 0) {
    Write-Log "Sign-in options restricted, potentially hiding Windows Hello Face."
    exit 1  # Non-compliant: Restricted sign-in options
}

# 6. Check Event Logs for Errors
$events = Get-WinEvent -LogName "Microsoft-Windows-HelloForBusiness/Operational" -MaxEvents 10 -ErrorAction SilentlyContinue | Where-Object { $_.LevelDisplayName -eq "Error" }
if ($events) {
    Write-Log "Found Windows Hello errors in event logs: $($events.Message -join '; ')"
    exit 1  # Non-compliant: Errors detected
}

# If all checks pass, Windows Hello Face is functional
Write-Log "Windows Hello Face sign-in option appears functional."
exit 0  # Compliant