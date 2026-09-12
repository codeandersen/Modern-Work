<#
.SYNOPSIS
    Sets the desktop wallpaper and lock screen image on Windows 11 Pro devices.

.DESCRIPTION
    This script is designed to be deployed via Microsoft Intune as a Platform Script.
    It downloads wallpaper and/or lock screen images from a URL (e.g., Azure Blob Storage)
    and applies them via registry keys. This is the recommended approach for Windows 11 Pro
    devices where the Personalization CSP (./Vendor/MSFT/Personalization) is not supported.

    The script:
    1. Downloads the image(s) from the specified URL(s) to a local folder.
    2. Sets the desktop wallpaper via the PersonalizationCSP registry key + SystemParametersInfo Win32 API
       + per-user registry keys for all loaded profiles.
    3. Sets the lock screen image via the PersonalizationCSP registry key (requires SYSTEM context).
    4. Optionally prevents users from changing the wallpaper.

    IMPORTANT: The Personalization CSP (./Vendor/MSFT/Personalization) and the Group Policy path
    (HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization) both require Enterprise/Education.
    This script uses the PersonalizationCSP registry key at:
    HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP
    which works on Windows 11 Pro.

.NOTES
    File Name  : Set-DesktopWallpaperAndLockScreen.ps1
    Author     : CodeAndersen
    Version    : 1.1
    Created    : 2026-02-18
    Requires   : Windows 11 Pro, PowerShell 5.1+
    Context    : Run as SYSTEM in Intune (required for lock screen and policy enforcement)

    Intune deployment settings:
    - Run this script using the logged-on credentials: No
    - Run script in 64-bit PowerShell: Yes
    - Enforce script signature check: No (or sign the script)

.LINK
    https://learn.microsoft.com/en-us/windows/client-management/mdm/personalization-csp
    https://learn.microsoft.com/en-us/mem/intune/fundamentals/manage-shell-scripts
#>

#region Configuration
# ============================================================================
# CONFIGURE THESE VARIABLES BEFORE DEPLOYMENT
# ============================================================================

# URL to the desktop wallpaper image (e.g., Azure Blob Storage SAS URL or public URL)
# Supported formats: .jpg, .jpeg, .png, .bmp
# Set to $null or empty string to skip setting the desktop wallpaper
$WallpaperUrl = "https://xyz.blob.core.windows.net/brandingse/BG.jpg"

# URL to the lock screen image
# Set to $null or empty string to skip setting the lock screen
$LockScreenUrl = "https://xyz.blob.core.windows.net/brandingse/BG.jpg"

# Local folder where images will be stored on the device
$ImageFolder = "$env:ProgramData\CompanyImages"

# Wallpaper fit style: Fill, Fit, Stretch, Tile, Center, Span
$WallpaperStyle = "Center"

# Prevent users from changing the desktop wallpaper (set to $true to enforce)
# Note: The lock screen set via PersonalizationCSP is already enforced by Windows
# and cannot be changed by users through the Settings app.
$PreventWallpaperChange = $false

#endregion Configuration

#region Functions
# ============================================================================
# FUNCTIONS
# ============================================================================

function Write-Log {
    <#
    .SYNOPSIS
        Writes a log entry to the Intune Management Extension log directory.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("Info", "Warning", "Error")]
        [string]$Level = "Info"
    )

    $LogFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\Set-DesktopWallpaperAndLockScreen.log"
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"

    # Create log directory if it doesn't exist
    $LogDir = Split-Path -Path $LogFile -Parent
    if (-not (Test-Path -Path $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
    }

    Add-Content -Path $LogFile -Value $LogEntry -Encoding UTF8
    Write-Output $LogEntry
}

function Get-ImageFromUrl {
    <#
    .SYNOPSIS
        Downloads an image from a URL to a local path.
    .OUTPUTS
        Returns $true if the download was successful, $false otherwise.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Url,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    try {
        # Create destination directory if it doesn't exist
        $DestDir = Split-Path -Path $DestinationPath -Parent
        if (-not (Test-Path -Path $DestDir)) {
            New-Item -Path $DestDir -ItemType Directory -Force | Out-Null
            Write-Log -Message "Created directory: $DestDir"
        }

        # Use TLS 1.2 for secure connections
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        # Download the image
        Write-Log -Message "Downloading image from: $Url"
        $WebClient = New-Object System.Net.WebClient
        $WebClient.DownloadFile($Url, $DestinationPath)
        $WebClient.Dispose()

        if (Test-Path -Path $DestinationPath) {
            $FileSize = (Get-Item -Path $DestinationPath).Length
            Write-Log -Message "Image downloaded successfully to: $DestinationPath ($FileSize bytes)"
            return $true
        }
        else {
            Write-Log -Message "Download completed but file not found at: $DestinationPath" -Level "Error"
            return $false
        }
    }
    catch {
        Write-Log -Message "Failed to download image from $Url. Error: $($_.Exception.Message)" -Level "Error"
        return $false
    }
}

function Set-DesktopWallpaper {
    <#
    .SYNOPSIS
        Sets the desktop wallpaper for all current and future users using the
        SystemParametersInfo API and registry keys.
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$ImagePath,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Fill", "Fit", "Stretch", "Tile", "Center", "Span")]
        [string]$Style
    )

    # Map style names to registry values
    # WallpaperStyle: 10=Fill, 6=Fit, 2=Stretch, 0=Tile/Center/Span
    # TileWallpaper:  0=No tile, 1=Tile
    $StyleMap = @{
        "Fill"    = @{ WallpaperStyle = "10"; TileWallpaper = "0" }
        "Fit"     = @{ WallpaperStyle = "6";  TileWallpaper = "0" }
        "Stretch" = @{ WallpaperStyle = "2";  TileWallpaper = "0" }
        "Tile"    = @{ WallpaperStyle = "0";  TileWallpaper = "1" }
        "Center"  = @{ WallpaperStyle = "0";  TileWallpaper = "0" }
        "Span"    = @{ WallpaperStyle = "22"; TileWallpaper = "0" }
    }

    try {
        # Set wallpaper via PersonalizationCSP registry key (works on Windows 11 Pro)
        $CspRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
        if (-not (Test-Path -Path $CspRegPath)) {
            New-Item -Path $CspRegPath -Force | Out-Null
            Write-Log -Message "Created registry key: $CspRegPath"
        }
        Set-ItemProperty -Path $CspRegPath -Name "DesktopImagePath" -Value $ImagePath -Type String -Force
        Set-ItemProperty -Path $CspRegPath -Name "DesktopImageUrl" -Value $ImagePath -Type String -Force
        Set-ItemProperty -Path $CspRegPath -Name "DesktopImageStatus" -Value 1 -Type DWord -Force
        Write-Log -Message "Set desktop wallpaper via PersonalizationCSP registry key"

        # Set wallpaper for the DEFAULT user profile (applies to new users)
        $DefaultRegPath = "Registry::HKEY_USERS\.DEFAULT\Control Panel\Desktop"
        Set-ItemProperty -Path $DefaultRegPath -Name Wallpaper -Value $ImagePath -Force
        Set-ItemProperty -Path $DefaultRegPath -Name WallpaperStyle -Value $StyleMap[$Style].WallpaperStyle -Force
        Set-ItemProperty -Path $DefaultRegPath -Name TileWallpaper -Value $StyleMap[$Style].TileWallpaper -Force
        Write-Log -Message "Set wallpaper in DEFAULT user profile registry"

        # Set wallpaper for all currently loaded user profiles (logged-in users)
        $LoadedProfiles = Get-ChildItem "Registry::HKEY_USERS" |
            Where-Object { $_.PSChildName -match '^S-1-5-21-\d+-\d+-\d+-\d+$' }

        foreach ($UserProfile in $LoadedProfiles) {
            $UserRegPath = "Registry::HKEY_USERS\$($UserProfile.PSChildName)\Control Panel\Desktop"
            if (Test-Path $UserRegPath) {
                Set-ItemProperty -Path $UserRegPath -Name Wallpaper -Value $ImagePath -Force
                Set-ItemProperty -Path $UserRegPath -Name WallpaperStyle -Value $StyleMap[$Style].WallpaperStyle -Force
                Set-ItemProperty -Path $UserRegPath -Name TileWallpaper -Value $StyleMap[$Style].TileWallpaper -Force
                Write-Log -Message "Set wallpaper for user profile: $($UserProfile.PSChildName)"
            }
        }

        # Use SystemParametersInfo to apply wallpaper immediately for the current session
        # Note: When running as SYSTEM, this applies to the SYSTEM desktop.
        # The registry changes above ensure it applies when users log in.
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class Wallpaper {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);

    public const int SPI_SETDESKWALLPAPER = 0x0014;
    public const int SPIF_UPDATEINIFILE = 0x01;
    public const int SPIF_SENDCHANGE = 0x02;

    public static void SetWallpaper(string path) {
        SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, path, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
    }
}
"@
        [Wallpaper]::SetWallpaper($ImagePath)
        Write-Log -Message "Called SystemParametersInfo to set wallpaper: $ImagePath (Style: $Style)"

        return $true
    }
    catch {
        Write-Log -Message "Failed to set desktop wallpaper. Error: $($_.Exception.Message)" -Level "Error"
        return $false
    }
}

function Set-LockScreenImage {
    <#
    .SYNOPSIS
        Sets the lock screen image via the PersonalizationCSP registry key.
        Requires SYSTEM context. Works on Windows 11 Pro (unlike the Policies path
        or the Personalization CSP which both require Enterprise/Education).

    .NOTES
        Registry path: HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP
        Required values:
        - LockScreenImagePath (REG_SZ) = local file path to the image
        - LockScreenImageUrl  (REG_SZ) = local file path to the image (same value)
        - LockScreenImageStatus (REG_DWORD) = 1 (indicates image is ready)
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$ImagePath
    )

    try {
        # Use the PersonalizationCSP registry key (works on Windows 11 Pro)
        # This is NOT the same as:
        #   - ./Vendor/MSFT/Personalization (MDM CSP - Enterprise/Education only)
        #   - HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization (GP - Enterprise/Education only)
        $CspRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"

        if (-not (Test-Path -Path $CspRegPath)) {
            New-Item -Path $CspRegPath -Force | Out-Null
            Write-Log -Message "Created registry key: $CspRegPath"
        }

        # Set the lock screen image path and status
        Set-ItemProperty -Path $CspRegPath -Name "LockScreenImagePath" -Value $ImagePath -Type String -Force
        Set-ItemProperty -Path $CspRegPath -Name "LockScreenImageUrl" -Value $ImagePath -Type String -Force
        Set-ItemProperty -Path $CspRegPath -Name "LockScreenImageStatus" -Value 1 -Type DWord -Force
        Write-Log -Message "Set lock screen image via PersonalizationCSP: $ImagePath"

        # Disable Windows Spotlight on the lock screen to ensure our custom image is used
        # This uses per-user registry for all loaded profiles
        $LoadedProfiles = Get-ChildItem "Registry::HKEY_USERS" |
            Where-Object { $_.PSChildName -match '^S-1-5-21-\d+-\d+-\d+-\d+$' }

        foreach ($UserProfile in $LoadedProfiles) {
            $SpotlightRegPath = "Registry::HKEY_USERS\$($UserProfile.PSChildName)\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
            if (-not (Test-Path -Path $SpotlightRegPath)) {
                New-Item -Path $SpotlightRegPath -Force | Out-Null
            }
            Set-ItemProperty -Path $SpotlightRegPath -Name "DisableWindowsSpotlightFeatures" -Value 1 -Type DWord -Force
            Set-ItemProperty -Path $SpotlightRegPath -Name "DisableWindowsSpotlightOnLockScreen" -Value 1 -Type DWord -Force
            Write-Log -Message "Disabled Windows Spotlight for user profile: $($UserProfile.PSChildName)"
        }

        # Also set for the default profile (new users)
        $DefaultSpotlightPath = "Registry::HKEY_USERS\.DEFAULT\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
        if (-not (Test-Path -Path $DefaultSpotlightPath)) {
            New-Item -Path $DefaultSpotlightPath -Force | Out-Null
        }
        Set-ItemProperty -Path $DefaultSpotlightPath -Name "DisableWindowsSpotlightFeatures" -Value 1 -Type DWord -Force
        Set-ItemProperty -Path $DefaultSpotlightPath -Name "DisableWindowsSpotlightOnLockScreen" -Value 1 -Type DWord -Force
        Write-Log -Message "Disabled Windows Spotlight for DEFAULT user profile"

        return $true
    }
    catch {
        Write-Log -Message "Failed to set lock screen image. Error: $($_.Exception.Message)" -Level "Error"
        return $false
    }
}

function Set-PreventChanges {
    <#
    .SYNOPSIS
        Prevents users from changing the desktop wallpaper via ActiveDesktop policy registry key.

    .NOTES
        The lock screen image set via PersonalizationCSP is already enforced by Windows and
        cannot be changed by the user through the Settings app. No additional prevention is needed.
        The NoChangingLockScreen GP registry key (HKLM:\SOFTWARE\Policies\...) requires
        Enterprise/Education and is not used here.
    #>
    param (
        [bool]$PreventWallpaper = $false
    )

    try {
        if ($PreventWallpaper) {
            # Prevent changing desktop wallpaper via ActiveDesktop policy registry key
            $WpPolicyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop"
            if (-not (Test-Path -Path $WpPolicyPath)) {
                New-Item -Path $WpPolicyPath -Force | Out-Null
            }
            Set-ItemProperty -Path $WpPolicyPath -Name "NoChangingWallPaper" -Value 1 -Type DWord -Force
            Write-Log -Message "Prevented users from changing desktop wallpaper"
        }

        return $true
    }
    catch {
        Write-Log -Message "Failed to set prevention policies. Error: $($_.Exception.Message)" -Level "Error"
        return $false
    }
}

#endregion Functions

#region Main
# ============================================================================
# MAIN EXECUTION
# ============================================================================

Write-Log -Message "=========================================="
Write-Log -Message "Starting Set-DesktopWallpaperAndLockScreen"
Write-Log -Message "=========================================="
Write-Log -Message "Running as: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
Write-Log -Message "OS: $(Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty Caption)"

$ExitCode = 0

# --- Desktop Wallpaper ---
if (-not [string]::IsNullOrWhiteSpace($WallpaperUrl)) {
    $WallpaperFileName = [System.IO.Path]::GetFileName(($WallpaperUrl -split '\?')[0])
    $WallpaperLocalPath = Join-Path -Path $ImageFolder -ChildPath $WallpaperFileName

    Write-Log -Message "Processing desktop wallpaper..."
    $Downloaded = Get-ImageFromUrl -Url $WallpaperUrl -DestinationPath $WallpaperLocalPath

    if ($Downloaded) {
        $WallpaperSet = Set-DesktopWallpaper -ImagePath $WallpaperLocalPath -Style $WallpaperStyle
        if (-not $WallpaperSet) {
            Write-Log -Message "Failed to apply desktop wallpaper" -Level "Error"
            $ExitCode = 1
        }
    }
    else {
        Write-Log -Message "Skipping wallpaper - download failed" -Level "Error"
        $ExitCode = 1
    }
}
else {
    Write-Log -Message "No wallpaper URL configured - skipping desktop wallpaper"
}

# --- Lock Screen ---
if (-not [string]::IsNullOrWhiteSpace($LockScreenUrl)) {
    $LockScreenFileName = [System.IO.Path]::GetFileName(($LockScreenUrl -split '\?')[0])
    $LockScreenLocalPath = Join-Path -Path $ImageFolder -ChildPath $LockScreenFileName

    Write-Log -Message "Processing lock screen image..."

    # Only download if it's a different URL/file than the wallpaper
    if ($LockScreenUrl -eq $WallpaperUrl) {
        Write-Log -Message "Lock screen uses same image as wallpaper - reusing downloaded file"
        $LockScreenLocalPath = $WallpaperLocalPath
        $Downloaded = $true
    }
    else {
        $Downloaded = Get-ImageFromUrl -Url $LockScreenUrl -DestinationPath $LockScreenLocalPath
    }

    if ($Downloaded) {
        $LockScreenSet = Set-LockScreenImage -ImagePath $LockScreenLocalPath
        if (-not $LockScreenSet) {
            Write-Log -Message "Failed to apply lock screen image" -Level "Error"
            $ExitCode = 1
        }
    }
    else {
        Write-Log -Message "Skipping lock screen - download failed" -Level "Error"
        $ExitCode = 1
    }
}
else {
    Write-Log -Message "No lock screen URL configured - skipping lock screen"
}

# --- Prevent Changes (optional) ---
if ($PreventWallpaperChange) {
    Write-Log -Message "Applying change prevention policies..."
    Set-PreventChanges -PreventWallpaper $PreventWallpaperChange
}

Write-Log -Message "=========================================="
Write-Log -Message "Script completed with exit code: $ExitCode"
Write-Log -Message "=========================================="

exit $ExitCode

#endregion Main
