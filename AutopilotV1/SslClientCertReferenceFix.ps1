<#
        .SYNOPSIS
        Fixes the SSL Client Certificate Reference for Microsoft Intune-managed devices.

        .DESCRIPTION
        This script verifies and corrects the SSL Client Certificate Reference in the Windows Registry
        for Intune-managed devices. It ensures the correct format using the device's certificate thumbprint
        and proper certificate store path. The script includes comprehensive logging and error handling.

        Key functions:
        - Retrieves Intune enrollment GUID from registry
        - Gets certificate thumbprint from enrollment
        - Verifies and updates SSL Client Certificate Reference
        - Provides detailed logging of all operations

        .NOTES
        Version:        1.0
        Author:         Hans Christian Andersen @codeandersen
        Creation Date:  2025-01-27
        Purpose/Change: Initial script development
        Repository:     https://github.com/codeandersen/intune

        Registry paths used:
        - HKLM:\SOFTWARE\Microsoft\Enrollments
        - HKLM:\SOFTWARE\Microsoft\Enrollments\[GUID]
        - HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\[GUID]

        .EXAMPLE
        C:\PS> .\SslClientCertReferenceFix.ps1
        Runs the script to verify and fix the SSL Client Certificate Reference.

        .LINK
        https://github.com/codeandersen/intune

        .NOTES
        The script requires administrative privileges to modify registry entries.
        Logging is performed in the Intune Management Extension log directory.
#>


# Initialize Logging
$LogPath = "$env:SystemDrive\ProgramData\Microsoft\IntuneManagementExtension\Logs"
$LogFile = "IntuneSslClientCertReferenceFix_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:LogFilePath = Join-Path -Path $LogPath -ChildPath $LogFile

# Create log directory if it doesn't exist
if (-not (Test-Path -Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

# Function for consistent logging
Function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Information', 'Warning', 'Error')]
        [string]$Level = 'Information',
        [Parameter(Mandatory = $false)]
        [int]$Step = 0
    )
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $StepText = if ($Step -gt 0) { "Step $Step : " } else { "" }
    $LogEntry = "$TimeStamp [$Level] $StepText$Message"
    
    # Write to log file
    Add-Content -Path $script:LogFilePath -Value $LogEntry
    
    # Enhanced console output with clear step indication
    $ConsolePrefix = if ($Step -gt 0) { 
        Write-Host "Step $Step" -NoNewline -ForegroundColor Cyan
        Write-Host " | " -NoNewline -ForegroundColor White
    }
    
    # Write message to console with appropriate color
    switch ($Level) {
        'Error' { Write-Host $Message -ForegroundColor Red }
        'Warning' { Write-Host $Message -ForegroundColor Yellow }
        'Information' { Write-Host $Message -ForegroundColor Green }
    }
}

Write-Log "Script execution started" -Level Information -Step 1

#Get guid of the intune enrollment
$ProviderRegistryPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Enrollments"
$ProviderPropertyName = "ProviderID"
$ProviderPropertyValue = "MS DM Server"

Write-Log "Searching for Intune enrollment GUID" -Level Information -Step 2

try {
    $GUID = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Enrollments" | 
        Get-ItemProperty | 
        Where-Object { $_.$ProviderPropertyName -eq $ProviderPropertyValue } |
        Select-Object -ExpandProperty PSChildName
    
    if ($GUID) {
        Write-Log "Successfully found Intune enrollment GUID: $GUID" -Level Information -Step 2
    } else {
        Write-Log "No Intune enrollment GUID found" -Level Error -Step 2
        Exit 1
    }
} catch {
    Write-Log "Error while searching for Intune enrollment GUID: $_" -Level Error -Step 2
    Exit 1
}

Write-Log "Retrieving certificate information" -Level Information -Step 3
try {
    $SslClientCertReference = Get-ItemPropertyValue HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$GUID -name SslClientCertReference -ErrorAction Stop
} catch {
    Write-Log "SslClientCertReference not found, setting to empty string" -Level Warning -Step 3
    $SslClientCertReference = ""
}
$Thumbprint = Get-ItemPropertyValue HKLM:\SOFTWARE\Microsoft\Enrollments\$GUID -Name DMPCertThumbPrint
$cert = Get-ChildItem Cert:\LocalMachine\My\ | Where-Object {$_.Issuer -Like "*Intune MDM*"}

Write-Log "Current SslClientCertReference: $SslClientCertReference" -Level Information -Step 3
Write-Log "Certificate Thumbprint: $Thumbprint" -Level Information -Step 3

Write-Log "Checking certificate reference" -Level Information -Step 4
if($SslClientCertReference -notlike "*$thumbprint*"){
    Write-Log "SslClientCertReference is NOT set correctly" -Level Warning -Step 4
    Write-Log "Current Reference: $SslClientCertReference" -Level Warning -Step 4
    Write-Log "Expected Thumbprint: $Thumbprint" -Level Warning -Step 4
    Write-Log "Setting SslClientCertReference to $Thumbprint" -Level Information -Step 5
    try {
        New-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$GUID -Name SslClientCertReference -PropertyType String -Value "MY;System;$Thumbprint" -Force
        Write-Log "Successfully updated SslClientCertReference" -Level Information -Step 5
    }
    catch {
        Write-Log "Failed to update SslClientCertReference: $_" -Level Error -Step 5
    }
} else {
    Write-Log "SslClientCertReference is set correctly" -Level Information -Step 4
}

Write-Log "Script execution completed" -Level Information -Step 6
