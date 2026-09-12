<#
        .SYNOPSIS
        Fixes the SSL Client Certificate Search Criteria for Microsoft Intune-managed devices.

        .DESCRIPTION
        This script verifies and corrects the SSL Client Certificate Search Criteria in the Windows Registry
        for Intune-managed devices. It ensures the correct format using the device's EntDMID and proper
        certificate store path. The script includes comprehensive logging and error handling.

        Key functions:
        - Retrieves Intune enrollment GUID from registry
        - Gets EntDMID from DMClient settings
        - Verifies and updates SSL Client Certificate Search Criteria
        - Provides detailed logging of all operations

        .NOTES
        Version:        1.0
        Author:         Hans Christian Andersen @codeandersen
        Creation Date:  2025-01-27
        Purpose/Change: Initial script development
        Repository:     https://github.com/codeandersen/intune

        Registry paths used:
        - HKLM:\SOFTWARE\Microsoft\Enrollments
        - HKLM:\SOFTWARE\Microsoft\Enrollments\[GUID]\DMClient\MS DM Server
        - HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\[GUID]\Protected

        .EXAMPLE
        C:\PS> .\SslClientCertSearchCriteriaFix.ps1
        Runs the script to verify and fix the SSL Client Certificate Search Criteria.

        .LINK
        https://github.com/codeandersen/intune

        .NOTES
        The script requires administrative privileges to modify registry entries.
        Logging is performed in the Intune Management Extension log directory.
#>

# Initialize Logging
$LogPath = "$env:SystemDrive\ProgramData\Microsoft\IntuneManagementExtension\Logs"
$LogFile = "IntuneSslClientCertSearchCriteriaFix_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
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

Write-Log "Script started - Checking and fixing Intune SSL Client Certificate Search Criteria" -Level Information -Step 1

#Get guid of the intune enrollment
$ProviderRegistryPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Enrollments"
$ProviderPropertyName = "ProviderID"
$ProviderPropertyValue = "MS DM Server"

Write-Log "Searching for Intune enrollment GUID in registry" -Level Information -Step 2

try {
    $GUID = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Enrollments" | 
        Get-ItemProperty | 
        Where-Object { $_.$ProviderPropertyName -eq $ProviderPropertyValue } |
        Select-Object -ExpandProperty PSChildName
    
    if ($GUID) {
        Write-Log "Successfully found Intune enrollment GUID: $GUID" -Level Information -Step 2
    } else {
        Write-Log "No Intune enrollment GUID found in registry" -Level Error -Step 2
        Exit 1
    }
} catch {
    Write-Log "Error accessing registry to find Intune enrollment GUID: $_" -Level Error -Step 2
    Exit 1
}

Write-Log "Retrieving EntDMID and Search Criteria from registry" -Level Information -Step 3
try {
    $entdmid = Get-ItemPropertyValue "HKLM:\SOFTWARE\Microsoft\Enrollments\$GUID\DMClient\MS DM Server" -Name EntDMID -ErrorAction Stop
    Write-Log "Successfully retrieved EntDMID: $entdmid" -Level Information -Step 3
} catch {
    Write-Log "Failed to retrieve EntDMID from registry: $_" -Level Error -Step 3
    Exit 1
}

try {
    $SslClientCertSearchCriteria = Get-ItemPropertyValue "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$GUID\Protected" -Name SslClientCertSearchCriteria -ErrorAction Stop
    Write-Log "Current Search Criteria found in registry" -Level Information -Step 3
} catch {
    Write-Log "Search Criteria not found in registry, will need to be created" -Level Warning -Step 3
    $SslClientCertSearchCriteria = ""
}

$SslClientCertSearchCriteriaGood = "Subject=CN%3d$entdmid&Stores=MY%5CSystem"

Write-Log "Analyzing Search Criteria configuration:" -Level Information -Step 4
Write-Log "Current value: $SslClientCertSearchCriteria" -Level Information -Step 4
Write-Log "Required value: $SslClientCertSearchCriteriaGood" -Level Information -Step 4

if($SslClientCertSearchCriteria -ne $SslClientCertSearchCriteriaGood){
    Write-Log "Search Criteria needs to be updated" -Level Warning -Step 4
    Write-Log "Attempting to set correct Search Criteria in registry" -Level Information -Step 5
    try {
        New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$GUID\Protected" -Name SslClientCertSearchCriteria -PropertyType String -Value $SslClientCertSearchCriteriaGood -Force
        Write-Log "Successfully updated Search Criteria in registry" -Level Information -Step 5
    }
    catch {
        Write-Log "Failed to update Search Criteria in registry: $_" -Level Error -Step 5
        Exit 1
    }
} else {
    Write-Log "Search Criteria is already correctly configured" -Level Information -Step 4
}

Write-Log "Script completed - SSL Client Certificate Search Criteria verification and update process finished" -Level Information -Step 6
