<#
        .SYNOPSIS
        Sets or updates the GroupTag for a device in Microsoft Autopilot using Microsoft Graph API.

        .DESCRIPTION
        This script utilizes Microsoft Graph API to manage Autopilot device GroupTags. It supports modern authentication
        methods and is designed for Microsoft Intune device management workflows. The script is part of a larger
        Microsoft Intune automation solution.

        .PARAMETER clientId
        The Entra ID Application (Client) ID for Microsoft Graph authentication.

        .PARAMETER clientSecret
        The Entra ID Application Client Secret for Microsoft Graph authentication.
        Note: Consider using certificate-based authentication for production environments.

        .PARAMETER TenantId
        The Entra ID Directory (Tenant) ID for Microsoft Graph authentication.

        .PARAMETER GroupTag
        The GroupTag to set for the device in Autopilot.

        .PARAMETER csvFile
        The path to the CSV file containing the device serial numbers.

        .NOTES
        Version:        1.0
        Author:         Hans Christian Andersen @codeandersen
        Creation Date:  2025-01-13
        Purpose/Change: Initial script development
        Repository:     https://github.com/codeandersen/intune
        
        Required application permissions for app registration:
        - DeviceManagementConfiguration.ReadWrite.All
        - DeviceManagementManagedDevices.ReadWrite.All
        - DeviceManagementServiceConfig.ReadWrite.All

        .EXAMPLE
        C:\PS> Set-AutopilotDeviceGroupTag.ps1 -clientId "your-client-id" -clientSecret "your-client-secret" -TenantId "your-tenant-id" -GroupTag "group-tag" -csvFile "path-to-csv-file"

        .LINK        
#>


param(
    [string]$clientId,
    [string]$clientSecret,
    [string]$TenantId,
    [string]$GroupTag,
    [string]$csvFile
)


# Check and install required modules
$modules = @('Microsoft.Graph.Authentication', 'WindowsAutoPilotIntune')
foreach ($module in $modules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Force -Scope CurrentUser
    }
}

# Import the modules
Import-Module Microsoft.Graph.Authentication
Import-Module WindowsAutoPilotIntune

$SecuredPasswordPassword = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force

$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $clientId, $SecuredPasswordPassword

Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential

function Write-LogMessage {
    param(
        [string]$serialNumber,
        [string]$autopilotId,
        [string]$message
    )
    $logFilePath = "./AutopilotLog.csv"
    if (-not (Test-Path $logFilePath)) {
        "date,serialnumber,autopilotid,message" | Out-File -FilePath $logFilePath
    }
    $logEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$serialNumber,$autopilotId,$message"
    Add-Content -Path $logFilePath -Value $logEntry
}

# Import the CSV file
if (Test-Path $csvFile) {
    $devices = Import-Csv -Path $csvFile

    foreach ($device in $devices) {
        $serialNumber = $device.SerialNumber

        # Check if the device exists
        $AutopilotDevice = Get-AutopilotDevice -serial $serialNumber
        if (-not $AutopilotDevice) {
            # Log to CSV if the device does not exist
            Write-LogMessage -serialNumber $serialNumber -autopilotId "" -message "Device not found in Autopilot"
            Write-Host "Autopilot device not found for serial number: $serialNumber"
            continue
        }

        # Set the group tag if the device exists
        Write-Host "Processing serial number: $serialNumber"
        $currentGroupTag = $AutopilotDevice.groupTag
        if ($currentGroupTag -eq $GroupTag) {
            Write-LogMessage -serialNumber $serialNumber -autopilotId $AutopilotDevice.id -message "GroupTag already set to $GroupTag"
            continue
        }

        Set-AutopilotDevice -id $AutopilotDevice.id -groupTag $GroupTag

        # Log to CSV
        Write-LogMessage -serialNumber $serialNumber -autopilotId $AutopilotDevice.id -message "GroupTag set to $GroupTag"
    }
} else {
    Write-Error "The CSV file path '$csvFile' does not exist."
}
