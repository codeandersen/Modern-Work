<#
        .SYNOPSIS
        Checks if a device exists in Autopilot using its serial number and outputs the device's Autopilot ID and GroupTag if found.

        .DESCRIPTION
        This script connects to Microsoft Graph and checks for the existence of a device in Autopilot based on the provided serial number. If the device exists, it outputs the device's Autopilot ID and GroupTag.

        .PARAMETER clientId
        The Client ID for authentication.

        .PARAMETER clientSecret
        The Client Secret for authentication.

        .PARAMETER TenantId
        The Tenant ID for Microsoft Graph.

        .PARAMETER serialnumber
        The serial number of the device to check in Autopilot.

        .EXAMPLE
        C:\PS> Check-AutopilotDeviceGroupTag.ps1 -clientId "your-client-id" -clientSecret "your-client-secret" -TenantId "your-tenant-id" -serialnumber "computer serial number"

        .COPYRIGHT
        MIT License, feel free to distribute and use as you like, please leave author information.

        .LINK
        BLOG: http://www.hcconsult.dk
        Twitter: @dk_hcandersen

        .DISCLAIMER
        This script is provided AS-IS, with no warranty - Use at own risk.
#>

param(
    [string]$clientId,
    [string]$clientSecret,
    [string]$TenantId,
    [string]$serialnumber
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

# Convert client secret to secure string
$SecuredPassword = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
$ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $clientId, $SecuredPassword

# Connect to Microsoft Graph
Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientSecretCredential

# Check for the device
$AutopilotDevice = Get-AutopilotDevice -serial $serialnumber
$date = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
if ($AutopilotDevice) {
    #$currentGroupTag = $AutopilotDevice.groupTag
    Write-Host "$date Device with serial number $serialnumber exists in Autopilot with Autopilot ID $($AutopilotDevice.id) and GroupTag $($AutopilotDevice.groupTag)"
} else {
    Write-Host "$date Device with serial number $serialnumber doesn't exist in Autopilot"
}
