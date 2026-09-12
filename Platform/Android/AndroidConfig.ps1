Install-Module Microsoft.Graph -Scope CurrentUser
Get-MgOrganization
Disconnect-MgGraph

Connect-MgGraph -Scopes "DeviceManagementConfiguration.ReadWrite.All"

Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All"

policyId = 2a616524-05fe-4394-9d36-c608d4d2279c

Get-MgDeviceManagementDeviceConfiguration -DeviceConfigurationId "b54ded64-fe1f-4870-91ab-ef82a23fe60f"


Get-MgDeviceManagementDeviceConfiguration | Select-Object Id, DisplayName














# Get your current token
$token = (Get-MgContext).AccessToken

# Set headers
$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type"  = "application/json"
}

# Replace with your actual config ID
$configId = "2a616524-05fe-4394-9d36-c608d4d2279c"

# Send GET request
$response = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$configId" -Headers $headers -Method GET

# View the result
$response | Format-List






$profile = @{
  "@odata.type" = "#microsoft.graph.androidDeviceOwnerOEMConfigConfiguration"
  displayName   = "OEMConfig Demo"
  description   = "Created via PowerShell"
  packageId     = "com.samsung.android.knox.kpecore"
  payloadJson   = '{ "com.samsung.android.knox.kpecore": { "policy": { "wifi": { "ssid": "CorpWiFi", "password": "supersecret" }}}}'
}

New-MgDeviceManagementDeviceConfiguration -BodyParameter $profile

