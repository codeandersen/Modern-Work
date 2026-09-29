
#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [guid]$GroupId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [switch]$UseDeviceCodeForIntune
)

$ErrorActionPreference = 'Stop'

$script:GraphScopes = @(
    'GroupMember.Read.All',
    'Device.Read.All',
    'DeviceManagementManagedDevices.Read.All',
    'DeviceManagementServiceConfig.Read.All'
)

$script:CsvColumns = @(
    'DeviceName', 'EntraObjectId', 'AzureAdDeviceId', 'OperatingSystem', 'OperatingSystemVersion',
    'TrustType', 'LastSignIn', 'IntuneManaged', 'SerialNumber', 'InAutopilot', 'MatchMethod',
    'AutopilotDeviceId', 'ZtdIdOnEntraObject', 'GroupTag', 'Manufacturer', 'Model', 'EnrollmentState',
    'EnrollmentProfileName', 'ProfileAssignmentStatus', 'ProfileAssignedDateTime', 'AutopilotLastContacted'
)

function Initialize-GraphModules {
    $module = 'Microsoft.Graph.Authentication'
    if (-not (Get-Module -ListAvailable -Name $module)) {
        throw "Required module '$module' is not installed. Install it with: Install-Module $module -Scope CurrentUser"
    }
    Import-Module $module -ErrorAction Stop
}

function Connect-GraphSession {
    param(
        [Parameter(Mandatory)][string[]]$Scopes,
        [string]$ClientId,
        [string]$TenantId,
        [switch]$UseDeviceCode
    )
    $connectParams = @{
        Scopes       = $Scopes
        ContextScope = 'Process'
        NoWelcome    = $true
    }
    if ($UseDeviceCode) { $connectParams.UseDeviceCode = $true }
    if (-not [string]::IsNullOrWhiteSpace($ClientId)) { $connectParams.ClientId = $ClientId }
    if (-not [string]::IsNullOrWhiteSpace($TenantId)) { $connectParams.TenantId = $TenantId }
    Connect-MgGraph @connectParams | Out-Null

    $ctx = Get-MgContext
    Write-Host "Connected as $($ctx.Account) (TenantId: $($ctx.TenantId))"
    Write-Host "Scopes: $(@($ctx.Scopes) -join ', ')"
}

function Get-GraphProperty {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    return $Object.$Name
}

function Invoke-GraphWithRetry {
    param([scriptblock]$Script, [int]$MaxRetries = 5)
    $attempt = 0
    while ($true) {
        try {
            return & $Script
        }
        catch {
            $attempt++
            $msg = "$_"
            $isThrottle = $msg -match '429|503|throttl|TooManyRequests|ServiceUnavailable'
            if (-not $isThrottle -or $attempt -ge $MaxRetries) { throw }
            $delay = [Math]::Min(60, [Math]::Pow(2, $attempt))
            Write-Warning "Throttled/transient error (attempt $attempt of $MaxRetries), retrying in $delay s..."
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-GraphAllPages {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$Headers
    )
    $items = New-Object 'System.Collections.Generic.List[object]'
    $next = $Uri
    while ($next) {
        $requestParams = @{ Method = 'GET'; Uri = $next; OutputType = 'PSObject' }
        if ($Headers) { $requestParams.Headers = $Headers }
        $resp = Invoke-GraphWithRetry { Invoke-MgGraphRequest @requestParams }
        if ($null -eq $resp) { throw "Graph request to '$next' returned no response; aborting export." }
        $hasValue = $false
        if ($resp -is [System.Collections.IDictionary]) { $hasValue = $resp.Contains('value') }
        elseif ($null -ne $resp.PSObject.Properties['value']) { $hasValue = $true }
        if (-not $hasValue) { throw "Graph response from '$next' has no 'value' collection; aborting export." }
        foreach ($v in @($resp.value)) {
            if ($null -ne $v) { [void]$items.Add($v) }
        }
        $nextLink = Get-GraphProperty -Object $resp -Name '@odata.nextLink'
        $next = if ($nextLink) { [string]$nextLink } else { $null }
    }
    return $items
}

function Get-ZtdIdFromPhysicalIds {
    param($Device)
    $physicalIds = @(Get-GraphProperty -Object $Device -Name 'physicalIds')
    foreach ($physicalId in $physicalIds) {
        $s = [string]$physicalId
        if ($s -match '^\[ZTDID\]:(.+)$') {
            return $Matches[1].Trim()
        }
    }
    return $null
}

function Test-IsGuid {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $g = [guid]::Empty
    return [guid]::TryParse($Value, [ref]$g)
}

function Get-GroupDeviceMembers {
    param([guid]$Id)
    $select = 'id,deviceId,displayName,operatingSystem,operatingSystemVersion,trustType,approximateLastSignInDateTime,physicalIds,enrollmentProfileName'
    $uri = "https://graph.microsoft.com/v1.0/groups/$Id/members/microsoft.graph.device?`$select=$select&`$top=999&`$count=true"
    return Get-GraphAllPages -Uri $uri -Headers @{ ConsistencyLevel = 'eventual' }
}

function Get-ManagedDevices {
    $uri = 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$select=id,azureADDeviceId,serialNumber&$top=999'
    return Get-GraphAllPages -Uri $uri
}

function Get-AutopilotDevices {
    $uri = 'https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities?$top=200'
    return Get-GraphAllPages -Uri $uri
}

function Add-IndexEntry {
    param(
        [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[object]]]$Index,
        [string]$Key,
        $Item
    )
    if ([string]::IsNullOrWhiteSpace($Key)) { return }
    if (-not $Index.ContainsKey($Key)) {
        $Index[$Key] = New-Object 'System.Collections.Generic.List[object]'
    }
    [void]$Index[$Key].Add($Item)
}

function Resolve-AutopilotMatch {
    param(
        $EntraDevice,
        [string]$ZtdId,
        $ApById,
        $ApByAadId,
        $ApBySerial,
        $ManagedByAadId
    )

    if (-not [string]::IsNullOrWhiteSpace($ZtdId)) {
        $candidates = $null
        if ($ApById.ContainsKey($ZtdId)) { $candidates = $ApById[$ZtdId] }
        if ($null -ne $candidates -and $candidates.Count -eq 1) {
            return @{ Device = $candidates[0]; Method = 'ZtdId' }
        }
        if ($null -ne $candidates -and $candidates.Count -gt 1) {
            return @{ Device = $null; Method = 'AmbiguousZtdId' }
        }
    }

    $entraDeviceId = [string](Get-GraphProperty -Object $EntraDevice -Name 'deviceId')
    if (Test-IsGuid $entraDeviceId) {
        $candidates = $null
        if ($ApByAadId.ContainsKey($entraDeviceId)) { $candidates = $ApByAadId[$entraDeviceId] }
        if ($null -ne $candidates -and $candidates.Count -eq 1) {
            return @{ Device = $candidates[0]; Method = 'AzureAdDeviceId' }
        }
        if ($null -ne $candidates -and $candidates.Count -gt 1) {
            return @{ Device = $null; Method = 'AmbiguousDeviceId' }
        }
    }

    $serial = $null
    $serialFallbackOk = $false
    if (Test-IsGuid $entraDeviceId -and $ManagedByAadId.ContainsKey($entraDeviceId)) {
        $managed = $ManagedByAadId[$entraDeviceId]
        $serials = @(
            $managed |
                ForEach-Object { [string](Get-GraphProperty -Object $_ -Name 'serialNumber') } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -Unique
        )
        if ($serials.Count -eq 1) {
            $serial = $serials[0]
            $serialFallbackOk = $true
        }
    }
    if ($serialFallbackOk) {
        $candidates = $null
        if ($ApBySerial.ContainsKey($serial)) { $candidates = $ApBySerial[$serial] }
        if ($null -ne $candidates -and $candidates.Count -eq 1) {
            $candidate = $candidates[0]
            $linkedIds = @(
                [string](Get-GraphProperty -Object $candidate -Name 'azureAdDeviceId'),
                [string](Get-GraphProperty -Object $candidate -Name 'azureActiveDirectoryDeviceId')
            ) | Where-Object { Test-IsGuid $_ }
            $linkedIdMismatch = $linkedIds.Count -gt 0 -and -not ($linkedIds | Where-Object { $_ -ieq $entraDeviceId })
            $method = if ($linkedIdMismatch) { 'SerialNumberLinkedIdMismatch' } else { 'SerialNumber' }
            return @{ Device = $candidate; Method = $method }
        }
        elseif ($null -ne $candidates -and $candidates.Count -gt 1) {
            return @{ Device = $null; Method = 'AmbiguousSerial' }
        }
    }

    return @{ Device = $null; Method = 'None' }
}

if ($UseDeviceCodeForIntune -and [string]::IsNullOrWhiteSpace($ClientId)) {
    throw '-UseDeviceCodeForIntune requires -ClientId (device code is used only for the stage 2 Intune/Autopilot connection); aborting before authentication.'
}

Initialize-GraphModules

$script:EntraScopes = @('GroupMember.Read.All', 'Device.Read.All')
$script:IntuneScopes = @('DeviceManagementManagedDevices.Read.All', 'DeviceManagementServiceConfig.Read.All')

$useTwoStageAuth = -not [string]::IsNullOrWhiteSpace($ClientId)
$firstTenantId = $null

if ($useTwoStageAuth) {
    Write-Host 'Stage 1: connecting with the default Graph PowerShell client to read group membership...'
    Connect-GraphSession -Scopes $script:EntraScopes -TenantId $TenantId
    $ctx1 = Get-MgContext
    if ($null -eq $ctx1 -or [string]::IsNullOrWhiteSpace($ctx1.TenantId)) {
        throw 'Stage 1 connection returned no tenant context; aborting.'
    }
    $firstTenantId = [string]$ctx1.TenantId
    if (-not [string]::IsNullOrWhiteSpace($TenantId) -and $firstTenantId -ine $TenantId) {
        throw "Stage 1 connected tenant '$firstTenantId' does not match requested -TenantId '$TenantId'; aborting before group fetch."
    }
}
else {
    Connect-GraphSession -Scopes $script:GraphScopes -TenantId $TenantId
}

Write-Host "Fetching group device members ($GroupId)..."
$entraDevices = Get-GroupDeviceMembers -Id $GroupId
Write-Host "Group device members: $($entraDevices.Count)"

$outDir = Split-Path -Parent $OutputPath
if (-not [string]::IsNullOrWhiteSpace($outDir) -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

if ($entraDevices.Count -eq 0) {
    $header = '"' + ($script:CsvColumns -join '","') + '"'
    $absOut = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    [System.IO.File]::WriteAllText($absOut, $header + "`r`n", [System.Text.UTF8Encoding]::new($false))
    Write-Host "No device members in group. Header-only CSV written to $OutputPath"
    return
}

if ($useTwoStageAuth) {
    if ($UseDeviceCodeForIntune) {
        Write-Host 'Stage 2: reconnecting with the supplied ClientId using device code. Device code instructions will appear in this console - follow them to sign in.'
    }
    else {
        Write-Host 'Stage 2: reconnecting with the supplied ClientId to read Intune and Autopilot data...'
    }
    Disconnect-MgGraph | Out-Null
    Connect-GraphSession -Scopes $script:IntuneScopes -ClientId $ClientId -TenantId $firstTenantId -UseDeviceCode:$UseDeviceCodeForIntune
    $ctx2 = Get-MgContext
    if ($null -eq $ctx2 -or [string]::IsNullOrWhiteSpace($ctx2.TenantId)) {
        throw 'Stage 2 connection returned no tenant context; aborting export.'
    }
    if ($ctx2.TenantId -ine $firstTenantId) {
        throw "Stage 2 tenant '$($ctx2.TenantId)' does not match stage 1 tenant '$firstTenantId'; aborting export."
    }
}

Write-Host "Fetching Intune managed devices..."
$managedDevices = Get-ManagedDevices
Write-Host "Managed devices: $($managedDevices.Count)"

Write-Host "Fetching Autopilot device identities (beta)..."
$autopilotDevices = Get-AutopilotDevices
Write-Host "Autopilot devices: $($autopilotDevices.Count)"

$apById = New-Object 'System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[object]]' ([StringComparer]::OrdinalIgnoreCase)
$apByAadId = New-Object 'System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[object]]' ([StringComparer]::OrdinalIgnoreCase)
$apBySerial = New-Object 'System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[object]]' ([StringComparer]::OrdinalIgnoreCase)

foreach ($ap in $autopilotDevices) {
    $apId = [string](Get-GraphProperty -Object $ap -Name 'id')
    if (Test-IsGuid $apId) { Add-IndexEntry -Index $apById -Key $apId -Item $ap }
    $aad1 = [string](Get-GraphProperty -Object $ap -Name 'azureAdDeviceId')
    $aad2 = [string](Get-GraphProperty -Object $ap -Name 'azureActiveDirectoryDeviceId')
    if (Test-IsGuid $aad1) { Add-IndexEntry -Index $apByAadId -Key $aad1 -Item $ap }
    if (Test-IsGuid $aad2 -and $aad2 -ne $aad1) { Add-IndexEntry -Index $apByAadId -Key $aad2 -Item $ap }
    $serial = [string](Get-GraphProperty -Object $ap -Name 'serialNumber')
    if (-not [string]::IsNullOrWhiteSpace($serial)) { Add-IndexEntry -Index $apBySerial -Key $serial -Item $ap }
}

$managedByAadId = New-Object 'System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[object]]' ([StringComparer]::OrdinalIgnoreCase)
foreach ($md in $managedDevices) {
    $aadId = [string](Get-GraphProperty -Object $md -Name 'azureADDeviceId')
    if (Test-IsGuid $aadId) { Add-IndexEntry -Index $managedByAadId -Key $aadId -Item $md }
}

$rows = New-Object 'System.Collections.Generic.List[object]'
foreach ($device in $entraDevices) {
    $ztdId = Get-ZtdIdFromPhysicalIds -Device $device
    $match = Resolve-AutopilotMatch -EntraDevice $device -ZtdId $ztdId -ApById $apById -ApByAadId $apByAadId -ApBySerial $apBySerial -ManagedByAadId $managedByAadId
    $ap = $match.Device
    $method = [string]$match.Method
    $inAutopilot = if ($method -like 'Ambiguous*') { 'Ambiguous' } elseif ($null -ne $ap) { 'True' } else { 'False' }

    $entraDeviceId = [string](Get-GraphProperty -Object $device -Name 'deviceId')
    $intuneManaged = (Test-IsGuid $entraDeviceId) -and $managedByAadId.ContainsKey($entraDeviceId)
    $entraSerial = $null
    if (Test-IsGuid $entraDeviceId -and $managedByAadId.ContainsKey($entraDeviceId)) {
        $serials = @(
            $managedByAadId[$entraDeviceId] |
                ForEach-Object { [string](Get-GraphProperty -Object $_ -Name 'serialNumber') } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -Unique
        )
        if ($serials.Count -eq 1) { $entraSerial = $serials[0] }
    }

    $apSerial = $null
    if ($null -ne $ap) {
        $apSerial = [string](Get-GraphProperty -Object $ap -Name 'serialNumber')
        if ([string]::IsNullOrWhiteSpace($apSerial)) { $apSerial = $null }
    }
    $serialNumber = if ($apSerial) { $apSerial } elseif ($entraSerial) { $entraSerial } else { $null }

    $profileName = [string](Get-GraphProperty -Object $device -Name 'enrollmentProfileName')
    if ([string]::IsNullOrWhiteSpace($profileName)) { $profileName = $null }
    $apAssignmentStatus = if ($null -ne $ap) { [string](Get-GraphProperty -Object $ap -Name 'deploymentProfileAssignmentStatus') } else { $null }

    $rows.Add([PSCustomObject][ordered]@{
        DeviceName               = [string](Get-GraphProperty -Object $device -Name 'displayName')
        EntraObjectId            = [string](Get-GraphProperty -Object $device -Name 'id')
        AzureAdDeviceId          = $entraDeviceId
        OperatingSystem          = [string](Get-GraphProperty -Object $device -Name 'operatingSystem')
        OperatingSystemVersion   = [string](Get-GraphProperty -Object $device -Name 'operatingSystemVersion')
        TrustType                = [string](Get-GraphProperty -Object $device -Name 'trustType')
        LastSignIn               = [string](Get-GraphProperty -Object $device -Name 'approximateLastSignInDateTime')
        IntuneManaged            = "$intuneManaged"
        SerialNumber             = $serialNumber
        InAutopilot              = $inAutopilot
        MatchMethod              = $method
        AutopilotDeviceId        = if ($null -ne $ap) { [string](Get-GraphProperty -Object $ap -Name 'id') } else { $null }
        ZtdIdOnEntraObject       = $ztdId
        GroupTag                 = if ($null -ne $ap) { [string](Get-GraphProperty -Object $ap -Name 'groupTag') } else { $null }
        Manufacturer             = if ($null -ne $ap) { [string](Get-GraphProperty -Object $ap -Name 'manufacturer') } else { $null }
        Model                    = if ($null -ne $ap) { [string](Get-GraphProperty -Object $ap -Name 'model') } else { $null }
        EnrollmentState          = if ($null -ne $ap) { [string](Get-GraphProperty -Object $ap -Name 'enrollmentState') } else { $null }
        EnrollmentProfileName    = $profileName
        ProfileAssignmentStatus  = $apAssignmentStatus
        ProfileAssignedDateTime  = if ($null -ne $ap) { [string](Get-GraphProperty -Object $ap -Name 'deploymentProfileAssignedDateTime') } else { $null }
        AutopilotLastContacted   = if ($null -ne $ap) { [string](Get-GraphProperty -Object $ap -Name 'lastContactedDateTime') } else { $null }
    }) | Out-Null
}

$rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "Exported $($rows.Count) device row(s) to $OutputPath"
