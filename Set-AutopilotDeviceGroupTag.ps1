<#
.SYNOPSIS
    WPF GUI to set Windows Autopilot device group tags (delegated Graph auth).

.DESCRIPTION
    Interactive tool for large tenants: connect with Client ID / Tenant ID, load Autopilot
    devices with progress, client-side pagination/sort/filter, serial search, CSV import
    (serial numbers only), and apply one Group Tag to selected devices.

.EXAMPLE
    powershell.exe -STA -NoProfile -File .\Set-AutopilotDeviceGroupTag.ps1

.NOTES
    Run in STA mode (Windows PowerShell 5.1 default is STA; PowerShell 7 may need -STA).

    Modules: Microsoft.Graph.Authentication only (Invoke-MgGraphRequest).
    WindowsAutoPilotIntune is NOT required.

    Entra app (delegated interactive browser - no device code, no client secret):
    - API permission: DeviceManagementServiceConfig.ReadWrite.All (Delegated) + admin consent
    - Authentication -> Mobile and desktop applications -> http://localhost
    - Authentication -> Allow public client flows = Yes
    - Do not use a client secret for this GUI

    User needs Intune rights to manage Autopilot devices.

    CSV import columns: SerialNumber (aliases: Serial, Serial Number)
    Group tag always comes from the UI text box.

.COPYRIGHT
    MIT License. Author info: http://www.hcconsult.dk

.DISCLAIMER
    Provided AS-IS, with no warranty - Use at own risk.
#>

#Requires -Version 5.1

# StrictMode disabled: WPF event/scriptblock variable capture is unreliable under StrictMode
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Strongly typed row: avoids PSCustomObject/Sort-Object reflection issues on large tenants (20k+ devices)
if (-not ('AutopilotRow' -as [type])) {
    Add-Type -TypeDefinition @'
public class AutopilotRow
{
    public string Id { get; set; }
    public string SerialNumber { get; set; }
    public string GroupTag { get; set; }
    public string Model { get; set; }
    public string Manufacturer { get; set; }
}
'@
}

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = if (Get-Command pwsh -ErrorAction SilentlyContinue) { (Get-Command pwsh).Source } else { (Get-Command powershell).Source }
    $psi.Arguments = "-STA -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $psi.UseShellExecute = $false
    [void][System.Diagnostics.Process]::Start($psi)
    return
}

$script:ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$script:SuccessLogPath = Join-Path $script:ScriptRoot 'AutopilotGroupTagSuccess.csv'
$script:FailedLogPath = Join-Path $script:ScriptRoot 'AutopilotGroupTagFailed.csv'
$script:GraphScope = 'DeviceManagementServiceConfig.ReadWrite.All'
$script:GraphPageSize = 200

$script:DeviceCache = New-Object 'System.Collections.Generic.List[AutopilotRow]'
$script:ViewList = New-Object 'System.Collections.Generic.List[AutopilotRow]'
$script:ImportedRows = New-Object 'System.Collections.Generic.List[AutopilotRow]'
$script:Connected = $false
$script:CancelLoad = $false
$script:Busy = $false
$script:PageIndex = 0
$script:PageSize = 100
$script:SortColumn = 'SerialNumber'
$script:SortAscending = $true
$script:SearchText = ''

#region Helpers

function Invoke-UiPump {
    # Runs pending dispatcher work so the window repaints during long UI-thread loops
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Windows.Threading.DispatcherPriority]::Background, [Action]{}
    )
}

function Write-UiLog {
    param([string]$Message, [System.Drawing.Color]$Color = [System.Drawing.Color]::Black)
    if ($null -ne $script:lstLog -and -not $script:lstLog.Dispatcher.CheckAccess()) {
        $script:_uiLogMessage = $Message
        $script:lstLog.Dispatcher.Invoke([Action]{
            Write-UiLog -Message $script:_uiLogMessage
        }) | Out-Null
        return
    }
    if ($null -eq $script:lstLog) { return }
    $ts = (Get-Date).ToString('HH:mm:ss')
    [void]$script:lstLog.Items.Insert(0, "[$ts] $Message")
    if ($script:lstLog.Items.Count -gt 500) { [void]$script:lstLog.Items.RemoveAt($script:lstLog.Items.Count - 1) }
}

function Write-AutopilotLog {
    param(
        [ValidateSet('Success', 'Failed')][string]$Status,
        [string]$SerialNumber,
        [string]$GroupTag,
        [string]$Message
    )
    $entry = [PSCustomObject]@{
        Time         = (Get-Date).ToString('o')
        Status       = $Status
        SerialNumber = $SerialNumber
        GroupTag     = $GroupTag
        Message      = $Message
    }
    $path = if ($Status -eq 'Success') { $script:SuccessLogPath } else { $script:FailedLogPath }
    $entry | Export-Csv -Path $path -Append -NoTypeInformation -Encoding UTF8
}

function Initialize-GraphModules {
    # Graph REST via Invoke-MgGraphRequest - only Authentication module is required.
    $module = 'Microsoft.Graph.Authentication'
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-UiLog "Installing module $module..."
        Install-Module -Name $module -Force -Scope CurrentUser -AllowClobber -ErrorAction Stop
    }
    Import-Module $module -ErrorAction Stop
}

function Set-AutopilotDeviceGroupTagGraph {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][AllowEmptyString()][string]$GroupTag
    )
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities/$DeviceId/updateDeviceProperties"
    $body = @{ groupTag = $GroupTag }
    Invoke-GraphWithRetry {
        Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType 'application/json'
    }
}

function Set-UiBusy {
    param([bool]$Busy)
    $script:Busy = $Busy
    $enabled = -not $Busy
    foreach ($c in @(
            $script:txtClientId, $script:txtTenantId, $script:btnConnect, $script:btnDisconnect,
            $script:btnLoadAll, $script:btnCancelLoad, $script:btnSearch, $script:btnClearSearch,
            $script:btnPrev, $script:btnNext, $script:cmbPageSize, $script:btnApply, $script:btnApplyImported, $script:btnRemoveImported, $script:btnImportCsv,
            $script:txtSearch, $script:txtGroupTag, $script:grid
        )) {
        if ($null -eq $c) { continue }
        if ($c -eq $script:btnCancelLoad) {
            $c.IsEnabled = $Busy -and $script:CancelLoad -eq $false
            continue
        }
        if ($c -eq $script:btnConnect) { $c.IsEnabled = $enabled -and -not $script:Connected; continue }
        if ($c -eq $script:btnDisconnect) { $c.IsEnabled = $enabled -and $script:Connected; continue }
        if ($c -eq $script:btnApplyImported -or $c -eq $script:btnRemoveImported) {
            $c.IsEnabled = $enabled -and $script:Connected -and $script:ImportedRows.Count -gt 0
            continue
        }
        if ($c -in @($script:btnLoadAll, $script:btnSearch, $script:btnClearSearch, $script:btnPrev, $script:btnNext, $script:cmbPageSize, $script:btnApply, $script:btnImportCsv, $script:txtSearch, $script:txtGroupTag, $script:grid)) {
            $c.IsEnabled = $enabled -and $script:Connected
            continue
        }
        $c.IsEnabled = $enabled
    }
    $script:btnCancelLoad.IsEnabled = $Busy
}

function Update-StatusBar {
    param([string]$Text)
    if ($null -ne $script:lblStatus -and -not $script:lblStatus.Dispatcher.CheckAccess()) {
        $script:_uiStatusText = $Text
        $script:lblStatus.Dispatcher.Invoke([Action]{
            Update-StatusBar -Text $script:_uiStatusText
        }) | Out-Null
        return
    }
    if ($null -eq $script:lblStatus) { return }
    $script:lblStatus.Text = $Text
}

function Update-Progress {
    param([int]$Value = 0, [int]$Maximum = 100, [bool]$StyleMarquee = $false)
    if ($null -ne $script:progress -and -not $script:progress.Dispatcher.CheckAccess()) {
        $script:_uiProgressValue = $Value
        $script:_uiProgressMaximum = $Maximum
        $script:_uiProgressMarquee = $StyleMarquee
        $script:progress.Dispatcher.Invoke([Action]{
            Update-Progress -Value $script:_uiProgressValue -Maximum $script:_uiProgressMaximum -StyleMarquee:$script:_uiProgressMarquee
        }) | Out-Null
        return
    }
    if ($null -eq $script:progress) { return }
    if ($StyleMarquee) {
        $script:progress.IsIndeterminate = $true
    }
    else {
        $script:progress.IsIndeterminate = $false
        $max = [Math]::Max(1, $Maximum)
        $script:progress.Maximum = $max
        $val = [Math]::Min([Math]::Max(0, $Value), $max)
        $script:progress.Value = $val
    }
}

function Get-GraphProperty {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [hashtable] -or $Object -is [System.Collections.IDictionary]) {
        if ($Object.ContainsKey($Name)) { return $Object[$Name] }
        return $null
    }
    return $Object.$Name
}

function ConvertTo-DeviceRow {
    param($Device)
    $row = New-Object AutopilotRow
    $row.Id = [string](Get-GraphProperty -Object $Device -Name 'id')
    $row.SerialNumber = [string](Get-GraphProperty -Object $Device -Name 'serialNumber')
    $row.GroupTag = [string](Get-GraphProperty -Object $Device -Name 'groupTag')
    $row.Model = [string](Get-GraphProperty -Object $Device -Name 'model')
    $row.Manufacturer = [string](Get-GraphProperty -Object $Device -Name 'manufacturer')
    return $row
}

function Get-RowSortKey {
    param([AutopilotRow]$Row, [string]$Column)
    if ($null -eq $Row) { return '' }
    $value = switch ($Column) {
        'GroupTag' { $Row.GroupTag }
        'Model' { $Row.Model }
        'Manufacturer' { $Row.Manufacturer }
        default { $Row.SerialNumber }
    }
    if ($null -eq $value) { return '' }
    return [string]$value
}

function Get-GraphAuthHelpMessage {
    param([string]$ErrorText)
    $ctx = $null
    try { $ctx = Get-MgContext } catch { }
    $scopes = if ($ctx -and $ctx.Scopes) { ($ctx.Scopes -join ', ') } else { '(none / not connected)' }
    $account = if ($ctx) { $ctx.Account } else { '(unknown)' }

    $hint = @"
Signed-in account: $account
Token scopes: $scopes

This usually means one of:

0) Wrong account signed in (very common with WAM)
   - Log line "Connected as" must be MEMBER UPN
   - If you see #EXT# you are on a GUEST - Intune Autopilot often returns 401
   - Disconnect, Connect, pick the member account in the picker (not the guest)

1) Delegated permission not in the token
   - App registration -> API permissions -> Microsoft Graph -> Delegated:
     DeviceManagementServiceConfig.ReadWrite.All
   - Click Grant admin consent
   - Disconnect in the GUI, then Connect again (so a new token is issued)

2) User has no Intune rights (very common with 401 from DeviceEnrollmentFE)
   - The SAME UPN as "Connected as" must have Intune Administrator (or equivalent)
   - Assigning the role to a different account than the one in the GUI does nothing

3) Tenant Intune
   - Tenant must have Intune licensed/configured

Original error:
$ErrorText
"@
    return $hint
}

function Test-GraphAutopilotAccess {
    $uri = 'https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?$top=1'
    Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject | Out-Null
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
            $isThrottle = $msg -match '429|throttl|TooManyRequests'
            if (-not $isThrottle -or $attempt -ge $MaxRetries) { throw }
            $delay = [Math]::Min(60, [Math]::Pow(2, $attempt))
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-AutopilotDeviceBySerialGraph {
    param([string]$Serial)
    $escaped = $Serial.Replace("'", "''")
    $filter = "contains(serialNumber,'$escaped')"
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=$([uri]::EscapeDataString($filter))&`$top=50"
    $resp = Invoke-GraphWithRetry { Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject }
    $values = @($resp.value)
    foreach ($v in $values) {
        ConvertTo-DeviceRow -Device $v
    }
}

function Sync-AllAutopilotDevices {
    $script:CancelLoad = $false
    $list = New-Object 'System.Collections.Generic.List[AutopilotRow]'
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$top=$($script:GraphPageSize)"
    $page = 0
    while ($uri) {
        if ($script:CancelLoad) { throw 'Load cancelled by user.' }
        $page++
        $resp = Invoke-GraphWithRetry { Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject }
        $values = @(Get-GraphProperty -Object $resp -Name 'value')
        foreach ($v in $values) {
            if ($null -eq $v) { continue }
            $list.Add((ConvertTo-DeviceRow -Device $v)) | Out-Null
        }
        # Load runs on the UI thread - update progress directly (BeginInvoke + StrictMode cannot see $count/$page).
        $loadedCount = $list.Count
        $pageNum = $page
        Update-Progress -Value $loadedCount -Maximum ([Math]::Max($loadedCount + $script:GraphPageSize, $loadedCount + 1))
        Update-StatusBar "Loading devices... $loadedCount (page $pageNum)"
        Invoke-UiPump
        $next = Get-GraphProperty -Object $resp -Name '@odata.nextLink'
        if ($next) { $uri = [string]$next } else { $uri = $null }
    }
    return $list
}

function Update-FilterAndSort {
    # Build filtered+sorted view for the grid (in-memory; no Graph calls).
    # Uses List[T].Sort with a comparison delegate: fast and avoids Sort-Object on 20k PSObjects.
    $view = New-Object 'System.Collections.Generic.List[AutopilotRow]'
    $search = [string]$script:SearchText

    if ([string]::IsNullOrWhiteSpace($search)) {
        for ($i = 0; $i -lt $script:DeviceCache.Count; $i++) {
            $device = $script:DeviceCache[$i]
            if ($null -ne $device) { [void]$view.Add($device) }
        }
    }
    else {
        $needle = $search.Trim()
        for ($i = 0; $i -lt $script:DeviceCache.Count; $i++) {
            $device = $script:DeviceCache[$i]
            if ($null -eq $device) { continue }
            $serial = [string]$device.SerialNumber
            if ([string]::IsNullOrEmpty($serial)) { continue }
            if ($serial.IndexOf($needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                [void]$view.Add($device)
            }
        }
    }

    $sortColumn = [string]$script:SortColumn
    if ([string]::IsNullOrWhiteSpace($sortColumn)) { $sortColumn = 'SerialNumber' }

    if ($view.Count -gt 1) {
        # Key-array sort in .NET: no per-comparison scriptblock, fast at 20k+ rows
        $rows = $view.ToArray()
        $count = $rows.Length
        $keys = New-Object 'System.String[]' $count
        for ($i = 0; $i -lt $count; $i++) {
            $keys[$i] = Get-RowSortKey -Row $rows[$i] -Column $sortColumn
        }
        [System.Array]::Sort($keys, $rows, [System.StringComparer]::OrdinalIgnoreCase)
        if (-not $script:SortAscending) {
            [System.Array]::Reverse($rows)
        }
        $view = New-Object 'System.Collections.Generic.List[AutopilotRow]'
        for ($i = 0; $i -lt $count; $i++) {
            [void]$view.Add($rows[$i])
        }
    }

    $script:ViewList = $view
}

function Get-PageCount {
    $n = $script:ViewList.Count
    if ($n -le 0) { return 1 }
    return [int][Math]::Ceiling($n / [double]$script:PageSize)
}

function Show-CurrentPage {
    if ($null -ne $script:grid -and -not $script:grid.Dispatcher.CheckAccess()) {
        $script:grid.Dispatcher.Invoke([Action]{ Show-CurrentPage }) | Out-Null
        return
    }

    $pageCount = Get-PageCount
    if ($script:PageIndex -ge $pageCount) { $script:PageIndex = [Math]::Max(0, $pageCount - 1) }
    if ($script:PageIndex -lt 0) { $script:PageIndex = 0 }

    $start = $script:PageIndex * $script:PageSize
    $take = [Math]::Min($script:PageSize, [Math]::Max(0, $script:ViewList.Count - $start))

    $pageList = New-Object 'System.Collections.Generic.List[AutopilotRow]'
    for ($i = 0; $i -lt $take; $i++) {
        $row = $script:ViewList[$start + $i]
        if ($null -eq $row) { continue }
        [void]$pageList.Add($row)
    }
    $script:grid.ItemsSource = $null
    $script:grid.ItemsSource = $pageList

    $sortDir = if ($script:SortAscending) { [System.ComponentModel.ListSortDirection]::Ascending } else { [System.ComponentModel.ListSortDirection]::Descending }
    foreach ($col in $script:grid.Columns) {
        $col.SortDirection = $null
        if ($col.SortMemberPath -eq $script:SortColumn) {
            $col.SortDirection = $sortDir
        }
    }

    $totalCache = $script:DeviceCache.Count
    $filtered = $script:ViewList.Count
    $script:lblPage.Text = "Page $($script:PageIndex + 1) of $pageCount  |  Showing $take  |  Filtered $filtered of $totalCache"
    $script:btnPrev.IsEnabled = $script:Connected -and -not $script:Busy -and ($script:PageIndex -gt 0)
    $script:btnNext.IsEnabled = $script:Connected -and -not $script:Busy -and ($script:PageIndex -lt $pageCount - 1)
}

function Update-DeviceView {
    param([switch]$ResetPage)
    Update-StatusBar 'Sorting / filtering...'
    Update-Progress -Value 0 -Maximum 100 -StyleMarquee $true
    Update-FilterAndSort
    if ($ResetPage) { $script:PageIndex = 0 }
    Show-CurrentPage
    Update-Progress -Value 100 -Maximum 100
    Update-StatusBar "Ready. Cache: $($script:DeviceCache.Count) devices."
}

function Get-SelectedDeviceRows {
    # HashSet dedupe: avoids O(n^2) growth when many rows are selected
    $rows = New-Object 'System.Collections.Generic.List[AutopilotRow]'
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($item in $script:grid.SelectedItems) {
        if ($null -eq $item) { continue }
        $id = [string]$item.Id
        if ([string]::IsNullOrEmpty($id)) { continue }
        if (-not $seen.Add($id)) { continue }
        $row = New-Object AutopilotRow
        $row.Id = $id
        $row.SerialNumber = [string]$item.SerialNumber
        $row.GroupTag = [string]$item.GroupTag
        [void]$rows.Add($row)
    }
    return $rows
}

function Update-ImportedButton {
    if ($null -eq $script:btnApplyImported) { return }
    $n = if ($null -ne $script:ImportedRows) { $script:ImportedRows.Count } else { 0 }
    $script:btnApplyImported.Content = "Apply to all imported ($n)"
    $script:btnApplyImported.IsEnabled = $script:Connected -and -not $script:Busy -and $n -gt 0
    if ($null -ne $script:btnRemoveImported) {
        $script:btnRemoveImported.Content = "Remove tag from all imported ($n)"
        $script:btnRemoveImported.IsEnabled = $script:Connected -and -not $script:Busy -and $n -gt 0
    }
}

function Invoke-ApplyGroupTag {
    param(
        [System.Collections.Generic.List[AutopilotRow]]$Devices,
        [AllowEmptyString()][string]$GroupTag
    )
    $isRemove = [string]::IsNullOrEmpty($GroupTag)
    $stamp = (Get-Date).ToString('yyyy-MM-dd_HH-mm')
    $script:SuccessLogPath = Join-Path $script:ScriptRoot "AutopilotGroupTagSuccess_$stamp.csv"
    $script:FailedLogPath = Join-Path $script:ScriptRoot "AutopilotGroupTagFailed_$stamp.csv"

    $ok = 0
    $fail = 0
    $skipped = 0
    $cancelled = $false
    $script:CancelLoad = $false
    try {
        Set-UiBusy -Busy $true
        $script:btnCancelLoad.IsEnabled = $true

        $cacheById = New-Object 'System.Collections.Generic.Dictionary[string,AutopilotRow]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($c in $script:DeviceCache) {
            if ($null -ne $c -and $c.Id -and -not $cacheById.ContainsKey($c.Id)) { $cacheById[$c.Id] = $c }
        }

        $i = 0
        $total = $Devices.Count
        foreach ($dev in $Devices) {
            if ($script:CancelLoad) { $cancelled = $true; break }
            $i++
            Update-Progress -Value $i -Maximum $total
            if ($isRemove) {
                Update-StatusBar "Removing tag ($i / $total): $($dev.SerialNumber)"
            }
            else {
                Update-StatusBar "Applying tag ($i / $total): $($dev.SerialNumber)"
            }
            Invoke-UiPump
            try {
                $oldTag = if ($null -eq $dev.GroupTag) { '' } else { [string]$dev.GroupTag }
                if ($oldTag -eq $GroupTag) {
                    if ($isRemove) {
                        Write-UiLog "Skip $($dev.SerialNumber) (no group tag)"
                    }
                    else {
                        Write-UiLog "Skip $($dev.SerialNumber) (already '$GroupTag')"
                    }
                    $skipped++
                    continue
                }
                Set-AutopilotDeviceGroupTagGraph -DeviceId $dev.Id -GroupTag $GroupTag
                $cacheRow = $null
                if ($cacheById.TryGetValue($dev.Id, [ref]$cacheRow) -and $null -ne $cacheRow) {
                    $cacheRow.GroupTag = $GroupTag
                }
                $dev.GroupTag = $GroupTag
                if ($isRemove) {
                    Write-AutopilotLog -Status Success -SerialNumber $dev.SerialNumber -GroupTag $GroupTag -Message "Group tag removed (was '$oldTag')"
                    Write-UiLog "REMOVED $($dev.SerialNumber) (was '$oldTag')"
                }
                else {
                    Write-AutopilotLog -Status Success -SerialNumber $dev.SerialNumber -GroupTag $GroupTag -Message 'Group tag updated'
                    Write-UiLog "OK $($dev.SerialNumber) -> $GroupTag"
                }
                $ok++
            }
            catch {
                $fail++
                Write-AutopilotLog -Status Failed -SerialNumber $dev.SerialNumber -GroupTag $GroupTag -Message "$_"
                Write-UiLog "FAIL $($dev.SerialNumber): $_"
            }
        }
        Show-CurrentPage
        $donePrefix = if ($isRemove) { 'Remove done.' } else { 'Apply done.' }
        Update-StatusBar "$donePrefix Success: $ok  Skipped: $skipped  Failed: $fail$(if ($cancelled) { ' (cancelled)' })"
        Write-UiLog "Logs: $($script:SuccessLogPath) / $($script:FailedLogPath)"
    }
    finally {
        $script:CancelLoad = $false
        Set-UiBusy -Busy $false
        Update-ImportedButton
    }
    return [PSCustomObject]@{
        Ok        = $ok
        Fail      = $fail
        Skipped   = $skipped
        Cancelled = $cancelled
        Stamp     = $stamp
    }
}

function Start-BackgroundWork {
    param(
        [scriptblock]$Work,
        [scriptblock]$OnSuccess,
        [scriptblock]$OnError
    )
    Set-UiBusy -Busy $true
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'MTA'
    $runspace.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    [void]$ps.AddScript($Work.ToString())
    $handle = $ps.BeginInvoke()

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(200)
    $timer.Add_Tick({
        if (-not $handle.IsCompleted) { return }
        $timer.Stop()
        try {
            $result = $ps.EndInvoke($handle)
            if ($ps.HadErrors) {
                $err = ($ps.Streams.Error | ForEach-Object { "$_" }) -join '; '
                if (-not $err) { $err = 'Background work failed.' }
                & $OnError $err
            }
            else {
                & $OnSuccess $result
            }
        }
        catch {
            & $OnError "$_"
        }
        finally {
            $ps.Dispose()
            $runspace.Close()
            $runspace.Dispose()
            Set-UiBusy -Busy $false
        }
    }.GetNewClosure())
    $timer.Start()
}

#endregion

#region UI

[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Autopilot Group Tag"
    Height="800" Width="1200"
    MinHeight="640" MinWidth="960"
    WindowStartupLocation="CenterScreen"
    Background="#F5F6FA">

    <Window.Resources>
        <!-- Accent color -->
        <SolidColorBrush x:Key="AccentBrush" Color="#0078D4"/>
        <SolidColorBrush x:Key="AccentHoverBrush" Color="#106EBE"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#D1D5DB"/>
        <SolidColorBrush x:Key="CardBrush" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#6B7280"/>

        <!-- Button style -->
        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{StaticResource AccentHoverBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="DangerButton" TargetType="Button">
            <Setter Property="Background" Value="#C42B1C"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#A4262C"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SecondaryButton" TargetType="Button">
            <Setter Property="Background" Value="#E5E7EB"/>
            <Setter Property="Foreground" Value="#374151"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#D1D5DB"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- TextBox style -->
        <Style TargetType="TextBox">
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>

        <!-- ComboBox style -->
        <Style TargetType="ComboBox">
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
        </Style>

        <!-- Label style -->
        <Style TargetType="Label">
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Foreground" Value="#374151"/>
            <Setter Property="Padding" Value="0,0,0,4"/>
            <Setter Property="FontWeight" Value="Medium"/>
        </Style>

        <!-- ListBox style -->
        <Style TargetType="ListBox">
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Foreground" Value="#374151"/>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <TextBlock Grid.Row="0" Text="Autopilot Group Tag"
                   FontSize="22" FontWeight="Bold" Foreground="#1F2937"
                   Margin="0,0,0,12"/>

        <!-- Connection card -->
        <Border Grid.Row="1" Background="{StaticResource CardBrush}"
                CornerRadius="6" Padding="16,12"
                BorderBrush="{StaticResource BorderBrush}" BorderThickness="1"
                Margin="0,0,0,10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Label Grid.Row="0" Grid.Column="0" Content="Client ID:" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <TextBox Grid.Row="0" Grid.Column="1" Name="txtClientId" FontFamily="Consolas" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <Button Grid.Row="0" Grid.Column="2" Name="btnConnect" Content="Connect" Style="{StaticResource PrimaryButton}" VerticalAlignment="Center" Margin="0,0,8,0"/>
                <Button Grid.Row="0" Grid.Column="3" Name="btnDisconnect" Content="Disconnect" Style="{StaticResource SecondaryButton}" IsEnabled="False" VerticalAlignment="Center"/>
                <Label Grid.Row="1" Grid.Column="0" Content="Tenant ID:" VerticalAlignment="Center" Margin="0,8,8,0"/>
                <TextBox Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="3" Name="txtTenantId" FontFamily="Consolas" VerticalAlignment="Center" Margin="0,8,0,0"/>
            </Grid>
        </Border>

        <!-- Actions card -->
        <Border Grid.Row="2" Background="{StaticResource CardBrush}"
                CornerRadius="6" Padding="16,12"
                BorderBrush="{StaticResource BorderBrush}" BorderThickness="1"
                Margin="0,0,0,10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <StackPanel Grid.Row="0" Orientation="Horizontal">
                    <Button Name="btnLoadAll" Content="Load all devices" Style="{StaticResource PrimaryButton}" IsEnabled="False"/>
                    <Button Name="btnCancelLoad" Content="Cancel" Style="{StaticResource SecondaryButton}" IsEnabled="False" Margin="8,0,0,0"/>
                    <Label Content="Serial:" VerticalAlignment="Center" Margin="16,0,8,0"/>
                    <TextBox Name="txtSearch" Width="200" IsEnabled="False" VerticalAlignment="Center"/>
                    <Button Name="btnSearch" Content="Search" Style="{StaticResource PrimaryButton}" IsEnabled="False" Margin="8,0,0,0"/>
                    <Button Name="btnClearSearch" Content="Clear" Style="{StaticResource SecondaryButton}" IsEnabled="False" Margin="8,0,0,0"/>
                    <Button Name="btnImportCsv" Content="Import CSV..." Style="{StaticResource SecondaryButton}" IsEnabled="False" Margin="8,0,0,0"/>
                </StackPanel>
                <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,10,0,0">
                    <Label Content="Group tag:" VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBox Name="txtGroupTag" Width="200" IsEnabled="False" VerticalAlignment="Center"/>
                    <Button Name="btnApply" Content="Apply to selected" Style="{StaticResource PrimaryButton}" IsEnabled="False" Margin="8,0,0,0"/>
                    <Button Name="btnApplyImported" Content="Apply to all imported (0)" Style="{StaticResource PrimaryButton}" IsEnabled="False" Margin="8,0,0,0" ToolTip="Apply the group tag to every device found in the last CSV import (all pages)"/>
                    <Button Name="btnRemoveImported" Content="Remove tag from all imported (0)" Style="{StaticResource DangerButton}" IsEnabled="False" Margin="8,0,0,0" ToolTip="Clear the group tag on every device found in the last CSV import (all pages)"/>
                    <Label Content="Page size:" VerticalAlignment="Center" Margin="16,0,8,0"/>
                    <ComboBox Name="cmbPageSize" Width="80" IsEnabled="False" VerticalAlignment="Center"/>
                    <Button Name="btnPrev" Content="&lt;" Style="{StaticResource SecondaryButton}" IsEnabled="False" Width="36" Margin="8,0,0,0"/>
                    <Button Name="btnNext" Content="&gt;" Style="{StaticResource SecondaryButton}" IsEnabled="False" Width="36" Margin="8,0,0,0"/>
                    <TextBlock Name="lblPage" Text="Page 0 of 0" VerticalAlignment="Center" Foreground="{StaticResource TextSecondary}" FontSize="12" Margin="12,0,0,0"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- DataGrid card -->
        <Border Grid.Row="3" Background="{StaticResource CardBrush}"
                CornerRadius="6"
                BorderBrush="{StaticResource BorderBrush}" BorderThickness="1">
            <DataGrid Name="grid"
                      AutoGenerateColumns="False"
                      IsReadOnly="True"
                      CanUserSortColumns="True"
                      CanUserReorderColumns="True"
                      CanUserResizeColumns="True"
                      SelectionMode="Extended"
                      SelectionUnit="FullRow"
                      EnableRowVirtualization="True"
                      IsEnabled="False"
                      GridLinesVisibility="Horizontal"
                      HorizontalGridLinesBrush="#F3F4F6"
                      HeadersVisibility="Column"
                      AlternatingRowBackground="#F9FAFB"
                      RowBackground="White"
                      BorderThickness="0"
                      FontSize="12.5"
                      VerticalScrollBarVisibility="Auto"
                      HorizontalScrollBarVisibility="Auto">
                <DataGrid.ColumnHeaderStyle>
                    <Style TargetType="DataGridColumnHeader">
                        <Setter Property="Background" Value="#F3F4F6"/>
                        <Setter Property="Foreground" Value="#374151"/>
                        <Setter Property="FontWeight" Value="SemiBold"/>
                        <Setter Property="FontSize" Value="12"/>
                        <Setter Property="Padding" Value="10,8"/>
                        <Setter Property="BorderBrush" Value="#E5E7EB"/>
                        <Setter Property="BorderThickness" Value="0,0,1,1"/>
                        <Setter Property="Cursor" Value="Hand"/>
                    </Style>
                </DataGrid.ColumnHeaderStyle>
                <DataGrid.CellStyle>
                    <Style TargetType="DataGridCell">
                        <Setter Property="Padding" Value="10,6"/>
                        <Setter Property="BorderThickness" Value="0"/>
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="DataGridCell">
                                    <Border Padding="{TemplateBinding Padding}" Background="{TemplateBinding Background}">
                                        <ContentPresenter VerticalAlignment="Center"/>
                                    </Border>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                        <Style.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Background" Value="#DBEAFE"/>
                                <Setter Property="Foreground" Value="#1F2937"/>
                            </Trigger>
                        </Style.Triggers>
                    </Style>
                </DataGrid.CellStyle>
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Serial Number" Binding="{Binding SerialNumber}" SortMemberPath="SerialNumber" Width="*"/>
                    <DataGridTextColumn Header="Group Tag" Binding="{Binding GroupTag}" SortMemberPath="GroupTag" Width="*"/>
                    <DataGridTextColumn Header="Model" Binding="{Binding Model}" SortMemberPath="Model" Width="*"/>
                    <DataGridTextColumn Header="Manufacturer" Binding="{Binding Manufacturer}" SortMemberPath="Manufacturer" Width="*"/>
                </DataGrid.Columns>
            </DataGrid>
        </Border>

        <!-- Log card -->
        <Border Grid.Row="4" Background="{StaticResource CardBrush}"
                CornerRadius="6" Padding="12,8"
                BorderBrush="{StaticResource BorderBrush}" BorderThickness="1"
                Margin="0,8,0,0">
            <StackPanel>
                <Label Content="Log" Padding="0,0,0,4"/>
                <ListBox Name="lstLog" Height="110"/>
            </StackPanel>
        </Border>

        <!-- Status bar -->
        <Border Grid.Row="5" Background="{StaticResource CardBrush}"
                CornerRadius="4" Padding="12,8"
                BorderBrush="{StaticResource BorderBrush}" BorderThickness="1"
                Margin="0,8,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Name="lblStatus" Text="Enter Client ID and Tenant ID, then Connect."
                           VerticalAlignment="Center" Foreground="{StaticResource TextSecondary}" FontSize="12"
                           TextTrimming="CharacterEllipsis"/>
                <ProgressBar Grid.Column="1" Name="progress" Width="220" Height="8"
                             Minimum="0" Maximum="100" Value="0"
                             VerticalAlignment="Center" Margin="12,0,0,0"
                             Foreground="{StaticResource AccentBrush}" Background="#E5E7EB" BorderThickness="0"/>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$script:txtClientId    = $window.FindName('txtClientId')
$script:txtTenantId    = $window.FindName('txtTenantId')
$script:btnConnect     = $window.FindName('btnConnect')
$script:btnDisconnect  = $window.FindName('btnDisconnect')
$script:btnLoadAll     = $window.FindName('btnLoadAll')
$script:btnCancelLoad  = $window.FindName('btnCancelLoad')
$script:txtSearch      = $window.FindName('txtSearch')
$script:btnSearch      = $window.FindName('btnSearch')
$script:btnClearSearch = $window.FindName('btnClearSearch')
$script:btnImportCsv   = $window.FindName('btnImportCsv')
$script:txtGroupTag    = $window.FindName('txtGroupTag')
$script:btnApply       = $window.FindName('btnApply')
$script:btnApplyImported = $window.FindName('btnApplyImported')
$script:btnRemoveImported = $window.FindName('btnRemoveImported')
$script:cmbPageSize    = $window.FindName('cmbPageSize')
$script:btnPrev        = $window.FindName('btnPrev')
$script:btnNext        = $window.FindName('btnNext')
$script:lblPage        = $window.FindName('lblPage')
$script:grid           = $window.FindName('grid')
$script:lstLog         = $window.FindName('lstLog')
$script:progress       = $window.FindName('progress')
$script:lblStatus      = $window.FindName('lblStatus')

@('50', '100', '200', '300', '400', '500') | ForEach-Object { [void]$script:cmbPageSize.Items.Add($_) }
$script:cmbPageSize.SelectedItem = '100'

#endregion

#region Events

$script:btnConnect.Add_Click({
    try {
        $clientId = $script:txtClientId.Text.Trim()
        $tenantId = $script:txtTenantId.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($clientId) -or [string]::IsNullOrWhiteSpace($tenantId)) {
            [System.Windows.Forms.MessageBox]::Show('Client ID and Tenant ID are required.', 'Connect', 'OK', 'Warning') | Out-Null
            return
        }
        Set-UiBusy -Busy $true
        Update-StatusBar 'Loading modules...'
        Initialize-GraphModules
        Update-StatusBar 'Sign in with your account (browser)...'
        Write-UiLog 'Connecting (delegated interactive)...'
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
        try {
            # Process-scoped context avoids reusing another session's token/account
            Connect-MgGraph -ClientId $clientId -TenantId $tenantId -Scopes $script:GraphScope -ContextScope Process -NoWelcome
        }
        catch {
            $errText = "$_"
            if ($errText -match 'AADSTS7000218|client_assertion|client_secret') {
                throw @"
Entra ID rejected sign-in (AADSTS7000218): the app is not configured as a public client for interactive browser login.

In Entra admin center -> App registrations -> your app:
1. Authentication -> Add a platform -> Mobile and desktop applications
   - Enable http://localhost
2. Authentication -> Advanced settings -> Allow public client flows = Yes
3. API permissions -> Microsoft Graph -> Delegated:
   DeviceManagementServiceConfig.ReadWrite.All -> Grant admin consent
4. Do not use a client secret for this GUI (delegated public client / browser)

Original error: $errText
"@
            }
            throw
        }
        $ctx = Get-MgContext
        $scopeList = @($ctx.Scopes)
        $account = [string]$ctx.Account
        Write-UiLog "Connected as $account"
        Write-UiLog "TenantId: $($ctx.TenantId)"
        Write-UiLog "Scopes: $($scopeList -join ', ')"
        if ($account -match '#EXT#') {
            Write-UiLog 'WARNING: Signed in as a GUEST (#EXT#). Intune Autopilot often returns 401 for guests.'
            [System.Windows.Forms.MessageBox]::Show(
                "You are signed in as a GUEST account:`n$account`n`nIntune Autopilot APIs often return 401 for B2B guests, even with Intune Administrator.`n`nDisconnect and Connect again.",
                'Guest account detected',
                'OK',
                'Warning'
            ) | Out-Null
        }
        elseif ($account -and $account -notmatch '@') {
            Write-UiLog "WARNING: Unexpected account format: $account"
        }
        $hasAutopilotScope = $scopeList | Where-Object {
            $_ -eq 'DeviceManagementServiceConfig.ReadWrite.All' -or
            $_ -eq 'DeviceManagementServiceConfig.Read.All' -or
            $_ -like '*DeviceManagementServiceConfig.ReadWrite.All*' -or
            $_ -like '*DeviceManagementServiceConfig.Read.All*'
        }
        if (-not $hasAutopilotScope) {
            Write-UiLog 'WARNING: Token is missing DeviceManagementServiceConfig.ReadWrite.All - grant admin consent and reconnect.'
            [System.Windows.Forms.MessageBox]::Show(
                "Connected as $($ctx.Account), but the access token does not include DeviceManagementServiceConfig.ReadWrite.All.`n`nScopes:`n$($scopeList -join "`n")`n`nIn Entra: app -> API permissions -> add Delegated DeviceManagementServiceConfig.ReadWrite.All -> Grant admin consent -> Disconnect and Connect again.",
                'Missing Graph scope',
                'OK',
                'Warning'
            ) | Out-Null
        }
        else {
            try {
                Update-StatusBar 'Verifying Autopilot API access...'
                Test-GraphAutopilotAccess
                Write-UiLog 'Autopilot API probe OK.'
            }
            catch {
                $help = Get-GraphAuthHelpMessage -ErrorText "$_"
                Write-UiLog "Autopilot API probe failed: $_"
                [System.Windows.Forms.MessageBox]::Show(
                    "Signed in, but Intune/Autopilot API returned an error (often 401).`n`n$help",
                    'Autopilot access check failed',
                    'OK',
                    'Warning'
                ) | Out-Null
            }
        }
        $script:Connected = $true
        $script:DeviceCache.Clear()
        $script:ViewList.Clear()
        Show-CurrentPage
        Update-StatusBar "Connected as $($ctx.Account). Load all devices, search, or import CSV."
    }
    catch {
        $script:Connected = $false
        Write-UiLog "Connect failed: $_"
        [System.Windows.Forms.MessageBox]::Show("Connect failed:`n`n$_", 'Entra sign-in error', 'OK', 'Error') | Out-Null
        Update-StatusBar 'Not connected. Check public client settings in Entra if AADSTS7000218.'
    }
    finally {
        Set-UiBusy -Busy $false
        $script:btnConnect.IsEnabled = -not $script:Connected
        $script:btnDisconnect.IsEnabled = $script:Connected
        foreach ($c in @($script:btnLoadAll, $script:btnSearch, $script:btnClearSearch, $script:btnImportCsv, $script:btnApply, $script:txtSearch, $script:txtGroupTag, $script:cmbPageSize, $script:grid)) {
            $c.IsEnabled = $script:Connected
        }
        Update-ImportedButton
    }
})

$script:btnDisconnect.Add_Click({
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    catch { }
    $script:Connected = $false
    $script:DeviceCache.Clear()
    $script:ViewList.Clear()
    $script:ImportedRows.Clear()
    Show-CurrentPage
    Write-UiLog 'Disconnected.'
    Update-StatusBar 'Disconnected. Enter Client ID / Tenant ID and Connect.'
    $script:btnConnect.IsEnabled = $true
    $script:btnDisconnect.IsEnabled = $false
    foreach ($c in @($script:btnLoadAll, $script:btnSearch, $script:btnClearSearch, $script:btnImportCsv, $script:btnApply, $script:txtSearch, $script:txtGroupTag, $script:cmbPageSize, $script:grid, $script:btnPrev, $script:btnNext)) {
        $c.IsEnabled = $false
    }
    Update-ImportedButton
})

$script:btnCancelLoad.Add_Click({
    $script:CancelLoad = $true
    Write-UiLog 'Cancel requested...'
})

$script:btnLoadAll.Add_Click({
    if (-not $script:Connected) { return }
    $script:CancelLoad = $false
    Set-UiBusy -Busy $true
    $script:btnCancelLoad.IsEnabled = $true
    Update-Progress -Value 0 -Maximum 100 -StyleMarquee $true
    Update-StatusBar 'Loading all Autopilot devices from Graph...'
    Write-UiLog 'Starting full device load...'

    # Keep Graph calls on UI thread via async-ish DoEvents loop for simpler token context
    # (MgGraph context is not reliably available in other runspaces)
    try {
        $list = Sync-AllAutopilotDevices
        if ($null -eq $list) {
            $list = New-Object 'System.Collections.Generic.List[AutopilotRow]'
        }
        $script:DeviceCache = $list
        $script:SearchText = ''
        $script:txtSearch.Text = ''
        $deviceTotal = $script:DeviceCache.Count
        Write-UiLog "Fetched $deviceTotal devices. Building view..."
        Update-DeviceView -ResetPage
        Write-UiLog "Loaded $deviceTotal devices into memory."
        Update-StatusBar "Loaded $deviceTotal devices."
        Update-Progress -Value 100 -Maximum 100
    }
    catch {
        $errRecord = $_
        $msg = $errRecord.ToString()
        if ($errRecord.InvocationInfo) {
            $msg = "$msg`n`nAt: $($errRecord.InvocationInfo.PositionMessage)"
        }
        Write-UiLog "Load failed: $msg"
        if ($msg -match '401|Unauthorized|Forbidden|DeviceEnrollmentFE') {
            $msg = Get-GraphAuthHelpMessage -ErrorText $msg
        }
        [System.Windows.Forms.MessageBox]::Show("Load failed:`n`n$msg", 'Error', 'OK', 'Error') | Out-Null
        Update-StatusBar 'Load failed or cancelled.'
    }
    finally {
        $script:CancelLoad = $false
        Set-UiBusy -Busy $false
        $script:btnCancelLoad.IsEnabled = $false
        $script:btnConnect.IsEnabled = -not $script:Connected
        $script:btnDisconnect.IsEnabled = $script:Connected
    }
})

function Invoke-SerialSearch {
    if (-not $script:Connected) { return }
    $q = $script:txtSearch.Text.Trim()
    $script:SearchText = $q

    if ($script:DeviceCache.Count -gt 0) {
        Update-StatusBar 'Filtering in-memory cache...'
        Update-DeviceView -ResetPage
        Write-UiLog "Local filter: '$q' -> $($script:ViewList.Count) hit(s)."
        return
    }

    if ([string]::IsNullOrWhiteSpace($q)) {
        [System.Windows.Forms.MessageBox]::Show('Enter a serial number, or use Load all devices first.', 'Search', 'OK', 'Information') | Out-Null
        return
    }

    try {
        Set-UiBusy -Busy $true
        Update-Progress -Value 0 -Maximum 100 -StyleMarquee $true
        Update-StatusBar "Searching Graph for serial containing '$q'..."
        $hits = @(Get-AutopilotDeviceBySerialGraph -Serial $q)
        $script:ViewList = New-Object 'System.Collections.Generic.List[AutopilotRow]'
        $knownIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        for ($i = 0; $i -lt $script:DeviceCache.Count; $i++) {
            $cached = $script:DeviceCache[$i]
            if ($null -ne $cached -and $cached.Id) { [void]$knownIds.Add($cached.Id) }
        }
        foreach ($h in $hits) {
            if ($null -eq $h) { continue }
            [void]$script:ViewList.Add($h)
            if ($h.Id -and $knownIds.Add($h.Id)) { [void]$script:DeviceCache.Add($h) }
        }
        $script:SearchText = ''
        $script:PageIndex = 0
        Show-CurrentPage
        Write-UiLog "Graph search '$q' -> $($hits.Count) hit(s). (Load all for full browse/sort.)"
        Update-Progress -Value 100 -Maximum 100
    }
    catch {
        Write-UiLog "Search failed: $_"
        [System.Windows.Forms.MessageBox]::Show("Search failed:`n$_", 'Error', 'OK', 'Error') | Out-Null
    }
    finally {
        Set-UiBusy -Busy $false
    }
}

$script:btnSearch.Add_Click({ Invoke-SerialSearch })
$script:txtSearch.Add_KeyDown({
    param($eventSender, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Return) {
        $e.Handled = $true
        Invoke-SerialSearch
    }
})

$script:btnClearSearch.Add_Click({
    $script:txtSearch.Text = ''
    $script:SearchText = ''
    if ($script:DeviceCache.Count -gt 0) {
        Update-DeviceView -ResetPage
    }
    else {
        $script:ViewList.Clear()
        Show-CurrentPage
    }
    Write-UiLog 'Search cleared.'
})

$script:cmbPageSize.Add_SelectionChanged({
    if ($script:Busy) { return }
    if ($null -eq $script:cmbPageSize.SelectedItem) { return }
    $script:PageSize = [int]$script:cmbPageSize.SelectedItem
    $script:PageIndex = 0
    Show-CurrentPage
})

$script:btnPrev.Add_Click({
    if ($script:PageIndex -gt 0) {
        $script:PageIndex--
        Show-CurrentPage
    }
})

$script:btnNext.Add_Click({
    if ($script:PageIndex -lt (Get-PageCount) - 1) {
        $script:PageIndex++
        Show-CurrentPage
    }
})

$script:grid.Add_Sorting({
    param($s, $e)
    $e.Handled = $true
    if (-not $script:Connected -or $script:Busy) { return }
    if ($script:DeviceCache.Count -eq 0) { return }
    $colName = [string]$e.Column.SortMemberPath
    if ([string]::IsNullOrEmpty($colName)) { return }
    if ($script:SortColumn -eq $colName) {
        $script:SortAscending = -not $script:SortAscending
    }
    else {
        $script:SortColumn = $colName
        $script:SortAscending = $true
    }
    Write-UiLog "Sort by $($script:SortColumn) $(if ($script:SortAscending) { 'asc' } else { 'desc' }) (all filtered rows)"
    Update-DeviceView -ResetPage
})

$script:btnApply.Add_Click({
    if (-not $script:Connected) { return }
    $tag = $script:txtGroupTag.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($tag)) {
        [System.Windows.Forms.MessageBox]::Show('Enter a Group tag.', 'Apply', 'OK', 'Warning') | Out-Null
        return
    }
    $selected = @(Get-SelectedDeviceRows)
    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Select one or more devices in the grid.', 'Apply', 'OK', 'Warning') | Out-Null
        return
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Set group tag '$tag' on $($selected.Count) device(s)?",
        'Confirm',
        'YesNo',
        'Question'
    )
    if ($confirm -ne 'Yes') { return }

    $devices = New-Object 'System.Collections.Generic.List[AutopilotRow]'
    foreach ($d in $selected) { if ($null -ne $d) { [void]$devices.Add($d) } }
    $r = Invoke-ApplyGroupTag -Devices $devices -GroupTag $tag
    [System.Windows.Forms.MessageBox]::Show("Done.`nSuccess: $($r.Ok)`nSkipped: $($r.Skipped)`nFailed: $($r.Fail)$(if ($r.Cancelled) { "`n(cancelled)" })`nLog folder: $script:ScriptRoot`nFiles: AutopilotGroupTag*_$($r.Stamp).csv", 'Apply', 'OK', 'Information') | Out-Null
})

$script:btnApplyImported.Add_Click({
    if (-not $script:Connected) { return }
    $tag = $script:txtGroupTag.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($tag)) {
        [System.Windows.Forms.MessageBox]::Show('Enter a Group tag.', 'Apply', 'OK', 'Warning') | Out-Null
        return
    }
    if ($null -eq $script:ImportedRows -or $script:ImportedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Import a CSV first.', 'Apply', 'OK', 'Warning') | Out-Null
        return
    }
    $n = $script:ImportedRows.Count
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Set group tag '$tag' on ALL $n device(s) from the last CSV import?`n`nThis includes devices on every page, not only the visible one.",
        'Confirm',
        'YesNo',
        'Warning'
    )
    if ($confirm -ne 'Yes') { return }
    $devices = New-Object 'System.Collections.Generic.List[AutopilotRow]'
    foreach ($d in $script:ImportedRows) { if ($null -ne $d) { [void]$devices.Add($d) } }
    $r = Invoke-ApplyGroupTag -Devices $devices -GroupTag $tag
    [System.Windows.Forms.MessageBox]::Show("Done.`nSuccess: $($r.Ok)`nSkipped: $($r.Skipped)`nFailed: $($r.Fail)$(if ($r.Cancelled) { "`n(cancelled)" })`nLog folder: $script:ScriptRoot`nFiles: AutopilotGroupTag*_$($r.Stamp).csv", 'Apply', 'OK', 'Information') | Out-Null
})

$script:btnRemoveImported.Add_Click({
    if (-not $script:Connected) { return }
    if ($null -eq $script:ImportedRows -or $script:ImportedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Import a CSV first.', 'Apply', 'OK', 'Warning') | Out-Null
        return
    }
    $n = $script:ImportedRows.Count
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Remove the group tag from ALL $n device(s) from the last CSV import?`n`nThis includes devices on every page, not only the visible one. Devices without a tag are skipped.",
        'Confirm',
        'YesNo',
        'Warning'
    )
    if ($confirm -ne 'Yes') { return }
    $devices = New-Object 'System.Collections.Generic.List[AutopilotRow]'
    foreach ($d in $script:ImportedRows) { if ($null -ne $d) { [void]$devices.Add($d) } }
    $r = Invoke-ApplyGroupTag -Devices $devices -GroupTag ''
    [System.Windows.Forms.MessageBox]::Show("Done.`nRemoved: $($r.Ok)`nSkipped: $($r.Skipped)`nFailed: $($r.Fail)$(if ($r.Cancelled) { "`n(cancelled)" })`nLog folder: $script:ScriptRoot`nFiles: AutopilotGroupTag*_$($r.Stamp).csv", 'Remove group tag', 'OK', 'Information') | Out-Null
})

$script:btnImportCsv.Add_Click({
    if (-not $script:Connected) { return }
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $ofd.Title = 'Import serial numbers'
    if ($ofd.ShowDialog() -ne 'OK') { return }

    try {
        Set-UiBusy -Busy $true
        Update-Progress -Value 0 -Maximum 100 -StyleMarquee $true
        $csv = Import-Csv -Path $ofd.FileName
        if (-not $csv) { throw 'CSV is empty.' }

        $props = $csv[0].PSObject.Properties.Name
        $serialProp = $props | Where-Object { $_ -match '^(SerialNumber|Serial Number|Serial)$' } | Select-Object -First 1
        if (-not $serialProp) {
            throw "CSV must include a SerialNumber column (or Serial / Serial Number). Found: $($props -join ', ')"
        }

        $serials = @(
            $csv | ForEach-Object { [string]$_.$serialProp } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                ForEach-Object { $_.Trim() } |
                Select-Object -Unique
        )
        Write-UiLog "CSV: $($serials.Count) unique serial(s) from $($ofd.FileName)"

        $cacheMap = @{}
        foreach ($d in $script:DeviceCache) {
            if ($d.SerialNumber) { $cacheMap[$d.SerialNumber.ToLowerInvariant()] = $d }
        }

        $found = New-Object 'System.Collections.Generic.List[AutopilotRow]'
        $missing = New-Object 'System.Collections.Generic.List[string]'
        $i = 0
        foreach ($sn in $serials) {
            $i++
            Update-Progress -Value $i -Maximum $serials.Count
            Update-StatusBar "Resolving CSV serials ($i / $($serials.Count))"
            Invoke-UiPump
            $key = $sn.ToLowerInvariant()
            if ($cacheMap.ContainsKey($key)) {
                $found.Add($cacheMap[$key]) | Out-Null
                continue
            }
            $hits = @(Get-AutopilotDeviceBySerialGraph -Serial $sn)
            $exact = $hits | Where-Object { $_.SerialNumber -and $_.SerialNumber.Equals($sn, [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
            if (-not $exact -and $hits.Count -eq 1) { $exact = $hits[0] }
            if ($exact) {
                $found.Add($exact) | Out-Null
                if (-not $cacheMap.ContainsKey($exact.SerialNumber.ToLowerInvariant())) {
                    $script:DeviceCache.Add($exact) | Out-Null
                    $cacheMap[$exact.SerialNumber.ToLowerInvariant()] = $exact
                }
            }
            else {
                $missing.Add($sn) | Out-Null
            }
        }

        # Merge found into cache (do not wipe full load); show found set in the grid
        $script:ViewList = New-Object 'System.Collections.Generic.List[AutopilotRow]'
        foreach ($f in $found) {
            if ($null -ne $f) { [void]$script:ViewList.Add($f) }
        }
        $script:SearchText = ''
        $script:txtSearch.Text = ''
        $script:PageIndex = 0
        Show-CurrentPage
        $script:ImportedRows = $found
        Update-ImportedButton

        $msg = "Found: $($found.Count)`nNot found: $($missing.Count)"
        if ($missing.Count -gt 0 -and $missing.Count -le 20) {
            $msg += "`n`nMissing:`n$($missing -join "`n")"
        }
        elseif ($missing.Count -gt 20) {
            $msg += "`n`n(First 20 missing)`n$(($missing | Select-Object -First 20) -join "`n")"
        }
        if ($missing.Count -gt 0) {
            $notFoundPath = Join-Path $script:ScriptRoot ("AutopilotNotFound_{0}.csv" -f (Get-Date).ToString('yyyy-MM-dd_HH-mm'))
            try {
                $missing | ForEach-Object { [PSCustomObject]@{ SerialNumber = $_ } } | Export-Csv -Path $notFoundPath -NoTypeInformation -Encoding UTF8
                Write-UiLog "Not-found serials written to $notFoundPath"
                $msg += "`n`nNot-found list saved to:`n$notFoundPath"
            }
            catch {
                Write-UiLog "Could not write not-found CSV: $_"
                $notFoundPath = $null
            }
        }
        Write-UiLog "CSV resolve done. Found $($found.Count), missing $($missing.Count)."
        Update-StatusBar "CSV: $($found.Count) found. Set Group tag and 'Apply to all imported', 'Remove tag from all imported', or select rows."
        [System.Windows.Forms.MessageBox]::Show($msg + "`n`nEnter Group tag and click 'Apply to all imported ($($found.Count))' to tag, or 'Remove tag from all imported ($($found.Count))' to clear the tag. Or select rows and use 'Apply to selected'.", 'CSV import', 'OK', 'Information') | Out-Null
    }
    catch {
        Write-UiLog "CSV import failed: $_"
        [System.Windows.Forms.MessageBox]::Show("CSV import failed:`n$_", 'Error', 'OK', 'Error') | Out-Null
    }
    finally {
        Set-UiBusy -Busy $false
    }
})

$window.Add_Closing({
    if ($script:Connected) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    }
})

#endregion

Write-UiLog 'Ready (build 2026-09-30 WPF). Enter Client ID and Tenant ID, then press Connect.'
[void]$window.ShowDialog()
