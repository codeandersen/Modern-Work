<#
.SYNOPSIS
    WinForms GUI to set Windows Autopilot device group tags (delegated Graph auth).

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

# StrictMode disabled: WinForms event/scriptblock variable capture is unreliable under StrictMode
$ErrorActionPreference = 'Stop'

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
$script:Connected = $false
$script:CancelLoad = $false
$script:Busy = $false
$script:PageIndex = 0
$script:PageSize = 100
$script:SortColumn = 'SerialNumber'
$script:SortAscending = $true
$script:SearchText = ''

#region Helpers

function Write-UiLog {
    param([string]$Message, [System.Drawing.Color]$Color = [System.Drawing.Color]::Black)
    if ($null -ne $script:lstLog -and $script:lstLog.InvokeRequired) {
        $script:_uiLogMessage = $Message
        $script:lstLog.Invoke([Action]{
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
        [Parameter(Mandatory)][string]$GroupTag
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
            $script:btnPrev, $script:btnNext, $script:cmbPageSize, $script:btnApply, $script:btnImportCsv,
            $script:txtSearch, $script:txtGroupTag, $script:grid
        )) {
        if ($null -eq $c) { continue }
        if ($c -eq $script:btnCancelLoad) {
            $c.Enabled = $Busy -and $script:CancelLoad -eq $false
            continue
        }
        if ($c -eq $script:btnConnect) { $c.Enabled = $enabled -and -not $script:Connected; continue }
        if ($c -eq $script:btnDisconnect) { $c.Enabled = $enabled -and $script:Connected; continue }
        if ($c -in @($script:btnLoadAll, $script:btnSearch, $script:btnClearSearch, $script:btnPrev, $script:btnNext, $script:cmbPageSize, $script:btnApply, $script:btnImportCsv, $script:txtSearch, $script:txtGroupTag, $script:grid)) {
            $c.Enabled = $enabled -and $script:Connected
            continue
        }
        $c.Enabled = $enabled
    }
    $script:btnCancelLoad.Enabled = $Busy
}

function Update-StatusBar {
    param([string]$Text)
    if ($null -ne $script:lblStatus -and $script:lblStatus.InvokeRequired) {
        $script:_uiStatusText = $Text
        $script:lblStatus.Invoke([Action]{
            Update-StatusBar -Text $script:_uiStatusText
        }) | Out-Null
        return
    }
    if ($null -eq $script:lblStatus) { return }
    $script:lblStatus.Text = $Text
}

function Update-Progress {
    param([int]$Value = 0, [int]$Maximum = 100, [bool]$StyleMarquee = $false)
    if ($null -ne $script:progress -and $script:progress.InvokeRequired) {
        $script:_uiProgressValue = $Value
        $script:_uiProgressMaximum = $Maximum
        $script:_uiProgressMarquee = $StyleMarquee
        $script:progress.Invoke([Action]{
            Update-Progress -Value $script:_uiProgressValue -Maximum $script:_uiProgressMaximum -StyleMarquee:$script:_uiProgressMarquee
        }) | Out-Null
        return
    }
    if ($null -eq $script:progress) { return }
    if ($StyleMarquee) {
        $script:progress.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
    }
    else {
        $script:progress.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
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
        [System.Windows.Forms.Application]::DoEvents()
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
    if ($null -ne $script:grid -and $script:grid.InvokeRequired) {
        $script:grid.Invoke([Action]{ Show-CurrentPage }) | Out-Null
        return
    }

    $pageCount = Get-PageCount
    if ($script:PageIndex -ge $pageCount) { $script:PageIndex = [Math]::Max(0, $pageCount - 1) }
    if ($script:PageIndex -lt 0) { $script:PageIndex = 0 }

    $start = $script:PageIndex * $script:PageSize
    $take = [Math]::Min($script:PageSize, [Math]::Max(0, $script:ViewList.Count - $start))

    $script:grid.SuspendLayout()
    $script:grid.DataSource = $null
    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add('Id', [string])
    [void]$table.Columns.Add('SerialNumber', [string])
    [void]$table.Columns.Add('GroupTag', [string])
    [void]$table.Columns.Add('Model', [string])
    [void]$table.Columns.Add('Manufacturer', [string])
    for ($i = 0; $i -lt $take; $i++) {
        $row = $script:ViewList[$start + $i]
        if ($null -eq $row) { continue }
        [void]$table.Rows.Add($row.Id, $row.SerialNumber, $row.GroupTag, $row.Model, $row.Manufacturer)
    }
    $script:grid.DataSource = $table
    if ($script:grid.Columns['Id']) { $script:grid.Columns['Id'].Visible = $false }
    foreach ($colName in @('SerialNumber', 'GroupTag', 'Model', 'Manufacturer')) {
        if ($script:grid.Columns[$colName]) {
            $script:grid.Columns[$colName].SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic
        }
    }
    $script:grid.ResumeLayout()

    $totalCache = $script:DeviceCache.Count
    $filtered = $script:ViewList.Count
    $script:lblPage.Text = "Page $($script:PageIndex + 1) of $pageCount  |  Showing $take  |  Filtered $filtered of $totalCache"
    $script:btnPrev.Enabled = $script:Connected -and -not $script:Busy -and ($script:PageIndex -gt 0)
    $script:btnNext.Enabled = $script:Connected -and -not $script:Busy -and ($script:PageIndex -lt $pageCount - 1)
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

    $gridRows = New-Object 'System.Collections.Generic.List[System.Windows.Forms.DataGridViewRow]'
    foreach ($gridRow in $script:grid.SelectedRows) { [void]$gridRows.Add($gridRow) }
    foreach ($cell in $script:grid.SelectedCells) { [void]$gridRows.Add($cell.OwningRow) }

    foreach ($gridRow in $gridRows) {
        if ($null -eq $gridRow -or $gridRow.IsNewRow) { continue }
        $id = [string]$gridRow.Cells['Id'].Value
        if ([string]::IsNullOrEmpty($id)) { continue }
        if (-not $seen.Add($id)) { continue }
        $row = New-Object AutopilotRow
        $row.Id = $id
        $row.SerialNumber = [string]$gridRow.Cells['SerialNumber'].Value
        $row.GroupTag = [string]$gridRow.Cells['GroupTag'].Value
        [void]$rows.Add($row)
    }
    return $rows
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

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 200
    $timer.Add_Tick({
        if (-not $handle.IsCompleted) { return }
        $timer.Stop()
        $timer.Dispose()
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

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Autopilot Group Tag'
$form.Size = New-Object System.Drawing.Size(1100, 760)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(960, 640)

# Connection panel (taller; full-width ID fields so GUIDs are not clipped)
$grpConn = New-Object System.Windows.Forms.GroupBox
$grpConn.Text = 'Connection (delegated)'
$grpConn.Location = New-Object System.Drawing.Point(12, 8)
$grpConn.Size = New-Object System.Drawing.Size(1060, 118)
$grpConn.Anchor = 'Top,Left,Right'
$form.Controls.Add($grpConn)

$lblClient = New-Object System.Windows.Forms.Label
$lblClient.Text = 'Client ID'
$lblClient.Location = New-Object System.Drawing.Point(12, 28)
$lblClient.Size = New-Object System.Drawing.Size(70, 20)
$grpConn.Controls.Add($lblClient)

$script:txtClientId = New-Object System.Windows.Forms.TextBox
$script:txtClientId.Location = New-Object System.Drawing.Point(90, 24)
$script:txtClientId.Size = New-Object System.Drawing.Size(720, 24)
$script:txtClientId.Anchor = 'Top,Left,Right'
$script:txtClientId.Font = New-Object System.Drawing.Font('Consolas', 9)
$grpConn.Controls.Add($script:txtClientId)

$script:btnConnect = New-Object System.Windows.Forms.Button
$script:btnConnect.Text = 'Connect'
$script:btnConnect.Location = New-Object System.Drawing.Point(830, 22)
$script:btnConnect.Size = New-Object System.Drawing.Size(100, 28)
$script:btnConnect.Anchor = 'Top,Right'
$grpConn.Controls.Add($script:btnConnect)

$script:btnDisconnect = New-Object System.Windows.Forms.Button
$script:btnDisconnect.Text = 'Disconnect'
$script:btnDisconnect.Location = New-Object System.Drawing.Point(940, 22)
$script:btnDisconnect.Size = New-Object System.Drawing.Size(100, 28)
$script:btnDisconnect.Anchor = 'Top,Right'
$script:btnDisconnect.Enabled = $false
$grpConn.Controls.Add($script:btnDisconnect)

$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Text = 'Tenant ID'
$lblTenant.Location = New-Object System.Drawing.Point(12, 68)
$lblTenant.Size = New-Object System.Drawing.Size(70, 20)
$grpConn.Controls.Add($lblTenant)

$script:txtTenantId = New-Object System.Windows.Forms.TextBox
$script:txtTenantId.Location = New-Object System.Drawing.Point(90, 64)
$script:txtTenantId.Size = New-Object System.Drawing.Size(950, 24)
$script:txtTenantId.Anchor = 'Top,Left,Right'
$script:txtTenantId.Font = New-Object System.Drawing.Font('Consolas', 9)
$grpConn.Controls.Add($script:txtTenantId)

# Toolbar
$pnlTools = New-Object System.Windows.Forms.Panel
$pnlTools.Location = New-Object System.Drawing.Point(12, 136)
$pnlTools.Size = New-Object System.Drawing.Size(1060, 70)
$pnlTools.Anchor = 'Top,Left,Right'
$form.Controls.Add($pnlTools)

$script:btnLoadAll = New-Object System.Windows.Forms.Button
$script:btnLoadAll.Text = 'Load all devices'
$script:btnLoadAll.Location = New-Object System.Drawing.Point(0, 4)
$script:btnLoadAll.Width = 130
$script:btnLoadAll.Enabled = $false
$pnlTools.Controls.Add($script:btnLoadAll)

$script:btnCancelLoad = New-Object System.Windows.Forms.Button
$script:btnCancelLoad.Text = 'Cancel'
$script:btnCancelLoad.Location = New-Object System.Drawing.Point(140, 4)
$script:btnCancelLoad.Width = 80
$script:btnCancelLoad.Enabled = $false
$pnlTools.Controls.Add($script:btnCancelLoad)

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = 'Serial'
$lblSearch.Location = New-Object System.Drawing.Point(240, 8)
$lblSearch.AutoSize = $true
$pnlTools.Controls.Add($lblSearch)

$script:txtSearch = New-Object System.Windows.Forms.TextBox
$script:txtSearch.Location = New-Object System.Drawing.Point(285, 4)
$script:txtSearch.Width = 180
$script:txtSearch.Enabled = $false
$pnlTools.Controls.Add($script:txtSearch)

$script:btnSearch = New-Object System.Windows.Forms.Button
$script:btnSearch.Text = 'Search'
$script:btnSearch.Location = New-Object System.Drawing.Point(475, 4)
$script:btnSearch.Width = 80
$script:btnSearch.Enabled = $false
$pnlTools.Controls.Add($script:btnSearch)

$script:btnClearSearch = New-Object System.Windows.Forms.Button
$script:btnClearSearch.Text = 'Clear'
$script:btnClearSearch.Location = New-Object System.Drawing.Point(560, 4)
$script:btnClearSearch.Width = 70
$script:btnClearSearch.Enabled = $false
$pnlTools.Controls.Add($script:btnClearSearch)

$script:btnImportCsv = New-Object System.Windows.Forms.Button
$script:btnImportCsv.Text = 'Import CSV...'
$script:btnImportCsv.Location = New-Object System.Drawing.Point(650, 4)
$script:btnImportCsv.Width = 110
$script:btnImportCsv.Enabled = $false
$pnlTools.Controls.Add($script:btnImportCsv)

$lblTag = New-Object System.Windows.Forms.Label
$lblTag.Text = 'Group tag'
$lblTag.Location = New-Object System.Drawing.Point(0, 40)
$lblTag.AutoSize = $true
$pnlTools.Controls.Add($lblTag)

$script:txtGroupTag = New-Object System.Windows.Forms.TextBox
$script:txtGroupTag.Location = New-Object System.Drawing.Point(70, 36)
$script:txtGroupTag.Width = 200
$script:txtGroupTag.Enabled = $false
$pnlTools.Controls.Add($script:txtGroupTag)

$script:btnApply = New-Object System.Windows.Forms.Button
$script:btnApply.Text = 'Apply to selected'
$script:btnApply.Location = New-Object System.Drawing.Point(280, 34)
$script:btnApply.Width = 130
$script:btnApply.Enabled = $false
$pnlTools.Controls.Add($script:btnApply)

$lblPageSize = New-Object System.Windows.Forms.Label
$lblPageSize.Text = 'Page size'
$lblPageSize.Location = New-Object System.Drawing.Point(430, 40)
$lblPageSize.AutoSize = $true
$pnlTools.Controls.Add($lblPageSize)

$script:cmbPageSize = New-Object System.Windows.Forms.ComboBox
$script:cmbPageSize.DropDownStyle = 'DropDownList'
$script:cmbPageSize.Location = New-Object System.Drawing.Point(500, 36)
$script:cmbPageSize.Width = 70
@('50', '100', '200', '300', '400', '500') | ForEach-Object { [void]$script:cmbPageSize.Items.Add($_) }
$script:cmbPageSize.SelectedItem = '100'
$script:cmbPageSize.Enabled = $false
$pnlTools.Controls.Add($script:cmbPageSize)

$script:btnPrev = New-Object System.Windows.Forms.Button
$script:btnPrev.Text = '<'
$script:btnPrev.Location = New-Object System.Drawing.Point(590, 34)
$script:btnPrev.Width = 40
$script:btnPrev.Enabled = $false
$pnlTools.Controls.Add($script:btnPrev)

$script:btnNext = New-Object System.Windows.Forms.Button
$script:btnNext.Text = '>'
$script:btnNext.Location = New-Object System.Drawing.Point(635, 34)
$script:btnNext.Width = 40
$script:btnNext.Enabled = $false
$pnlTools.Controls.Add($script:btnNext)

$script:lblPage = New-Object System.Windows.Forms.Label
$script:lblPage.Text = 'Page 0 of 0'
$script:lblPage.Location = New-Object System.Drawing.Point(685, 40)
$script:lblPage.AutoSize = $true
$pnlTools.Controls.Add($script:lblPage)

# Grid
$script:grid = New-Object System.Windows.Forms.DataGridView
$script:grid.Location = New-Object System.Drawing.Point(12, 214)
$script:grid.Size = New-Object System.Drawing.Size(1060, 316)
$script:grid.Anchor = 'Top,Bottom,Left,Right'
$script:grid.ReadOnly = $true
$script:grid.AllowUserToAddRows = $false
$script:grid.AllowUserToDeleteRows = $false
$script:grid.SelectionMode = 'FullRowSelect'
$script:grid.MultiSelect = $true
$script:grid.AutoSizeColumnsMode = 'Fill'
$script:grid.RowHeadersVisible = $false
$script:grid.Enabled = $false
$form.Controls.Add($script:grid)

# Log
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = 'Log'
$lblLog.Location = New-Object System.Drawing.Point(12, 538)
$lblLog.Anchor = 'Bottom,Left'
$lblLog.AutoSize = $true
$form.Controls.Add($lblLog)

$script:lstLog = New-Object System.Windows.Forms.ListBox
$script:lstLog.Location = New-Object System.Drawing.Point(12, 558)
$script:lstLog.Size = New-Object System.Drawing.Size(1060, 80)
$script:lstLog.Anchor = 'Bottom,Left,Right'
$form.Controls.Add($script:lstLog)

# Status / progress
$script:progress = New-Object System.Windows.Forms.ProgressBar
$script:progress.Location = New-Object System.Drawing.Point(12, 648)
$script:progress.Size = New-Object System.Drawing.Size(1060, 18)
$script:progress.Anchor = 'Bottom,Left,Right'
$form.Controls.Add($script:progress)

$script:lblStatus = New-Object System.Windows.Forms.Label
$script:lblStatus.Text = 'Enter Client ID and Tenant ID, then Connect.'
$script:lblStatus.Location = New-Object System.Drawing.Point(12, 670)
$script:lblStatus.AutoSize = $true
$script:lblStatus.Anchor = 'Bottom,Left'
$form.Controls.Add($script:lblStatus)

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
        $script:btnConnect.Enabled = -not $script:Connected
        $script:btnDisconnect.Enabled = $script:Connected
        foreach ($c in @($script:btnLoadAll, $script:btnSearch, $script:btnClearSearch, $script:btnImportCsv, $script:btnApply, $script:txtSearch, $script:txtGroupTag, $script:cmbPageSize, $script:grid)) {
            $c.Enabled = $script:Connected
        }
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
    Show-CurrentPage
    Write-UiLog 'Disconnected.'
    Update-StatusBar 'Disconnected. Enter Client ID / Tenant ID and Connect.'
    $script:btnConnect.Enabled = $true
    $script:btnDisconnect.Enabled = $false
    foreach ($c in @($script:btnLoadAll, $script:btnSearch, $script:btnClearSearch, $script:btnImportCsv, $script:btnApply, $script:txtSearch, $script:txtGroupTag, $script:cmbPageSize, $script:grid, $script:btnPrev, $script:btnNext)) {
        $c.Enabled = $false
    }
})

$script:btnCancelLoad.Add_Click({
    $script:CancelLoad = $true
    Write-UiLog 'Cancel requested...'
})

$script:btnLoadAll.Add_Click({
    if (-not $script:Connected) { return }
    $script:CancelLoad = $false
    Set-UiBusy -Busy $true
    $script:btnCancelLoad.Enabled = $true
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
        $script:btnCancelLoad.Enabled = $false
        $script:btnConnect.Enabled = -not $script:Connected
        $script:btnDisconnect.Enabled = $script:Connected
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
    if ($e.KeyCode -eq 'Enter') {
        $e.SuppressKeyPress = $true
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

$script:cmbPageSize.Add_SelectedIndexChanged({
    if ($script:Busy) { return }
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

$script:grid.Add_ColumnHeaderMouseClick({
    param($eventSender, $e)
    if (-not $script:Connected -or $script:Busy) { return }
    if ($script:DeviceCache.Count -eq 0) { return }
    $col = $script:grid.Columns[$e.ColumnIndex]
    if (-not $col -or $col.Name -eq 'Id') { return }
    if ($script:SortColumn -eq $col.Name) {
        $script:SortAscending = -not $script:SortAscending
    }
    else {
        $script:SortColumn = $col.Name
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

    try {
        Set-UiBusy -Busy $true
        $ok = 0
        $fail = 0
        $i = 0
        $total = $selected.Count
        foreach ($dev in $selected) {
            $i++
            Update-Progress -Value $i -Maximum $total
            Update-StatusBar "Applying tag ($i / $total): $($dev.SerialNumber)"
            try {
                if ($dev.GroupTag -eq $tag) {
                    Write-UiLog "Skip $($dev.SerialNumber) (already '$tag')"
                    $ok++
                    continue
                }
                Set-AutopilotDeviceGroupTagGraph -DeviceId $dev.Id -GroupTag $tag
                # update cache
                foreach ($c in $script:DeviceCache) {
                    if ($c.Id -eq $dev.Id) { $c.GroupTag = $tag; break }
                }
                Write-AutopilotLog -Status Success -SerialNumber $dev.SerialNumber -GroupTag $tag -Message 'Group tag updated'
                Write-UiLog "OK $($dev.SerialNumber) -> $tag"
                $ok++
            }
            catch {
                $fail++
                Write-AutopilotLog -Status Failed -SerialNumber $dev.SerialNumber -GroupTag $tag -Message "$_"
                Write-UiLog "FAIL $($dev.SerialNumber): $_"
            }
        }
        Update-DeviceView
        Update-StatusBar "Apply done. Success: $ok  Failed: $fail"
        [System.Windows.Forms.MessageBox]::Show("Done.`nSuccess: $ok`nFailed: $fail", 'Apply', 'OK', 'Information') | Out-Null
    }
    finally {
        Set-UiBusy -Busy $false
    }
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
        $script:grid.SelectAll()

        $msg = "Found: $($found.Count)`nNot found: $($missing.Count)"
        if ($missing.Count -gt 0 -and $missing.Count -le 20) {
            $msg += "`n`nMissing:`n$($missing -join "`n")"
        }
        elseif ($missing.Count -gt 20) {
            $msg += "`n`n(First 20 missing)`n$(($missing | Select-Object -First 20) -join "`n")"
        }
        Write-UiLog "CSV resolve done. Found $($found.Count), missing $($missing.Count)."
        Update-StatusBar "CSV: $($found.Count) found. Set Group tag and Apply to selected."
        [System.Windows.Forms.MessageBox]::Show($msg + "`n`nEnter Group tag and click Apply to selected.", 'CSV import', 'OK', 'Information') | Out-Null
    }
    catch {
        Write-UiLog "CSV import failed: $_"
        [System.Windows.Forms.MessageBox]::Show("CSV import failed:`n$_", 'Error', 'OK', 'Error') | Out-Null
    }
    finally {
        Set-UiBusy -Busy $false
    }
})

$form.Add_FormClosing({
    if ($script:Connected) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
    }
})

#endregion

Write-UiLog 'Ready (build 2026-03-18c, typed rows). Enter Client ID and Tenant ID, then press Connect.'
[System.Windows.Forms.Application]::EnableVisualStyles()
[void]$form.ShowDialog()
