<#
.DESCRIPTION
    Maps network drives on Entra ID joined machines using Cloud Kerberos Trust.
    Group membership is resolved from the Windows logon token (whoami /groups)
    instead of LDAP, which is not available on Entra ID joined devices.

    Deploy via Intune as a PowerShell script (runs as SYSTEM) which will create
    a scheduled task that executes at each user logon.

.NOTES
    Adapted for Entra ID joined + Cloud Kerberos Trust environments.
    Based on intune-drive-mapping-generator pattern by Nicola Suter.
    https://tech.nicolonsky.ch

    PREREQUISITES:
    - Windows Hello for Business or Primary Refresh Token (PRT) must be active
    - File server must be configured for Cloud Kerberos Trust
    - Group names in GroupFilter must match exactly as shown in 'whoami /groups'
      (the part after the backslash, e.g. "LANDBOFORENING\SEC plante1" -> "SEC plante1")
    - Run 'whoami /groups' as the target user to verify exact group names before deploying
#>

[CmdletBinding()]
Param()

###########################################################################################
# Logging setup
###########################################################################################

$logPath = Join-Path $env:temp "DriveMapping-EntraID.log"

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $logPath -Value $entry -Force
    switch ($Level) {
        "WARNING" { Write-Warning $Message }
        "ERROR"   { Write-Error   $Message }
        default   { Write-Output  $entry }
    }
}

# Rotate log if over 2 MB to prevent unbounded growth
if (Test-Path $logPath) {
    if ((Get-Item $logPath).Length -gt 2MB) {
        $archivePath = $logPath -replace '\.log$', '-archive.log'
        Move-Item -Path $logPath -Destination $archivePath -Force
        Write-Log "Log rotated. Previous log archived to: $archivePath"
    }
}

Write-Log "===== DriveMapper-EntraID started ====="
Write-Log "Script version: 1.0 | Host: $env:COMPUTERNAME | User: $env:USERNAME"

###########################################################################################
# Drive mapping configuration
#
# DriveLetter : Single letter (no colon)
# Path        : UNC path. Use $env:USERNAME for per-user home drives.
# Label       : Volume label shown in File Explorer
# GroupFilter : AD group sAMAccountName as it appears AFTER the backslash in whoami /groups
#               Leave empty ("") to map the drive for ALL users regardless of group membership
###########################################################################################

$driveMappingConfig = @(
    # All users
    [PSCustomObject]@{ DriveLetter = "P"; Path = "\\landboforening.dk\data01\brugere\$env:USERNAME"; Label = "Brugere";        GroupFilter = "" },
    [PSCustomObject]@{ DriveLetter = "P"; Path = "\\landboforening.dk\data01\software";              Label = "Software";        GroupFilter = "" },

    # Group-based drives
    [PSCustomObject]@{ DriveLetter = "I"; Path = "\\landboforening.dk\data01\bedrift";               Label = "Bedrift";         GroupFilter = "Bedriftlosning drev i Folle" },
    [PSCustomObject]@{ DriveLetter = "R"; Path = "\\landboforening.dk\DATA01\djursrevision";         Label = "Revision";        GroupFilter = "DL-Revision" },
    [PSCustomObject]@{ DriveLetter = "F"; Path = "\\landboforening.dk\data01\faelles";               Label = "Faelles";         GroupFilter = "SEC Foelle ansatte" },
    [PSCustomObject]@{ DriveLetter = "F"; Path = "\\landboforening.dk\DATA01\lb+_faelles";           Label = "LB Faelles";      GroupFilter = "SEC Foelle ansatte" },
    [PSCustomObject]@{ DriveLetter = "G"; Path = "\\landboforening.dk\data01\admin";                 Label = "Admin";           GroupFilter = "Sekretariat Foelle drev" },
    [PSCustomObject]@{ DriveLetter = "G"; Path = "\\landboforening.dk\data01\plante1";               Label = "Plante1";         GroupFilter = "SEC plante1" },
    [PSCustomObject]@{ DriveLetter = "J"; Path = "\\landboforening.dk\data01\plante2";               Label = "Plante2";         GroupFilter = "SEC plante2" },
    [PSCustomObject]@{ DriveLetter = "N"; Path = "\\landboforening.dk\data01\natur_miljo";           Label = "Natur/Miljo";     GroupFilter = "SEC natur_miljo" },
    [PSCustomObject]@{ DriveLetter = "M"; Path = "\\landboforening.dk\data01\plante_adm";            Label = "Plante Adm";      GroupFilter = "SEC plante_adm" },
    [PSCustomObject]@{ DriveLetter = "K"; Path = "\\landboforening.dk\DATA01\okologi";               Label = "Okologi";         GroupFilter = "SEC Okologi drev" },
    [PSCustomObject]@{ DriveLetter = "L"; Path = "\\landboforening.dk\DATA01\Arbejdsmiljogruppen";   Label = "Arbejdsmiljo";    GroupFilter = "SEC Arbejdsmiljogruppen" },
    [PSCustomObject]@{ DriveLetter = "M"; Path = "\\landboforening.dk\data01\okono_faelles";         Label = "Okonomi";         GroupFilter = "Okonomi drev" },
    [PSCustomObject]@{ DriveLetter = "N"; Path = "\\landboforening.dk\data01\admin";                 Label = "Administration";  GroupFilter = "Administration drev" },
    [PSCustomObject]@{ DriveLetter = "Y"; Path = "\\landboforening.dk\data01\messe";                 Label = "Messe";           GroupFilter = "SEC Messe" }
)

# Set to $true to automatically remove mapped drives that are not in the config above
$removeStaleDrives = $false

###########################################################################################
# Helper functions
###########################################################################################

function Test-RunningAsSystem {
    [CmdletBinding()]
    param()
    process {
        return [bool]($(whoami -user) -match "S-1-5-18")
    }
}

# Retrieves group membership from the Windows logon token.
# On Entra ID joined machines with Cloud Kerberos Trust, cloud-synced security groups
# are present in the token when the user is authenticated via a Primary Refresh Token (PRT).
# Returns an array of group names with the domain prefix stripped.
function Get-TokenGroupMembership {
    [CmdletBinding()]
    param()
    process {
        try {
            Write-Log "Retrieving group memberships from Windows logon token..."

            # whoami /groups /fo csv returns: "Group Name","Type","SID","Attributes"
            $raw = whoami /groups /fo csv 2>$null | ConvertFrom-Csv

            if (-not $raw) {
                Write-Log "No groups returned by whoami /groups - user may not have a valid logon token yet" -Level WARNING
                return @()
            }

            $groups = foreach ($entry in $raw) {
                $groupName = $entry.'Group Name'
                # Strip domain prefix: "LANDBOFORENING\SEC plante1" -> "SEC plante1"
                if ($groupName -match '\\') {
                    $groupName.Split('\')[1]
                }
                else {
                    $groupName
                }
            }

            Write-Log "Token groups resolved ($($groups.Count) total):"
            foreach ($g in ($groups | Sort-Object)) {
                Write-Log "  - $g" -Level DEBUG
            }

            return $groups
        }
        catch {
            Write-Log "Failed to retrieve token group membership: $($_.Exception.Message)" -Level ERROR
            return @()
        }
    }
}

# Returns only drives that apply to the current user based on group membership
function Get-ApplicableDrives {
    [CmdletBinding()]
    param(
        [array]$DriveMappingConfig,
        [array]$GroupMemberships
    )
    process {
        $result = foreach ($drive in $DriveMappingConfig) {
            if ([string]::IsNullOrEmpty($drive.GroupFilter)) {
                Write-Log "Drive $($drive.DriveLetter): '$($drive.Path)' applies to ALL users" -Level DEBUG
                $drive
            }
            elseif ($GroupMemberships -contains $drive.GroupFilter) {
                Write-Log "Drive $($drive.DriveLetter): '$($drive.Path)' - group match: '$($drive.GroupFilter)'" -Level DEBUG
                $drive
            }
            else {
                Write-Log "Drive $($drive.DriveLetter): '$($drive.Path)' - SKIPPED (not member of '$($drive.GroupFilter)')" -Level DEBUG
            }
        }
        return $result
    }
}

# Writes a diagnostic summary of the Entra ID join state and Kerberos status
function Write-DiagnosticInfo {
    [CmdletBinding()]
    param()
    process {
        Write-Log "--- Diagnostic Information ---"
        Write-Log "Computer : $env:COMPUTERNAME"
        Write-Log "User     : $env:USERDOMAIN\$env:USERNAME"
        Write-Log "UPN      : $(whoami /upn 2>$null)"

        # dsregcmd /status - check join state and PRT/Kerberos
        try {
            $dsreg = dsregcmd /status 2>$null
            $azureAdJoined   = ($dsreg | Select-String "AzureAdJoined\s*:\s*(\S+)")  -replace '.*:\s*', ''
            $domainJoined    = ($dsreg | Select-String "DomainJoined\s*:\s*(\S+)")   -replace '.*:\s*', ''
            $prtStatus       = ($dsreg | Select-String "AzureAdPrt\s*:\s*(\S+)")     -replace '.*:\s*', ''
            $kerberosCloud   = ($dsreg | Select-String "CloudTGT\s*:\s*(\S+)")       -replace '.*:\s*', ''

            Write-Log "AzureAdJoined    : $azureAdJoined"
            Write-Log "DomainJoined     : $domainJoined"
            Write-Log "AzureAdPrt       : $prtStatus"
            Write-Log "CloudTGT         : $kerberosCloud"

            if ($prtStatus -notmatch "YES") {
                Write-Log "WARNING: No valid PRT detected. Cloud Kerberos tickets cannot be acquired. Drive mapping may fail." -Level WARNING
            }
            if ($kerberosCloud -notmatch "YES") {
                Write-Log "WARNING: CloudTGT not present. Verify Cloud Kerberos Trust is configured correctly." -Level WARNING
            }
        }
        catch {
            Write-Log "Could not retrieve dsregcmd status: $($_.Exception.Message)" -Level WARNING
        }

        # Check network connectivity to the file server domain
        try {
            $domain = "landboforening.dk"
            Write-Log "Testing DNS resolution for domain: $domain"
            $resolved = [System.Net.Dns]::GetHostAddresses($domain) | Select-Object -First 1
            Write-Log "DNS resolved '$domain' -> $($resolved.IPAddressToString)"
        }
        catch {
            Write-Log "DNS resolution failed for '$domain': $($_.Exception.Message)" -Level WARNING
        }

        Write-Log "--- End Diagnostic Information ---"
    }
}

###########################################################################################
# Main drive mapping logic (runs in user context only)
###########################################################################################

Write-Log "Running as SYSTEM: $(Test-RunningAsSystem)"

if (-not (Test-RunningAsSystem)) {

    # Write diagnostic info to log for troubleshooting
    Write-DiagnosticInfo

    # Get group memberships from token
    $groupMemberships = Get-TokenGroupMembership

    # Filter config to applicable drives
    $applicableDrives = Get-ApplicableDrives -DriveMappingConfig $driveMappingConfig -GroupMemberships $groupMemberships
    Write-Log "Applicable drives for this user: $($applicableDrives.Count)"

    # Snapshot currently mapped drives
    $psDrives = Get-PSDrive | Where-Object {
        $_.Provider.Name -eq "FileSystem" -and $_.Root -notin @("$env:SystemDrive\", "D:\")
    } | Select-Object @{N = "DriveLetter"; E = { $_.Name } }, @{N = "Path"; E = { $_.DisplayRoot } }

    Write-Log "Currently mapped drives: $($psDrives.Count)"
    foreach ($d in $psDrives) {
        Write-Log "  Existing: $($d.DriveLetter): -> $($d.Path)" -Level DEBUG
    }

    foreach ($drive in $applicableDrives) {
        try {
            # Expand environment variables in path (e.g. $env:USERNAME for home drives)
            if ($drive.Path -match '\$env:') {
                $expandedPath = $ExecutionContext.InvokeCommand.ExpandString($drive.Path)
                Write-Log "Expanded path '$($drive.Path)' -> '$expandedPath'" -Level DEBUG
                $drive.Path = $expandedPath
            }

            if ($null -eq $drive.Label) { $drive.Label = "" }

            $exists = $psDrives | Where-Object {
                $_.Path -eq $drive.Path -or $_.DriveLetter -eq $drive.DriveLetter
            }
            $process = $true

            if ($null -ne $exists -and ($exists.Path -eq $drive.Path -and $exists.DriveLetter -eq $drive.DriveLetter)) {
                Write-Log "Drive '$($drive.DriveLetter):' '$($drive.Path)' already mapped correctly - skipping"
                $process = $false
            }
            else {
                if ($null -ne $exists) {
                    Write-Log "Drive '$($drive.DriveLetter):' exists but with wrong config - remapping" -Level WARNING
                }
                # Remove any conflicting mapping on this letter or path
                Get-PSDrive | Where-Object {
                    $_.DisplayRoot -eq $drive.Path -or $_.Name -eq $drive.DriveLetter
                } | ForEach-Object {
                    Write-Log "Removing existing drive '$($_.Name):' before remapping" -Level DEBUG
                    Remove-PSDrive $_ -ErrorAction SilentlyContinue
                }
            }

            if ($process) {
                Write-Log "Mapping '$($drive.DriveLetter):' -> '$($drive.Path)' (Label: '$($drive.Label)')"

                # Verify path is reachable before attempting to map
                $reachable = Test-Path $drive.Path
                if (-not $reachable) {
                    Write-Log "Path '$($drive.Path)' is NOT reachable. Check Kerberos ticket and permissions." -Level WARNING
                }

                $null = New-PSDrive -PSProvider FileSystem -Name $drive.DriveLetter -Root $drive.Path `
                    -Description $drive.Label -Persist -Scope global -ErrorAction Stop

                # Set friendly label visible in File Explorer
                (New-Object -ComObject Shell.Application).NameSpace("$($drive.DriveLetter):").Self.Name = $drive.Label

                Write-Log "Successfully mapped '$($drive.DriveLetter):' -> '$($drive.Path)'"
            }
        }
        catch {
            Write-Log "FAILED mapping '$($drive.DriveLetter):' -> '$($drive.Path)': $($_.Exception.Message)" -Level ERROR

            # Extra diagnostics on failure
            try {
                $pingResult = Test-Connection -ComputerName ($drive.Path -replace '\\\\([^\\]+)\\.*', '$1') -Count 1 -Quiet
                Write-Log "Network ping to file server: $pingResult" -Level DEBUG
            }
            catch {
                Write-Log "Could not ping file server: $($_.Exception.Message)" -Level DEBUG
            }
        }
    }

    # Remove stale drives not present in the current config
    if ($removeStaleDrives -and $null -ne $psDrives) {
        $diff = Compare-Object -ReferenceObject $applicableDrives -DifferenceObject $psDrives `
            -Property "DriveLetter" -PassThru | Where-Object { $_.SideIndicator -eq "=>" }
        foreach ($staleDrive in $diff) {
            Write-Log "Removing unassigned stale drive '$($staleDrive.DriveLetter):'" -Level WARNING
            Remove-SmbMapping -LocalPath "$($staleDrive.DriveLetter):" -Force -UpdateProfile
        }
    }

    # Ensure all mapped drives are marked as persistent in the registry
    $null = Get-ChildItem -Path HKCU:\Network -ErrorAction SilentlyContinue | ForEach-Object {
        New-ItemProperty -Name ConnectionType -Value 1 -Path $_.PSPath -Force -ErrorAction SilentlyContinue
    }

    Write-Log "Drive mapping completed. Log file: $logPath"
}

###########################################################################################
# End of user-context execution
###########################################################################################

Write-Log "===== DriveMapper-EntraID finished ====="

#!SCHTASKCOMESHERE!#

###########################################################################################
# If running as SYSTEM (Intune IME), create a recurring logon scheduled task
###########################################################################################

if (Test-RunningAsSystem) {

    $systemLogPath = Join-Path $env:temp "IntuneDriveMappingScheduledTask.log"
    Start-Transcript -Path $systemLogPath
    Write-Output "Running as SYSTEM --> creating scheduled task for user logon"

    # Save the user-context portion of the script (everything before the marker)
    $currentScript  = Get-Content -Path $PSCommandPath
    $schtaskScript  = $currentScript[(0)..($currentScript.IndexOf("#!SCHTASKCOMESHERE!#") - 1)]

    $scriptSavePath = Join-Path $env:ProgramData "intune-drive-mapping-generator"
    if (-not (Test-Path $scriptSavePath)) {
        New-Item -ItemType Directory -Path $scriptSavePath -Force | Out-Null
    }

    $scriptPath = Join-Path $scriptSavePath "DriveMapping.ps1"
    $schtaskScript | Out-File -FilePath $scriptPath -Force
    Write-Output "Script saved to: $scriptPath"

    # VBScript wrapper - suppresses the PowerShell console window at logon
    $vbsDummyScript = @"
Dim shell,fso,file
Set shell=CreateObject("WScript.Shell")
Set fso=CreateObject("Scripting.FileSystemObject")
strPath=WScript.Arguments.Item(0)
If fso.FileExists(strPath) Then
    set file=fso.GetFile(strPath)
    strCMD="powershell -nologo -executionpolicy ByPass -command " & Chr(34) & "&{" & file.ShortPath & "}" & Chr(34)
    shell.Run strCMD,0
End If
"@

    $vbsPath = Join-Path $scriptSavePath "IntuneDriveMapping-VBSHelper.vbs"
    $vbsDummyScript | Out-File -FilePath $vbsPath -Force
    Write-Output "VBS helper saved to: $vbsPath"

    $wscriptPath = Join-Path $env:SystemRoot "System32\wscript.exe"

    $schtaskName        = "IntuneDriveMapping"
    $schtaskDescription = "Map network drives on Entra ID joined machines (Cloud Kerberos Trust)"

    $trigger   = New-ScheduledTaskTrigger -AtLogOn
    # Run in the context of all interactive users (Builtin\Users)
    $principal = New-ScheduledTaskPrincipal -GroupId "S-1-5-32-545" -Id "Author"
    $action    = New-ScheduledTaskAction -Execute $wscriptPath -Argument "`"$vbsPath`" `"$scriptPath`""
    $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

    $null = Register-ScheduledTask -TaskName $schtaskName -Trigger $trigger -Action $action `
        -Principal $principal -Settings $settings -Description $schtaskDescription -Force

    Write-Output "Scheduled task '$schtaskName' registered successfully"

    Start-ScheduledTask -TaskName $schtaskName

    Stop-Transcript
}

###########################################################################################
# Done
###########################################################################################
