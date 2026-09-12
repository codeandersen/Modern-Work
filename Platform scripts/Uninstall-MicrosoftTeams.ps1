<#
.SYNOPSIS
    Uninstalls Microsoft Teams for all users on the device.

.DESCRIPTION
    This script removes Microsoft Teams AppX packages for all users on the device.
    It includes comprehensive logging and is designed to run as an Intune platform script.

.NOTES
    File Name      : Uninstall-MicrosoftTeams.ps1
    Author         : IT Support
    Prerequisite   : PowerShell 5.1 or later, Administrator privileges
    Created        : $(Get-Date -Format "yyyy-MM-dd")
    
.EXAMPLE
    .\Uninstall-MicrosoftTeams.ps1
#>

# Initialize variables
$LogDirectory = "$env:SystemDrive\STARK_ITSupport\Logs"
$LogFile = "$LogDirectory\Uninstall-MicrosoftTeams_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$ScriptName = "Uninstall-MicrosoftTeams"

# Function to write log entries
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$TimeStamp] [$Level] $Message"
    
    # Write to console
    switch ($Level) {
        "ERROR" { Write-Host $LogEntry -ForegroundColor Red }
        "WARNING" { Write-Host $LogEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $LogEntry -ForegroundColor Green }
        default { Write-Host $LogEntry -ForegroundColor White }
    }
    
    # Write to log file
    try {
        Add-Content -Path $LogFile -Value $LogEntry -ErrorAction Stop
    }
    catch {
        Write-Host "Failed to write to log file: $_" -ForegroundColor Red
    }
}

# Function to create log directory if it doesn't exist
function Initialize-LogDirectory {
    try {
        if (-not (Test-Path -Path $LogDirectory)) {
            New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
            Write-Log "Created log directory: $LogDirectory" -Level "INFO"
        }
        else {
            Write-Log "Log directory already exists: $LogDirectory" -Level "INFO"
        }
    }
    catch {
        Write-Host "Failed to create log directory: $_" -ForegroundColor Red
        exit 1
    }
}

# Function to get Microsoft Teams packages
function Get-TeamsPackages {
    try {
        Write-Log "Searching for Microsoft Teams packages..." -Level "INFO"
        $TeamsPackages = Get-AppxPackage *MSTeams* -AllUsers -ErrorAction Stop
        
        if ($TeamsPackages) {
            Write-Log "Found $($TeamsPackages.Count) Microsoft Teams package(s)" -Level "INFO"
            foreach ($Package in $TeamsPackages) {
                Write-Log "Package: $($Package.Name) - Version: $($Package.Version) - User: $($Package.PackageUserInformation.UserSecurityId)" -Level "INFO"
            }
            return $TeamsPackages
        }
        else {
            Write-Log "No Microsoft Teams packages found" -Level "INFO"
            return $null
        }
    }
    catch {
        Write-Log "Error searching for Microsoft Teams packages: $_" -Level "ERROR"
        return $null
    }
}

# Function to uninstall Microsoft Teams
function Remove-TeamsPackages {
    param(
        [Parameter(Mandatory = $true)]
        $Packages
    )
    
    $SuccessCount = 0
    $FailureCount = 0
    
    foreach ($Package in $Packages) {
        try {
            Write-Log "Attempting to remove package: $($Package.Name)" -Level "INFO"
            $Package | Remove-AppxPackage -AllUsers -ErrorAction Stop
            Write-Log "Successfully removed package: $($Package.Name)" -Level "SUCCESS"
            $SuccessCount++
        }
        catch {
            Write-Log "Failed to remove package $($Package.Name): $_" -Level "ERROR"
            $FailureCount++
        }
    }
    
    return @{
        Success = $SuccessCount
        Failure = $FailureCount
    }
}

# Main execution
try {
    Write-Log "Starting $ScriptName script execution" -Level "INFO"
    Write-Log "Running as user: $env:USERNAME" -Level "INFO"
    Write-Log "Computer name: $env:COMPUTERNAME" -Level "INFO"
    
    # Initialize logging
    Initialize-LogDirectory
    
    # Get Microsoft Teams packages
    $TeamsPackages = Get-TeamsPackages
    
    if ($TeamsPackages) {
        # Uninstall Microsoft Teams packages
        Write-Log "Beginning Microsoft Teams uninstallation process..." -Level "INFO"
        $Results = Remove-TeamsPackages -Packages $TeamsPackages
        
        # Report results
        Write-Log "Uninstallation completed. Success: $($Results.Success), Failures: $($Results.Failure)" -Level "INFO"
        
        if ($Results.Failure -eq 0) {
            Write-Log "Microsoft Teams successfully uninstalled from all users" -Level "SUCCESS"
            exit 0
        }
        else {
            Write-Log "Microsoft Teams uninstallation completed with some failures" -Level "WARNING"
            exit 1
        }
    }
    else {
        Write-Log "No Microsoft Teams packages found to uninstall" -Level "INFO"
        exit 0
    }
}
catch {
    Write-Log "Unexpected error during script execution: $_" -Level "ERROR"
    exit 1
}
finally {
    Write-Log "Script execution completed" -Level "INFO"
    Write-Log "Log file location: $LogFile" -Level "INFO"
}
