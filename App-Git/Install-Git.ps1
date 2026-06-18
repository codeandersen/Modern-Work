#Requires -RunAsAdministrator

$ErrorActionPreference = 'Stop'
$logDir  = "C:\appinstalls"
$logFile = "$logDir\Install-Git.log"

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$timestamp  $Message" | Out-File -FilePath $logFile -Append -Encoding utf8
    Write-Output "$timestamp  $Message"
}

try {
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir | Out-Null
    }

    Write-Log "Starting Git installation via winget."

    # Ensure winget is available
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) {
        throw "winget not found. Ensure App Installer is installed."
    }

    Write-Log "winget found at: $($winget.Source)"

    # Run the install
    $result = winget install --id Git.Git -e --source winget --silent --accept-package-agreements --accept-source-agreements
    Write-Log "winget exit code: $LASTEXITCODE"
    Write-Log "winget output: $result"

    # -1978335189 (0x8A150011) = already installed, treat as success
    if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne -1978335189) {
        throw "winget returned exit code $LASTEXITCODE"
    }

    # Refresh PATH
    $env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path', 'User')

    # Verify git.exe exists at expected path
    $gitExe = "${env:ProgramFiles}\Git\cmd\git.exe"
    if (Test-Path $gitExe) {
        $version = & $gitExe --version
        Write-Log "Git installed successfully. Version: $version"
        exit 0
    } else {
        throw "git.exe not found at expected path after install."
    }
}
catch {
    Write-Log "ERROR: $_"
    exit 1
}
