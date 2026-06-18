$gitExe = "${env:ProgramFiles}\Git\cmd\git.exe"

if (Test-Path $gitExe) {
    $version = & $gitExe --version 2>&1
    Write-Output "Git detected: $version"
    exit 0
} else {
    exit 1
}
