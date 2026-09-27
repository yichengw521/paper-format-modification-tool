param(
    [int]$Port = 8080,
    [string]$Storage,
    [switch]$SkipBuild
)

$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$workspaceRoot = Split-Path -Parent $projectRoot

function Find-ApplicationJar {
    Get-ChildItem -LiteralPath (Join-Path $projectRoot 'backend\server\target') -Filter 'paper-format-server-*.jar' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch '\.original\.jar$' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

if (-not $Storage) {
    $Storage = Join-Path $workspaceRoot 'storage'
}

$applicationJar = Find-ApplicationJar
if (-not $SkipBuild -or -not $applicationJar) {
    & (Join-Path $projectRoot 'build-app.ps1')
    if ($LASTEXITCODE -ne 0) {
        throw "Application build failed with exit code $LASTEXITCODE."
    }
    $applicationJar = Find-ApplicationJar
}

if (-not $applicationJar) {
    throw 'No runnable paper-format-server JAR was found after the build.'
}

$env:PAPER_FORMAT_PORT = $Port
$env:PAPER_FORMAT_STORAGE = $Storage

Write-Host 'Paper Format Modification Tool is starting at:'
Write-Host "http://127.0.0.1:$Port"
Write-Host "Server package: $($applicationJar.Name)"
Write-Host 'Press Ctrl+C to stop the service.'
& java -jar $applicationJar.FullName
