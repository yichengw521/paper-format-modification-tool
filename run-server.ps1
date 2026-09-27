param(
    [int]$Port = 8080,
    [string]$Storage
)

$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$workspaceRoot = Split-Path -Parent $projectRoot

if (-not $Storage) {
    $Storage = Join-Path $workspaceRoot 'storage'
}

$env:PAPER_FORMAT_PORT = $Port
$env:PAPER_FORMAT_STORAGE = $Storage

Push-Location $projectRoot
try {
    & mvn -q -pl backend/server -am package -DskipTests
    if ($LASTEXITCODE -ne 0) {
        throw "Backend build failed with exit code $LASTEXITCODE."
    }
    $applicationJar = Get-ChildItem -LiteralPath '.\backend\server\target' -Filter 'paper-format-server-*.jar' |
        Where-Object { $_.Name -notmatch '\.original\.jar$' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if (-not $applicationJar) {
        throw 'No runnable paper-format-server JAR was found after the build.'
    }
    Write-Host "Server package: $($applicationJar.Name)"
    & java -jar $applicationJar.FullName
}
finally {
    Pop-Location
}
