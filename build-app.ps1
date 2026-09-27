$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$frontendRoot = Join-Path $projectRoot 'frontend'
$staticRoot = Join-Path $projectRoot 'backend\server\src\main\resources\static'
$npmCache = Join-Path $projectRoot 'work\npm-cache'

New-Item -ItemType Directory -Path $npmCache -Force | Out-Null

Push-Location $frontendRoot
try {
    & npm.cmd --cache $npmCache ci
    if ($LASTEXITCODE -ne 0) {
        throw "Frontend dependency installation failed with exit code $LASTEXITCODE."
    }
    & npm.cmd run build
    if ($LASTEXITCODE -ne 0) {
        throw "Frontend build failed with exit code $LASTEXITCODE."
    }
}
finally {
    Pop-Location
}

if (Test-Path -LiteralPath $staticRoot) {
    Get-ChildItem -LiteralPath $staticRoot -Force | Remove-Item -Recurse -Force
}
else {
    New-Item -ItemType Directory -Path $staticRoot -Force | Out-Null
}
Copy-Item -Path (Join-Path $frontendRoot 'dist\*') -Destination $staticRoot -Recurse -Force

Push-Location $projectRoot
try {
    & mvn.cmd -q -pl backend/server -am package -DskipTests
    if ($LASTEXITCODE -ne 0) {
        throw "Backend build failed with exit code $LASTEXITCODE."
    }
}
finally {
    Pop-Location
}

Write-Host 'Application build completed.'
