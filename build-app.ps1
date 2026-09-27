$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$frontendRoot = Join-Path $projectRoot 'frontend'
$staticRoot = Join-Path $projectRoot 'backend\server\src\main\resources\static'
$npmCache = Join-Path $projectRoot 'work\npm-cache'
$frontendBuildRoot = $frontendRoot

New-Item -ItemType Directory -Path $npmCache -Force | Out-Null

$localVite = Join-Path $frontendRoot 'node_modules\.bin\vite.cmd'
if (-not (Test-Path -LiteralPath $localVite)) {
    # A running Vite process on Windows can lock its native module while npm ci replaces
    # node_modules. Build from an isolated generated directory so IDEA may stay running.
    $frontendBuildRoot = Join-Path $projectRoot ("work\frontend-build-" + $PID)
    New-Item -ItemType Directory -Path $frontendBuildRoot -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $frontendRoot 'package.json') -Destination $frontendBuildRoot
    Copy-Item -LiteralPath (Join-Path $frontendRoot 'package-lock.json') -Destination $frontendBuildRoot
    Copy-Item -LiteralPath (Join-Path $frontendRoot 'vite.config.js') -Destination $frontendBuildRoot
    Copy-Item -LiteralPath (Join-Path $frontendRoot 'index.html') -Destination $frontendBuildRoot
    Copy-Item -LiteralPath (Join-Path $frontendRoot 'src') -Destination $frontendBuildRoot -Recurse
}

Push-Location $frontendBuildRoot
try {
    $nodeModules = Join-Path $frontendBuildRoot 'node_modules'
    if (-not (Test-Path -LiteralPath $nodeModules)) {
        & npm.cmd --cache $npmCache ci
        if ($LASTEXITCODE -ne 0) {
            throw "Frontend dependency installation failed with exit code $LASTEXITCODE."
        }
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
Copy-Item -Path (Join-Path $frontendBuildRoot 'dist\*') -Destination $staticRoot -Recurse -Force

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
