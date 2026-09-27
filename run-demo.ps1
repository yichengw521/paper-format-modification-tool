param(
    [string]$Template,
    [string]$Source,
    [string]$Output,
    [string]$Report
)

$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$workspaceRoot = Split-Path -Parent $projectRoot
$coreRoot = Join-Path $projectRoot 'backend\core'

if (-not $Template) {
    $Template = Join-Path $workspaceRoot '6-0毕业设计说明书_模板_20251016.docx'
}
if (-not $Source) {
    $Source = Join-Path $workspaceRoot '说明书.docx'
}
if (-not $Output) {
    $Output = Join-Path $workspaceRoot 'outputs\formatted-manual.docx'
}
if (-not $Report) {
    $Report = Join-Path $workspaceRoot 'outputs\format-report.json'
}

$maven = (Get-Command mvn -ErrorAction Stop).Source
$java = (Get-Command java -ErrorAction Stop).Source
$classpathFile = Join-Path $coreRoot 'target\runtime-classpath.txt'

Push-Location $coreRoot
try {
    & $maven -q compile dependency:build-classpath "-Dmdep.outputFile=$classpathFile"
    if ($LASTEXITCODE -ne 0) {
        throw "Maven 编译失败，退出码：$LASTEXITCODE"
    }

    $dependencyClasspath = Get-Content -LiteralPath $classpathFile -Raw
    $runtimeClasspath = "$(Join-Path $coreRoot 'target\classes');$($dependencyClasspath.Trim())"
    & $java -cp $runtimeClasspath com.paperformat.core.PaperFormatCli $Template $Source $Output $Report
    if ($LASTEXITCODE -ne 0) {
        throw "格式处理失败，退出码：$LASTEXITCODE"
    }
}
finally {
    Pop-Location
}

Write-Host "已生成：$Output"
Write-Host "修改报告：$Report"
