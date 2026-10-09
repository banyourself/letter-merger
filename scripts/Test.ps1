param([switch]$IncludeUI, [string]$MSBuildPath = '')
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
if (-not $MSBuildPath) {
    $command = Get-Command MSBuild.exe -ErrorAction SilentlyContinue
    if (-not $command) { throw 'Run this from Visual Studio Developer PowerShell with the .NET Framework 4.8 targeting pack installed.' }
    $MSBuildPath = $command.Source
}
& (Join-Path $PSScriptRoot 'Build.ps1') -MSBuildPath $MSBuildPath
& $MSBuildPath (Join-Path $projectRoot 'tests\LetterMerger.Tests.csproj') /t:Rebuild /p:Configuration=Release /nologo /verbosity:minimal
if ($LASTEXITCODE -ne 0) { throw 'Test build failed.' }
$runner = Join-Path $projectRoot 'tests\bin\Release\LetterMerger.Tests.exe'
$runtimeRoot = Join-Path $projectRoot 'Runtime'
$resultRoot = Join-Path $projectRoot ('TestResults\Run_' + [Guid]::NewGuid().ToString('N'))
& $runner --core $runtimeRoot (Join-Path $resultRoot 'Core')
if ($LASTEXITCODE -ne 0) { throw 'Core regression checks failed.' }
if ($IncludeUI) {
    & $runner --ui $runtimeRoot (Join-Path $resultRoot 'UI')
    if ($LASTEXITCODE -ne 0) { throw 'Interactive UI regression checks failed.' }
}
Write-Output ('Fictional test results: ' + $resultRoot)
