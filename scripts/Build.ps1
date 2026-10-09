param([string]$MSBuildPath = '')
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
if (-not $MSBuildPath) {
    $command = Get-Command MSBuild.exe -ErrorAction SilentlyContinue
    if (-not $command) { throw 'Run this from Visual Studio Developer PowerShell with the .NET Framework 4.8 targeting pack installed.' }
    $MSBuildPath = $command.Source
}
& $MSBuildPath (Join-Path $projectRoot 'src\LetterMerger.csproj') /t:Rebuild /p:Configuration=Release /nologo /verbosity:minimal
if ($LASTEXITCODE -ne 0) { throw 'Application build failed.' }
$application = Join-Path $projectRoot 'src\bin\Release\LetterMerger.exe'
$runtimeRoot = Join-Path $projectRoot 'Runtime'
foreach ($folder in @('Data', 'Photos', 'Templates', 'Output')) {
    [void][System.IO.Directory]::CreateDirectory((Join-Path $runtimeRoot $folder))
}
Copy-Item -LiteralPath $application -Destination (Join-Path $runtimeRoot 'LetterMerger.exe') -Force
Write-Output 'Built Runtime\LetterMerger.exe. The existing input files and outputs were preserved.'
