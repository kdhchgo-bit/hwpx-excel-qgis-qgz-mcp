param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('qgis_health', 'qgz_inspect', 'qgz_audit_sources', 'qgz_rebase_paths', 'qgis_list_algorithms', 'qgis_algorithm_help', 'qgis_run_algorithm')]
    [string]$Tool,
    [string]$Json = '{}',
    [switch]$Pretty
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$runtimePathFile = Join-Path $PSScriptRoot 'runtime-path.txt'
$localPython = $null
if ($env:OFFICE_GIS_MCP_PYTHON) {
    $localPython = $env:OFFICE_GIS_MCP_PYTHON
} elseif (Test-Path -LiteralPath $runtimePathFile -PathType Leaf) {
    $localPython = (Get-Content -LiteralPath $runtimePathFile -Raw -Encoding UTF8).Trim()
} else {
    $repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..'))
    $localPython = Join-Path $repositoryRoot '.venv\Scripts\python.exe'
}
if (-not (Test-Path -LiteralPath $localPython -PathType Leaf)) {
    throw "Local office/GIS Python runtime was not found. Run scripts\install.ps1 or set OFFICE_GIS_MCP_PYTHON."
}

$utf8NoBom = [Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = $utf8NoBom
$OutputEncoding = $utf8NoBom
$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
$invokeArguments = @('-m', 'office_gis_mcp.invoke', $Tool)
if ($Pretty) { $invokeArguments += '--pretty' }
$Json | & $localPython @invokeArguments
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
