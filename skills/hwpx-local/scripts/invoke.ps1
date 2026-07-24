param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('hwpx_health', 'hwpx_validate', 'hwpx_inspect', 'hwpx_analyze_tables', 'hwpx_analyze_paragraph_hierarchy', 'hwpx_normalize_paragraph_hierarchy', 'hwpx_extract_text', 'hwpx_find_text', 'hwpx_replace_text', 'hwpx_native_open_check', 'hwpx_export_pdf')]
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
