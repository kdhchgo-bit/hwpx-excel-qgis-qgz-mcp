param(
    [string]$QgisProcessPath = 'D:\bin\qgis_process-qgis-ltr.bat',
    [switch]$SkipSkills
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $projectRoot

uv sync --python 3.11
$python = Join-Path $projectRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python)) {
    throw "Virtual environment Python was not created: $python"
}

foreach ($name in @('hwpx-local', 'excel-local', 'qgis-local')) {
    codex mcp remove $name 2>$null
}

codex mcp add hwpx-local --env PYTHONUTF8=1 --env PYTHONIOENCODING=utf-8 -- $python -m office_gis_mcp.server hwpx
codex mcp add excel-local --env PYTHONUTF8=1 --env PYTHONIOENCODING=utf-8 -- $python -m office_gis_mcp.server excel
codex mcp add qgis-local --env PYTHONUTF8=1 --env PYTHONIOENCODING=utf-8 --env "QGIS_PROCESS_PATH=$QgisProcessPath" -- $python -m office_gis_mcp.server qgis

if (-not $SkipSkills) {
    $userProfile = [Environment]::GetFolderPath('UserProfile')
    $skillInstallRoot = Join-Path $userProfile '.agents\skills'
    New-Item -ItemType Directory -Path $skillInstallRoot -Force | Out-Null

    foreach ($name in @('hwpx-local', 'excel-local', 'qgis-local')) {
        $sourceSkill = Join-Path $projectRoot "skills\$name"
        $destinationSkill = Join-Path $skillInstallRoot $name
        if (-not (Test-Path -LiteralPath $sourceSkill -PathType Container)) {
            throw "Bundled skill was not found: $sourceSkill"
        }
        New-Item -ItemType Directory -Path $destinationSkill -Force | Out-Null
        Get-ChildItem -LiteralPath $sourceSkill -Force | ForEach-Object {
            Copy-Item -LiteralPath $_.FullName -Destination $destinationSkill -Recurse -Force
        }
        $runtimePath = Join-Path $destinationSkill 'scripts\runtime-path.txt'
        [IO.File]::WriteAllText($runtimePath, $python, [Text.UTF8Encoding]::new($false))
    }

    Write-Host "Installed Codex skills to $skillInstallRoot"
}

Write-Host "Installed MCP servers from $projectRoot"
codex mcp list
