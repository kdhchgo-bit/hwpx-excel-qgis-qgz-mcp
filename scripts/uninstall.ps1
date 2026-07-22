$ErrorActionPreference = 'Continue'
foreach ($name in @('hwpx-local', 'excel-local', 'qgis-local')) {
    codex mcp remove $name
}

$userProfile = [Environment]::GetFolderPath('UserProfile')
$skillInstallRoot = Join-Path $userProfile '.agents\skills'
foreach ($name in @('hwpx-local', 'excel-local', 'qgis-local')) {
    $skillPath = Join-Path $skillInstallRoot $name
    if (Test-Path -LiteralPath $skillPath -PathType Container) {
        Remove-Item -LiteralPath $skillPath -Recurse -Force
    }
}
codex mcp list
