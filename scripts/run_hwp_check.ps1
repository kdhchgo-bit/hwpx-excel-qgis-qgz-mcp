$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $projectRoot '.venv\Scripts\python.exe'
$artifactRoot = Join-Path $projectRoot 'test-artifacts'
$stdout = Join-Path $artifactRoot 'hwp_native_stdout.txt'
$stderr = Join-Path $artifactRoot 'hwp_native_stderr.txt'

Remove-Item -LiteralPath $stdout -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $stderr -Force -ErrorAction SilentlyContinue

$process = Start-Process `
    -FilePath $python `
    -ArgumentList @((Join-Path $PSScriptRoot 'verify_hwp_native.py')) `
    -WorkingDirectory $projectRoot `
    -WindowStyle Hidden `
    -RedirectStandardOutput $stdout `
    -RedirectStandardError $stderr `
    -PassThru

$finished = $process.WaitForExit(60000)
if (-not $finished) {
    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    throw 'HWP native verification timed out after 60 seconds'
}

Write-Output "EXIT=$($process.ExitCode)"
if (Test-Path -LiteralPath $stdout) {
    Get-Content -Raw -LiteralPath $stdout
}
if (Test-Path -LiteralPath $stderr) {
    Get-Content -Raw -LiteralPath $stderr
}
exit $process.ExitCode
