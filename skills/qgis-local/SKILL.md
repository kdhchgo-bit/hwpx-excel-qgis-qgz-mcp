---
name: qgis-local
description: Inspect and audit local QGIS QGS/QGZ projects, rebase layer paths, discover Processing algorithms, inspect algorithm schemas, and run qgis_process on Windows. Use when the user names a .qgs or .qgz project, asks about broken GIS sources or project portability, or needs a QGIS Processing operation with explicit output files.
---

# QGIS Local

Use the verified local QGIS engine directly; it does not depend on an MCP server being visible in the current task.

## Workflow

1. Resolve the exact `.qgs` or `.qgz` project and run `qgz_inspect`.
2. Run `qgz_audit_sources` before changing project paths.
3. For path rebasing, use an explicit new `output_path`; keep `overwrite=false` unless authorized.
4. Before a Processing run, use `qgis_algorithm_help` to confirm parameter names and output types.
5. Supply explicit output paths in `parameters`, run the algorithm, and verify every created dataset. Reinspect a modified project copy.

## Run a tool

Build JSON with PowerShell and call the bundled wrapper:

```powershell
$skillRoot = Join-Path ([Environment]::GetFolderPath('UserProfile')) '.agents\skills\qgis-local'
$runner = Join-Path $skillRoot 'scripts\invoke.ps1'
$payload = @{ path = 'D:\GIS\사업.qgz'; include_layers = $true; max_layers = 200 } | ConvertTo-Json -Compress -Depth 30
& $runner -Tool qgz_inspect -Json $payload -Pretty
```

Do not use long `python -c` commands for Korean paths. Check the returned `ok` field and treat a nonzero exit code as failure.

## Tools

- `qgis_health`: `{}`
- `qgz_inspect`: `path`, optional `include_layers`, `max_layers`
- `qgz_audit_sources`: `path`
- `qgz_rebase_paths`: `path`, `old_root`, `new_root`, optional `output_path`, `overwrite`
- `qgis_list_algorithms`: optional `search`, `max_results`
- `qgis_algorithm_help`: `algorithm_id`
- `qgis_run_algorithm`: `algorithm_id`, `parameters`, optional `project_path`, `ellipsoid`, `load_plugins`, `timeout_seconds`

Do not infer algorithm parameters from memory. Read the installed algorithm help first because providers and schemas can differ by QGIS installation.
