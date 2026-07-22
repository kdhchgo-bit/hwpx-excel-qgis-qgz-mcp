---
name: excel-local
description: Inspect, read, search, copy-edit, recalculate, or export local Microsoft Excel workbooks on Windows. Use for .xlsx, .xlsm, .xltx, and .xltm work, especially when formulas, formatting, macros, native Excel fidelity, exact ranges, or PDF output must be preserved and verified.
---

# Excel Local

Use the verified local Excel engine directly; it does not depend on an MCP server being visible in the current task.

## Workflow

1. Resolve the exact workbook named by the user and run `excel_inspect`.
2. Read the smallest relevant range with formulas visible (`data_only=false`). Use `excel_find` when the location is unknown.
3. Write to an explicit new `output_path`. Keep `overwrite=false` unless the user explicitly authorizes replacing that exact output.
4. Prefer `engine=com` for native Excel fidelity, formulas, macros, and recalculation. Use `openpyxl` only when native Excel behavior is unnecessary.
5. Reopen the output with `excel_inspect` and `excel_read_range`; report the full path and verified cells.

## Run a tool

Build JSON with PowerShell and call the bundled wrapper:

```powershell
$skillRoot = Join-Path ([Environment]::GetFolderPath('UserProfile')) '.agents\skills\excel-local'
$runner = Join-Path $skillRoot 'scripts\invoke.ps1'
$payload = @{ path = 'D:\자료\계산서.xlsx'; sheet_name = '집계'; cell_range = 'A1:F20'; data_only = $false } | ConvertTo-Json -Compress -Depth 20
& $runner -Tool excel_read_range -Json $payload -Pretty
```

Do not use long `python -c` commands for Korean paths. Check the returned `ok` field and treat a nonzero exit code as failure.

## Tools

- `excel_health`: `{}`
- `excel_inspect`: `path`, optional `include_defined_names`
- `excel_read_range`: `path`, `sheet_name`, `cell_range`, optional `data_only`, `include_number_formats`
- `excel_find`: `path`, `query`, optional `regex`, `sheet_name`, `max_results`, `max_scanned_cells`
- `excel_write_range`: `path`, `sheet_name`, `start_cell`, rectangular `values`, optional `output_path`, `copy_style_from`, `engine`, `recalculate`, `overwrite`
- `excel_recalculate`: `path`, optional `output_path`, `overwrite`
- `excel_export_pdf`: `path`, `output_pdf`, optional `sheet_name`, `overwrite`

`copy_style_from` is one source cell or range copied onto the target before values are written. Keep macro-enabled output extensions unchanged so VBA can be preserved.
