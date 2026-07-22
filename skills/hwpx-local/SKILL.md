---
name: hwpx-local
description: Inspect, validate, search, extract, copy-edit, native-open, or export local Hangul HWPX files on Windows. Use when the user names an .hwpx file, asks to preserve Hangul layout, replace HWPX text, check package integrity, extract document text, or convert HWPX to PDF with installed Hancom Hangul.
---

# HWPX Local

Use the verified local HWPX engine directly; it does not depend on an MCP server being visible in the current task.

## Workflow

1. Resolve the exact source file named by the user.
2. Run `hwpx_validate` and `hwpx_inspect` before any edit.
3. For an edit, use an explicit new `output_path`. Keep `overwrite=false` unless the user explicitly authorizes replacing that exact output.
4. Run `hwpx_validate`, `hwpx_inspect`, or `hwpx_find_text` on the output and report its full path.
5. Use native Hangul operations only for open verification or PDF export; they run hidden and close after use.

This skill supports `.hwpx`, not legacy binary `.hwp` editing.

## Run a tool

Build JSON with PowerShell and call the bundled wrapper:

```powershell
$skillRoot = Join-Path ([Environment]::GetFolderPath('UserProfile')) '.agents\skills\hwpx-local'
$runner = Join-Path $skillRoot 'scripts\invoke.ps1'
$payload = @{ path = 'D:\문서\보고서.hwpx'; paragraph_preview_limit = 30 } | ConvertTo-Json -Compress -Depth 20
& $runner -Tool hwpx_inspect -Json $payload -Pretty
```

Do not use long `python -c` commands for Korean paths. Check the returned `ok` field and treat a nonzero exit code as failure.

## Tools

- `hwpx_health`: `{}`
- `hwpx_validate`: `path`
- `hwpx_inspect`: `path`, optional `paragraph_preview_limit`
- `hwpx_extract_text`: `path`, optional `output_path`, `overwrite`
- `hwpx_find_text`: `path`, `query`, optional `regex`, `max_results`
- `hwpx_replace_text`: `path`, `old`, `new`, optional `output_path`, `max_replacements`, `overwrite`
- `hwpx_native_open_check`: `path`
- `hwpx_export_pdf`: `path`, `output_pdf`, optional `overwrite`

`hwpx_replace_text` handles literal text split across HWPX runs, but it is not a general style editor. For font/color/table-layout surgery, inspect the package XML and make the smallest verified package-level change instead of forcing this replacement tool.
