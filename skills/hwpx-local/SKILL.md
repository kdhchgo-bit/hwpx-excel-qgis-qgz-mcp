---
name: hwpx-local
description: Inspect, validate, search, extract, copy-edit, native-open, or export local Hangul HWPX files on Windows, including table morphology and report paragraph hierarchy analysis, hanging-indent/style normalization, and interleave-page exclusion. Use when the user names an .hwpx file, asks to preserve Hangul layout, edit or understand a table, normalize numbered outline levels such as 1./가./1)/가), distinguish divider pages, replace HWPX text, check package integrity, extract document text, or convert HWPX to PDF with installed Hancom Hangul.
---

# HWPX Local

Use the verified local HWPX engine directly; it does not depend on an MCP server being visible in the current task.

## Workflow

1. Resolve the exact source file named by the user.
2. Run `hwpx_validate` and `hwpx_inspect` before any edit.
3. For table work, run `hwpx_analyze_tables` before editing. Read [references/table-morphology.md](references/table-morphology.md) before structural or border changes.
4. For report outline or indentation work, run `hwpx_analyze_paragraph_hierarchy` first and read [references/paragraph-hierarchy.md](references/paragraph-hierarchy.md). Keep interleave exclusion enabled unless the user wants divider pages included. Run normalization as a dry run before applying it.
5. For an edit, use an explicit new `output_path`. Keep `overwrite=false` unless the user explicitly authorizes replacing that exact output.
6. Run `hwpx_validate`, `hwpx_inspect`, or `hwpx_find_text` on the output and report its full path.
7. Use native Hangul operations only for open verification or PDF export; they run hidden and close after use.

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
- `hwpx_analyze_tables`: `path`, optional zero-based `table_index`, `include_cells`, `max_cells`
- `hwpx_analyze_paragraph_hierarchy`: `path`, optional `include_tables`, `exclude_interleaves`, `interleave_min_score`, `reference_paragraphs`, `max_paragraphs`
- `hwpx_normalize_paragraph_hierarchy`: `path`, optional `output_path`, `include_tables`, `exclude_interleaves`, `interleave_min_score`, `reference_paragraphs`, `min_level_confidence`, `min_style_consensus`, `apply_character_style`, `dry_run`, `max_changes`, `overwrite`
- `hwpx_extract_text`: `path`, optional `output_path`, `overwrite`
- `hwpx_find_text`: `path`, `query`, optional `regex`, `max_results`
- `hwpx_replace_text`: `path`, `old`, `new`, optional `output_path`, `max_replacements`, `overwrite`
- `hwpx_native_open_check`: `path`
- `hwpx_export_pdf`: `path`, `output_pdf`, optional `overwrite`

`hwpx_replace_text` handles literal text split across HWPX runs, but it is not a general style editor. For font/color/table-layout surgery, inspect the package XML and make the smallest verified package-level change instead of forcing this replacement tool.

Treat `tc@header` as explicit evidence. Treat inferred header and bottom-form rows as scored candidates and preserve their evidence and warnings. Address merged cells through their reported anchor and `covers` coordinates; never assume every logical grid coordinate owns a separate `tc` element.

For numbered report paragraphs, reuse a verified same-document style instead of inventing spaces or indentation values. Preserve visible-marker/heading conflicts for review, exclude table and interleave paragraphs by default, and never apply a normalization batch without reviewing the dry-run plan.
