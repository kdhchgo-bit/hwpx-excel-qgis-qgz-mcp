# Paragraph hierarchy and interleave pages

Use this reference for report-outline work such as `1.`, `가.`, `1)`, and `가)` paragraphs.

## Analyze before editing

Run `hwpx_analyze_paragraph_hierarchy` first. It reports stable paragraph IDs such as `S01-P0042`, page IDs such as `S01-PG0003`, visible marker levels, HWPX heading levels, paragraph properties, styles, run character properties, margins, and canonical same-document styles.

The default level map is:

| Level | Visible markers |
|---|---|
| 1 | `1.` |
| 2 | `가.` |
| 3 | `1)` or `(1)` |
| 4 | `가)` or `(가)` |

Treat a visible-marker/heading-level conflict as review-only. Do not automatically change it.

## Learn from the document

Prefer an existing correctly formatted paragraph in the same document over invented margin or font values. The analyzer chooses the majority signature per level from:

- effective `paraPrIDRef`
- `styleIDRef`
- text-weighted `charPrIDRef`

Use `reference_paragraphs`, for example `{"1":"S01-P0012","2":"S01-P0017"}`, when the majority is not the intended style.

## Interleave toggle

Keep `exclude_interleaves=true` for reports and permit packages containing divider pages. The detector scores sparse pages using centering, large or bold title text, chapter/appendix cues, explicit page boundaries, and table presence. A table strongly lowers the score so an application form is not casually treated as a divider.

Interleave classification is heuristic. Review each returned page score, evidence, and text preview. Set `exclude_interleaves=false` only when divider-page paragraphs should participate in style learning and normalization.

## Normalize safely

Run `hwpx_normalize_paragraph_hierarchy` with its default `dry_run=true` first. Review `changes`, canonical source paragraphs, confidence, consensus, and interleave pages. Then set `dry_run=false` and provide an explicit `output_path`.

Defaults deliberately:

- exclude table paragraphs
- exclude interleave pages
- require level confidence of at least `0.75`
- require same-level style consensus of at least `0.60`
- replace only the dominant base run style and preserve minority inline emphasis
- create a separate output copy

Never insert literal spaces to simulate continuation-line alignment. Reuse the canonical paragraph property containing `margin/intent` and `margin/left` so first-line and continuation-line positions remain HWPX paragraph formatting.

## Example

```powershell
$payload = @{
    path = 'D:\문서\보고서.hwpx'
    exclude_interleaves = $true
    dry_run = $true
} | ConvertTo-Json -Compress -Depth 20
& $runner -Tool hwpx_normalize_paragraph_hierarchy -Json $payload -Pretty
```

After applying, run `hwpx_validate` and preferably `hwpx_native_open_check` or PDF export for visual verification.
