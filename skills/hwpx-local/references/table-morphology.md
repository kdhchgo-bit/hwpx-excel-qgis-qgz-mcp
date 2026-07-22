# HWPX table morphology

Use this reference before changing table structure, merged cells, row roles, fills, or borders.

## Interpretation rules

- Build the logical grid from each cell's `cellAddr` and `cellSpan`. A merged cell has one anchor `tc`; the other covered coordinates are not independent cells.
- Prefer `tc@header=1` as explicit header evidence. `repeatHeader` only says that a title band may repeat across pages; it does not identify every header cell by itself.
- Treat `inferred_header_rows` as heuristic. The score compares leading rows with likely body rows using fill, border style, alignment, boldness, merge shape, and label-versus-number patterns.
- Treat `footer_form_rows` as heuristic. It looks for a distinct bottom band such as totals, notes, dates, signatures, approvals, wide merged cells, unusual height, fill, or borders.
- A row can be `mixed_header_body` when a row header or vertically merged header shares the row with normal content.
- Check `warnings`, `row_roles`, and each cell's `role_source` before applying edits.

## Border model

- Resolve each cell's `borderFillIDRef` through `Contents/header.xml`.
- Classify a cell side as `outer` only when the cell span reaches the table's physical top, bottom, left, or right boundary. Otherwise classify it as `internal`.
- Inspect border profiles separately for `header`, `body`, `footer`, and `mixed` regions.
- Treat `shared_boundary_mismatches` as review targets, not automatic corruption. Two adjacent cells can define different styles for the same physical boundary and Hancom rendering rules may choose one side.
- Keep the table-level `borderFillIDRef` separate from per-cell border fills; report both rather than silently substituting one for the other.

## Editing guardrails

1. Select the table by zero-based `index` and verify its section and dimensions.
2. Modify only anchor cells listed in `cells`; use `covers` to protect merged coordinates.
3. Preserve header/body/footer role boundaries unless the user asks for a redesign.
4. When copying a border, copy the complete border-fill definition or reuse its ID consistently; do not change one visual side without checking the adjacent cell's shared boundary.
5. Re-run `hwpx_analyze_tables`, `hwpx_validate`, and a native open or PDF render after structural edits.
