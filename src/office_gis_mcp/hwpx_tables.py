from __future__ import annotations

import json
import re
import statistics
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from lxml import etree

from .common import local_name, natural_key, read_member, require_file, validate_archive

XML_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, recover=False, huge_tree=False)
SIDES = ("left", "right", "top", "bottom")
SIDE_ELEMENTS = {
    "left": "leftBorder",
    "right": "rightBorder",
    "top": "topBorder",
    "bottom": "bottomBorder",
}
FOOTER_CUE = re.compile(
    r"^(합계|총계|소계|계|주\)|비고|작성|확인|서명|날인|담당|검토|승인|일자)(\s*[:：]|\s*$)|\((서명|날인)\)|^\d{0,4}\s*년\s*\d{0,2}\s*월\s*\d{0,2}\s*일\s*$"
)
NUMERIC_LIKE = re.compile(r"^[\s\d.,+\-/%():~㎡㎥㎜㎝㎞mmcmkm년월일]+$", flags=re.IGNORECASE)


def _section_names(zf: zipfile.ZipFile) -> list[str]:
    names = [
        info.filename
        for info in zf.infolist()
        if re.fullmatch(r"Contents/section\d+\.xml", info.filename, flags=re.IGNORECASE)
    ]
    return sorted(names, key=natural_key)


def _parse_xml(payload: bytes, name: str) -> etree._Element:
    try:
        return etree.fromstring(payload, parser=XML_PARSER)
    except etree.XMLSyntaxError as exc:
        raise ValueError(f"Invalid XML in {name}: {exc}") from exc


def _direct_child(element: etree._Element, name: str) -> etree._Element | None:
    return next((child for child in element if local_name(child.tag) == name), None)


def _attr_int(element: etree._Element | None, name: str, default: int = 0) -> int:
    if element is None:
        return default
    try:
        return int(element.get(name, str(default)))
    except (TypeError, ValueError):
        return default


def _truthy(value: str | None) -> bool:
    return (value or "").strip().casefold() in {"1", "true", "yes", "on"}


def _nearest_cell(element: etree._Element) -> etree._Element | None:
    return next((ancestor for ancestor in element.iterancestors() if local_name(ancestor.tag) == "tc"), None)


def _owned_descendants(cell: etree._Element, name: str) -> list[etree._Element]:
    return [
        item
        for item in cell.iter()
        if local_name(item.tag) == name and _nearest_cell(item) is cell
    ]


def _parse_fill(border_fill: etree._Element) -> dict[str, Any]:
    fill_brush = next((item for item in border_fill.iter() if local_name(item.tag) == "fillBrush"), None)
    if fill_brush is None:
        return {"kind": "none", "face_color": None}
    win_brush = next((item for item in fill_brush.iter() if local_name(item.tag) == "winBrush"), None)
    if win_brush is not None:
        return {
            "kind": "solid_or_pattern",
            "face_color": win_brush.get("faceColor"),
            "hatch_color": win_brush.get("hatchColor"),
            "hatch_style": win_brush.get("hatchStyle"),
            "alpha": win_brush.get("alpha"),
        }
    gradation = next((item for item in fill_brush.iter() if local_name(item.tag) == "gradation"), None)
    if gradation is not None:
        return {
            "kind": "gradation",
            "face_color": None,
            "type": gradation.get("type"),
            "angle": gradation.get("angle"),
            "center_x": gradation.get("centerX"),
            "center_y": gradation.get("centerY"),
            "step": gradation.get("step"),
        }
    image = next((item for item in fill_brush.iter() if local_name(item.tag) == "imgBrush"), None)
    if image is not None:
        return {"kind": "image", "face_color": None, "mode": image.get("mode")}
    return {"kind": "other", "face_color": None}


def _parse_header_styles(root: etree._Element) -> tuple[dict[str, dict[str, Any]], dict[str, str], dict[str, bool]]:
    border_fills: dict[str, dict[str, Any]] = {}
    for element in root.iter():
        if local_name(element.tag) != "borderFill":
            continue
        border_id = element.get("id")
        if border_id is None:
            continue
        sides: dict[str, dict[str, str | None]] = {}
        for side, child_name in SIDE_ELEMENTS.items():
            child = _direct_child(element, child_name)
            sides[side] = {
                "type": child.get("type") if child is not None else None,
                "width": child.get("width") if child is not None else None,
                "color": child.get("color") if child is not None else None,
            }
        border_fills[border_id] = {
            "id": border_id,
            "three_d": _truthy(element.get("threeD")),
            "shadow": _truthy(element.get("shadow")),
            "center_line": element.get("centerLine"),
            "break_cell_separate_line": _truthy(element.get("breakCellSeparateLine")),
            "sides": sides,
            "fill": _parse_fill(element),
        }

    paragraph_alignments: dict[str, str] = {}
    for element in root.iter():
        if local_name(element.tag) != "paraPr" or element.get("id") is None:
            continue
        align = next((item for item in element.iter() if local_name(item.tag) == "align"), None)
        paragraph_alignments[element.get("id", "")] = (align.get("horizontal") if align is not None else None) or "UNKNOWN"

    bold_character_styles: dict[str, bool] = {}
    for element in root.iter():
        if local_name(element.tag) != "charPr" or element.get("id") is None:
            continue
        bold_character_styles[element.get("id", "")] = any(local_name(item.tag) == "bold" for item in element)
    return border_fills, paragraph_alignments, bold_character_styles


def _cell_text(cell: etree._Element) -> str:
    return "".join((item.text or "") for item in _owned_descendants(cell, "t"))


def _cell_alignments(cell: etree._Element, paragraph_alignments: dict[str, str]) -> list[str]:
    values = {
        paragraph_alignments.get(paragraph.get("paraPrIDRef", ""), "UNKNOWN")
        for paragraph in _owned_descendants(cell, "p")
    }
    return sorted(values)


def _cell_bold_ratio(cell: etree._Element, bold_character_styles: dict[str, bool]) -> float:
    total = 0
    bold = 0
    for run in _owned_descendants(cell, "run"):
        length = sum(len(item.text or "") for item in run.iter() if local_name(item.tag) == "t" and _nearest_cell(item) is cell)
        if length <= 0:
            continue
        total += length
        if bold_character_styles.get(run.get("charPrIDRef", ""), False):
            bold += length
    return round(bold / total, 4) if total else 0.0


def _mode(values: list[Any]) -> Any:
    filtered = [value for value in values if value not in (None, "")]
    return Counter(filtered).most_common(1)[0][0] if filtered else None


def _median(values: list[float]) -> float | None:
    return float(statistics.median(values)) if values else None


def _numeric_like(text: str) -> bool:
    stripped = text.strip()
    return bool(stripped and any(char.isdigit() for char in stripped) and NUMERIC_LIKE.fullmatch(stripped))


def _row_cells(cells: list[dict[str, Any]], row: int) -> list[dict[str, Any]]:
    return [cell for cell in cells if cell["row"] <= row < cell["row"] + cell["row_span"]]


def _row_metrics(cells: list[dict[str, Any]], row: int) -> dict[str, Any]:
    anchors = _row_cells(cells, row)
    count = len(anchors)
    nonempty = [cell for cell in anchors if cell["text"].strip()]
    heights = [float(cell["height"]) for cell in anchors if cell["height"] is not None]
    return {
        "anchor_count": count,
        "explicit_header_ratio": sum(1 for cell in anchors if cell["explicit_header"]) / count if count else 0.0,
        "border_fill_mode": _mode([cell["border_fill_id"] for cell in anchors]),
        "fill_mode": _mode([cell["fill_signature"] for cell in anchors]),
        "center_ratio": sum(1 for cell in anchors if "CENTER" in cell["paragraph_alignments"]) / count if count else 0.0,
        "bold_ratio": sum(1 for cell in anchors if cell["bold_ratio"] >= 0.5) / count if count else 0.0,
        "merged_ratio": sum(1 for cell in anchors if cell["row_span"] > 1 or cell["col_span"] > 1) / count if count else 0.0,
        "nonempty_ratio": len(nonempty) / count if count else 0.0,
        "numeric_ratio": sum(1 for cell in nonempty if cell["numeric_like"]) / len(nonempty) if nonempty else 0.0,
        "median_height": _median(heights),
        "texts": [cell["text"] for cell in nonempty],
    }


def _pool_metrics(cells: list[dict[str, Any]], rows: list[int]) -> dict[str, Any]:
    anchors = {cell["anchor"]: cell for row in rows for cell in _row_cells(cells, row)}
    values = list(anchors.values())
    nonempty = [cell for cell in values if cell["text"].strip()]
    heights = [float(cell["height"]) for cell in values if cell["height"] is not None]
    return {
        "border_fill_mode": _mode([cell["border_fill_id"] for cell in values]),
        "fill_mode": _mode([cell["fill_signature"] for cell in values]),
        "center_ratio": sum(1 for cell in values if "CENTER" in cell["paragraph_alignments"]) / len(values) if values else 0.0,
        "bold_ratio": sum(1 for cell in values if cell["bold_ratio"] >= 0.5) / len(values) if values else 0.0,
        "numeric_ratio": sum(1 for cell in nonempty if cell["numeric_like"]) / len(nonempty) if nonempty else 0.0,
        "median_height": _median(heights),
    }


def _header_score(row: int, metrics: dict[str, Any], body: dict[str, Any], repeat_header: bool) -> tuple[float, list[str]]:
    score = 0.0
    evidence: list[str] = []
    if row == 0:
        score += 0.05
        evidence.append("leading_row")
    if row == 0 and repeat_header:
        score += 0.2
        evidence.append("table_repeat_header")
    if metrics["fill_mode"] and body["fill_mode"] and metrics["fill_mode"] != body["fill_mode"]:
        score += 0.25
        evidence.append("fill_differs_from_body")
    if metrics["border_fill_mode"] and body["border_fill_mode"] and metrics["border_fill_mode"] != body["border_fill_mode"]:
        score += 0.1
        evidence.append("border_style_differs_from_body")
    if metrics["center_ratio"] >= 0.6 and metrics["center_ratio"] >= body["center_ratio"] + 0.25:
        score += 0.15
        evidence.append("centered_more_than_body")
    if metrics["bold_ratio"] >= 0.6 and metrics["bold_ratio"] >= body["bold_ratio"] + 0.25:
        score += 0.15
        evidence.append("bold_more_than_body")
    if metrics["merged_ratio"] >= 0.2:
        score += 0.1
        evidence.append("merged_header_pattern")
    if metrics["nonempty_ratio"] >= 0.5:
        score += 0.05
        evidence.append("mostly_nonempty")
    if metrics["numeric_ratio"] <= 0.25 and body["numeric_ratio"] >= 0.4:
        score += 0.1
        evidence.append("labels_above_numeric_body")
    return round(min(score, 1.0), 3), evidence


def _footer_score(metrics: dict[str, Any], body: dict[str, Any]) -> tuple[float, list[str]]:
    score = 0.05
    evidence = ["bottom_band"]
    cue_texts = [text.strip() for text in metrics["texts"] if len(text.strip()) <= 120 and FOOTER_CUE.search(text.strip())]
    if cue_texts:
        score += 0.4
        evidence.append("footer_or_form_text")
    fill_differs = bool(metrics["fill_mode"] and body["fill_mode"] and metrics["fill_mode"] != body["fill_mode"])
    border_differs = bool(
        metrics["border_fill_mode"]
        and body["border_fill_mode"]
        and metrics["border_fill_mode"] != body["border_fill_mode"]
    )
    if fill_differs:
        score += 0.25
        evidence.append("fill_differs_from_body")
    if border_differs:
        score += 0.12
        evidence.append("border_style_differs_from_body")
    if metrics["merged_ratio"] >= 0.5:
        score += 0.2
        evidence.append("wide_merged_form_row")
    if (
        metrics["median_height"] is not None
        and body["median_height"] is not None
        and metrics["median_height"] >= body["median_height"] * 1.25
    ):
        score += 0.15
        evidence.append("height_differs_from_body")
    if metrics["nonempty_ratio"] <= 0.5 and (fill_differs or border_differs):
        score += 0.1
        evidence.append("blank_form_cells_with_distinct_style")
    if metrics["center_ratio"] >= body["center_ratio"] + 0.4 or metrics["bold_ratio"] >= body["bold_ratio"] + 0.4:
        score += 0.1
        evidence.append("alignment_or_weight_differs_from_body")
    return round(min(score, 1.0), 3), evidence


def _border_signature(style: dict[str, Any] | None, side: str) -> dict[str, Any]:
    if style is None:
        return {"type": "UNRESOLVED", "width": None, "color": None}
    raw = style.get("sides", {}).get(side, {})
    return {
        "type": (raw.get("type") or "UNSPECIFIED").upper(),
        "width": raw.get("width"),
        "color": raw.get("color"),
    }


def _summarize_border_records(records: list[dict[str, Any]]) -> dict[str, Any]:
    styles: dict[str, dict[str, Any]] = {}
    side_counts = Counter(record["side"] for record in records)
    visible = 0
    unresolved = 0
    for record in records:
        signature = record["signature"]
        border_type = signature["type"]
        if border_type == "UNRESOLVED":
            unresolved += 1
        elif border_type not in {"NONE", "UNSPECIFIED"}:
            visible += 1
        key = json.dumps(signature, ensure_ascii=False, sort_keys=True)
        entry = styles.setdefault(key, {"signature": signature, "count": 0, "examples": []})
        entry["count"] += 1
        if len(entry["examples"]) < 8:
            entry["examples"].append({"cell": record["cell"], "side": record["side"]})
    return {
        "cell_side_count": len(records),
        "visible_count": visible,
        "hidden_or_none_count": len(records) - visible - unresolved,
        "unresolved_count": unresolved,
        "side_counts": dict(sorted(side_counts.items())),
        "styles": sorted(styles.values(), key=lambda item: (-item["count"], json.dumps(item["signature"], sort_keys=True))),
    }


def _border_profiles(
    cells: list[dict[str, Any]],
    occupancy: dict[tuple[int, int], str],
    row_count: int,
    column_count: int,
) -> dict[str, Any]:
    by_anchor = {cell["anchor"]: cell for cell in cells}
    region_records: dict[str, dict[str, list[dict[str, Any]]]] = defaultdict(lambda: {"outer": [], "internal": []})
    region_cells: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for cell in cells:
        region = cell["role"] if cell["role"] in {"header", "body", "footer"} else "mixed"
        region_cells[region].append(cell)
        outer_by_side = {
            "left": cell["col"] == 0,
            "right": cell["col"] + cell["col_span"] >= column_count,
            "top": cell["row"] == 0,
            "bottom": cell["row"] + cell["row_span"] >= row_count,
        }
        for side in SIDES:
            scope = "outer" if outer_by_side[side] else "internal"
            region_records[region][scope].append(
                {"cell": cell["anchor"], "side": side, "signature": _border_signature(cell["border_style"], side)}
            )

    regions: dict[str, Any] = {}
    for region in ("header", "body", "footer", "mixed"):
        fill_counts: dict[str, dict[str, Any]] = {}
        for cell in region_cells.get(region, []):
            signature = cell["fill_signature"]
            entry = fill_counts.setdefault(
                signature,
                {
                    "border_fill_id": cell["border_fill_id"],
                    "fill": cell["border_style"].get("fill") if cell["border_style"] else None,
                    "count": 0,
                },
            )
            entry["count"] += 1
        regions[region] = {
            "cell_count": len(region_cells.get(region, [])),
            "fills": sorted(fill_counts.values(), key=lambda item: -item["count"]),
            "outer": _summarize_border_records(region_records[region]["outer"]),
            "internal": _summarize_border_records(region_records[region]["internal"]),
        }

    mismatch_count = 0
    mismatches: list[dict[str, Any]] = []

    def compare_boundary(
        orientation: str,
        row: int,
        column: int,
        first_anchor: str | None,
        first_side: str,
        second_anchor: str | None,
        second_side: str,
    ) -> None:
        nonlocal mismatch_count
        if not first_anchor or not second_anchor or first_anchor == second_anchor:
            return
        first = by_anchor[first_anchor]
        second = by_anchor[second_anchor]
        first_signature = _border_signature(first["border_style"], first_side)
        second_signature = _border_signature(second["border_style"], second_side)
        if first_signature == second_signature:
            return
        mismatch_count += 1
        if len(mismatches) < 200:
            mismatches.append(
                {
                    "orientation": orientation,
                    "row": row,
                    "column": column,
                    "first_cell": first_anchor,
                    "first_side": first_side,
                    "first_border": first_signature,
                    "second_cell": second_anchor,
                    "second_side": second_side,
                    "second_border": second_signature,
                }
            )

    for row in range(row_count):
        for column in range(1, column_count):
            compare_boundary(
                "vertical",
                row,
                column,
                occupancy.get((row, column - 1)),
                "right",
                occupancy.get((row, column)),
                "left",
            )
    for row in range(1, row_count):
        for column in range(column_count):
            compare_boundary(
                "horizontal",
                row,
                column,
                occupancy.get((row - 1, column)),
                "bottom",
                occupancy.get((row, column)),
                "top",
            )
    return {
        "regions": regions,
        "shared_boundary_mismatch_count": mismatch_count,
        "shared_boundary_mismatches": mismatches,
        "shared_boundary_mismatches_truncated": mismatch_count > len(mismatches),
    }


def _analyze_table(
    table: etree._Element,
    section_name: str,
    section_table_index: int,
    global_index: int,
    border_fills: dict[str, dict[str, Any]],
    paragraph_alignments: dict[str, str],
    bold_character_styles: dict[str, bool],
    include_cells: bool,
    max_cells: int,
) -> dict[str, Any]:
    declared_rows = _attr_int(table, "rowCnt")
    declared_columns = _attr_int(table, "colCnt")
    repeat_header = _truthy(table.get("repeatHeader"))
    table_border_fill_id = table.get("borderFillIDRef")
    cells: list[dict[str, Any]] = []
    occupancy: dict[tuple[int, int], str] = {}
    overlaps: list[dict[str, Any]] = []

    for fallback_row, row_element in enumerate(child for child in table if local_name(child.tag) == "tr"):
        fallback_column = 0
        for cell_element in (child for child in row_element if local_name(child.tag) == "tc"):
            address = _direct_child(cell_element, "cellAddr")
            span = _direct_child(cell_element, "cellSpan")
            size = _direct_child(cell_element, "cellSz")
            row = _attr_int(address, "rowAddr", fallback_row)
            column = _attr_int(address, "colAddr", fallback_column)
            row_span = max(1, _attr_int(span, "rowSpan", 1))
            column_span = max(1, _attr_int(span, "colSpan", 1))
            anchor = f"R{row}C{column}"
            border_fill_id = cell_element.get("borderFillIDRef")
            border_style = border_fills.get(border_fill_id or "")
            fill_signature = json.dumps(
                border_style.get("fill") if border_style else {"kind": "unresolved", "face_color": None},
                ensure_ascii=False,
                sort_keys=True,
            )
            text = _cell_text(cell_element)
            vertical = _direct_child(cell_element, "subList")
            cell = {
                "anchor": anchor,
                "row": row,
                "col": column,
                "row_span": row_span,
                "col_span": column_span,
                "covers": [[covered_row, covered_col] for covered_row in range(row, row + row_span) for covered_col in range(column, column + column_span)],
                "merged": row_span > 1 or column_span > 1,
                "explicit_header": _truthy(cell_element.get("header")),
                "text": text,
                "text_length": len(text),
                "numeric_like": _numeric_like(text),
                "border_fill_id": border_fill_id,
                "border_style": border_style,
                "fill_signature": fill_signature,
                "paragraph_alignments": _cell_alignments(cell_element, paragraph_alignments),
                "vertical_alignment": vertical.get("vertAlign") if vertical is not None else None,
                "bold_ratio": _cell_bold_ratio(cell_element, bold_character_styles),
                "width": _attr_int(size, "width") if size is not None and size.get("width") is not None else None,
                "height": _attr_int(size, "height") if size is not None and size.get("height") is not None else None,
            }
            cells.append(cell)
            for covered_row in range(row, row + row_span):
                for covered_column in range(column, column + column_span):
                    prior = occupancy.get((covered_row, covered_column))
                    if prior and prior != anchor:
                        overlaps.append({"row": covered_row, "col": covered_column, "first": prior, "second": anchor})
                    else:
                        occupancy[(covered_row, covered_column)] = anchor
            fallback_column = max(fallback_column, column + column_span)

    observed_rows = max((cell["row"] + cell["row_span"] for cell in cells), default=0)
    observed_columns = max((cell["col"] + cell["col_span"] for cell in cells), default=0)
    row_count = max(declared_rows, observed_rows)
    column_count = max(declared_columns, observed_columns)
    gaps = [[row, column] for row in range(row_count) for column in range(column_count) if (row, column) not in occupancy]
    out_of_declared = [
        cell["anchor"]
        for cell in cells
        if declared_rows and declared_columns and (cell["row"] + cell["row_span"] > declared_rows or cell["col"] + cell["col_span"] > declared_columns)
    ]

    row_metrics = {row: _row_metrics(cells, row) for row in range(row_count)}
    explicit_header_rows = [row for row, metrics in row_metrics.items() if metrics["explicit_header_ratio"] >= 0.5]
    baseline_rows = list(range(1, max(1, row_count - 1))) or list(range(row_count))
    body_baseline = _pool_metrics(cells, baseline_rows)
    inferred_header_rows: list[int] = []
    header_scores: dict[int, dict[str, Any]] = {}
    if row_count > 1:
        for row in range(min(3, row_count - 1)):
            if row in explicit_header_rows:
                header_scores[row] = {"score": 1.0, "evidence": ["explicit_header_cells"]}
                continue
            score, evidence = _header_score(row, row_metrics[row], body_baseline, repeat_header)
            header_scores[row] = {"score": score, "evidence": evidence}
            if score >= 0.45:
                inferred_header_rows.append(row)
            else:
                break
    header_rows = sorted(set(explicit_header_rows) | set(inferred_header_rows))

    footer_rows: list[int] = []
    footer_scores: dict[int, dict[str, Any]] = {}
    if row_count > 1:
        for row in range(row_count - 1, max(-1, row_count - 4), -1):
            if row in header_rows:
                break
            preceding_rows = [candidate for candidate in range(row) if candidate not in header_rows]
            preceding = _pool_metrics(cells, preceding_rows)
            score, evidence = _footer_score(row_metrics[row], preceding)
            footer_scores[row] = {"score": score, "evidence": evidence}
            threshold = 0.4 if not footer_rows else 0.35
            if score >= threshold:
                footer_rows.append(row)
            else:
                break
    footer_rows.sort()

    header_row_set = set(header_rows)
    footer_row_set = set(footer_rows)
    for cell in cells:
        covered_rows = set(range(cell["row"], cell["row"] + cell["row_span"]))
        header_overlap = covered_rows & header_row_set
        footer_overlap = covered_rows & footer_row_set
        if cell["explicit_header"]:
            role = "header"
            role_source = "explicit_cell_header"
        elif footer_overlap:
            role = "footer" if footer_overlap == covered_rows else "mixed_footer_body"
            role_source = "inferred_footer_band"
        elif header_overlap:
            role = "header" if header_overlap == covered_rows else "mixed_header_body"
            role_source = "inferred_header_band"
        else:
            role = "body"
            role_source = "remaining_table_content"
        cell["role"] = role
        cell["role_source"] = role_source
        cell["header_overlap"] = "full" if header_overlap == covered_rows and covered_rows else "partial" if header_overlap else "none"
        cell["footer_overlap"] = "full" if footer_overlap == covered_rows and covered_rows else "partial" if footer_overlap else "none"

    row_roles: list[dict[str, Any]] = []
    body_rows: list[int] = []
    for row in range(row_count):
        row_anchor_ids = {occupancy.get((row, column)) for column in range(column_count)} - {None}
        roles = {next(cell["role"] for cell in cells if cell["anchor"] == anchor) for anchor in row_anchor_ids}
        if row in footer_row_set:
            role = "footer"
        elif roles and roles <= {"header"}:
            role = "header"
        elif "header" in roles and len(roles) > 1:
            role = "mixed_header_body"
        else:
            role = "body"
        if role in {"body", "mixed_header_body"}:
            body_rows.append(row)
        row_roles.append(
            {
                "row": row,
                "role": role,
                "header_inference": header_scores.get(row),
                "footer_inference": footer_scores.get(row),
            }
        )

    merged_cells = [
        {
            "anchor": cell["anchor"],
            "row_span": cell["row_span"],
            "col_span": cell["col_span"],
            "covers": cell["covers"],
            "role": cell["role"],
            "header_overlap": cell["header_overlap"],
            "footer_overlap": cell["footer_overlap"],
        }
        for cell in cells
        if cell["merged"]
    ]
    borders = _border_profiles(cells, occupancy, row_count, column_count)
    borders["table_default_border_fill_id"] = table_border_fill_id
    borders["table_default_border_fill"] = border_fills.get(table_border_fill_id or "")

    detailed_cells = []
    if include_cells:
        for cell in cells[:max_cells]:
            detailed_cells.append(
                {
                    "anchor": cell["anchor"],
                    "row": cell["row"],
                    "col": cell["col"],
                    "row_span": cell["row_span"],
                    "col_span": cell["col_span"],
                    "covers": cell["covers"],
                    "merged": cell["merged"],
                    "explicit_header": cell["explicit_header"],
                    "role": cell["role"],
                    "role_source": cell["role_source"],
                    "header_overlap": cell["header_overlap"],
                    "footer_overlap": cell["footer_overlap"],
                    "text": cell["text"][:1000],
                    "text_length": cell["text_length"],
                    "text_truncated": cell["text_length"] > 1000,
                    "numeric_like": cell["numeric_like"],
                    "border_fill_id": cell["border_fill_id"],
                    "border_fill": cell["border_style"],
                    "paragraph_alignments": cell["paragraph_alignments"],
                    "vertical_alignment": cell["vertical_alignment"],
                    "bold_ratio": cell["bold_ratio"],
                    "width": cell["width"],
                    "height": cell["height"],
                }
            )

    warnings = []
    if overlaps:
        warnings.append("Cell spans overlap in the logical occupancy grid.")
    if gaps:
        warnings.append("The declared/observed grid contains coordinates not covered by any cell.")
    if out_of_declared:
        warnings.append("Some cells extend beyond the table's declared rowCnt/colCnt.")
    if inferred_header_rows and not any(cell["explicit_header"] for cell in cells):
        warnings.append("Header rows are heuristic because no cell has an explicit header flag.")
    if footer_rows:
        warnings.append("Bottom-form/footer rows are heuristic candidates, not an explicit HWPX semantic role.")

    return {
        "index": global_index,
        "section": section_name,
        "section_table_index": section_table_index,
        "declared_rows": declared_rows,
        "declared_columns": declared_columns,
        "grid_rows": row_count,
        "grid_columns": column_count,
        "repeat_header": repeat_header,
        "anchor_cell_count": len(cells),
        "occupied_coordinate_count": len(occupancy),
        "structure": {
            "merged_cell_count": len(merged_cells),
            "merged_cells": merged_cells[:1000],
            "merged_cells_truncated": len(merged_cells) > 1000,
            "overlap_count": len(overlaps),
            "overlaps": overlaps[:200],
            "overlaps_truncated": len(overlaps) > 200,
            "gap_count": len(gaps),
            "gaps": gaps[:200],
            "gaps_truncated": len(gaps) > 200,
            "out_of_declared_grid": out_of_declared,
        },
        "roles": {
            "explicit_header_cell_count": sum(1 for cell in cells if cell["explicit_header"]),
            "explicit_header_cells": [cell["anchor"] for cell in cells if cell["explicit_header"]],
            "header_cells": [cell["anchor"] for cell in cells if cell["role"] == "header"],
            "body_cells": [cell["anchor"] for cell in cells if cell["role"] == "body"],
            "footer_form_cells": [cell["anchor"] for cell in cells if cell["role"] == "footer"],
            "mixed_role_cells": [cell["anchor"] for cell in cells if cell["role"].startswith("mixed_")],
            "explicit_header_rows": explicit_header_rows,
            "inferred_header_rows": inferred_header_rows,
            "header_rows": header_rows,
            "body_rows": body_rows,
            "footer_form_rows": footer_rows,
            "row_roles": row_roles,
        },
        "borders": borders,
        "cells": detailed_cells if include_cells else None,
        "cells_truncated": include_cells and len(cells) > len(detailed_cells),
        "warnings": warnings,
    }


def hwpx_analyze_tables(
    path: str,
    table_index: int | None = None,
    include_cells: bool = True,
    max_cells: int = 2000,
) -> dict[str, Any]:
    """Analyze HWPX table topology, merged cells, semantic row roles, fills, and inner/outer borders."""
    source = require_file(path, {".hwpx"})
    if table_index is not None and table_index < 0:
        raise ValueError("table_index must be zero or greater")
    max_cells = max(0, min(max_cells, 10_000))
    with zipfile.ZipFile(source, "r") as zf:
        archive = validate_archive(zf)
        names = set(zf.namelist())
        if "Contents/header.xml" not in names:
            raise ValueError("HWPX package does not contain Contents/header.xml")
        header = _parse_xml(read_member(zf, "Contents/header.xml"), "Contents/header.xml")
        border_fills, paragraph_alignments, bold_character_styles = _parse_header_styles(header)
        table_refs: list[tuple[str, int, etree._Element]] = []
        for section_name in _section_names(zf):
            section = _parse_xml(read_member(zf, section_name), section_name)
            section_tables = [item for item in section.iter() if local_name(item.tag) == "tbl"]
            table_refs.extend((section_name, index, table) for index, table in enumerate(section_tables))
        if table_index is not None and table_index >= len(table_refs):
            raise IndexError(f"table_index {table_index} is out of range for {len(table_refs)} tables")
        selected = range(len(table_refs)) if table_index is None else [table_index]
        tables = [
            _analyze_table(
                table_refs[index][2],
                table_refs[index][0],
                table_refs[index][1],
                index,
                border_fills,
                paragraph_alignments,
                bold_character_styles,
                include_cells,
                max_cells,
            )
            for index in selected
        ]
    return {
        "path": str(Path(source)),
        "table_count": len(table_refs),
        "selected_table_count": len(tables),
        "analysis_version": 1,
        "method": {
            "topology": "cellAddr and cellSpan logical occupancy grid",
            "explicit_header": "tc@header",
            "inferred_header": "leading-row score from repeatHeader, fill, border, alignment, bold, merge, and content patterns",
            "footer_form": "bottom-band score from labels, fill, border, merge, height, alignment, and blank form cells",
            "borders": "borderFillIDRef resolved through Contents/header.xml and grouped by semantic region plus physical outer/internal position",
        },
        "tables": tables,
        **archive,
    }
