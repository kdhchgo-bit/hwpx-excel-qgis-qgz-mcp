from __future__ import annotations

import re
import statistics
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from lxml import etree

from .common import (
    local_name,
    natural_key,
    output_copy_path,
    read_member,
    require_file,
    rewrite_zip,
    validate_archive,
)

XML_PARSER = etree.XMLParser(
    resolve_entities=False, no_network=True, recover=False, huge_tree=False
)
LEVEL_PATTERNS: tuple[tuple[str, int, re.Pattern[str]], ...] = (
    ("number_dot", 1, re.compile(r"^\s*(?P<marker>\d{1,3}[.．])(?!\d)\s*(?=\S)")),
    ("hangul_dot", 2, re.compile(r"^\s*(?P<marker>[가-힣][.．])\s*(?=\S)")),
    ("number_paren", 3, re.compile(r"^\s*(?P<marker>\d{1,3}\))\s*(?=\S)")),
    ("number_wrapped", 3, re.compile(r"^\s*(?P<marker>\(\d{1,3}\))\s*(?=\S)")),
    ("hangul_paren", 4, re.compile(r"^\s*(?P<marker>[가-힣]\))\s*(?=\S)")),
    ("hangul_wrapped", 4, re.compile(r"^\s*(?P<marker>\([가-힣]\))\s*(?=\S)")),
)
INTERLEAVE_CUE = re.compile(
    r"(?:^|\s)(?:간\s*지|제\s*\d+\s*[장편부]|부\s*록|별\s*첨|첨\s*부|목\s*차|차\s*례|SECTION|PART)(?:\s|$)",
    flags=re.IGNORECASE,
)
TRUTHY = {"1", "true", "yes", "on"}


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


def _truthy(value: str | None) -> bool:
    return (value or "").strip().lower() in TRUTHY


def _integer(value: str | None) -> int | None:
    try:
        return int(value) if value not in (None, "") else None
    except ValueError:
        return None


def _attribute(element: etree._Element | None, *names: str) -> str | None:
    if element is None:
        return None
    lowered = {key.lower(): value for key, value in element.attrib.items()}
    for name in names:
        if name in element.attrib:
            return element.get(name)
        if name.lower() in lowered:
            return lowered[name.lower()]
    return None


def _direct_child(element: etree._Element | None, name: str) -> etree._Element | None:
    if element is None:
        return None
    return next((child for child in element if local_name(child.tag) == name), None)


def _descendant(element: etree._Element | None, name: str) -> etree._Element | None:
    if element is None:
        return None
    return next(
        (
            child
            for child in element.iter()
            if child is not element and local_name(child.tag) == name
        ),
        None,
    )


def _child_or_descendant(
    element: etree._Element | None, name: str
) -> etree._Element | None:
    child = _direct_child(element, name)
    return child if child is not None else _descendant(element, name)


def _nearest_ancestor(element: etree._Element, name: str) -> etree._Element | None:
    parent = element.getparent()
    while parent is not None:
        if local_name(parent.tag) == name:
            return parent
        parent = parent.getparent()
    return None


def _owned_descendants(paragraph: etree._Element, name: str) -> list[etree._Element]:
    return [
        item
        for item in paragraph.iter()
        if local_name(item.tag) == name and _nearest_ancestor(item, "p") is paragraph
    ]


def _paragraph_text(paragraph: etree._Element) -> str:
    return "".join((item.text or "") for item in _owned_descendants(paragraph, "t"))


def _parse_margin(para_property: etree._Element) -> dict[str, str | None]:
    margin = _child_or_descendant(para_property, "margin")
    values: dict[str, str | None] = {}
    for name in ("intent", "left", "right", "prev", "next"):
        child = _direct_child(margin, name)
        values[name] = (
            _attribute(child, "value")
            if child is not None
            else _attribute(margin, name)
        )
    return values


def _parse_header(
    header: etree._Element,
) -> tuple[
    dict[str, dict[str, Any]], dict[str, dict[str, Any]], dict[str, dict[str, Any]]
]:
    paragraph_properties: dict[str, dict[str, Any]] = {}
    character_properties: dict[str, dict[str, Any]] = {}
    styles: dict[str, dict[str, Any]] = {}

    for element in header.iter():
        element_name = local_name(element.tag)
        element_id = element.get("id")
        if element_name == "paraPr" and element_id is not None:
            align = _child_or_descendant(element, "align")
            heading = _child_or_descendant(element, "heading")
            line_spacing = _child_or_descendant(element, "lineSpacing")
            break_setting = _child_or_descendant(element, "breakSetting")
            paragraph_properties[element_id] = {
                "id": element_id,
                "alignment": _attribute(align, "horizontal"),
                "vertical_alignment": _attribute(align, "vertical"),
                "heading_type": (_attribute(heading, "type") or "NONE").upper(),
                "heading_id_ref": _attribute(heading, "idRef"),
                "heading_level_raw": _integer(_attribute(heading, "level")),
                "margin": _parse_margin(element),
                "line_spacing": {
                    "type": _attribute(line_spacing, "type", "lineSpacingType"),
                    "value": _attribute(line_spacing, "value", "lineSpacing"),
                },
                "page_break_before": _truthy(
                    _attribute(break_setting, "pageBreakBefore")
                ),
                "tab_property_id_ref": element.get("tabPrIDRef"),
            }
        elif element_name == "charPr" and element_id is not None:
            font_ref = _child_or_descendant(element, "fontRef")
            character_properties[element_id] = {
                "id": element_id,
                "height": _integer(element.get("height")),
                "text_color": element.get("textColor"),
                "shade_color": element.get("shadeColor"),
                "bold": any(local_name(child.tag) == "bold" for child in element),
                "italic": any(local_name(child.tag) == "italic" for child in element),
                "underline": any(
                    local_name(child.tag) == "underline" for child in element
                ),
                "font_ref": dict(font_ref.attrib) if font_ref is not None else {},
            }
        elif element_name == "style" and element_id is not None:
            styles[element_id] = {
                "id": element_id,
                "name": element.get("name"),
                "english_name": element.get("engName"),
                "type": element.get("type"),
                "para_property_id_ref": element.get("paraPrIDRef"),
                "character_property_id_ref": element.get("charPrIDRef"),
                "next_style_id_ref": element.get("nextStyleIDRef"),
            }
    return paragraph_properties, character_properties, styles


def _level_evidence(
    text: str, paragraph_property: dict[str, Any] | None
) -> dict[str, Any]:
    marker_kind = None
    marker = None
    marker_level = None
    for kind, level, pattern in LEVEL_PATTERNS:
        match = pattern.match(text)
        if match:
            marker_kind = kind
            marker = match.group("marker")
            marker_level = level
            break

    heading_type = (paragraph_property or {}).get("heading_type", "NONE")
    heading_level_raw = (paragraph_property or {}).get("heading_level_raw")
    explicit_level = None
    if heading_type not in (None, "", "NONE") and heading_level_raw is not None:
        explicit_level = heading_level_raw + 1

    evidence: list[str] = []
    conflict = False
    if marker_level is not None:
        evidence.append(f"text_marker:{marker_kind}")
    if explicit_level is not None:
        evidence.append(f"heading:{heading_type.lower()}:{explicit_level}")

    if marker_level is not None and explicit_level is not None:
        if marker_level == explicit_level:
            level = marker_level
            confidence = 1.0
            evidence.append("marker_heading_agree")
        else:
            level = marker_level
            confidence = 0.6
            conflict = True
            evidence.append(f"marker_heading_conflict:{marker_level}!={explicit_level}")
    elif marker_level is not None:
        level = marker_level
        confidence = 0.9
    elif explicit_level is not None:
        level = explicit_level
        confidence = 0.85
    else:
        level = None
        confidence = 0.0

    return {
        "level": level,
        "confidence": confidence,
        "marker": marker,
        "marker_kind": marker_kind,
        "marker_level": marker_level,
        "explicit_heading_level": explicit_level,
        "level_conflict": conflict,
        "evidence": evidence,
    }


def _run_metrics(
    paragraph: etree._Element,
    style: dict[str, Any] | None,
    character_properties: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    weights: Counter[str] = Counter()
    bold_characters = 0
    total_characters = 0
    maximum_height = None
    for run in _owned_descendants(paragraph, "run"):
        text = "".join(
            (item.text or "")
            for item in run.iter()
            if local_name(item.tag) == "t" and _nearest_ancestor(item, "p") is paragraph
        )
        if not text:
            continue
        char_id = run.get("charPrIDRef") or (style or {}).get(
            "character_property_id_ref"
        )
        if char_id:
            weights[char_id] += len(text)
            shape = character_properties.get(char_id, {})
            if shape.get("bold"):
                bold_characters += len(text)
            height = shape.get("height")
            if height is not None:
                maximum_height = (
                    height if maximum_height is None else max(maximum_height, height)
                )
        total_characters += len(text)
    dominant = (
        weights.most_common(1)[0][0]
        if weights
        else (style or {}).get("character_property_id_ref")
    )
    dominant_shape = character_properties.get(dominant or "", {})
    if maximum_height is None:
        maximum_height = dominant_shape.get("height")
    return {
        "dominant_character_property_id_ref": dominant,
        "character_property_ids": sorted(weights, key=natural_key),
        "character_property_weights": dict(weights),
        "bold_ratio": round(bold_characters / total_characters, 4)
        if total_characters
        else 0.0,
        "maximum_character_height": maximum_height,
    }


def _first_line_position(paragraph: etree._Element) -> int | None:
    line_segment = next(iter(_owned_descendants(paragraph, "lineseg")), None)
    return _integer(_attribute(line_segment, "vertpos", "vertPos"))


def _paragraph_record(
    paragraph: etree._Element,
    section_name: str,
    section_number: int,
    section_paragraph: int,
    document_paragraph: int,
    paragraph_properties: dict[str, dict[str, Any]],
    character_properties: dict[str, dict[str, Any]],
    styles: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    text = _paragraph_text(paragraph)
    style_id = paragraph.get("styleIDRef")
    style = styles.get(style_id or "")
    direct_para_id = paragraph.get("paraPrIDRef")
    effective_para_id = direct_para_id or (style or {}).get("para_property_id_ref")
    para_property = paragraph_properties.get(effective_para_id or "")
    levels = _level_evidence(text, para_property)
    runs = _run_metrics(paragraph, style, character_properties)
    inside_table = _nearest_ancestor(paragraph, "tbl") is not None
    return {
        "id": f"S{section_number:02d}-P{section_paragraph:04d}",
        "section": section_name,
        "section_number": section_number,
        "section_paragraph": section_paragraph,
        "document_paragraph": document_paragraph,
        "text": text,
        "inside_table": inside_table,
        "contains_table": any(
            local_name(item.tag) == "tbl" for item in paragraph.iter()
        ),
        "contains_picture": any(
            local_name(item.tag) in {"pic", "picture"} for item in paragraph.iter()
        ),
        "page_break_after": _truthy(paragraph.get("pageBreak")),
        "page_break_before": bool((para_property or {}).get("page_break_before")),
        "first_line_vertical_position": _first_line_position(paragraph),
        "para_property_id_ref": direct_para_id,
        "effective_para_property_id_ref": effective_para_id,
        "style_id_ref": style_id,
        "style_name": (style or {}).get("name"),
        "paragraph_property": para_property,
        **runs,
        **levels,
        "page_id": None,
        "excluded_reason": None,
    }


def _assign_pages(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    pages: list[dict[str, Any]] = []
    by_section: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for record in records:
        if not record["inside_table"]:
            by_section[record["section_number"]].append(record)

    for section_number in sorted(by_section):
        page_number = 1
        pending_break = False
        previous_vertical = None
        current: dict[str, Any] | None = None
        for record in by_section[section_number]:
            vertical = record["first_line_vertical_position"]
            boundary_reason = None
            if current is None:
                boundary_reason = "section_start"
            elif pending_break:
                boundary_reason = "previous_paragraph_page_break"
            elif record["page_break_before"]:
                boundary_reason = "paragraph_page_break_before"
            elif (
                vertical is not None
                and previous_vertical is not None
                and vertical + 1000 < previous_vertical
            ):
                boundary_reason = "line_position_reset"

            if boundary_reason is not None:
                if current is not None:
                    page_number += 1
                current = {
                    "id": f"S{section_number:02d}-PG{page_number:04d}",
                    "section_number": section_number,
                    "page_number": page_number,
                    "boundary_reason": boundary_reason,
                    "paragraph_ids": [],
                }
                pages.append(current)

            if current is None:  # defensive; section_start always initializes it
                continue
            record["page_id"] = current["id"]
            current["paragraph_ids"].append(record["id"])
            pending_break = record["page_break_after"]
            if vertical is not None:
                previous_vertical = vertical
    return pages


def _score_interleaves(
    pages: list[dict[str, Any]],
    records: list[dict[str, Any]],
    minimum_score: float,
) -> None:
    by_id = {record["id"]: record for record in records}
    heights = [
        record["maximum_character_height"]
        for record in records
        if not record["inside_table"]
        and record["text"].strip()
        and record["maximum_character_height"] is not None
    ]
    median_height = float(statistics.median(heights)) if heights else None

    for index, page in enumerate(pages):
        page_records = [by_id[item] for item in page["paragraph_ids"]]
        nonempty = [record for record in page_records if record["text"].strip()]
        text = "\n".join(record["text"].strip() for record in nonempty)
        character_count = len(text)
        centered = sum(
            1
            for record in nonempty
            if ((record.get("paragraph_property") or {}).get("alignment") or "").upper()
            == "CENTER"
        )
        center_ratio = centered / len(nonempty) if nonempty else 0.0
        total_text_characters = sum(len(record["text"]) for record in nonempty)
        bold_ratio = (
            sum(len(record["text"]) * record["bold_ratio"] for record in nonempty)
            / total_text_characters
            if total_text_characters
            else 0.0
        )
        maximum_height = max(
            (
                record["maximum_character_height"]
                for record in nonempty
                if record["maximum_character_height"] is not None
            ),
            default=None,
        )
        contains_table = any(record["contains_table"] for record in page_records)
        contains_picture = any(record["contains_picture"] for record in page_records)
        candidate_count = sum(1 for record in nonempty if record["level"] is not None)
        cue_match = INTERLEAVE_CUE.search(text)
        next_boundary = (
            pages[index + 1]["boundary_reason"] if index + 1 < len(pages) else None
        )

        score = 0.0
        evidence: list[str] = []
        if nonempty and len(nonempty) <= 5 and character_count <= 160:
            score += 0.3
            evidence.append("sparse_page")
        elif nonempty and len(nonempty) <= 8 and character_count <= 260:
            score += 0.15
            evidence.append("moderately_sparse_page")
        if center_ratio >= 0.6:
            score += 0.2
            evidence.append("mostly_centered")
        elif center_ratio >= 0.35:
            score += 0.1
            evidence.append("partly_centered")
        if median_height and maximum_height and maximum_height >= median_height * 1.35:
            score += 0.2
            evidence.append("large_title_text")
        if bold_ratio >= 0.6:
            score += 0.1
            evidence.append("mostly_bold")
        if cue_match:
            score += 0.2
            evidence.append(f"divider_cue:{cue_match.group(0).strip()}")
        if candidate_count <= 1:
            score += 0.05
            evidence.append("few_hierarchy_candidates")
        if (
            page["boundary_reason"]
            in {"previous_paragraph_page_break", "paragraph_page_break_before"}
            or next_boundary == "previous_paragraph_page_break"
        ):
            score += 0.1
            evidence.append("explicit_page_boundary")
        if contains_picture and not contains_table:
            score += 0.05
            evidence.append("picture_present")
        if contains_table:
            score -= 0.45
            evidence.append("table_present_penalty")
        if character_count > 300 or len(nonempty) > 10:
            score -= 0.3
            evidence.append("dense_page_penalty")
        score = round(max(0.0, min(score, 1.0)), 4)
        page.update(
            {
                "nonempty_paragraph_count": len(nonempty),
                "character_count": character_count,
                "hierarchy_candidate_count": candidate_count,
                "center_ratio": round(center_ratio, 4),
                "bold_ratio": round(bold_ratio, 4),
                "maximum_character_height": maximum_height,
                "document_median_character_height": median_height,
                "contains_table": contains_table,
                "contains_picture": contains_picture,
                "interleave_score": score,
                "is_interleave": bool(nonempty and score >= minimum_score),
                "interleave_evidence": evidence,
                "text_preview": text[:300],
            }
        )


def _format_signature(
    record: dict[str, Any],
) -> tuple[str | None, str | None, str | None]:
    return (
        record.get("effective_para_property_id_ref"),
        record.get("style_id_ref"),
        record.get("dominant_character_property_id_ref"),
    )


def _signature_dict(
    signature: tuple[str | None, str | None, str | None],
) -> dict[str, str | None]:
    return {
        "para_property_id_ref": signature[0],
        "style_id_ref": signature[1],
        "character_property_id_ref": signature[2],
    }


def _analyze_levels(
    records: list[dict[str, Any]],
    include_tables: bool,
    exclude_interleaves: bool,
    pages: list[dict[str, Any]],
    reference_paragraphs: dict[str, str] | None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    interleave_pages = {page["id"] for page in pages if page["is_interleave"]}
    references = {
        int(level): paragraph_id
        for level, paragraph_id in (reference_paragraphs or {}).items()
    }
    by_id = {record["id"]: record for record in records}
    warnings: list[str] = []

    for record in records:
        if record["inside_table"] and not include_tables:
            record["excluded_reason"] = "table"
        elif exclude_interleaves and record.get("page_id") in interleave_pages:
            record["excluded_reason"] = "interleave"
        elif record["level_conflict"]:
            record["excluded_reason"] = "level_conflict"

    levels: list[dict[str, Any]] = []
    canonical_by_level: dict[int, dict[str, Any]] = {}
    for level in sorted(
        {record["level"] for record in records if record["level"] is not None}
    ):
        candidates = [
            record
            for record in records
            if record["level"] == level
            and record["excluded_reason"] is None
            and record["confidence"] >= 0.75
            and record["text"].strip()
        ]
        canonical = None
        alternatives: list[dict[str, Any]] = []
        reference_id = references.get(level)
        if reference_id:
            reference = by_id.get(reference_id)
            if reference is None:
                raise ValueError(
                    f"Reference paragraph was not found: level {level} -> {reference_id}"
                )
            if reference["inside_table"] and not include_tables:
                raise ValueError(
                    f"Reference paragraph is inside a table while include_tables=false: {reference_id}"
                )
            signature = _format_signature(reference)
            canonical = {
                **_signature_dict(signature),
                "source_paragraph_id": reference_id,
                "selection": "explicit_reference",
                "support_count": 1,
                "support_ratio": 1.0,
            }
        elif candidates:
            signatures = [_format_signature(record) for record in candidates]
            counts = Counter(signatures)
            signature, support = counts.most_common(1)[0]
            source = next(
                record
                for record in candidates
                if _format_signature(record) == signature
            )
            canonical = {
                **_signature_dict(signature),
                "source_paragraph_id": source["id"],
                "selection": "document_majority",
                "support_count": support,
                "support_ratio": round(support / len(candidates), 4),
            }
            for alternative, count in counts.most_common():
                alternative_source = next(
                    record
                    for record in candidates
                    if _format_signature(record) == alternative
                )
                alternatives.append(
                    {
                        **_signature_dict(alternative),
                        "count": count,
                        "ratio": round(count / len(candidates), 4),
                        "example_paragraph_id": alternative_source["id"],
                    }
                )
        if canonical:
            canonical_by_level[level] = canonical
            if (
                canonical["selection"] == "document_majority"
                and canonical["support_ratio"] < 0.6
            ):
                warnings.append(
                    f"Level {level} has weak style consensus ({canonical['support_ratio']:.2f}); automatic normalization will skip it by default."
                )
        levels.append(
            {
                "level": level,
                "candidate_count": len(candidates),
                "excluded_count": sum(
                    1
                    for record in records
                    if record["level"] == level
                    and record["excluded_reason"] is not None
                ),
                "canonical": canonical,
                "alternatives": alternatives,
            }
        )

    for record in records:
        canonical = canonical_by_level.get(record["level"])
        record["format_signature"] = _signature_dict(_format_signature(record))
        record["matches_canonical"] = bool(
            canonical
            and _format_signature(record)
            == (
                canonical["para_property_id_ref"],
                canonical["style_id_ref"],
                canonical["character_property_id_ref"],
            )
        )
        record["normalization_candidate"] = bool(
            canonical
            and record["level"] is not None
            and record["confidence"] >= 0.75
            and record["excluded_reason"] is None
            and not record["matches_canonical"]
        )

    anomalies: list[dict[str, Any]] = []
    previous = None
    for record in records:
        if (
            record["level"] is None
            or record["excluded_reason"] is not None
            or record["confidence"] < 0.75
        ):
            continue
        if previous and record["level"] > previous["level"] + 1:
            anomalies.append(
                {
                    "type": "level_jump",
                    "from_paragraph_id": previous["id"],
                    "from_level": previous["level"],
                    "to_paragraph_id": record["id"],
                    "to_level": record["level"],
                }
            )
        previous = record
    if any(record["level_conflict"] for record in records):
        warnings.append(
            "Some paragraphs have conflicting visible markers and HWPX heading levels; they are review-only and excluded from automatic normalization."
        )
    if exclude_interleaves and interleave_pages:
        warnings.append(
            "Interleave detection is heuristic; review page scores and evidence before applying large batches."
        )
    return levels, anomalies, warnings


def hwpx_analyze_paragraph_hierarchy(
    path: str,
    include_tables: bool = False,
    exclude_interleaves: bool = True,
    interleave_min_score: float = 0.65,
    reference_paragraphs: dict[str, str] | None = None,
    max_paragraphs: int = 3000,
) -> dict[str, Any]:
    """Analyze HWPX report hierarchy, paragraph/run styles, hanging-indent morphology, and interleave pages."""
    source = require_file(path, {".hwpx"})
    interleave_min_score = max(0.0, min(float(interleave_min_score), 1.0))
    max_paragraphs = max(0, min(int(max_paragraphs), 50_000))
    with zipfile.ZipFile(source, "r") as zf:
        archive = validate_archive(zf)
        if "Contents/header.xml" not in zf.namelist():
            raise ValueError("HWPX package does not contain Contents/header.xml")
        header = _parse_xml(
            read_member(zf, "Contents/header.xml"), "Contents/header.xml"
        )
        paragraph_properties, character_properties, styles = _parse_header(header)
        records: list[dict[str, Any]] = []
        document_paragraph = 0
        sections = _section_names(zf)
        for section_number, section_name in enumerate(sections, start=1):
            section = _parse_xml(read_member(zf, section_name), section_name)
            paragraphs = [
                item for item in section.iter() if local_name(item.tag) == "p"
            ]
            for section_paragraph, paragraph in enumerate(paragraphs, start=1):
                document_paragraph += 1
                records.append(
                    _paragraph_record(
                        paragraph,
                        section_name,
                        section_number,
                        section_paragraph,
                        document_paragraph,
                        paragraph_properties,
                        character_properties,
                        styles,
                    )
                )

    pages = _assign_pages(records)
    _score_interleaves(pages, records, interleave_min_score)
    levels, anomalies, warnings = _analyze_levels(
        records,
        include_tables,
        exclude_interleaves,
        pages,
        reference_paragraphs,
    )
    scoped = [
        record for record in records if include_tables or not record["inside_table"]
    ]
    displayed = scoped[:max_paragraphs]
    return {
        "path": str(Path(source)),
        "analysis_version": 1,
        "settings": {
            "include_tables": include_tables,
            "exclude_interleaves": exclude_interleaves,
            "interleave_min_score": interleave_min_score,
            "reference_paragraphs": reference_paragraphs or {},
        },
        "method": {
            "levels": "visible markers (1., 가., 1), 가)) reconciled with paraPr heading type/level",
            "paragraph_format": "p@paraPrIDRef and p@styleIDRef resolved through Contents/header.xml",
            "character_format": "text-weighted run@charPrIDRef with minority inline styles preserved by the normalizer",
            "canonical_style": "same-document majority per level unless reference_paragraphs overrides it",
            "pages": "section starts, paragraph page breaks, page-break-before, and top-level line-position resets",
            "interleaves": "sparse-page score from centering, title size, bold, divider cues, page boundaries, and table penalties",
        },
        "section_count": len(sections),
        "paragraph_count": len(records),
        "scoped_paragraph_count": len(scoped),
        "table_paragraph_count": sum(1 for record in records if record["inside_table"]),
        "hierarchy_candidate_count": sum(
            1 for record in scoped if record["level"] is not None
        ),
        "normalization_candidate_count": sum(
            1 for record in scoped if record["normalization_candidate"]
        ),
        "interleave_page_count": sum(1 for page in pages if page["is_interleave"]),
        "pages": pages,
        "levels": levels,
        "hierarchy_anomalies": anomalies,
        "paragraphs": displayed,
        "paragraphs_truncated": len(scoped) > len(displayed),
        "warnings": warnings,
        **archive,
    }


def _apply_plan_to_paragraph(
    paragraph: etree._Element, change: dict[str, Any], apply_character_style: bool
) -> bool:
    changed = False
    target = change["to"]
    source = change["from"]
    target_para = target.get("para_property_id_ref")
    target_style = target.get("style_id_ref")
    target_char = target.get("character_property_id_ref")
    source_char = source.get("character_property_id_ref")

    if target_para is not None and paragraph.get("paraPrIDRef") != target_para:
        paragraph.set("paraPrIDRef", target_para)
        changed = True
    if target_style is not None and paragraph.get("styleIDRef") != target_style:
        paragraph.set("styleIDRef", target_style)
        changed = True
    if (
        apply_character_style
        and target_char is not None
        and source_char is not None
        and target_char != source_char
    ):
        for run in _owned_descendants(paragraph, "run"):
            if not any(
                (item.text or "")
                for item in run.iter()
                if local_name(item.tag) == "t"
                and _nearest_ancestor(item, "p") is paragraph
            ):
                continue
            if run.get("charPrIDRef") == source_char:
                run.set("charPrIDRef", target_char)
                changed = True
    return changed


def hwpx_normalize_paragraph_hierarchy(
    path: str,
    output_path: str | None = None,
    include_tables: bool = False,
    exclude_interleaves: bool = True,
    interleave_min_score: float = 0.65,
    reference_paragraphs: dict[str, str] | None = None,
    min_level_confidence: float = 0.75,
    min_style_consensus: float = 0.6,
    apply_character_style: bool = True,
    dry_run: bool = True,
    max_changes: int = 0,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Plan or apply same-document paragraph hierarchy styles to a separate HWPX copy."""
    source = require_file(path, {".hwpx"})
    min_level_confidence = max(0.0, min(float(min_level_confidence), 1.0))
    min_style_consensus = max(0.0, min(float(min_style_consensus), 1.0))
    analysis = hwpx_analyze_paragraph_hierarchy(
        str(source),
        include_tables=include_tables,
        exclude_interleaves=exclude_interleaves,
        interleave_min_score=interleave_min_score,
        reference_paragraphs=reference_paragraphs,
        max_paragraphs=50_000,
    )
    canonical_by_level = {
        item["level"]: item["canonical"]
        for item in analysis["levels"]
        if item["canonical"] is not None
    }
    plans: list[dict[str, Any]] = []
    skipped_low_consensus: set[int] = set()
    for record in analysis["paragraphs"]:
        level = record["level"]
        canonical = canonical_by_level.get(level)
        if (
            canonical is None
            or record["excluded_reason"] is not None
            or record["matches_canonical"]
        ):
            continue
        if record["confidence"] < min_level_confidence:
            continue
        if (
            canonical["selection"] != "explicit_reference"
            and canonical["support_ratio"] < min_style_consensus
        ):
            skipped_low_consensus.add(level)
            continue
        plans.append(
            {
                "paragraph_id": record["id"],
                "section": record["section"],
                "section_paragraph": record["section_paragraph"],
                "level": level,
                "confidence": record["confidence"],
                "text_preview": record["text"][:240],
                "from": record["format_signature"],
                "to": {
                    "para_property_id_ref": canonical["para_property_id_ref"],
                    "style_id_ref": canonical["style_id_ref"],
                    "character_property_id_ref": canonical["character_property_id_ref"],
                },
                "canonical_source_paragraph_id": canonical["source_paragraph_id"],
                "canonical_support_ratio": canonical["support_ratio"],
            }
        )
    if max_changes > 0:
        plans = plans[: max(0, int(max_changes))]

    warnings = list(analysis["warnings"])
    if skipped_low_consensus:
        warnings.append(
            "Skipped levels below min_style_consensus: "
            + ", ".join(str(level) for level in sorted(skipped_low_consensus))
        )
    if dry_run or not plans:
        return {
            "source": str(source),
            "output": None,
            "dry_run": dry_run,
            "planned_change_count": len(plans),
            "applied_change_count": 0,
            "changes": plans,
            "settings": analysis["settings"]
            | {
                "min_level_confidence": min_level_confidence,
                "min_style_consensus": min_style_consensus,
                "apply_character_style": apply_character_style,
                "max_changes": max_changes,
            },
            "interleave_pages": [
                page for page in analysis["pages"] if page["is_interleave"]
            ],
            "warnings": warnings,
            "message": "Dry run only; no output file was created."
            if dry_run
            else "No eligible paragraph style changes were found; no output file was created.",
        }

    output = output_copy_path(source, output_path, "문단정규화", overwrite)
    changes_by_section: dict[str, dict[int, dict[str, Any]]] = defaultdict(dict)
    for change in plans:
        changes_by_section[change["section"]][change["section_paragraph"]] = change
    updates: dict[str, bytes] = {}
    applied: list[dict[str, Any]] = []
    with zipfile.ZipFile(source, "r") as zf:
        validate_archive(zf)
        for section_name, section_changes in changes_by_section.items():
            root = _parse_xml(read_member(zf, section_name), section_name)
            paragraphs = [item for item in root.iter() if local_name(item.tag) == "p"]
            section_changed = False
            for section_paragraph, change in section_changes.items():
                if section_paragraph < 1 or section_paragraph > len(paragraphs):
                    raise IndexError(
                        f"Paragraph locator is out of range after re-open: {section_name} #{section_paragraph}"
                    )
                if _apply_plan_to_paragraph(
                    paragraphs[section_paragraph - 1], change, apply_character_style
                ):
                    applied.append(change)
                    section_changed = True
            if section_changed:
                updates[section_name] = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=None
                )
    if not applied:
        return {
            "source": str(source),
            "output": None,
            "dry_run": False,
            "planned_change_count": len(plans),
            "applied_change_count": 0,
            "changes": plans,
            "warnings": warnings,
            "message": "The plan did not produce any XML attribute changes; no output file was created.",
        }

    rewrite_zip(source, output, updates, overwrite)
    with zipfile.ZipFile(output, "r") as zf:
        output_archive = validate_archive(zf)
        for name in ["Contents/header.xml", *_section_names(zf)]:
            _parse_xml(read_member(zf, name), name)
    return {
        "source": str(source),
        "output": str(output),
        "dry_run": False,
        "planned_change_count": len(plans),
        "applied_change_count": len(applied),
        "changed_sections": sorted(updates, key=natural_key),
        "changes": applied,
        "output_bytes": output.stat().st_size,
        "valid_output": output_archive["bad_member"] is None,
        "interleave_pages": [
            page for page in analysis["pages"] if page["is_interleave"]
        ],
        "warnings": warnings,
        **output_archive,
    }
