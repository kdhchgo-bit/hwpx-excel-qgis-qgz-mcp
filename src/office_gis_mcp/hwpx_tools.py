from __future__ import annotations

import csv
import locale
import os
import re
import subprocess
import time
import zipfile
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
from .hwpx_paragraphs import (
    hwpx_analyze_paragraph_hierarchy,
    hwpx_normalize_paragraph_hierarchy,
)
from .hwpx_tables import hwpx_analyze_tables

XML_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, recover=False, huge_tree=False)


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


def _paragraphs(root: etree._Element) -> list[etree._Element]:
    return [element for element in root.iter() if local_name(element.tag) == "p"]


def _paragraph_text(paragraph: etree._Element) -> str:
    return "".join((element.text or "") for element in paragraph.iter() if local_name(element.tag) == "t")


def hwpx_health() -> dict[str, Any]:
    """Return the HWPX server capability summary."""
    hwp_exe = _find_hwp_exe()
    return {
        "server": "hwpx-local",
        "mode": "HWPX ZIP/XML package processing",
        "hwp_exe": str(hwp_exe) if hwp_exe else None,
        "hwp_native_available": bool(os.name == "nt" and hwp_exe),
        "source_overwrite": False,
        "features": [
            "validate",
            "inspect",
            "analyze_table_morphology",
            "analyze_paragraph_hierarchy",
            "normalize_paragraph_hierarchy",
            "extract_text",
            "find_text",
            "replace_text_across_runs",
            "native_open_check",
            "native_export_pdf",
        ],
    }


def _find_hwp_exe() -> Path | None:
    configured = os.environ.get("HWP_EXE_PATH")
    candidates = [
        configured,
        r"C:\Program Files (x86)\Hnc\Office 2022\HOffice120\Bin\Hwp.exe",
        r"C:\Program Files (x86)\Hnc\Office 2020\HOffice110\Bin\Hwp.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return Path(candidate).resolve()
    return None


def _hwp_process_ids() -> set[int]:
    if os.name != "nt":
        return set()
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq Hwp.exe", "/FO", "CSV", "/NH"],
        text=True,
        encoding=locale.getpreferredencoding(False),
        errors="replace",
        capture_output=True,
        check=False,
    )
    process_ids: set[int] = set()
    for row in csv.reader(result.stdout.splitlines()):
        if len(row) >= 2 and row[0].lower() == "hwp.exe":
            try:
                process_ids.add(int(row[1]))
            except ValueError:
                pass
    return process_ids


def _open_hwp_session():
    if os.name != "nt":
        raise RuntimeError("Hancom HWP COM is only available on Windows")
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    existing_process_ids = _hwp_process_ids()
    hwp = None
    launched = None
    first_error = None
    try:
        hwp = win32com.client.DispatchEx("HWPFrame.HwpObject")
    except Exception as exc:  # pragma: no cover - installation state dependent
        first_error = exc
        executable = _find_hwp_exe()
        if executable is None:
            pythoncom.CoUninitialize()
            raise RuntimeError(f"HWP COM failed and Hwp.exe was not found: {exc}") from exc
        startup = subprocess.STARTUPINFO()
        startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startup.wShowWindow = subprocess.SW_HIDE
        launched = subprocess.Popen([str(executable)], startupinfo=startup)
        time.sleep(2)
        try:
            hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        except Exception:
            if launched.poll() is None:
                launched.terminate()
            pythoncom.CoUninitialize()
            raise RuntimeError(f"Unable to create HWP COM session: {first_error}") from first_error
    try:
        hwp.XHwpWindows.Item(0).Visible = False
    except Exception:
        pass
    try:
        registered = bool(hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule"))
        if not registered:
            raise RuntimeError("Hancom file-path security module registration failed")
    except Exception:
        pass
    owned_process_ids = _hwp_process_ids() - existing_process_ids
    if launched is not None:
        owned_process_ids.add(launched.pid)
    return pythoncom, hwp, launched, owned_process_ids


def _close_hwp_session(
    pythoncom: Any,
    hwp: Any,
    launched: subprocess.Popen[Any] | None,
    owned_process_ids: set[int],
) -> None:
    try:
        hwp.Clear(1)
    except Exception:
        pass
    try:
        hwp.Quit()
    except Exception:
        pass
    if launched is not None:
        try:
            launched.wait(timeout=5)
        except subprocess.TimeoutExpired:
            launched.terminate()
    pythoncom.CoUninitialize()
    time.sleep(1)
    remaining = _hwp_process_ids()
    for process_id in sorted(owned_process_ids & remaining):
        subprocess.run(
            ["taskkill", "/PID", str(process_id), "/T", "/F"],
            text=True,
            capture_output=True,
            check=False,
        )


def hwpx_native_open_check(path: str) -> dict[str, Any]:
    """Open an HWPX in installed Hancom Hangul, read plain text, and close it without saving."""
    source = require_file(path, {".hwpx"})
    pythoncom, hwp, launched, owned_process_ids = _open_hwp_session()
    opened = False
    text = ""
    try:
        opened = bool(hwp.Open(str(source), "", ""))
        if not opened:
            raise RuntimeError(f"Hancom Hangul did not open the HWPX file: {source}")
        text = hwp.GetTextFile("TEXT", "") or ""
    finally:
        _close_hwp_session(pythoncom, hwp, launched, owned_process_ids)
    return {
        "path": str(source),
        "opened": opened,
        "text_character_count": len(text),
        "text_preview": text[:1000],
        "hwp_exe": str(_find_hwp_exe()) if _find_hwp_exe() else None,
    }


def hwpx_export_pdf(path: str, output_pdf: str, overwrite: bool = False) -> dict[str, Any]:
    """Open an HWPX in installed Hancom Hangul and export it to PDF in a hidden COM session."""
    source = require_file(path, {".hwpx"})
    output = Path(output_pdf).expanduser().resolve()
    if output.suffix.lower() != ".pdf":
        raise ValueError("output_pdf must end with .pdf")
    if output.exists() and not overwrite:
        raise FileExistsError(f"Output already exists: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()
    pythoncom, hwp, launched, owned_process_ids = _open_hwp_session()
    opened = False
    saved = False
    try:
        opened = bool(hwp.Open(str(source), "", ""))
        if not opened:
            raise RuntimeError(f"Hancom Hangul did not open the HWPX file: {source}")
        saved = bool(hwp.SaveAs(str(output), "PDF", ""))
    finally:
        _close_hwp_session(pythoncom, hwp, launched, owned_process_ids)
    if not saved or not output.is_file() or output.stat().st_size == 0:
        raise RuntimeError(f"Hancom Hangul did not create the requested PDF: {output}")
    return {
        "source": str(source),
        "output_pdf": str(output),
        "opened": opened,
        "saved": saved,
        "bytes": output.stat().st_size,
    }


def hwpx_validate(path: str) -> dict[str, Any]:
    """Validate HWPX ZIP integrity and parse its core XML files without modifying it."""
    source = require_file(path, {".hwpx"})
    with zipfile.ZipFile(source, "r") as zf:
        archive = validate_archive(zf)
        names = set(zf.namelist())
        sections = _section_names(zf)
        required = ["Contents/content.hpf", "Contents/header.xml"]
        missing = [name for name in required if name not in names]
        if not sections:
            missing.append("Contents/section*.xml")
        parsed: list[str] = []
        for name in [item for item in required if item in names] + sections:
            _parse_xml(read_member(zf, name), name)
            parsed.append(name)
        mimetype = zf.read("mimetype").decode("ascii", errors="replace") if "mimetype" in names else None
        return {
            "path": str(source),
            "valid": not missing and archive["bad_member"] is None,
            "missing": missing,
            "mimetype": mimetype,
            "section_count": len(sections),
            "parsed_xml": parsed,
            **archive,
        }


def hwpx_inspect(path: str, paragraph_preview_limit: int = 20) -> dict[str, Any]:
    """Inspect an HWPX document's sections, text, tables, images, and preview paragraphs."""
    source = require_file(path, {".hwpx"})
    paragraph_preview_limit = max(0, min(paragraph_preview_limit, 200))
    previews: list[dict[str, Any]] = []
    paragraph_count = 0
    table_count = 0
    character_count = 0
    with zipfile.ZipFile(source, "r") as zf:
        archive = validate_archive(zf)
        sections = _section_names(zf)
        for section_name in sections:
            root = _parse_xml(read_member(zf, section_name), section_name)
            table_count += sum(1 for item in root.iter() if local_name(item.tag) == "tbl")
            for local_index, paragraph in enumerate(_paragraphs(root), start=1):
                paragraph_count += 1
                text = _paragraph_text(paragraph)
                character_count += len(text)
                if text and len(previews) < paragraph_preview_limit:
                    previews.append(
                        {
                            "paragraph": paragraph_count,
                            "section": section_name,
                            "section_paragraph": local_index,
                            "text": text[:500],
                        }
                    )
        images = [
            info.filename
            for info in zf.infolist()
            if info.filename.lower().startswith("bindata/")
            and not info.is_dir()
        ]
        return {
            "path": str(source),
            "bytes": source.stat().st_size,
            "sections": sections,
            "section_count": len(sections),
            "paragraph_count": paragraph_count,
            "character_count": character_count,
            "table_count": table_count,
            "image_count": len(images),
            "image_entries": images[:200],
            "paragraph_preview": previews,
            **archive,
        }


def hwpx_extract_text(path: str, output_path: str | None = None, overwrite: bool = False) -> dict[str, Any]:
    """Extract HWPX paragraph text; optionally save it as a UTF-8 text file."""
    source = require_file(path, {".hwpx"})
    lines: list[str] = []
    with zipfile.ZipFile(source, "r") as zf:
        validate_archive(zf)
        for section_name in _section_names(zf):
            root = _parse_xml(read_member(zf, section_name), section_name)
            lines.extend(_paragraph_text(paragraph) for paragraph in _paragraphs(root))
    text = "\n".join(lines)
    saved_to = None
    if output_path:
        output = Path(output_path).expanduser().resolve()
        if output.exists() and not overwrite:
            raise FileExistsError(f"Output already exists: {output}")
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(text, encoding="utf-8")
        saved_to = str(output)
    return {
        "path": str(source),
        "paragraph_count": len(lines),
        "character_count": len(text),
        "saved_to": saved_to,
        "text": text if output_path is None else text[:4000],
        "text_truncated": bool(output_path and len(text) > 4000),
    }


def hwpx_find_text(path: str, query: str, regex: bool = False, max_results: int = 100) -> dict[str, Any]:
    """Find text in HWPX paragraphs and return paragraph-level snippets."""
    if not query:
        raise ValueError("query must not be empty")
    source = require_file(path, {".hwpx"})
    max_results = max(1, min(max_results, 1000))
    pattern = re.compile(query) if regex else None
    results: list[dict[str, Any]] = []
    total = 0
    paragraph_number = 0
    with zipfile.ZipFile(source, "r") as zf:
        validate_archive(zf)
        for section_name in _section_names(zf):
            root = _parse_xml(read_member(zf, section_name), section_name)
            for section_paragraph, paragraph in enumerate(_paragraphs(root), start=1):
                paragraph_number += 1
                text = _paragraph_text(paragraph)
                matches = list(pattern.finditer(text)) if pattern else []
                count = len(matches) if pattern else text.count(query)
                if count:
                    total += count
                    if len(results) < max_results:
                        first = matches[0].start() if matches else text.find(query)
                        start = max(0, first - 80)
                        end = min(len(text), first + max(len(query), 1) + 160)
                        results.append(
                            {
                                "paragraph": paragraph_number,
                                "section": section_name,
                                "section_paragraph": section_paragraph,
                                "count": count,
                                "snippet": text[start:end],
                            }
                        )
    return {"path": str(source), "query": query, "match_count": total, "results": results, "truncated": total > len(results)}


def _replace_across_text_nodes(paragraph: etree._Element, old: str, new: str, remaining: int | None) -> int:
    nodes = [element for element in paragraph.iter() if local_name(element.tag) == "t"]
    values = [element.text or "" for element in nodes]
    full = "".join(values)
    starts: list[int] = []
    offset = 0
    while True:
        found = full.find(old, offset)
        if found < 0:
            break
        starts.append(found)
        offset = found + len(old)
        if remaining is not None and len(starts) >= remaining:
            break
    if not starts:
        return 0
    boundaries: list[tuple[int, int]] = []
    cursor = 0
    for value in values:
        boundaries.append((cursor, cursor + len(value)))
        cursor += len(value)
    for start in reversed(starts):
        end = start + len(old)
        start_index = next(index for index, (_, node_end) in enumerate(boundaries) if start < node_end)
        end_index = next(index for index, (_, node_end) in enumerate(boundaries) if end <= node_end)
        start_base, _ = boundaries[start_index]
        end_base, _ = boundaries[end_index]
        start_local = start - start_base
        end_local = end - end_base
        if start_index == end_index:
            values[start_index] = values[start_index][:start_local] + new + values[start_index][end_local:]
        else:
            values[start_index] = values[start_index][:start_local] + new
            for index in range(start_index + 1, end_index):
                values[index] = ""
            values[end_index] = values[end_index][end_local:]
    for node, value in zip(nodes, values):
        node.text = value
    return len(starts)


def hwpx_replace_text(
    path: str,
    old: str,
    new: str,
    output_path: str | None = None,
    max_replacements: int = 0,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Replace literal HWPX text across run boundaries and create a separate output copy."""
    if not old:
        raise ValueError("old must not be empty")
    source = require_file(path, {".hwpx"})
    output = output_copy_path(source, output_path, "MCP수정", overwrite)
    cap = None if max_replacements <= 0 else max_replacements
    updates: dict[str, bytes] = {}
    replaced = 0
    changed_sections: list[str] = []
    with zipfile.ZipFile(source, "r") as zf:
        validate_archive(zf)
        for section_name in _section_names(zf):
            root = _parse_xml(read_member(zf, section_name), section_name)
            section_replaced = 0
            for paragraph in _paragraphs(root):
                remaining = None if cap is None else cap - replaced
                if remaining is not None and remaining <= 0:
                    break
                count = _replace_across_text_nodes(paragraph, old, new, remaining)
                replaced += count
                section_replaced += count
            if section_replaced:
                updates[section_name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=None)
                changed_sections.append(section_name)
        preview_name = "Preview/PrvText.txt"
        if replaced and preview_name in zf.namelist():
            preview = read_member(zf, preview_name).decode("utf-8", errors="replace")
            preview_cap = replaced if cap is not None else -1
            updates[preview_name] = preview.replace(old, new, preview_cap).encode("utf-8")
    if not replaced:
        return {
            "source": str(source),
            "output": None,
            "replacement_count": 0,
            "changed_sections": [],
            "message": "No matching text was found; no output file was created.",
        }
    rewrite_zip(source, output, updates, overwrite)
    validation = hwpx_validate(str(output))
    return {
        "source": str(source),
        "output": str(output),
        "replacement_count": replaced,
        "changed_sections": changed_sections,
        "valid_output": validation["valid"],
        "output_bytes": output.stat().st_size,
    }


def register_hwpx_tools(mcp: Any) -> None:
    for function in (
        hwpx_health,
        hwpx_validate,
        hwpx_inspect,
        hwpx_analyze_tables,
        hwpx_analyze_paragraph_hierarchy,
        hwpx_normalize_paragraph_hierarchy,
        hwpx_extract_text,
        hwpx_find_text,
        hwpx_replace_text,
        hwpx_native_open_check,
        hwpx_export_pdf,
    ):
        mcp.tool()(function)
