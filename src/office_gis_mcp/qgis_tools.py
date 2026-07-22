from __future__ import annotations

import json
import os
import re
import subprocess
import urllib.parse
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree

from .common import local_name, output_copy_path, read_member, require_file, rewrite_zip, validate_archive

XML_PARSER = etree.XMLParser(resolve_entities=False, no_network=True, recover=False, huge_tree=False)


def _find_qgis_process() -> Path | None:
    configured = os.environ.get("QGIS_PROCESS_PATH")
    candidates = [
        configured,
        r"D:\bin\qgis_process-qgis-ltr.bat",
        r"D:\bin\qgis_process-qgis.bat",
        r"C:\OSGeo4W\bin\qgis_process-qgis-ltr.bat",
        r"C:\OSGeo4W\bin\qgis_process-qgis.bat",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return Path(candidate).resolve()
    return None


def _run_process(args: list[str], input_text: str | None = None, timeout: int = 300) -> subprocess.CompletedProcess[str]:
    executable = _find_qgis_process()
    if executable is None:
        raise FileNotFoundError("qgis_process was not found. Set QGIS_PROCESS_PATH to its .bat or .exe path.")
    if executable.suffix.lower() in {".bat", ".cmd"}:
        command_line = subprocess.list2cmdline([str(executable), *args])
        command = [os.environ.get("COMSPEC", "cmd.exe"), "/d", "/s", "/c", command_line]
    else:
        command = [str(executable), *args]
    result = subprocess.run(
        command,
        input=input_text,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        timeout=max(1, min(timeout, 3600)),
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"qgis_process failed with exit code {result.returncode}\nSTDOUT:\n{result.stdout[-4000:]}\nSTDERR:\n{result.stderr[-4000:]}"
        )
    return result


def _json_output(result: subprocess.CompletedProcess[str]) -> Any:
    text = result.stdout.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start_candidates = [index for index in (text.find("{"), text.find("[")) if index >= 0]
        if start_candidates:
            return json.loads(text[min(start_candidates) :])
        raise ValueError(f"Expected JSON from qgis_process, got: {text[-2000:]}")


def _read_project(path: Path) -> tuple[bytes, str | None, dict[str, Any]]:
    if path.suffix.lower() == ".qgs":
        payload = path.read_bytes()
        return payload, None, {"entry_count": None, "uncompressed_bytes": None, "bad_member": None}
    with zipfile.ZipFile(path, "r") as zf:
        archive = validate_archive(zf)
        qgs_names = [name for name in zf.namelist() if name.lower().endswith(".qgs")]
        if len(qgs_names) != 1:
            raise ValueError(f"Expected exactly one .qgs entry in QGZ, found {len(qgs_names)}")
        return read_member(zf, qgs_names[0]), qgs_names[0], archive


def _parse_project(payload: bytes, name: str) -> etree._Element:
    try:
        return etree.fromstring(payload, parser=XML_PARSER)
    except etree.XMLSyntaxError as exc:
        raise ValueError(f"Invalid QGIS project XML in {name}: {exc}") from exc


def _first_text(root: etree._Element, paths: list[str]) -> str | None:
    for path in paths:
        item = root.find(path)
        if item is not None and item.text:
            return item.text
    return None


def _layer_info(layer: etree._Element) -> dict[str, Any]:
    def text(name: str) -> str | None:
        child = next((item for item in layer if local_name(item.tag) == name), None)
        return child.text if child is not None else None

    crs = None
    for child in layer.iter():
        if local_name(child.tag) == "authid" and child.text:
            crs = child.text
            break
    return {
        "id": text("id"),
        "name": text("layername"),
        "type": layer.get("type"),
        "geometry": layer.get("geometry"),
        "provider": text("provider"),
        "datasource": text("datasource"),
        "crs": crs,
    }


def _extract_local_path(datasource: str, project_dir: Path) -> tuple[str | None, str | None]:
    raw = datasource.strip()
    dbname = re.search(r"\bdbname\s*=\s*['\"]([^'\"]+)['\"]", raw, flags=re.IGNORECASE)
    candidate = dbname.group(1) if dbname else raw.split("|", 1)[0]
    candidate = candidate.strip().strip("'\"")
    if candidate.startswith("/vsizip/"):
        candidate = candidate[len("/vsizip/") :]
    parsed = urllib.parse.urlparse(candidate)
    if parsed.scheme.lower() in {"http", "https", "postgres", "wms", "wfs", "xyz"}:
        return None, "remote_or_service"
    if parsed.scheme.lower() == "file":
        candidate = urllib.parse.unquote(parsed.path)
        if re.match(r"^/[A-Za-z]:/", candidate):
            candidate = candidate[1:]
    if candidate.lower().startswith(("memory:", "virtual:", "ogr:", "pg:")):
        return None, "memory_or_provider"
    if not candidate or ("=" in candidate and not dbname):
        return None, "provider_expression"
    candidate = candidate.replace("/", os.sep)
    local = Path(candidate)
    if not local.is_absolute():
        local = project_dir / local
    return str(local.resolve(strict=False)), None


def qgis_health() -> dict[str, Any]:
    """Return QGIS/QGZ server capabilities and detected qgis_process version."""
    executable = _find_qgis_process()
    version = None
    error = None
    if executable:
        try:
            result = _run_process(["--version"], timeout=60)
            version = result.stdout.strip().splitlines()[0] if result.stdout.strip() else None
        except Exception as exc:  # pragma: no cover - installation dependent
            error = str(exc)
    return {
        "server": "qgis-local",
        "qgis_process": str(executable) if executable else None,
        "version": version,
        "error": error,
        "source_overwrite": False,
        "features": ["qgz_inspect", "qgz_audit_sources", "qgz_rebase_paths", "algorithm_list", "algorithm_help", "algorithm_run"],
    }


def qgz_inspect(path: str, include_layers: bool = True, max_layers: int = 1000) -> dict[str, Any]:
    """Inspect a QGS/QGZ project's version, CRS, canvas extent, layers, layouts, and archive."""
    source = require_file(path, {".qgs", ".qgz"})
    payload, qgs_entry, archive = _read_project(source)
    root = _parse_project(payload, qgs_entry or source.name)
    layers = [_layer_info(item) for item in root.iter() if local_name(item.tag) == "maplayer"]
    layout_names: list[str] = []
    for item in root.iter():
        if local_name(item.tag).lower() == "layout":
            name = item.get("name")
            if name and name not in layout_names:
                layout_names.append(name)
    extent_element = root.find("./mapcanvas/extent")
    extent = None
    if extent_element is not None:
        extent = {name: _first_text(extent_element, [name]) for name in ("xmin", "ymin", "xmax", "ymax")}
    project_crs = _first_text(
        root,
        ["./projectCrs/spatialrefsys/authid", "./mapcanvas/destinationsrs/spatialrefsys/authid"],
    )
    title = _first_text(root, ["./title"])
    return {
        "path": str(source),
        "bytes": source.stat().st_size,
        "format": source.suffix.lower().lstrip("."),
        "qgs_entry": qgs_entry,
        "qgis_version": root.get("version"),
        "project_name": root.get("projectname") or title,
        "project_crs": project_crs,
        "canvas_extent": extent,
        "layer_count": len(layers),
        "layers": layers[: max(0, min(max_layers, 5000))] if include_layers else None,
        "layers_truncated": include_layers and len(layers) > max_layers,
        "layout_count": len(layout_names),
        "layouts": layout_names,
        **archive,
    }


def qgz_audit_sources(path: str) -> dict[str, Any]:
    """Audit QGS/QGZ layer data sources and report missing local files."""
    source = require_file(path, {".qgs", ".qgz"})
    payload, qgs_entry, _ = _read_project(source)
    root = _parse_project(payload, qgs_entry or source.name)
    results: list[dict[str, Any]] = []
    for item in root.iter():
        if local_name(item.tag) != "maplayer":
            continue
        layer = _layer_info(item)
        datasource = layer.get("datasource") or ""
        local_path, skip_reason = _extract_local_path(datasource, source.parent)
        exists = Path(local_path).exists() if local_path else None
        results.append({**layer, "local_path": local_path, "exists": exists, "skip_reason": skip_reason})
    missing = [item for item in results if item["exists"] is False]
    return {
        "path": str(source),
        "layer_count": len(results),
        "local_source_count": sum(1 for item in results if item["local_path"]),
        "missing_count": len(missing),
        "all_local_sources_exist": not missing,
        "sources": results,
    }


def _replace_path_variants(text: str, old_root: str, new_root: str) -> tuple[str, int]:
    old_raw = str(Path(old_root).expanduser())
    new_raw = str(Path(new_root).expanduser())
    pairs = {
        old_raw: new_raw,
        old_raw.replace("\\", "/"): new_raw.replace("\\", "/"),
        old_raw.replace("/", "\\"): new_raw.replace("/", "\\"),
    }
    count = 0
    result = text
    for old, new in sorted(pairs.items(), key=lambda item: len(item[0]), reverse=True):
        if not old:
            continue
        result, hits = re.subn(re.escape(old), lambda _: new, result, flags=re.IGNORECASE)
        count += hits
    return result, count


def qgz_rebase_paths(
    path: str,
    old_root: str,
    new_root: str,
    output_path: str | None = None,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Replace an old root path with a new root in QGS/QGZ XML and create a separate copy."""
    if not old_root or not new_root:
        raise ValueError("old_root and new_root must not be empty")
    source = require_file(path, {".qgs", ".qgz"})
    output = output_copy_path(source, output_path, "경로수정", overwrite)
    payload, qgs_entry, _ = _read_project(source)
    text = payload.decode("utf-8-sig")
    replaced_text, count = _replace_path_variants(text, old_root, new_root)
    if count == 0:
        return {"source": str(source), "output": None, "replacement_count": 0, "message": "Old root was not found; no output created."}
    updated_payload = replaced_text.encode("utf-8")
    _parse_project(updated_payload, qgs_entry or source.name)
    if source.suffix.lower() == ".qgz":
        assert qgs_entry is not None
        rewrite_zip(source, output, {qgs_entry: updated_payload}, overwrite)
    else:
        if output.exists() and not overwrite:
            raise FileExistsError(f"Output already exists: {output}")
        output.write_bytes(updated_payload)
    return {
        "source": str(source),
        "output": str(output),
        "replacement_count": count,
        "output_bytes": output.stat().st_size,
        "inspection": qgz_inspect(str(output), include_layers=False),
    }


def qgis_list_algorithms(search: str | None = None, max_results: int = 200) -> dict[str, Any]:
    """List installed QGIS Processing algorithms, optionally filtered by text."""
    result = _run_process(["--json", "--skip-loading-plugins", "list"], timeout=180)
    data = _json_output(result)
    if isinstance(data, dict):
        raw_algorithms = data.get("algorithms", data)
    else:
        raw_algorithms = data
    flattened: list[dict[str, Any]] = []
    if isinstance(raw_algorithms, dict):
        for key, value in raw_algorithms.items():
            if isinstance(value, dict):
                flattened.append({"id": value.get("id", key), **value})
            else:
                flattened.append({"id": key, "name": str(value)})
    elif isinstance(raw_algorithms, list):
        for item in raw_algorithms:
            flattened.append(item if isinstance(item, dict) else {"name": str(item)})
    if search:
        needle = search.casefold()
        flattened = [item for item in flattened if needle in json.dumps(item, ensure_ascii=False).casefold()]
    max_results = max(1, min(max_results, 2000))
    return {"search": search, "match_count": len(flattened), "algorithms": flattened[:max_results], "truncated": len(flattened) > max_results}


def qgis_algorithm_help(algorithm_id: str) -> Any:
    """Return QGIS Processing help and parameter schema for an algorithm ID."""
    if not algorithm_id:
        raise ValueError("algorithm_id must not be empty")
    result = _run_process(["--json", "--skip-loading-plugins", "help", algorithm_id], timeout=180)
    return _json_output(result)


def qgis_run_algorithm(
    algorithm_id: str,
    parameters: dict[str, Any],
    project_path: str | None = None,
    ellipsoid: str | None = None,
    load_plugins: bool = False,
    timeout_seconds: int = 600,
) -> Any:
    """Run a QGIS Processing algorithm through JSON stdin and return its structured result."""
    if not algorithm_id:
        raise ValueError("algorithm_id must not be empty")
    payload: dict[str, Any] = {"inputs": parameters}
    if project_path:
        payload["project_path"] = str(require_file(project_path, {".qgs", ".qgz"}))
    if ellipsoid:
        payload["ellipsoid"] = ellipsoid
    args = ["--json"]
    if not load_plugins:
        args.append("--skip-loading-plugins")
    args.extend(["run", algorithm_id, "-"])
    result = _run_process(args, input_text=json.dumps(payload, ensure_ascii=False), timeout=timeout_seconds)
    return _json_output(result)


def register_qgis_tools(mcp: Any) -> None:
    for function in (
        qgis_health,
        qgz_inspect,
        qgz_audit_sources,
        qgz_rebase_paths,
        qgis_list_algorithms,
        qgis_algorithm_help,
        qgis_run_algorithm,
    ):
        mcp.tool()(function)
