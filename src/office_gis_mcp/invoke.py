from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Callable
from typing import Any

from .excel_tools import (
    excel_export_pdf,
    excel_find,
    excel_health,
    excel_inspect,
    excel_read_range,
    excel_recalculate,
    excel_write_range,
)
from .hwpx_tools import (
    hwpx_analyze_tables,
    hwpx_export_pdf,
    hwpx_extract_text,
    hwpx_find_text,
    hwpx_health,
    hwpx_inspect,
    hwpx_native_open_check,
    hwpx_replace_text,
    hwpx_validate,
)
from .qgis_tools import (
    qgis_algorithm_help,
    qgis_health,
    qgis_list_algorithms,
    qgis_run_algorithm,
    qgz_audit_sources,
    qgz_inspect,
    qgz_rebase_paths,
)

Tool = Callable[..., Any]

TOOLS: dict[str, Tool] = {
    function.__name__: function
    for function in (
        hwpx_health,
        hwpx_validate,
        hwpx_inspect,
        hwpx_analyze_tables,
        hwpx_extract_text,
        hwpx_find_text,
        hwpx_replace_text,
        hwpx_native_open_check,
        hwpx_export_pdf,
        excel_health,
        excel_inspect,
        excel_read_range,
        excel_find,
        excel_write_range,
        excel_recalculate,
        excel_export_pdf,
        qgis_health,
        qgz_inspect,
        qgz_audit_sources,
        qgz_rebase_paths,
        qgis_list_algorithms,
        qgis_algorithm_help,
        qgis_run_algorithm,
    )
}


def invoke_tool(tool_name: str, arguments: dict[str, Any] | None = None) -> dict[str, Any]:
    """Invoke one local office/GIS tool and return a JSON-serializable envelope."""
    if tool_name not in TOOLS:
        raise KeyError(f"Unknown tool: {tool_name}")
    if arguments is None:
        arguments = {}
    if not isinstance(arguments, dict):
        raise TypeError("Tool arguments must be a JSON object")
    return {"ok": True, "tool": tool_name, "result": TOOLS[tool_name](**arguments)}


def _load_arguments(raw: str) -> dict[str, Any]:
    raw = raw.lstrip("\ufeff")
    if not raw.strip():
        return {}
    value = json.loads(raw)
    if not isinstance(value, dict):
        raise TypeError("Tool arguments must be a JSON object")
    return value


def cli(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Invoke one local HWPX, Excel, or QGIS tool with JSON arguments")
    parser.add_argument("tool", choices=sorted(TOOLS))
    parser.add_argument("--json", dest="json_arguments", help="JSON object. If omitted, read UTF-8 JSON from stdin.")
    parser.add_argument("--pretty", action="store_true", help="Indent JSON output for humans")
    args = parser.parse_args(argv)

    try:
        raw = args.json_arguments if args.json_arguments is not None else sys.stdin.read()
        response = invoke_tool(args.tool, _load_arguments(raw))
    except Exception as exc:
        response = {
            "ok": False,
            "tool": args.tool,
            "error_type": type(exc).__name__,
            "error": str(exc),
        }
        print(json.dumps(response, ensure_ascii=False, default=str), file=sys.stderr)
        raise SystemExit(1) from exc

    print(json.dumps(response, ensure_ascii=False, indent=2 if args.pretty else None, default=str))


if __name__ == "__main__":
    cli()
