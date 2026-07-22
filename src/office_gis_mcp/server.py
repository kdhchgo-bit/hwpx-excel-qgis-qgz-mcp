from __future__ import annotations

import argparse
import sys

from mcp.server.fastmcp import FastMCP

from .excel_tools import register_excel_tools
from .hwpx_tools import register_hwpx_tools
from .qgis_tools import register_qgis_tools

SERVER_INSTRUCTIONS = {
    "hwpx": (
        "Local HWPX package tools. Prefer read-only inspect/find first. Mutating tools create a separate copy and never overwrite the source."
    ),
    "excel": (
        "Local Excel tools. Inspect the live workbook and exact sheet/range first. Mutating tools create a separate copy; use COM for maximum Excel fidelity."
    ),
    "qgis": (
        "Local QGIS/QGZ tools. Audit project layers and sources before edits. Rebase creates a copy. qgis_run_algorithm may create outputs requested by the caller."
    ),
}


def build_server(domain: str) -> FastMCP:
    if domain not in SERVER_INSTRUCTIONS:
        raise ValueError(f"Unknown domain: {domain}")
    mcp = FastMCP(name=f"{domain}-local", instructions=SERVER_INSTRUCTIONS[domain], json_response=True)
    if domain == "hwpx":
        register_hwpx_tools(mcp)
    elif domain == "excel":
        register_excel_tools(mcp)
    else:
        register_qgis_tools(mcp)
    return mcp


def cli() -> None:
    parser = argparse.ArgumentParser(description="Run a local HWPX, Excel, or QGIS MCP server")
    parser.add_argument("domain", choices=["hwpx", "excel", "qgis"])
    args = parser.parse_args()
    build_server(args.domain).run(transport="stdio")


if __name__ == "__main__":
    cli()
