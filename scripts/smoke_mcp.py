from __future__ import annotations

import argparse
import asyncio
import json
import os
import sys

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


async def check(domain: str) -> dict:
    env = dict(os.environ)
    env.setdefault("PYTHONUTF8", "1")
    if domain == "qgis":
        env.setdefault("QGIS_PROCESS_PATH", r"D:\bin\qgis_process-qgis-ltr.bat")
    params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "office_gis_mcp.server", domain],
        env=env,
    )
    async with stdio_client(params) as (reader, writer):
        async with ClientSession(reader, writer) as session:
            await session.initialize()
            tools = await session.list_tools()
            health = await session.call_tool(f"{domain}_health", {})
            return {
                "domain": domain,
                "tools": [tool.name for tool in tools.tools],
                "health_error": health.isError,
                "health": health.structuredContent,
            }


async def main(domains: list[str]) -> None:
    results = []
    for domain in domains:
        results.append(await check(domain))
    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("domains", nargs="*", default=["hwpx", "excel", "qgis"])
    args = parser.parse_args()
    asyncio.run(main(args.domains))
