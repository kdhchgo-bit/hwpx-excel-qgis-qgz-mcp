from __future__ import annotations

import json
from pathlib import Path

from office_gis_mcp.hwpx_tools import hwpx_export_pdf, hwpx_native_open_check


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    source = root / "test-artifacts" / "real_hwpx_mcp_edit.hwpx"
    output = root / "test-artifacts" / "real_hwpx_hancom_export.pdf"
    result = {
        "open": hwpx_native_open_check(str(source)),
        "pdf": hwpx_export_pdf(str(source), str(output), overwrite=True),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
