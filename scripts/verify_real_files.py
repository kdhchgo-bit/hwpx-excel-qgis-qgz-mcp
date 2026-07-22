from __future__ import annotations

import json
import os
from copy import copy
from pathlib import Path

from openpyxl import Workbook

from office_gis_mcp.excel_tools import excel_export_pdf, excel_inspect, excel_read_range, excel_write_range
from office_gis_mcp.hwpx_tools import (
    hwpx_export_pdf,
    hwpx_extract_text,
    hwpx_inspect,
    hwpx_native_open_check,
    hwpx_replace_text,
    hwpx_validate,
)
from office_gis_mcp.qgis_tools import (
    qgis_algorithm_help,
    qgis_health,
    qgis_list_algorithms,
    qgis_run_algorithm,
    qgz_audit_sources,
    qgz_inspect,
)

PROJECT_ROOT = Path(__file__).resolve().parents[1]
ARTIFACTS = PROJECT_ROOT / "test-artifacts"


def optional_path_from_env(name: str) -> Path | None:
    value = os.environ.get(name)
    return Path(value).expanduser().resolve() if value else None


HWPX_SAMPLE = optional_path_from_env("OFFICE_GIS_TEST_HWPX")
QGZ_SAMPLE = optional_path_from_env("OFFICE_GIS_TEST_QGZ")


def verify_hwpx() -> dict:
    if HWPX_SAMPLE is None or not HWPX_SAMPLE.is_file():
        return {"skipped": True, "reason": "Set OFFICE_GIS_TEST_HWPX to a local .hwpx file."}
    validation = hwpx_validate(str(HWPX_SAMPLE))
    inspection = hwpx_inspect(str(HWPX_SAMPLE), paragraph_preview_limit=5)
    extracted = hwpx_extract_text(str(HWPX_SAMPLE))
    nonempty = next((line for line in extracted["text"].splitlines() if line.strip()), "")
    edit_result = None
    if nonempty:
        old = nonempty[: min(8, len(nonempty))]
        output = ARTIFACTS / "real_hwpx_mcp_edit.hwpx"
        edit_result = hwpx_replace_text(
            str(HWPX_SAMPLE),
            old,
            f"{old}[MCP검증]",
            output_path=str(output),
            max_replacements=1,
            overwrite=True,
        )
    native = hwpx_native_open_check(str(edit_result["output"] if edit_result else HWPX_SAMPLE))
    pdf = ARTIFACTS / "real_hwpx_hancom_export.pdf"
    exported = hwpx_export_pdf(str(edit_result["output"] if edit_result else HWPX_SAMPLE), str(pdf), overwrite=True)
    return {
        "source": str(HWPX_SAMPLE),
        "valid": validation["valid"],
        "section_count": inspection["section_count"],
        "paragraph_count": inspection["paragraph_count"],
        "table_count": inspection["table_count"],
        "edit": edit_result,
        "native_open": native,
        "native_pdf": exported,
    }


def make_excel_sample(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "MCP검증"
    sheet["A1"] = "항목"
    sheet["B1"] = "값"
    sheet["A2"] = "원본"
    sheet["B2"] = 21
    sheet["A3"] = "계산"
    sheet["B3"] = "=B2*2"
    source_font = copy(sheet["A2"].font)
    source_font.bold = True
    source_font.color = "0000FF"
    sheet["A2"].font = source_font
    sheet.column_dimensions["A"].width = 20
    sheet.column_dimensions["B"].width = 15
    workbook.save(path)


def verify_excel() -> dict:
    source = ARTIFACTS / "excel_com_source.xlsx"
    output = ARTIFACTS / "excel_com_written.xlsx"
    pdf = ARTIFACTS / "excel_com_export.pdf"
    make_excel_sample(source)
    write = excel_write_range(
        str(source),
        "MCP검증",
        "A5",
        [["COM 입력", 123], ["수식", "=B5*2"]],
        output_path=str(output),
        copy_style_from="A2:B3",
        engine="com",
        recalculate=True,
        overwrite=True,
    )
    exported = excel_export_pdf(str(output), str(pdf), sheet_name="MCP검증", overwrite=True)
    return {
        "inspection": excel_inspect(str(output)),
        "range": excel_read_range(str(output), "MCP검증", "A5:B6"),
        "write": write,
        "pdf": exported,
    }


def verify_qgis() -> dict:
    health = qgis_health()
    project = None
    audit = None
    if QGZ_SAMPLE is not None and QGZ_SAMPLE.is_file():
        project = qgz_inspect(str(QGZ_SAMPLE), include_layers=True, max_layers=50)
        audit = qgz_audit_sources(str(QGZ_SAMPLE))
    algorithms = qgis_list_algorithms(search="buffer", max_results=20)
    help_result = qgis_algorithm_help("native:buffer")
    geojson = ARTIFACTS / "qgis_input.geojson"
    output = ARTIFACTS / "qgis_buffer.gpkg"
    geojson.write_text(
        json.dumps(
            {
                "type": "FeatureCollection",
                "name": "qgis_input",
                "crs": {"type": "name", "properties": {"name": "urn:ogc:def:crs:EPSG::5186"}},
                "features": [
                    {
                        "type": "Feature",
                        "properties": {"id": 1},
                        "geometry": {"type": "Point", "coordinates": [200000.0, 500000.0]},
                    }
                ],
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    if output.exists():
        output.unlink()
    run = qgis_run_algorithm(
        "native:buffer",
        {
            "INPUT": str(geojson),
            "DISTANCE": 10,
            "SEGMENTS": 5,
            "END_CAP_STYLE": 0,
            "JOIN_STYLE": 0,
            "MITER_LIMIT": 2,
            "DISSOLVE": False,
            "OUTPUT": str(output),
        },
        timeout_seconds=300,
    )
    return {
        "health": health,
        "real_project": project,
        "real_project_audit": audit,
        "buffer_search_matches": algorithms["match_count"],
        "buffer_help_available": bool(help_result),
        "buffer_run": run,
        "buffer_output": str(output),
        "buffer_output_exists": output.is_file(),
        "buffer_output_bytes": output.stat().st_size if output.is_file() else 0,
    }


def main() -> None:
    ARTIFACTS.mkdir(parents=True, exist_ok=True)
    results = {
        "hwpx": verify_hwpx(),
        "excel": verify_excel(),
        "qgis": verify_qgis(),
    }
    manifest = ARTIFACTS / "verification.json"
    manifest.write_text(json.dumps(results, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(json.dumps(results, ensure_ascii=False, indent=2, default=str))
    print(f"VERIFICATION_MANIFEST={manifest}")


if __name__ == "__main__":
    main()
