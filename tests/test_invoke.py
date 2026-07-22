from __future__ import annotations

import json

import pytest

from office_gis_mcp.invoke import TOOLS, _load_arguments, invoke_tool


def test_tool_registry_contains_all_domains():
    assert {"hwpx_health", "excel_health", "qgis_health"} <= TOOLS.keys()
    assert "hwpx_analyze_tables" in TOOLS


def test_invoke_health_returns_envelope():
    response = invoke_tool("excel_health")
    assert response["ok"] is True
    assert response["tool"] == "excel_health"
    assert response["result"]["server"] == "excel-local"


def test_hwpx_health_lists_table_morphology():
    response = invoke_tool("hwpx_health")
    assert "analyze_table_morphology" in response["result"]["features"]


def test_arguments_must_be_json_object():
    assert _load_arguments(json.dumps({"path": "sample.xlsx"})) == {"path": "sample.xlsx"}
    assert _load_arguments("\ufeff" + json.dumps({"path": "한글.xlsx"})) == {"path": "한글.xlsx"}
    with pytest.raises(TypeError, match="JSON object"):
        _load_arguments("[]")
