from __future__ import annotations

import zipfile

from office_gis_mcp.qgis_tools import qgz_audit_sources, qgz_inspect, qgz_rebase_paths


def make_qgz(path, datasource):
    xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<qgis projectname="MCP test" version="3.44.11-Solothurn">
  <projectCrs><spatialrefsys><authid>EPSG:5186</authid></spatialrefsys></projectCrs>
  <mapcanvas><extent><xmin>0</xmin><ymin>1</ymin><xmax>2</xmax><ymax>3</ymax></extent></mapcanvas>
  <projectlayers><maplayer type="vector" geometry="Point"><id>points</id><layername>테스트</layername><datasource>{datasource}</datasource><provider>ogr</provider><srs><spatialrefsys><authid>EPSG:5186</authid></spatialrefsys></srs></maplayer></projectlayers>
  <Layouts><Layout name="도면 1"/></Layouts>
</qgis>'''
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("project.qgs", xml.encode("utf-8"))


def test_qgz_inspect_audit_rebase(tmp_path):
    data_dir = tmp_path / "old"
    data_dir.mkdir()
    data_file = data_dir / "points.gpkg"
    data_file.write_bytes(b"placeholder")
    source = tmp_path / "sample.qgz"
    make_qgz(source, str(data_file))
    inspected = qgz_inspect(str(source))
    assert inspected["project_crs"] == "EPSG:5186"
    assert inspected["layer_count"] == 1
    assert inspected["layout_count"] == 1
    assert qgz_audit_sources(str(source))["missing_count"] == 0
    output = tmp_path / "rebased.qgz"
    result = qgz_rebase_paths(str(source), str(data_dir), str(tmp_path / "new"), str(output))
    assert result["replacement_count"] >= 1
    assert qgz_audit_sources(str(output))["missing_count"] == 1
