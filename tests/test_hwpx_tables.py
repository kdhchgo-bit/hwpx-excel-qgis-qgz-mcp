from __future__ import annotations

import zipfile

from office_gis_mcp.hwpx_tables import hwpx_analyze_tables


def _border_fill(border_id: int, width: str, fill: str, top_type: str = "SOLID") -> str:
    return f'''
    <hh:borderFill id="{border_id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">
      <hh:leftBorder type="SOLID" width="{width}" color="#000000"/>
      <hh:rightBorder type="SOLID" width="{width}" color="#000000"/>
      <hh:topBorder type="{top_type}" width="{width}" color="#000000"/>
      <hh:bottomBorder type="SOLID" width="{width}" color="#000000"/>
      <hc:fillBrush><hc:winBrush faceColor="{fill}" hatchColor="#000000" alpha="0"/></hc:fillBrush>
    </hh:borderFill>'''


def _cell(
    row: int,
    col: int,
    text: str,
    border_id: int,
    *,
    header: bool = False,
    row_span: int = 1,
    col_span: int = 1,
    center: bool = False,
    bold: bool = False,
    height: int = 1000,
) -> str:
    return f'''
      <hp:tc header="{int(header)}" borderFillIDRef="{border_id}">
        <hp:subList vertAlign="CENTER">
          <hp:p paraPrIDRef="{int(center)}"><hp:run charPrIDRef="{int(bold)}"><hp:t>{text}</hp:t></hp:run></hp:p>
        </hp:subList>
        <hp:cellAddr rowAddr="{row}" colAddr="{col}"/>
        <hp:cellSpan rowSpan="{row_span}" colSpan="{col_span}"/>
        <hp:cellSz width="2000" height="{height}"/>
      </hp:tc>'''


def make_morphology_hwpx(path, explicit_headers: bool = True):
    header = f'''<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">
  <hh:refList>
    <hh:borderFills itemCnt="3">
      {_border_fill(1, "0.12 mm", "#FFFFFF")}
      {_border_fill(2, "0.5 mm", "#DDEEFF")}
      {_border_fill(3, "0.7 mm", "#EEEEEE", top_type="DOUBLE")}
    </hh:borderFills>
    <hh:paraProperties itemCnt="2">
      <hh:paraPr id="0"><hh:align horizontal="LEFT" vertical="BASELINE"/></hh:paraPr>
      <hh:paraPr id="1"><hh:align horizontal="CENTER" vertical="BASELINE"/></hh:paraPr>
    </hh:paraProperties>
    <hh:charProperties itemCnt="2">
      <hh:charPr id="0" height="1000"/>
      <hh:charPr id="1" height="1000"><hh:bold/></hh:charPr>
    </hh:charProperties>
  </hh:refList>
</hh:head>'''
    flag = explicit_headers
    section = f'''<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p><hp:run>
    <hp:tbl rowCnt="4" colCnt="3" repeatHeader="1" borderFillIDRef="1">
      <hp:tr>
        {_cell(0, 0, "구분", 2, header=flag, col_span=2, center=True, bold=True)}
        {_cell(0, 2, "금액", 2, header=flag, center=True, bold=True)}
      </hp:tr>
      <hp:tr>
        {_cell(1, 0, "분류", 2, header=flag, row_span=2, center=True, bold=True)}
        {_cell(1, 1, "항목 A", 1)}
        {_cell(1, 2, "100", 1)}
      </hp:tr>
      <hp:tr>
        {_cell(2, 1, "항목 B", 1)}
        {_cell(2, 2, "200", 1)}
      </hp:tr>
      <hp:tr>
        {_cell(3, 0, "작성:        (서명)", 3, col_span=3, center=True, bold=True, height=1800)}
      </hp:tr>
    </hp:tbl>
  </hp:run></hp:p>
</hs:sec>'''
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("mimetype", "application/hwp+zip")
        zf.writestr("Contents/content.hpf", '<opf:package xmlns:opf="urn:opf"/>')
        zf.writestr("Contents/header.xml", header)
        zf.writestr("Contents/section0.xml", section)


def test_table_morphology_explicit_headers_and_footer(tmp_path):
    source = tmp_path / "morphology.hwpx"
    make_morphology_hwpx(source)
    result = hwpx_analyze_tables(str(source))
    assert result["table_count"] == 1
    table = result["tables"][0]
    assert (table["grid_rows"], table["grid_columns"]) == (4, 3)
    assert table["structure"]["merged_cell_count"] == 3
    assert table["structure"]["gap_count"] == 0
    assert table["roles"]["header_rows"] == [0]
    assert table["roles"]["inferred_header_rows"] == []
    assert table["roles"]["body_rows"] == [1, 2]
    assert table["roles"]["footer_form_rows"] == [3]
    assert table["roles"]["explicit_header_cells"] == ["R0C0", "R0C2", "R1C0"]
    assert table["roles"]["footer_form_cells"] == ["R3C0"]
    cells = {cell["anchor"]: cell for cell in table["cells"]}
    assert cells["R0C0"]["covers"] == [[0, 0], [0, 1]]
    assert cells["R1C0"]["role"] == "header"
    assert cells["R1C1"]["role"] == "body"
    assert cells["R3C0"]["role"] == "footer"
    assert table["borders"]["regions"]["header"]["outer"]["cell_side_count"] > 0
    assert table["borders"]["regions"]["header"]["internal"]["cell_side_count"] > 0
    assert table["borders"]["regions"]["body"]["outer"]["cell_side_count"] > 0
    assert table["borders"]["regions"]["header"]["fills"][0]["border_fill_id"] == "2"
    assert table["borders"]["regions"]["body"]["fills"][0]["border_fill_id"] == "1"
    assert table["borders"]["regions"]["footer"]["fills"][0]["border_fill_id"] == "3"
    assert table["borders"]["shared_boundary_mismatch_count"] > 0


def test_table_morphology_infers_unflagged_header(tmp_path):
    source = tmp_path / "heuristic.hwpx"
    make_morphology_hwpx(source, explicit_headers=False)
    table = hwpx_analyze_tables(str(source), table_index=0)["tables"][0]
    assert table["roles"]["explicit_header_cell_count"] == 0
    assert table["roles"]["inferred_header_rows"] == [0]
    assert table["roles"]["footer_form_rows"] == [3]
    assert any("heuristic" in warning.lower() for warning in table["warnings"])
