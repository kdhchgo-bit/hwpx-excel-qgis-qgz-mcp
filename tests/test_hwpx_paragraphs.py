from __future__ import annotations

import zipfile

from lxml import etree

from office_gis_mcp.hwpx_paragraphs import (
    hwpx_analyze_paragraph_hierarchy,
    hwpx_normalize_paragraph_hierarchy,
)


def _para_property(
    property_id: int,
    left: int,
    intent: int,
    *,
    alignment: str = "JUSTIFY",
    heading_type: str = "NONE",
    heading_level: int = 0,
) -> str:
    return f'''
      <hh:paraPr id="{property_id}" tabPrIDRef="0">
        <hh:align horizontal="{alignment}" vertical="BASELINE"/>
        <hh:heading type="{heading_type}" idRef="0" level="{heading_level}"/>
        <hh:breakSetting pageBreakBefore="0"/>
        <hh:margin>
          <hc:intent value="{intent}" unit="HWPUNIT"/>
          <hc:left value="{left}" unit="HWPUNIT"/>
          <hc:right value="0" unit="HWPUNIT"/>
          <hc:prev value="0" unit="HWPUNIT"/>
          <hc:next value="0" unit="HWPUNIT"/>
        </hh:margin>
        <hh:lineSpacing type="PERCENT" value="160"/>
      </hh:paraPr>'''


def _char_property(property_id: int, height: int, *, bold: bool = False) -> str:
    return f'''
      <hh:charPr id="{property_id}" height="{height}" textColor="#000000">
        {"<hh:bold/>" if bold else ""}
      </hh:charPr>'''


def _style(style_id: int, para_id: int, char_id: int, name: str) -> str:
    return f'<hh:style id="{style_id}" type="PARA" name="{name}" paraPrIDRef="{para_id}" charPrIDRef="{char_id}" nextStyleIDRef="0"/>'


def _paragraph(
    text: str,
    para_id: int,
    style_id: int,
    char_id: int,
    *,
    page_break: bool = False,
    extra_run: tuple[str, int] | None = None,
) -> str:
    extra = (
        f'<hp:run charPrIDRef="{extra_run[1]}"><hp:t>{extra_run[0]}</hp:t></hp:run>'
        if extra_run
        else ""
    )
    return f'''
  <hp:p paraPrIDRef="{para_id}" styleIDRef="{style_id}" pageBreak="{int(page_break)}">
    <hp:run charPrIDRef="{char_id}"><hp:t>{text}</hp:t></hp:run>{extra}
    <hp:linesegarray><hp:lineseg textpos="0" vertpos="1000"/></hp:linesegarray>
  </hp:p>'''


def _table_paragraph() -> str:
    return """
  <hp:p paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0">
      <hp:tbl rowCnt="1" colCnt="1">
        <hp:tr><hp:tc><hp:subList><hp:p paraPrIDRef="99" styleIDRef="99"><hp:run charPrIDRef="99"><hp:t>1. 표 내부 문단</hp:t></hp:run></hp:p></hp:subList></hp:tc></hp:tr>
      </hp:tbl>
    </hp:run>
  </hp:p>"""


def make_hierarchy_hwpx(path) -> None:
    header = f"""<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">
  <hh:refList>
    <hh:paraProperties itemCnt="7">
      {_para_property(0, 0, 0)}
      {_para_property(10, 1000, -1000)}
      {_para_property(11, 2000, -1000)}
      {_para_property(12, 3000, -1000)}
      {_para_property(13, 4000, -1000)}
      {_para_property(90, 0, 0, alignment="CENTER")}
      {_para_property(99, 0, 0)}
    </hh:paraProperties>
    <hh:charProperties itemCnt="8">
      {_char_property(0, 1000)}
      {_char_property(10, 1600, bold=True)}
      {_char_property(11, 1400, bold=True)}
      {_char_property(12, 1200)}
      {_char_property(13, 1000)}
      {_char_property(50, 1200, bold=True)}
      {_char_property(90, 2600, bold=True)}
      {_char_property(99, 900)}
    </hh:charProperties>
    <hh:styles itemCnt="7">
      {_style(0, 0, 0, "바탕글")}
      {_style(10, 10, 10, "수준1")}
      {_style(11, 11, 11, "수준2")}
      {_style(12, 12, 12, "수준3")}
      {_style(13, 13, 13, "수준4")}
      {_style(90, 90, 90, "간지제목")}
      {_style(99, 99, 99, "잘못된서식")}
    </hh:styles>
  </hh:refList>
</hh:head>"""
    section = f"""<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  {_paragraph("제 1 장 사업 개요", 90, 90, 90, page_break=True)}
  {_paragraph("1. 계획 개요", 10, 10, 10)}
  {_paragraph("2. 추진 방향", 10, 10, 10)}
  {_paragraph("가. 일반 사항", 11, 11, 11)}
  {_paragraph("나. 세부 사항", 11, 11, 11)}
  {_paragraph("1) 조사 범위", 12, 12, 12)}
  {_paragraph("2) 조사 방법", 12, 12, 12)}
  {_paragraph("3) 잘못된 세부 ", 99, 99, 99, extra_run=("강조", 50))}
  {_paragraph("가) 기존 자료", 13, 13, 13)}
  {_paragraph("나) 현장 자료", 13, 13, 13)}
  {_paragraph("다) 잘못된 상세", 99, 99, 99, page_break=True)}
  {_paragraph("1. 부록 자료", 90, 90, 90, page_break=True)}
  {_paragraph("신 청 서", 90, 90, 90)}
  {_table_paragraph()}
</hs:sec>"""
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("mimetype", "application/hwp+zip")
        zf.writestr("Contents/content.hpf", '<opf:package xmlns:opf="urn:opf"/>')
        zf.writestr("Contents/header.xml", header)
        zf.writestr("Contents/section0.xml", section)


def _paragraphs(path):
    with zipfile.ZipFile(path, "r") as zf:
        root = etree.fromstring(zf.read("Contents/section0.xml"))
    return [item for item in root.iter() if item.tag.rsplit("}", 1)[-1] == "p"]


def test_analyze_hierarchy_learns_styles_and_excludes_interleaves(tmp_path):
    source = tmp_path / "hierarchy.hwpx"
    make_hierarchy_hwpx(source)

    result = hwpx_analyze_paragraph_hierarchy(str(source))
    assert result["hierarchy_candidate_count"] == 11
    assert result["normalization_candidate_count"] == 2
    assert result["interleave_page_count"] == 2
    pages = {page["id"]: page for page in result["pages"]}
    assert pages["S01-PG0001"]["is_interleave"] is True
    assert pages["S01-PG0003"]["is_interleave"] is True
    assert pages["S01-PG0004"]["contains_table"] is True
    assert pages["S01-PG0004"]["is_interleave"] is False

    levels = {item["level"]: item for item in result["levels"]}
    assert levels[1]["canonical"]["para_property_id_ref"] == "10"
    assert levels[3]["canonical"]["para_property_id_ref"] == "12"
    assert levels[3]["canonical"]["support_ratio"] == 0.6667
    level_three = next(
        item for item in result["paragraphs"] if item["text"].startswith("1)")
    )
    assert level_three["paragraph_property"]["margin"] == {
        "intent": "-1000",
        "left": "3000",
        "right": "0",
        "prev": "0",
        "next": "0",
    }
    appendix = next(
        item for item in result["paragraphs"] if item["text"] == "1. 부록 자료"
    )
    assert appendix["excluded_reason"] == "interleave"


def test_interleave_exclusion_is_a_toggle(tmp_path):
    source = tmp_path / "toggle.hwpx"
    make_hierarchy_hwpx(source)

    included = hwpx_analyze_paragraph_hierarchy(str(source), exclude_interleaves=False)
    appendix = next(
        item for item in included["paragraphs"] if item["text"] == "1. 부록 자료"
    )
    assert appendix["excluded_reason"] is None
    assert appendix["normalization_candidate"] is True
    assert included["normalization_candidate_count"] == 3


def test_normalize_hierarchy_dry_run_and_apply_preserve_inline_emphasis(tmp_path):
    source = tmp_path / "normalize.hwpx"
    output = tmp_path / "normalized.hwpx"
    make_hierarchy_hwpx(source)

    preview = hwpx_normalize_paragraph_hierarchy(str(source))
    assert preview["dry_run"] is True
    assert preview["planned_change_count"] == 2
    assert preview["output"] is None
    assert not output.exists()

    result = hwpx_normalize_paragraph_hierarchy(
        str(source),
        output_path=str(output),
        dry_run=False,
    )
    assert result["valid_output"] is True
    assert result["applied_change_count"] == 2
    assert output.is_file()

    source_paragraphs = _paragraphs(source)
    output_paragraphs = _paragraphs(output)
    assert source_paragraphs[7].get("paraPrIDRef") == "99"
    assert output_paragraphs[7].get("paraPrIDRef") == "12"
    assert output_paragraphs[7].get("styleIDRef") == "12"
    output_runs = [
        item for item in output_paragraphs[7] if item.tag.rsplit("}", 1)[-1] == "run"
    ]
    assert [run.get("charPrIDRef") for run in output_runs] == ["12", "50"]
    assert output_paragraphs[10].get("paraPrIDRef") == "13"
    assert output_paragraphs[11].get("paraPrIDRef") == "90"
