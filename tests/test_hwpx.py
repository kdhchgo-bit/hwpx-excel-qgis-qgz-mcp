from __future__ import annotations

import zipfile

from office_gis_mcp.hwpx_tools import hwpx_extract_text, hwpx_find_text, hwpx_inspect, hwpx_replace_text, hwpx_validate


def make_hwpx(path):
    section = b'''<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="urn:hancom:section" xmlns:hp="urn:hancom:paragraph">
  <hp:p><hp:run><hp:t>Hello </hp:t></hp:run><hp:run><hp:t>HWPX</hp:t></hp:run></hp:p>
  <hp:tbl><hp:p><hp:run><hp:t>Table text</hp:t></hp:run></hp:p></hp:tbl>
</hs:sec>'''
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("mimetype", "application/hwp+zip")
        zf.writestr("Contents/content.hpf", b'<opf:package xmlns:opf="urn:opf"/>')
        zf.writestr("Contents/header.xml", b'<hh:head xmlns:hh="urn:head"/>')
        zf.writestr("Contents/section0.xml", section)
        zf.writestr("Preview/PrvText.txt", "Hello HWPX\nTable text")
        zf.writestr("BinData/image1.png", b"not-a-real-image")


def test_hwpx_inspect_find_replace(tmp_path):
    source = tmp_path / "sample.hwpx"
    make_hwpx(source)
    assert hwpx_validate(str(source))["valid"] is True
    inspected = hwpx_inspect(str(source))
    assert inspected["paragraph_count"] == 2
    assert inspected["table_count"] == 1
    assert inspected["image_count"] == 1
    assert hwpx_find_text(str(source), "Hello HWPX")["match_count"] == 1
    output = tmp_path / "edited.hwpx"
    result = hwpx_replace_text(str(source), "Hello HWPX", "안녕하세요 HWPX", str(output))
    assert result["replacement_count"] == 1
    assert "안녕하세요 HWPX" in hwpx_extract_text(str(output))["text"]
