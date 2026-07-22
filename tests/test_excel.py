from __future__ import annotations

from openpyxl import Workbook

from office_gis_mcp.excel_tools import excel_find, excel_inspect, excel_read_range, excel_write_range


def make_workbook(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "데이터"
    sheet["A1"] = "항목"
    sheet["B1"] = "값"
    sheet["A2"] = "유량"
    sheet["B2"] = 42
    sheet["B3"] = "=B2*2"
    sheet.merge_cells("D1:E1")
    workbook.save(path)


def test_excel_read_find_write(tmp_path):
    source = tmp_path / "sample.xlsx"
    make_workbook(source)
    inspected = excel_inspect(str(source))
    assert inspected["sheet_count"] == 1
    assert inspected["sheets"][0]["formula_cells"] == 1
    assert excel_read_range(str(source), "데이터", "A1:B2")["values"][1] == ["유량", 42]
    assert excel_find(str(source), "B2*2")["match_count"] == 1
    output = tmp_path / "edited.xlsx"
    result = excel_write_range(
        str(source),
        "데이터",
        "A4",
        [["수정", 100]],
        output_path=str(output),
        copy_style_from="A2",
        engine="openpyxl",
    )
    assert result["cells_written"] == 2
    assert excel_read_range(str(output), "데이터", "A4:B4")["values"] == [["수정", 100]]
