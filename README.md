# HWPX · Excel · QGIS Local Tools for Codex

[한국어](#한국어) · [English](#english)

> 한글 HWPX, Microsoft Excel, QGIS QGS/QGZ 파일을 Windows의 Codex에서 안전하게 다루기 위한 로컬 MCP 서버와 `$` 스킬 모음입니다.

> Local MCP servers and `$` skills for safe HWPX, Microsoft Excel, and QGIS QGS/QGZ workflows in Codex on Windows.

## 한국어

### 소개

이 프로젝트는 문서·스프레드시트·GIS 프로젝트를 Codex에서 직접 검사하고 편집할 수 있도록 세 가지 로컬 도구 세트를 제공합니다.

- `$hwpx-local`: HWPX 구조 검증, 문단·표·이미지 검사, 텍스트 검색·추출·치환, 한컴오피스 한글을 이용한 열기 확인 및 PDF 변환
- `$excel-local`: 통합 문서·시트·수식 검사, 범위 읽기와 검색, 서식 보존 입력, Excel 재계산 및 PDF 변환
- `$qgis-local`: QGS/QGZ 프로젝트와 레이어 검사, 누락된 데이터소스 감사, 경로 재설정, QGIS Processing 알고리즘 조회·실행

각 도구는 MCP 서버로 사용할 수도 있고, 현재 대화에 MCP가 표시되지 않을 때 `$` 스킬이 로컬 실행기를 직접 호출하도록 사용할 수도 있습니다.

### 안전한 작업 원칙

- 편집 작업은 기본적으로 원본을 덮어쓰지 않고 별도 복사본을 만듭니다.
- HWPX는 ZIP/XML 무결성을 검사하고, Excel은 필요할 때 설치된 Excel COM을 사용하며, QGZ는 내부 QGS XML과 데이터소스 경로를 검사합니다.
- 한글 경로와 JSON은 UTF-8로 전달합니다.
- 사용자 문서, 테스트 산출물, 가상환경과 로컬 실행 경로는 Git 저장소에 포함하지 않습니다.

### 요구 사항

- Windows 10/11
- [uv](https://docs.astral.sh/uv/) 및 Python 3.11
- Codex CLI
- 선택 기능:
  - 한컴오피스 한글: HWPX 네이티브 열기 확인 및 PDF 변환
  - Microsoft Excel: COM 기반 고충실도 편집·재계산·PDF 변환
  - QGIS: Processing 알고리즘 실행

### 설치

```powershell
git clone https://github.com/kdhchgo-bit/hwpx-excel-qgis-qgz-mcp.git
cd hwpx-excel-qgis-qgz-mcp
powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1
```

QGIS Processing 실행 파일이 기본 위치와 다르면 경로를 지정합니다.

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1 `
  -QgisProcessPath 'C:\OSGeo4W\bin\qgis_process-qgis-ltr.bat'
```

설치 스크립트는 프로젝트 가상환경을 만들고, 세 MCP 서버를 Codex에 등록하며, 세 `$` 스킬을 사용자 전역 스킬 폴더에 설치합니다. 설치 후 새 Codex 작업을 열거나 앱을 다시 시작한 다음 `$hwpx-local`, `$excel-local`, `$qgis-local`을 호출하세요.

### 사용 예시

```text
$hwpx-local D:\문서\보고서.hwpx를 검사하고 250mm를 350mm로 바꾼 복사본을 만들어줘.
$excel-local D:\자료\계산서.xlsx의 집계 시트 G3:H30을 읽고 기존 서식을 유지해 값을 채워줘.
$qgis-local D:\GIS\사업.qgz의 누락된 레이어 경로를 검사하고 경로를 수정한 복사본을 만들어줘.
```

자연어로 요청해도 스킬이 선택될 수 있지만, `$스킬명`을 명시하면 가장 확실합니다.

### 검증

```powershell
.\.venv\Scripts\python.exe -m pytest
.\.venv\Scripts\python.exe .\scripts\smoke_mcp.py
```

실제 로컬 파일을 이용하는 `verify_real_files.py`는 개인 문서와 설치된 데스크톱 앱을 사용하므로 공개 저장소의 일반 CI용 테스트가 아닙니다.

## English

### Overview

This project provides three local toolsets that let Codex inspect and edit office documents, spreadsheets, and GIS projects on Windows.

- `$hwpx-local`: validate HWPX packages; inspect paragraphs, tables, and images; search, extract, and replace text; verify opening in Hancom Hangul; export PDF
- `$excel-local`: inspect workbooks, sheets, and formulas; read and search ranges; write while preserving formatting; recalculate with Excel; export PDF
- `$qgis-local`: inspect QGS/QGZ projects and layers; audit missing data sources; rebase project paths; discover and run QGIS Processing algorithms

Each toolset works as a local MCP server. The included `$` skill can also call the verified local runner directly when the MCP tools are not visible in the current Codex task.

### Safety model

- Mutating operations create a separate output copy by default instead of overwriting the source.
- HWPX operations validate ZIP/XML integrity, Excel operations can use native Excel COM for fidelity, and QGZ operations inspect the embedded QGS XML and layer sources.
- Korean file paths and JSON payloads are passed as UTF-8.
- User documents, test artifacts, virtual environments, and machine-specific runtime paths are excluded from Git.

### Requirements

- Windows 10/11
- [uv](https://docs.astral.sh/uv/) and Python 3.11
- Codex CLI
- Optional native applications:
  - Hancom Hangul for native HWPX open checks and PDF export
  - Microsoft Excel for high-fidelity COM editing, recalculation, and PDF export
  - QGIS for Processing algorithms

### Install

```powershell
git clone https://github.com/kdhchgo-bit/hwpx-excel-qgis-qgz-mcp.git
cd hwpx-excel-qgis-qgz-mcp
powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1
```

Pass a custom QGIS Processing path when needed:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1 `
  -QgisProcessPath 'C:\OSGeo4W\bin\qgis_process-qgis-ltr.bat'
```

The installer creates the project virtual environment, registers all three MCP servers with Codex, and installs the three `$` skills in the user's global skill directory. Open a new Codex task or restart the app, then invoke `$hwpx-local`, `$excel-local`, or `$qgis-local`.

### Example prompts

```text
$hwpx-local inspect D:\docs\report.hwpx and create a separate copy that replaces 250mm with 350mm.
$excel-local read Summary!G3:H30 in D:\data\estimate.xlsx and fill values while preserving the existing format.
$qgis-local audit missing layer sources in D:\GIS\project.qgz and create a copy with rebased paths.
```

Natural-language routing can select a skill automatically, but explicitly naming `$skill-name` is the most reliable option.

### Verification

```powershell
.\.venv\Scripts\python.exe -m pytest
.\.venv\Scripts\python.exe .\scripts\smoke_mcp.py
```

`verify_real_files.py` uses private local files and installed desktop applications. It is intended for workstation acceptance testing, not generic public CI.

## Available tools

| Server / skill | Tools |
|---|---|
| `hwpx-local` | `hwpx_validate`, `hwpx_inspect`, `hwpx_extract_text`, `hwpx_find_text`, `hwpx_replace_text`, `hwpx_native_open_check`, `hwpx_export_pdf` |
| `excel-local` | `excel_inspect`, `excel_read_range`, `excel_find`, `excel_write_range`, `excel_recalculate`, `excel_export_pdf` |
| `qgis-local` | `qgz_inspect`, `qgz_audit_sources`, `qgz_rebase_paths`, `qgis_list_algorithms`, `qgis_algorithm_help`, `qgis_run_algorithm` |

## Technical references

- [Model Context Protocol Python SDK](https://github.com/modelcontextprotocol/python-sdk)
- [QGIS project files](https://docs.qgis.org/3.44/en/docs/user_manual/introduction/project_files.html)
- [QGIS QGS/QGZ file formats](https://docs.qgis.org/3.44/en/docs/user_manual/appendices/qgis_file_formats.html)
