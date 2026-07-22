from __future__ import annotations

import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Iterable

MAX_ARCHIVE_ENTRIES = 50_000
MAX_ARCHIVE_UNCOMPRESSED = 2 * 1024 * 1024 * 1024
MAX_XML_MEMBER = 256 * 1024 * 1024


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def require_file(path: str | os.PathLike[str], suffixes: Iterable[str] | None = None) -> Path:
    resolved = Path(path).expanduser().resolve()
    if not resolved.is_file():
        raise FileNotFoundError(f"File not found: {resolved}")
    if suffixes:
        allowed = {suffix.lower() for suffix in suffixes}
        if resolved.suffix.lower() not in allowed:
            raise ValueError(f"Expected one of {sorted(allowed)}, got: {resolved.suffix}")
    return resolved


def output_copy_path(
    source: Path,
    output_path: str | os.PathLike[str] | None,
    marker: str,
    overwrite: bool,
) -> Path:
    if output_path:
        output = Path(output_path).expanduser().resolve()
    else:
        output = source.with_name(f"{source.stem}_{marker}{source.suffix}")
    if output == source:
        raise ValueError("Refusing to overwrite the source file. Choose a separate output_path.")
    if output.exists() and not overwrite:
        raise FileExistsError(f"Output already exists (set overwrite=true to replace it): {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    return output


def copy_for_edit(source: Path, output: Path, overwrite: bool) -> None:
    if output.exists() and overwrite:
        output.unlink()
    shutil.copy2(source, output)


def validate_archive(zf: zipfile.ZipFile) -> dict[str, int | str | None]:
    infos = zf.infolist()
    if len(infos) > MAX_ARCHIVE_ENTRIES:
        raise ValueError(f"Archive has too many entries: {len(infos)}")
    total = 0
    for info in infos:
        member = info.filename.replace("\\", "/")
        parts = Path(member).parts
        if member.startswith("/") or ".." in parts:
            raise ValueError(f"Unsafe archive member: {info.filename}")
        total += info.file_size
        if total > MAX_ARCHIVE_UNCOMPRESSED:
            raise ValueError("Archive expands beyond the configured safety limit")
    bad_member = zf.testzip()
    return {
        "entry_count": len(infos),
        "uncompressed_bytes": total,
        "bad_member": bad_member,
    }


def read_member(zf: zipfile.ZipFile, name: str, max_bytes: int = MAX_XML_MEMBER) -> bytes:
    info = zf.getinfo(name)
    if info.file_size > max_bytes:
        raise ValueError(f"Archive member is too large to inspect safely: {name}")
    return zf.read(name)


def rewrite_zip(source: Path, output: Path, updates: dict[str, bytes], overwrite: bool) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=f".{output.stem}-", suffix=output.suffix, dir=output.parent)
    os.close(fd)
    temp = Path(temp_name)
    try:
        with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(temp, "w") as zout:
            validate_archive(zin)
            for info in zin.infolist():
                payload = updates.get(info.filename)
                if payload is None:
                    payload = read_member(zin, info.filename, max_bytes=MAX_ARCHIVE_UNCOMPRESSED)
                zout.writestr(info, payload)
        with zipfile.ZipFile(temp, "r") as check:
            result = validate_archive(check)
            if result["bad_member"]:
                raise ValueError(f"Rewritten archive failed CRC validation: {result['bad_member']}")
        if output.exists():
            if not overwrite:
                raise FileExistsError(f"Output already exists: {output}")
            output.unlink()
        temp.replace(output)
    finally:
        if temp.exists():
            temp.unlink()


def natural_key(value: str) -> list[str | int]:
    return [int(piece) if piece.isdigit() else piece.lower() for piece in re.split(r"(\d+)", value)]
