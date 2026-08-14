#!/usr/bin/env python3
"""Build build/ExcelVbaLib.xlam from source/ without Excel (pyOpenVBA)."""

from __future__ import annotations

import sys
from pathlib import Path

from pyopenvba import ExcelFile, VBAModuleKind

REPO = Path(__file__).resolve().parent.parent
SOURCE_ROOTS = [
    REPO / "source" / "Internal",
    REPO / "source" / "Api",
    REPO / "source" / "Menus",
]
THIS_WORKBOOK = REPO / "source" / "Menus" / "ThisWorkbook.cls"
OUT = REPO / "build" / "ExcelVbaLib.xlam"


def read_vba(path: Path) -> str:
    text = path.read_text(encoding="utf-8-sig")
    text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    if not text.endswith("\r\n"):
        text += "\r\n"
    return text


def this_workbook_body(path: Path) -> str:
    lines = []
    in_header = False
    for line in read_vba(path).split("\r\n"):
        if line.startswith("VERSION "):
            in_header = True
            continue
        if in_header and (line.startswith("BEGIN") or line == "END"):
            if line == "END":
                in_header = False
            continue
        if line.startswith("Attribute VB_"):
            continue
        lines.append(line)
    return "\r\n".join(lines).strip() + "\r\n"


def module_files() -> list[Path]:
    files: list[Path] = []
    for root in SOURCE_ROOTS:
        if not root.is_dir():
            continue
        for path in sorted(root.glob("*.bas")):
            files.append(path)
    return files


def build(out: Path = OUT) -> Path:
    files = module_files()
    if not files:
        raise SystemExit("No .bas files under source/Internal, Api, or Menus.")
    if not THIS_WORKBOOK.is_file():
        raise SystemExit(f"Missing {THIS_WORKBOOK}")

    out.parent.mkdir(parents=True, exist_ok=True)
    if out.exists():
        out.unlink()

    with ExcelFile.create_new(str(out)) as wb:
        project = wb.vba_project()
        project.name = "ExcelVbaLib"
        for extra in list(wb.module_names()):
            if extra not in ("ThisWorkbook", "Sheet1"):
                project.delete_module(extra)

        for path in files:
            print(f"  add {path.stem} from {path}")
            project.add_module(path.stem, read_vba(path), kind=VBAModuleKind.standard)

        wb.set_module("ThisWorkbook", this_workbook_body(THIS_WORKBOOK))
        wb.save()

    with ExcelFile(str(out)) as wb:
        names = wb.module_names()
        print("modules:", ", ".join(names))
        data = wb.get_module("modInternalData")
        if "Poisson_Inv(" in data:
            raise SystemExit("Built add-in still contains Poisson_Inv( — aborting.")
        if "Function RandomPoisson(" not in data:
            raise SystemExit("Built add-in is missing RandomPoisson — aborting.")
        if "Sub Workbook_Open(" not in wb.get_module("ThisWorkbook"):
            raise SystemExit("ThisWorkbook is missing Workbook_Open — aborting.")

    print(f"Saved {out} ({out.stat().st_size} bytes)")
    return out


if __name__ == "__main__":
    dest = Path(sys.argv[1]) if len(sys.argv) > 1 else OUT
    build(dest)
