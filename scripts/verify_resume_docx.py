#!/usr/bin/env python
"""Verify a generated resume DOCX structurally and, when possible, in Word."""

from __future__ import annotations

import argparse
import json
import platform
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from collections.abc import Sequence


WD_STATISTIC_PAGES = 2


def extract_docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml")
    root = ET.fromstring(xml)
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    chunks = [node.text or "" for node in root.findall(".//w:t", namespace)]
    return "\n".join(chunks)


def word_check(path: Path, export_pdf: Path | None = None) -> dict:
    if platform.system() != "Windows":
        raise RuntimeError("Word COM render check is only available on Windows with Microsoft Word installed.")

    import win32com.client  # type: ignore

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    doc = None
    try:
        doc = word.Documents.Open(str(path.resolve()), ReadOnly=True, AddToRecentFiles=False)
        doc.Repaginate()
        pages = int(doc.ComputeStatistics(WD_STATISTIC_PAGES))
        text = str(doc.Content.Text)
        if export_pdf:
            doc.ExportAsFixedFormat(str(export_pdf.resolve()), 17)
        return {"word_rendered": True, "pages": pages, "word_text_length": len(text)}
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit()


def verify_docx(
    path: Path,
    *,
    expect_one_page: bool = False,
    require_word: bool = False,
    export_pdf: Path | None = None,
    required_sections: Sequence[str] = (),
    required_keywords: Sequence[str] = (),
) -> tuple[dict, int]:
    failures: list[str] = []
    result = {
        "file": str(path),
        "exists": path.exists(),
        "word_rendered": False,
        "word_check_available": platform.system() == "Windows",
        "page_check": "not_checked",
        "pages": None,
        "structural_text_length": 0,
        "missing_sections": [],
        "missing_keywords": [],
    }

    if not path.exists():
        result["ok"] = False
        result["failures"] = ["file_missing"]
        return result, 2

    try:
        text = extract_docx_text(path)
        text_compact = re.sub(r"\s+", "", text)
        result["structural_text_length"] = len(text_compact)
    except Exception as exc:
        failures.append(f"cannot_extract_docx_text: {exc}")
        text_compact = ""

    missing_sections = [s for s in required_sections if s and s not in text_compact]
    missing_keywords = [k for k in required_keywords if k and k not in text_compact]
    result["missing_sections"] = missing_sections
    result["missing_keywords"] = missing_keywords
    if missing_sections:
        failures.append("missing_sections: " + ", ".join(missing_sections))
    if missing_keywords:
        failures.append("missing_keywords: " + ", ".join(missing_keywords))

    try:
        result.update(word_check(path, export_pdf))
        result["page_check"] = "word_rendered"
    except Exception as exc:
        result["word_error"] = str(exc)
        if require_word:
            failures.append(f"word_render_check_failed: {exc}")

    if expect_one_page and result.get("pages") is not None and result.get("pages") != 1:
        failures.append(f"expected_one_page_but_got_{result.get('pages')}")
    elif expect_one_page and result.get("pages") is None and require_word:
        failures.append("cannot_confirm_one_page_without_word")
    elif expect_one_page and result.get("pages") is None:
        result["page_check"] = "not_confirmed_without_word"

    if export_pdf:
        result["pdf_exported"] = export_pdf.exists()
        if not export_pdf.exists():
            failures.append("pdf_export_missing")

    result["ok"] = not failures
    result["failures"] = failures
    return result, 0 if not failures else 1


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", help="DOCX file to verify.")
    parser.add_argument("--expect-one-page", action="store_true")
    parser.add_argument("--require-word", action="store_true")
    parser.add_argument("--export-pdf", help="Optional PDF export path.")
    parser.add_argument("--required-section", action="append", default=[])
    parser.add_argument("--required-keyword", action="append", default=[])
    args = parser.parse_args()

    result, exit_code = verify_docx(
        Path(args.docx),
        expect_one_page=args.expect_one_page,
        require_word=args.require_word,
        export_pdf=Path(args.export_pdf) if args.export_pdf else None,
        required_sections=args.required_section,
        required_keywords=args.required_keyword,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
