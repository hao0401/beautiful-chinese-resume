#!/usr/bin/env python
"""Build a polished one-page Chinese resume DOCX from structured JSON."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "Missing dependency: python-docx. Install with `python -m pip install python-docx`."
    ) from exc


ACCENTS = {
    "campus": "2F6F73",
    "business": "31516A",
    "fresh": "3C7A89",
    "minimal": "404040",
}

SECTION_ORDER = ["教育背景", "项目经历", "实习经历", "技能证书", "校园经历", "获奖经历", "自我评价"]


def load_json(path: str) -> dict[str, Any]:
    if path == "-":
        return json.loads(sys.stdin.read().lstrip("\ufeff"))
    with open(path, "r", encoding="utf-8-sig") as fh:
        return json.load(fh)


def set_run_font(run, *, size: float | None = None, bold: bool = False, color: str | None = None) -> None:
    run.font.name = "Microsoft YaHei"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    run._element.rPr.rFonts.set(qn("w:ascii"), "Microsoft YaHei")
    run._element.rPr.rFonts.set(qn("w:hAnsi"), "Microsoft YaHei")
    if size:
        run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def set_paragraph_spacing(paragraph, *, before: float = 0, after: float = 0, line: float = 1.0) -> None:
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line


def add_bottom_border(paragraph, color: str = "D8E4E6") -> None:
    p = paragraph._p
    p_pr = p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)


def add_text(paragraph, text: str, *, size: float = 10, bold: bool = False, color: str | None = None):
    run = paragraph.add_run(text)
    set_run_font(run, size=size, bold=bold, color=color)
    return run


def normalize_sections(data: dict[str, Any]) -> list[dict[str, Any]]:
    sections = data.get("sections")
    if isinstance(sections, list):
        return [s for s in sections if s.get("title")]

    mapped: list[dict[str, Any]] = []
    key_map = [
        ("education", "教育背景"),
        ("projects", "项目经历"),
        ("internships", "实习经历"),
        ("experience", "实习经历"),
        ("skills", "技能证书"),
        ("certificates", "技能证书"),
        ("campus", "校园经历"),
        ("awards", "获奖经历"),
    ]
    for key, title in key_map:
        value = data.get(key)
        if not value:
            continue
        if key in {"skills", "certificates"} and all(isinstance(v, str) for v in value):
            items = [{"heading": title, "details": ["；".join(value)]}]
        elif isinstance(value, list):
            items = value
        else:
            items = [{"heading": title, "details": [str(value)]}]

        existing = next((s for s in mapped if s["title"] == title), None)
        if existing:
            existing["items"].extend(items)
        else:
            mapped.append({"title": title, "items": items})

    return mapped


def sort_sections(sections: list[dict[str, Any]]) -> list[dict[str, Any]]:
    order = {title: idx for idx, title in enumerate(SECTION_ORDER)}
    def section_sort_key(section: dict[str, Any]) -> int:
        return order.get(section["title"], 99)

    return sorted(sections, key = section_sort_key)


def add_header(doc: Document, data: dict[str, Any], accent: str) -> None:
    name = data.get("name") or "姓名"
    target_role = data.get("target_role") or data.get("role") or "求职方向待补充"
    target_company = data.get("target_company")
    show_company = bool(data.get("show_target_company"))
    contact = data.get("contact") or []
    if isinstance(contact, str):
        contact = [contact]

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, after=1)
    add_text(p, name, size=21, bold=True, color="202124")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p, after=1)
    role_line = f"求职意向：{target_role}"
    if target_company and show_company:
        role_line += f" | 目标公司：{target_company}"
    add_text(p, role_line, size=10.5, bold=True, color=accent)

    if contact:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p, after=5)
        add_text(p, " | ".join(str(item) for item in contact if item), size=9.2, color="5F6368")


def add_summary(doc: Document, summary: Any) -> None:
    if not summary:
        return
    if isinstance(summary, str):
        summary = [summary]
    p = doc.add_paragraph()
    set_paragraph_spacing(p, after=3)
    add_text(p, "；".join(str(item).strip() for item in summary if str(item).strip()), size=9.8, color="3C4043")


def add_section_title(doc: Document, title: str, accent: str) -> None:
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=3, after=2)
    add_text(p, title, size=11, bold=True, color=accent)
    add_bottom_border(p)


def add_item(doc: Document, item: Any) -> None:
    if isinstance(item, str):
        p = doc.add_paragraph()
        set_paragraph_spacing(p, after=1)
        add_text(p, "• " + item, size=9.7, color="202124")
        return

    heading = item.get("heading") or item.get("title") or ""
    date = item.get("date") or item.get("time") or ""
    subheading = item.get("subheading") or item.get("subtitle") or ""
    details = item.get("details") or item.get("bullets") or []
    if isinstance(details, str):
        details = [details]

    if heading or date:
        p = doc.add_paragraph()
        set_paragraph_spacing(p, before=1, after=0.5)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(18.0), WD_TAB_ALIGNMENT.RIGHT)
        add_text(p, str(heading), size=10, bold=True, color="202124")
        if date:
            add_text(p, "\t" + str(date), size=9.2, color="6B7280")

    if subheading:
        p = doc.add_paragraph()
        set_paragraph_spacing(p, after=0.5)
        add_text(p, str(subheading), size=9.5, color="4B5563")

    for detail in details:
        text = str(detail).strip()
        if not text:
            continue
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.15)
        p.paragraph_format.first_line_indent = Cm(-0.15)
        set_paragraph_spacing(p, after=0.2)
        add_text(p, "• " + text, size=9.35, color="202124")


def build_docx(data: dict[str, Any], output: Path, style: str) -> None:
    accent = ACCENTS.get(style, ACCENTS["campus"])
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.35)
    section.bottom_margin = Cm(1.25)
    section.left_margin = Cm(1.35)
    section.right_margin = Cm(1.35)

    normal = doc.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    normal.font.size = Pt(9.5)

    add_header(doc, data, accent)
    add_summary(doc, data.get("summary"))

    for section_data in sort_sections(normalize_sections(data)):
        items = section_data.get("items") or []
        if not items:
            continue
        add_section_title(doc, section_data["title"], accent)
        for item in items:
            add_item(doc, item)

    output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_json", help="Resume JSON path, or '-' for stdin.")
    parser.add_argument("output_docx", help="Output .docx path.")
    parser.add_argument("--style", choices=sorted(ACCENTS), default="campus")
    args = parser.parse_args()

    data = load_json(args.input_json)
    output = Path(args.output_docx)
    build_docx(data, output, args.style)
    print(f"wrote: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
