from __future__ import annotations

import json
from pathlib import Path

from docx import Document

from scripts.build_resume_docx import STYLE_PRESETS, build_docx, load_json
from scripts import verify_resume_docx
from scripts.verify_resume_docx import extract_docx_text


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "examples" / "sample_resume.json"


def test_sample_json_builds_docx(tmp_path: Path) -> None:
    data = load_json(str(SAMPLE))
    output = tmp_path / "sample.docx"

    build_docx(data, output, "campus")

    assert output.exists()
    text = extract_docx_text(output)
    assert "示例姓名" in text
    assert "教育背景" in text
    assert "项目经历" in text
    assert "Amazon" in text
    assert "Listing" in text


def test_all_style_presets_build(tmp_path: Path) -> None:
    data = json.loads(SAMPLE.read_text(encoding="utf-8-sig"))

    for style_name in STYLE_PRESETS:
        output = tmp_path / f"{style_name}.docx"
        build_docx(data, output, style_name)
        assert output.exists()

        document = Document(output)
        assert document.paragraphs[0].text == "示例姓名"


def test_load_json_accepts_utf8_bom(tmp_path: Path) -> None:
    path = tmp_path / "resume.json"
    path.write_text("\ufeff" + SAMPLE.read_text(encoding="utf-8"), encoding="utf-8")

    data = load_json(str(path))

    assert data["name"] == "示例姓名"


def test_verify_docx_reports_unconfirmed_page_count_without_word(tmp_path: Path, monkeypatch) -> None:
    data = load_json(str(SAMPLE))
    output = tmp_path / "sample.docx"
    build_docx(data, output, "campus")

    def no_word_check(path: Path, export_pdf: Path | None = None) -> dict:
        raise RuntimeError("Word is not available in this test")

    monkeypatch.setattr(verify_resume_docx, "word_check", no_word_check)
    result, exit_code = verify_resume_docx.verify_docx(
        output,
        expect_one_page=True,
        required_sections=["教育背景", "项目经历", "实习经历", "技能证书"],
        required_keywords=["Amazon", "Listing"],
    )

    assert exit_code == 0
    assert result["ok"] is True
    assert result["page_check"] == "not_confirmed_without_word"
    assert result["pages"] is None


def test_verify_docx_missing_keyword_fails(tmp_path: Path, monkeypatch) -> None:
    data = load_json(str(SAMPLE))
    output = tmp_path / "sample.docx"
    build_docx(data, output, "campus")

    monkeypatch.setattr(verify_resume_docx, "word_check", lambda path, export_pdf=None: {})
    result, exit_code = verify_resume_docx.verify_docx(output, required_keywords=["不存在的关键词"])

    assert exit_code == 1
    assert result["ok"] is False
    assert result["missing_keywords"] == ["不存在的关键词"]


def test_verify_docx_require_word_fails_when_word_unavailable(tmp_path: Path, monkeypatch) -> None:
    data = load_json(str(SAMPLE))
    output = tmp_path / "sample.docx"
    build_docx(data, output, "campus")

    def no_word_check(path: Path, export_pdf: Path | None = None) -> dict:
        raise RuntimeError("Word is not available in this test")

    monkeypatch.setattr(verify_resume_docx, "word_check", no_word_check)
    result, exit_code = verify_resume_docx.verify_docx(output, expect_one_page=True, require_word=True)

    assert exit_code == 1
    assert result["ok"] is False
    assert "cannot_confirm_one_page_without_word" in result["failures"]
