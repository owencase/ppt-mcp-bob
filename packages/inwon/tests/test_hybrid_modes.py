from __future__ import annotations

import platform
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Inches, Pt

from canva_ppt_mcp.com_editor import ComUnavailableError, design_fingerprint, edit_template_with_com
from canva_ppt_mcp.pipeline import create_presentation
from canva_ppt_mcp.routing import (
    confirm_presentation_mode,
    consume_mode_confirmation,
    infer_presentation_mode,
    list_user_templates,
    prepare_presentation_task,
    resolve_template,
)


def test_intent_recommendation_never_executes():
    assert infer_presentation_mode("IBM 소개 PPT 만들어줘") == "generate"
    assert infer_presentation_mode("이 디자인 그대로 내용만 IBM에 맞게 수정해줘") == "template_com"
    assert infer_presentation_mode("PPT 좀 부탁해") is None


def test_confirmation_token_is_required_and_one_time(tmp_path: Path):
    result = prepare_presentation_task("IBM 소개 PPT 만들어줘", str(tmp_path))
    assert result["requires_user_confirmation"] is True
    assert result["execution_blocked"] is True
    with pytest.raises(RuntimeError, match="MODE_CONFIRMATION_REQUIRED"):
        consume_mode_confirmation(requested_mode="generate", execution_token=None)
    confirmed = confirm_presentation_mode(result["confirmation_id"], "generate")
    consume_mode_confirmation(requested_mode="generate", execution_token=confirmed["execution_token"])
    with pytest.raises(RuntimeError, match="already-used"):
        consume_mode_confirmation(requested_mode="generate", execution_token=confirmed["execution_token"])


def test_wrong_mode_token_is_rejected(tmp_path: Path):
    result = prepare_presentation_task("템플릿 내용만 수정해줘", str(tmp_path))
    confirmed = confirm_presentation_mode(result["confirmation_id"], "template_com")
    with pytest.raises(RuntimeError, match="confirmed mode"):
        consume_mode_confirmation(requested_mode="generate", execution_token=confirmed["execution_token"])


def test_template_folder_is_the_only_source(tmp_path: Path):
    root = tmp_path / "template"; root.mkdir()
    prs = Presentation(); prs.save(root / "sample.pptx")
    assert list_user_templates(str(root))[0]["name"] == "sample.pptx"
    assert resolve_template("sample.pptx", str(root)) == (root / "sample.pptx").resolve()
    with pytest.raises(FileNotFoundError):
        resolve_template("../outside.pptx", str(root))


def test_design_fingerprint_ignores_text_content(tmp_path: Path):
    source = tmp_path / "a.pptx"
    prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
    run = box.text_frame.paragraphs[0].add_run(); run.text = "Old topic"; run.font.size = Pt(24)
    prs.save(source)
    before = design_fingerprint(source)
    prs = Presentation(source); prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].text = "New topic"
    prs.save(source)
    after = design_fingerprint(source)
    assert before == after


def _minimal_plan():
    return {
        "communication_job": "test",
        "design_system": {
            "palette": {
                "primary": "#102A43", "secondary": ["#D9EAF7"], "accent": "#00B8A9",
                "background_light": "#F7FBFF", "background_dark": "#071A2B"
            },
            "typography": {},
            "visual_motif": "test",
            "style_preset": "orbital",
            "layout_rotation": ["title", "two_column", "icon_rows", "closing"]
        },
        "slides": [
            {"title": "A", "layout": "title"},
            {"title": "B", "layout": "two_column", "items": [{"heading": "H", "body": "B"}]},
            {"title": "C", "layout": "closing"}
        ],
        "grounded": False,
        "language": "ko"
    }


def test_python_pptx_pipeline_refuses_template_path(tmp_path: Path):
    template = tmp_path / "template.pptx"; Presentation().save(template)
    with pytest.raises(ValueError, match="no longer accepted"):
        create_presentation(
            topic="test", output_path=str(tmp_path / "out.pptx"), template_path=str(template),
            content_json=_minimal_plan(), research_required=False, max_qa_rounds=1,
        )


def test_com_mode_fails_explicitly_without_windows(tmp_path: Path):
    if platform.system() == "Windows":
        pytest.skip("non-Windows behavior only")
    root = tmp_path / "template"; root.mkdir()
    prs = Presentation()
    for _ in range(3): prs.slides.add_slide(prs.slide_layouts[6])
    prs.save(root / "sample.pptx")
    with pytest.raises(ComUnavailableError, match="Windows"):
        edit_template_with_com(
            topic="test", template_name="sample.pptx", output_path=str(tmp_path / "out.pptx"),
            template_dir=str(root), research_required=False,
        )
