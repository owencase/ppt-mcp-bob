from __future__ import annotations

import json
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from canva_ppt_mcp.pipeline import create_presentation
from canva_ppt_mcp.routing import confirm_presentation_mode, prepare_presentation_task
from canva_ppt_mcp.template import inspect_template


def demo_plan() -> dict:
    return {
        "communication_job": "Hybrid MCP smoke test",
        "design_system": {
            "palette": {
                "primary": "#102A43", "secondary": ["#D9EAF7"], "accent": "#00B8A9",
                "background_light": "#F7FBFF", "background_dark": "#071A2B"
            },
            "typography": {},
            "visual_motif": "smoke",
            "style_preset": "orbital",
            "layout_rotation": ["title", "two_column", "icon_rows", "closing"]
        },
        "slides": [
            {"title": "Hybrid PPT MCP", "subtitle": "Smoke test", "layout": "title"},
            {"title": "python-pptx generation", "layout": "two_column", "items": [
                {"heading": "Generate", "body": "Creates a new deck from scratch."},
                {"heading": "Template", "body": "COM mode is intentionally separate."}
            ]},
            {"title": "Mode gate is required", "subtitle": "Confirm before execution", "layout": "closing"}
        ],
        "grounded": False,
        "language": "en"
    }


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="canva-ppt-smoke-") as temp:
        out = Path(temp)
        prepared = prepare_presentation_task("Create a PowerPoint from scratch", str(out / "template"))
        assert prepared["requires_user_confirmation"]
        confirmed = confirm_presentation_mode(prepared["confirmation_id"], "generate")
        assert confirmed["selected_mode"] == "generate"
        result = {"routing": "passed", "generation": "skipped (soffice not installed)"}
        if shutil.which("soffice"):
            auto = out / "auto.pptx"
            generated = create_presentation(
                topic="Hybrid PowerPoint MCP", output_path=str(auto), slide_count=3,
                language="en", content_json=demo_plan(), research_required=False,
            )
            assert generated["qa"]["passed"]
            assert inspect_template(auto)["slide_count"] == 3
            result["generation"] = "passed"
        print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
