from __future__ import annotations

import argparse
import json
from pathlib import Path

from .com_editor import edit_template_with_com
from .pipeline import create_presentation
from .qa import qa_loop
from .routing import infer_presentation_mode, list_user_templates
from .template import inspect_template


def _confirm_cli(expected: str) -> None:
    print("PPT 작업 방식을 반드시 확인합니다.")
    print("  1) 처음부터 새로 생성 (python-pptx)")
    print("  2) /template의 기존 디자인을 그대로 두고 내용만 수정 (PowerPoint COM)")
    answer = input("선택 [1/2]: ").strip().lower()
    selected = "generate" if answer in {"1", "generate", "생성"} else (
        "template_com" if answer in {"2", "template", "template_com", "수정"} else None
    )
    if selected != expected:
        raise SystemExit(f"선택한 모드({selected or '알 수 없음'})와 실행 명령({expected})이 다릅니다. 작업을 중단합니다.")


def main() -> None:
    parser = argparse.ArgumentParser(prog="canva-ppt")
    sub = parser.add_subparsers(dest="command", required=True)

    route = sub.add_parser("route", help="요청 문장에서 추천 모드만 판별하며 실행하지 않습니다")
    route.add_argument("--request", required=True)
    route.add_argument("--template-dir")

    create = sub.add_parser("create", help="처음부터 python-pptx로 생성")
    create.add_argument("--topic", required=True); create.add_argument("--output", required=True)
    create.add_argument("--audience", default=""); create.add_argument("--purpose", default="")
    create.add_argument("--slides", type=int, default=8); create.add_argument("--language", default="ko")
    create.add_argument("--content-json", help="DeckPlan JSON path; bypasses the LLM planner")
    create.add_argument("--research-text", help="UTF-8 text file used as grounded source material")
    create.add_argument("--research-documents", help="JSON list of {title,url,text} research documents")
    create.add_argument("--source-url", action="append", default=[])
    create.add_argument("--allow-generic-fallback", action="store_true")
    create.add_argument("--style", choices=["orbital", "editorial", "neon", "organic", "luxury", "geometric", "swiss"])
    create.add_argument("--yes-mode-confirmed", action="store_true", help="generate 모드를 사용자가 이미 명시적으로 확인함")

    edit = sub.add_parser("edit-template", help="/template 파일의 내용만 PowerPoint COM으로 수정")
    edit.add_argument("--topic", required=True); edit.add_argument("--template", required=True); edit.add_argument("--output", required=True)
    edit.add_argument("--template-dir"); edit.add_argument("--audience", default=""); edit.add_argument("--purpose", default="")
    edit.add_argument("--language", default="ko"); edit.add_argument("--research-text")
    edit.add_argument("--research-documents"); edit.add_argument("--source-url", action="append", default=[])
    edit.add_argument("--allow-generic-fallback", action="store_true"); edit.add_argument("--visible", action="store_true")
    edit.add_argument("--yes-mode-confirmed", action="store_true", help="template_com 모드를 사용자가 이미 명시적으로 확인함")

    inspect = sub.add_parser("inspect-template"); inspect.add_argument("path")
    templates = sub.add_parser("list-templates"); templates.add_argument("--template-dir")
    qa = sub.add_parser("qa"); qa.add_argument("path"); qa.add_argument("--fix", action="store_true")

    args = parser.parse_args()
    if args.command == "route":
        values = list_user_templates(args.template_dir)
        result = {
            "suggested_mode": infer_presentation_mode(args.request),
            "confirmation_required": True,
            "available_templates": values,
        }
    elif args.command == "create":
        if not args.yes_mode_confirmed:
            _confirm_cli("generate")
        content = json.loads(Path(args.content_json).read_text(encoding="utf-8")) if args.content_json else None
        research_text = Path(args.research_text).read_text(encoding="utf-8") if args.research_text else None
        research_documents = json.loads(Path(args.research_documents).read_text(encoding="utf-8")) if args.research_documents else None
        result = create_presentation(
            topic=args.topic, output_path=args.output, audience=args.audience, purpose=args.purpose,
            slide_count=args.slides, language=args.language, content_json=content,
            research_text=research_text, source_urls=args.source_url,
            research_required=not args.allow_generic_fallback, style_preference=args.style,
            research_documents=research_documents, template_path=None,
        )
    elif args.command == "edit-template":
        if not args.yes_mode_confirmed:
            _confirm_cli("template_com")
        research_text = Path(args.research_text).read_text(encoding="utf-8") if args.research_text else None
        research_documents = json.loads(Path(args.research_documents).read_text(encoding="utf-8")) if args.research_documents else None
        result = edit_template_with_com(
            topic=args.topic, template_name=args.template, output_path=args.output,
            audience=args.audience, purpose=args.purpose, language=args.language,
            research_text=research_text, source_urls=args.source_url,
            research_required=not args.allow_generic_fallback,
            research_documents=research_documents, template_dir=args.template_dir,
            visible=args.visible,
        )
    elif args.command == "inspect-template":
        result = inspect_template(args.path)
    elif args.command == "list-templates":
        result = list_user_templates(args.template_dir)
    else:
        result = qa_loop(args.path, args.path.rsplit(".", 1)[0] + "_qa", 3 if args.fix else 1, args.fix).model_dump()
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
