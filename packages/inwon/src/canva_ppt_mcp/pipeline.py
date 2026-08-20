from __future__ import annotations

import json
import os
from contextlib import contextmanager
from pathlib import Path
from urllib.parse import urlparse

from .planner import create_plan
from .qa import qa_loop
from .render import render_auto_deck, write_speaker_notes
from .semantic_qa import inspect_grounding
from .template import inspect_template


@contextmanager
def _output_lock(output: Path):
    lock = output.with_suffix(output.suffix + ".lock")
    handle = lock.open("a+b")
    acquired = False
    try:
        try:
            if os.name == "nt":
                import msvcrt
                handle.seek(0); handle.write(b"0"); handle.flush(); handle.seek(0)
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            acquired = True
        except OSError as exc:
            handle.close()
            raise RuntimeError(f"같은 출력 파일을 생성하는 작업이 이미 실행 중입니다: {output}") from exc
        handle.seek(0); handle.truncate(); handle.write(str(os.getpid()).encode("ascii")); handle.flush()
        yield
    finally:
        try:
            if not handle.closed:
                if os.name == "nt":
                    import msvcrt
                    handle.seek(0); msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
                handle.close()
        except OSError:
            pass
        if acquired:
            lock.unlink(missing_ok=True)


def _validate_inputs(topic: str, language: str, max_qa_rounds: int,
                     source_urls: list[str] | None,
                     research_documents: list[dict[str, str]] | None = None,
                     research_text: str | None = None) -> None:
    if not topic or not topic.strip():
        raise ValueError("topic must not be empty")
    if len(topic) > 180:
        raise ValueError("topic must be 180 characters or fewer")
    if language not in {"ko", "en"}:
        raise ValueError("language must be 'ko' or 'en'")
    if not 1 <= max_qa_rounds <= 6:
        raise ValueError("max_qa_rounds must be between 1 and 6")
    if len(source_urls or []) > 20:
        raise ValueError("source_urls accepts at most 20 values")
    for value in source_urls or []:
        if value == "user-provided":
            continue
        parsed = urlparse(value)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc:
            raise ValueError(f"invalid source URL: {value}")
    if len(research_documents or []) > 20:
        raise ValueError("research_documents accepts at most 20 documents")
    for index, document in enumerate(research_documents or [], 1):
        if not document.get("text", "").strip():
            raise ValueError(f"research_documents[{index}] has no text")
        if len(document.get("text", "").encode("utf-8")) > 2_000_000:
            raise ValueError(f"research_documents[{index}] exceeds 2 MB")
    if research_text and len(research_text.encode("utf-8")) > 2_000_000:
        raise ValueError("research_text exceeds 2 MB")


def _validate_output_root(output: Path) -> None:
    configured = os.getenv("PPT_MCP_OUTPUT_ROOT")
    if not configured:
        return
    root = Path(configured).expanduser().resolve()
    if output != root and root not in output.parents:
        raise ValueError(f"output_path must be inside PPT_MCP_OUTPUT_ROOT: {root}")


def create_presentation(*, topic: str, output_path: str, audience: str = "", purpose: str = "",
                        slide_count: int = 8, language: str = "ko", template_path: str | None = None,
                        content_json: dict | None = None, max_qa_rounds: int = 3,
                        research_text: str | None = None, source_urls: list[str] | None = None,
                        research_required: bool = True,
                        style_preference: str | None = None,
                        research_documents: list[dict[str, str]] | None = None) -> dict:
    _validate_inputs(topic, language, max_qa_rounds, source_urls, research_documents, research_text)
    output = Path(output_path).expanduser().resolve()
    _validate_output_root(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.suffix.lower() != ".pptx": raise ValueError("output_path must end in .pptx")
    if not 3 <= slide_count <= 30: raise ValueError("slide_count must be between 3 and 30")
    with _output_lock(output):
        plan = create_plan(topic, audience, purpose, slide_count, language, content_json,
                           research_text, source_urls, research_required, style_preference,
                           research_documents)
        qa_dir = output.with_name(output.stem + "_qa"); qa_dir.mkdir(parents=True, exist_ok=True)
        semantic_issues = inspect_grounding(plan, topic, research_required and content_json is None)
        (qa_dir / "deck-plan.json").write_text(plan.model_dump_json(indent=2), encoding="utf-8")
        (qa_dir / "semantic-report.json").write_text(json.dumps(
            {"passed": not any(x.severity == "error" for x in semantic_issues),
             "issues": [x.model_dump() for x in semantic_issues]}, ensure_ascii=False, indent=2), encoding="utf-8")
        if any(x.severity == "error" for x in semantic_issues):
            messages = "; ".join(x.message for x in semantic_issues)
            raise RuntimeError(f"Semantic QA failed: {messages}")
        if template_path:
            raise ValueError(
                "template_path is no longer accepted by the python-pptx generation pipeline. "
                "Use edit_template_with_com / template_com mode after explicit user confirmation."
            )
        manifest = {"mode": "generate", "engine": "python-pptx"}
        render_auto_deck(plan, str(output))
        write_speaker_notes(str(output), plan)
        report = qa_loop(str(output), str(qa_dir), max_rounds=max_qa_rounds, auto_fix=True)
        manifest.update({
            "output_path": str(output), "plan_path": str(qa_dir / "deck-plan.json"),
            "qa_report_path": str(qa_dir / "qa-report.json"), "qa": report.model_dump(),
            "semantic_report_path": str(qa_dir / "semantic-report.json"),
            "research_sources": [source.model_dump() for source in plan.research_sources],
            "evidence_claim_count": len(plan.evidence_claims),
        })
        (qa_dir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        if not report.passed:
            raise RuntimeError(f"QA failed; see {qa_dir / 'qa-report.json'}")
        return manifest


__all__ = ["create_presentation", "inspect_template", "qa_loop"]
