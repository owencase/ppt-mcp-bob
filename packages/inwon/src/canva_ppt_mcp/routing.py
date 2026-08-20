from __future__ import annotations

import re
import secrets
import time
from pathlib import Path
from typing import Literal

PresentationMode = Literal["generate", "template_com"]

_PENDING_CONFIRMATIONS: dict[str, dict] = {}
_EXECUTION_TOKENS: dict[str, dict] = {}
_TOKEN_TTL_SECONDS = 1800

GENERATE_PATTERNS = (
    r"(?:ppt|파워포인트|프레젠테이션).*(?:생성|만들|제작)",
    r"(?:생성|만들|제작).*(?:ppt|파워포인트|프레젠테이션)",
    r"처음부터.*(?:생성|만들|제작)",
    r"새(?:로|로운).*(?:ppt|파워포인트|프레젠테이션)",
)
TEMPLATE_PATTERNS = (
    r"템플릿.*(?:내용|텍스트|문구).*(?:수정|변경|바꿔|교체)",
    r"디자인.*그대로.*(?:내용|텍스트|문구)",
    r"내용만.*(?:수정|변경|바꿔|교체)",
    r"기존.*(?:ppt|파워포인트|프레젠테이션).*(?:수정|변경)",
    r"template.*(?:content|text).*(?:edit|replace|change)",
    r"keep.*design.*(?:content|text)",
)


def infer_presentation_mode(user_request: str) -> PresentationMode | None:
    """Infer a suggested mode. This never counts as user confirmation."""
    text = " ".join((user_request or "").split()).lower()
    if not text:
        return None
    template_score = sum(bool(re.search(pattern, text, re.I)) for pattern in TEMPLATE_PATTERNS)
    generate_score = sum(bool(re.search(pattern, text, re.I)) for pattern in GENERATE_PATTERNS)
    if template_score > generate_score:
        return "template_com"
    if generate_score > template_score:
        return "generate"
    return None


def template_root(configured: str | None = None) -> Path:
    if configured:
        return Path(configured).expanduser().resolve()
    # package is src/canva_ppt_mcp -> project root is parents[2]
    return Path(__file__).resolve().parents[2] / "template"


def list_user_templates(configured: str | None = None) -> list[dict[str, str]]:
    root = template_root(configured)
    root.mkdir(parents=True, exist_ok=True)
    values: list[dict[str, str]] = []
    for path in sorted(root.iterdir(), key=lambda p: p.name.lower()):
        if path.is_file() and path.suffix.lower() in {".pptx", ".pptm", ".potx", ".potm"}:
            values.append({"name": path.name, "path": str(path.resolve()), "stem": path.stem})
    return values


def resolve_template(template_name: str, configured: str | None = None) -> Path:
    if not template_name or not template_name.strip():
        raise ValueError("template_name is required for template_com mode")
    root = template_root(configured)
    root.mkdir(parents=True, exist_ok=True)
    requested = Path(template_name)
    # Security: template mode only accepts files from the configured /template folder.
    candidate = (root / requested.name).resolve()
    if root != candidate.parent:
        raise ValueError("template must be selected from the configured /template folder")
    if not candidate.is_file():
        by_stem = [p for p in root.iterdir() if p.is_file() and p.stem.lower() == requested.stem.lower()]
        if len(by_stem) == 1:
            candidate = by_stem[0].resolve()
        else:
            raise FileNotFoundError(f"template not found in {root}: {template_name}")
    if candidate.suffix.lower() not in {".pptx", ".pptm", ".potx", ".potm"}:
        raise ValueError("template must be .pptx, .pptm, .potx, or .potm")
    return candidate


def confirmation_question(suggested_mode: PresentationMode | None, templates: list[dict[str, str]]) -> str:
    if suggested_mode == "template_com":
        prefix = "요청은 기존 템플릿의 디자인을 유지하고 내용만 수정하는 작업으로 보입니다."
    elif suggested_mode == "generate":
        prefix = "요청은 처음부터 새 PPT를 생성하는 작업으로 보입니다."
    else:
        prefix = "PPT 작업 방식을 먼저 선택해야 합니다."
    template_hint = ""
    if templates:
        template_hint = " /template 폴더의 템플릿: " + ", ".join(x["name"] for x in templates[:8])
    return (
        f"{prefix} 진행 방식을 확인해주세요: "
        "① 처음부터 새로 생성(python-pptx) / "
        "② 기존 /template 템플릿의 디자인은 그대로 두고 내용만 수정(COM)."
        f"{template_hint}"
    )


def _purge_expired() -> None:
    now = time.time()
    for store in (_PENDING_CONFIRMATIONS, _EXECUTION_TOKENS):
        for key, value in list(store.items()):
            if now - value["created_at"] > _TOKEN_TTL_SECONDS:
                store.pop(key, None)


def prepare_presentation_task(user_request: str, template_dir: str | None = None) -> dict:
    _purge_expired()
    templates = list_user_templates(template_dir)
    suggested = infer_presentation_mode(user_request)
    confirmation_id = secrets.token_urlsafe(18)
    _PENDING_CONFIRMATIONS[confirmation_id] = {
        "created_at": time.time(),
        "user_request": user_request,
        "suggested_mode": suggested,
    }
    return {
        "status": "confirmation_required",
        "confirmation_id": confirmation_id,
        "suggested_mode": suggested,
        "requires_user_confirmation": True,
        "question": confirmation_question(suggested, templates),
        "available_templates": templates,
        "execution_blocked": True,
    }


def confirm_presentation_mode(confirmation_id: str, selected_mode: PresentationMode) -> dict:
    """Mint a one-time execution token after the assistant has received the user's choice."""
    _purge_expired()
    pending = _PENDING_CONFIRMATIONS.pop(confirmation_id, None)
    if not pending:
        raise RuntimeError("invalid or expired confirmation_id; ask the user for the mode again")
    if selected_mode not in {"generate", "template_com"}:
        raise ValueError("selected_mode must be 'generate' or 'template_com'")
    token = secrets.token_urlsafe(24)
    _EXECUTION_TOKENS[token] = {
        "created_at": time.time(),
        "mode": selected_mode,
        "user_request": pending["user_request"],
    }
    return {
        "status": "confirmed",
        "selected_mode": selected_mode,
        "execution_token": token,
        "one_time": True,
    }


def consume_mode_confirmation(*, requested_mode: PresentationMode, execution_token: str | None) -> dict:
    _purge_expired()
    if not execution_token:
        raise RuntimeError(
            "MODE_CONFIRMATION_REQUIRED: 먼저 prepare_presentation_task로 질문을 만들고, "
            "사용자 답변을 받은 뒤 confirm_presentation_mode를 호출해야 합니다."
        )
    confirmed = _EXECUTION_TOKENS.pop(execution_token, None)
    if not confirmed:
        raise RuntimeError("invalid, expired, or already-used execution_token")
    if confirmed["mode"] != requested_mode:
        raise RuntimeError(
            f"confirmed mode is {confirmed['mode']}, but requested operation is {requested_mode}; "
            "ask the user for the mode again"
        )
    return confirmed
