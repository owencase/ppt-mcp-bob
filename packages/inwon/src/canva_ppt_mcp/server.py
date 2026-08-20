from __future__ import annotations

from typing import Any, Literal

try:  # MCP Python SDK 2.x
    from mcp.server.mcpserver import MCPServer
except ImportError:  # MCP Python SDK 1.x compatibility
    from mcp.server.fastmcp import FastMCP as MCPServer

from .com_editor import edit_template_with_com as run_edit_template_with_com
from .pipeline import create_presentation as run_create_presentation
from .qa import qa_loop
from .routing import (
    confirm_presentation_mode as run_confirm_presentation_mode,
    consume_mode_confirmation,
    list_user_templates,
    prepare_presentation_task as run_prepare_presentation_task,
)
from .template import inspect_template as run_inspect_template


INSTRUCTIONS = """PowerPoint 작업은 반드시 MODE GATE를 먼저 통과한다.

[강제 순서]
1) 사용자가 PPT 생성/수정을 요청하면 어떤 실행 도구도 먼저 호출하지 말고 prepare_presentation_task(user_request)를 호출한다.
2) prepare 결과의 question을 사용자에게 그대로 또는 의미를 유지해 질문한다. 반드시 사용자가 둘 중 하나를 직접 선택해야 한다:
   - generate: 처음부터 새 PPT 생성 -> python-pptx
   - template_com: /template 폴더의 기존 템플릿 디자인은 그대로 두고 내용만 수정 -> Windows PowerPoint COM
3) 사용자의 답변을 받은 뒤에만 confirm_presentation_mode(confirmation_id, selected_mode)를 호출한다.
4) confirm 결과의 one-time execution_token을 해당 실행 도구에 전달한다. 토큰이 없거나 모드가 다르면 실행은 실패해야 한다.

[의도 추천 규칙]
- '~ 생성해줘', '~ 만들어줘', '처음부터 만들어줘'는 generate를 추천한다. 단, 추천만 하고 사용자 확인을 생략하지 않는다.
- '~ 템플릿의 내용만 수정해줘', '이 디자인 그대로 내용만 ~에 맞게 수정해줘', '기존 PPT 내용만 바꿔줘'는 template_com을 추천한다. 단, 추천만 하고 사용자 확인을 생략하지 않는다.
- 요청이 명확해 보여도 반드시 사용자에게 두 모드 중 하나를 확인한다.

[generate]
- create_presentation을 사용한다.
- python-pptx 기반 Research -> DeckPlan -> Semantic QA -> Render QA 파이프라인을 사용한다.
- template_path를 전달하지 않는다. 템플릿 작업을 python-pptx로 우회하지 않는다.

[template_com]
- /template 폴더의 파일만 대상으로 edit_template_presentation을 사용한다.
- 원본 템플릿을 덮어쓰지 않고 output_path에 복사/생성한 뒤 PowerPoint COM으로 텍스트 내용만 바꾼다.
- 슬라이드 수, 도형 수, 위치, 크기, 회전, 색, 선, 이미지/차트 등의 디자인 구조는 변경하지 않는다.
- 활성 창을 암묵적으로 수정하지 않고 명시적으로 연 출력 프레젠테이션 핸들을 타깃으로 사용한다.
- COM 호출이 RPC_E_CALL_REJECTED / SERVERCALL_RETRYLATER로 실패하면 제한된 횟수만 재시도한다.
- Windows + Microsoft PowerPoint + pywin32가 없으면 template_com을 실행할 수 없다고 명확히 보고한다.

생성 성공은 해당 모드의 검증을 통과한 경우에만 보고한다. 한국어 덱에는 영문 고정 UI 라벨을 남기지 않는다."""

try:
    mcp = MCPServer("canva-ppt", version="3.0.0", instructions=INSTRUCTIONS)
except TypeError:
    mcp = MCPServer("canva-ppt", instructions=INSTRUCTIONS)


@mcp.tool()
def prepare_presentation_task(user_request: str, template_dir: str | None = None) -> dict[str, Any]:
    """PPT 작업 전 필수 단계. 의도를 추천하되 실행하지 않고 사용자 확인 질문과 confirmation_id를 반환합니다."""
    return run_prepare_presentation_task(user_request, template_dir)


@mcp.tool()
def confirm_presentation_mode(
    confirmation_id: str,
    selected_mode: Literal["generate", "template_com"],
) -> dict[str, Any]:
    """사용자가 모드를 직접 선택한 뒤에만 호출. 1회용 execution_token을 발급합니다."""
    return run_confirm_presentation_mode(confirmation_id, selected_mode)


@mcp.tool()
def list_templates(template_dir: str | None = None) -> list[dict[str, str]]:
    """설정된 /template 폴더에서 COM 수정에 사용할 PowerPoint 템플릿을 나열합니다."""
    return list_user_templates(template_dir)


@mcp.tool()
def create_presentation(
    topic: str,
    output_path: str,
    execution_token: str | None = None,
    audience: str = "",
    purpose: str = "",
    slide_count: int = 8,
    language: str = "ko",
    content_json: dict[str, Any] | None = None,
    max_qa_rounds: int = 3,
    research_text: str | None = None,
    source_urls: list[str] | None = None,
    research_required: bool = True,
    style_preference: str | None = None,
    research_documents: list[dict[str, str]] | None = None,
) -> dict[str, Any]:
    """확인된 generate 모드에서만 python-pptx로 새 PPTX를 생성합니다."""
    consume_mode_confirmation(requested_mode="generate", execution_token=execution_token)
    return run_create_presentation(
        topic=topic,
        output_path=output_path,
        audience=audience,
        purpose=purpose,
        slide_count=slide_count,
        language=language,
        template_path=None,
        content_json=content_json,
        max_qa_rounds=max_qa_rounds,
        research_text=research_text,
        source_urls=source_urls,
        research_required=research_required,
        style_preference=style_preference,
        research_documents=research_documents,
    )


@mcp.tool()
def edit_template_presentation(
    topic: str,
    template_name: str,
    output_path: str,
    execution_token: str | None = None,
    audience: str = "",
    purpose: str = "",
    language: str = "ko",
    research_text: str | None = None,
    source_urls: list[str] | None = None,
    research_required: bool = True,
    research_documents: list[dict[str, str]] | None = None,
    template_dir: str | None = None,
    visible: bool = False,
) -> dict[str, Any]:
    """확인된 template_com 모드에서 /template 파일의 디자인을 보존하고 텍스트만 COM으로 수정합니다."""
    consume_mode_confirmation(requested_mode="template_com", execution_token=execution_token)
    return run_edit_template_with_com(
        topic=topic,
        template_name=template_name,
        output_path=output_path,
        audience=audience,
        purpose=purpose,
        language=language,
        research_text=research_text,
        source_urls=source_urls,
        research_required=research_required,
        research_documents=research_documents,
        template_dir=template_dir,
        visible=visible,
    )


@mcp.tool()
def inspect_template(template_path: str) -> dict[str, Any]:
    """PPTX/POTX의 슬라이드 유형, 색상, 폰트, 슬롯을 읽기 전용으로 분석합니다."""
    return run_inspect_template(template_path)


@mcp.tool()
def qa_presentation(pptx_path: str, max_rounds: int = 1, auto_fix: bool = False) -> dict[str, Any]:
    """기존 PPTX를 렌더링하고 오버플로, 겹침, 경계 이탈, placeholder를 검사합니다."""
    path = __import__("pathlib").Path(pptx_path)
    qa_dir = str(path.with_name(path.stem + "_qa"))
    return qa_loop(pptx_path, qa_dir, max_rounds, auto_fix).model_dump()


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
