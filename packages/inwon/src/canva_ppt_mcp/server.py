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
- COM 수정은 항상 실제 PowerPoint 창을 화면에 띄운 상태로 수행한다. 숨김/백그라운드 수정은 금지한다.
- 각 슬라이드로 이동하고 수정할 텍스트 상자를 선택한 뒤 내용을 바꾸며, 사용자가 변화를 볼 수 있도록 단계별 지연을 둔다.
- 텍스트 교체 시 AutoSize=2를 강제로 설정하지 않는다. 원래 템플릿의 AutoSize와 글꼴/크기/굵기/색을 보존한다.
- 1차 COM 편집 중에는 텍스트 오버플로/개별 텍스트 박스 오류로 작업을 중단하지 않는다. 가능한 모든 슬라이드를 먼저 끝까지 수정하고 저장한다.
- 전체 1차 편집이 끝난 뒤에만 post-QA를 수행한다. 문제가 있는 슬라이드/텍스트 박스만 최대 max_post_qa_rounds 범위에서 다시 수정한다.
- 같은 오류 signature가 반복되면 즉시 cycle breaker를 작동시켜 재수정을 멈춘다. 전체 파일을 처음부터 자동 재생성하지 않는다.
- 최종 QA에 미해결 문제가 남아도 결과 파일과 manifest를 반환하고 automatic_restart_blocked=true로 보고한다. 같은 작업을 자동으로 처음부터 다시 실행하지 않는다.
- Bob/MCP 클라이언트에서 COM 관련 운영 오류는 tool exception으로 밖에 던지지 않는다. 정상 JSON 응답에 completion_status, error_message, do_not_retry=true를 담고 종료한다.
- edit_template_presentation 결과에 do_not_retry=true가 있으면 같은 execution_token 또는 같은 작업을 자동 재호출하지 않는다. 사용자에게 현재 결과와 로그만 보고한다.
- post-QA에서 텍스트가 기존 상자를 넘으면 먼저 같은 사실만 더 짧게 재작성하고, 그래도 넘을 때만 최대 12.5%/4pt 범위에서 제한적으로 축소한다. 도형 크기/위치는 바꾸지 않는다.
- 디자인 QA는 PowerPoint의 theme↔RGB, implicit↔explicit run 정규화를 오류로 보지 않는다. 실제 도형/위치/크기/회전/fill/line/큰 타이포그래피 변화만 거부한다.
- 완료 후 결과 프레젠테이션을 PowerPoint에 열린 상태로 유지한다.
- COM 호출이 RPC_E_CALL_REJECTED / SERVERCALL_RETRYLATER로 실패하면 제한된 횟수만 재시도한다.
- Windows + Microsoft PowerPoint + pywin32가 없으면 template_com을 실행할 수 없다고 명확히 보고한다.

generate 모드는 검증 통과 시에만 성공으로 보고한다. template_com은 최종 QA가 실패해도 전체 자동 재시작을 하지 말고 completion_status와 post_validation을 사용자에게 보고한다. 한국어 덱에는 영문 고정 UI 라벨을 남기지 않는다."""

try:
    mcp = MCPServer("canva-ppt", version="3.4.0", instructions=INSTRUCTIONS)
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


def _normal_tool_result(payload: dict[str, Any], *, operation_completed: bool | None = None) -> dict[str, Any]:
    """Mark a domain failure as a normal MCP response so clients such as Bob do not retry the tool call."""
    result = dict(payload)
    result.setdefault("tool_call_succeeded", True)
    result.setdefault("mcp_transport_error", False)
    result.setdefault("do_not_retry", True)
    result.setdefault("automatic_restart_blocked", True)
    if operation_completed is not None:
        result.setdefault("operation_completed", operation_completed)
    return result


def _normal_tool_interruption(stage: str, exc: Exception, *, output_path: str | None = None) -> dict[str, Any]:
    return _normal_tool_result({
        "mode": "template_com",
        "completion_status": "interrupted_without_restart",
        "passed": False,
        "requires_manual_review": True,
        "stage": stage,
        "error_type": type(exc).__name__,
        "error_message": str(exc),
        "output_path": output_path,
        "retry_policy": "do_not_automatically_retry_same_tool_call",
        "user_message": (
            "COM 작업 중 문제가 발생했지만 MCP tool call은 정상 종료했습니다. "
            "같은 작업을 자동으로 처음부터 다시 실행하지 말고 현재 결과/로그를 확인하세요."
        ),
    }, operation_completed=False)


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
    step_delay: float = 0.55,
    max_post_qa_rounds: int = 2,
) -> dict[str, Any]:
    """Visible COM edit. Expected operational failures are returned as JSON, never as MCP tool errors."""
    try:
        consume_mode_confirmation(requested_mode="template_com", execution_token=execution_token)
    except Exception as exc:
        # A repeated Bob tool call with an already-consumed token must not become
        # another MCP exception/retry loop. Return a normal blocked result instead.
        return _normal_tool_interruption("mode_gate", exc, output_path=output_path)
    try:
        result = run_edit_template_with_com(
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
            step_delay=step_delay,
            max_post_qa_rounds=max_post_qa_rounds,
        )
        return _normal_tool_result(result, operation_completed=result.get("completion_status", "").startswith("completed"))
    except Exception as exc:
        # Final safety boundary for MCP clients: COM/setup/save/QA exceptions are
        # data in the response, not transport-level tool failures.
        return _normal_tool_interruption("template_com", exc, output_path=output_path)


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
