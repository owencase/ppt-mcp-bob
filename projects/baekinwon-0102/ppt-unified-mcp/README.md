# PPT Unified MCP

`ppt-to-python`과 `ppt-to-com`을 연결하는 MCP 진입점입니다.

- `generate`: 새 PPT용 `ppt_generate_*` 도구 7개
- `edit`: 핵심 도구 10개와 동적 게이트웨이 3개로 COM 기능 155개 제공
- `all`: 두 도구군 모두 노출

실행 모듈은 `ppt_unified_mcp.server`이며, `PPT_UNIFIED_MODE`로 모드를 선택합니다. `edit` 모드는 기본적으로 `PPT_EDIT_PROFILE=compact`이고, 기존 도구 155개를 직접 노출하려면 `full`로 변경합니다. 기본 생성 작업 경로는 이 디렉터리의 `workspace/`입니다.

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install "."
$env:PPT_UNIFIED_MODE = "generate"
.venv\Scripts\python -m ppt_unified_mcp.server
```

전체 아키텍처, 도구, IBM Bob 설정과 문제 해결은 상위 [`README.md`](../README.md)를 참고하십시오.
