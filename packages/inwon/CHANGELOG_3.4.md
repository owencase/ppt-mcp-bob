# 3.4.0 변경 내역

## Bob/MCP 오류 루프 방지

- `edit_template_presentation`의 최종 MCP 경계에서 예상 가능한 운영 오류를 예외가 아닌 정상 JSON으로 반환합니다.
- 반환값에 `tool_call_succeeded=true`, `mcp_transport_error=false`, `do_not_retry=true`, `automatic_restart_blocked=true`를 추가했습니다.
- mode gate 토큰이 이미 사용되었거나 잘못된 경우에도 Bob에서 tool error를 만들지 않고 재시도 금지 상태로 반환합니다.
- 동일 작업을 자동으로 처음부터 다시 실행하지 않도록 Skill 규칙을 강화했습니다.

## Complete-first COM 편집 강화

- 개별 텍스트 박스 오류뿐 아니라 슬라이드 접근/슬롯 수집 오류를 해당 슬라이드 오류로 기록하고 다음 슬라이드로 진행합니다.
- 1차 저장과 repair 저장 오류를 `SAVE_ERROR`로 누적하고 가능한 범위에서 후속 QA를 계속합니다.
- post-QA 자체 오류는 `QA_RUNTIME_ERROR`, repair 런타임 오류는 `QA_REPAIR_RUNTIME_ERROR`로 기록합니다.
- PowerPoint가 열린 이후 세션 오류가 발생하면 가능한 경우 현재 상태를 저장하고 `interrupted_with_partial_result`로 반환합니다.
- semantic QA는 COM 작업 전 중단 조건이 아니라 최종 검증 문제로 이동했습니다.

## 유지되는 기능

- PowerPoint 실시간 표시/슬라이드 이동/텍스트 상자 선택/step delay/완료 후 창 유지
- AutoSize 강제 축소 금지
- 전체 편집 후 post-QA
- 문제 슬라이드만 제한적으로 재수정
- 동일 오류 signature cycle breaker
- 디자인 구조 보존 검사

## 검증

- pytest: 24 passed
- `scripts/run_tests.py`: passed
- `scripts/smoke_test.py`: routing + generation passed
- 실제 Windows PowerPoint COM GUI E2E는 실행 환경상 별도 Windows 검증이 필요합니다.
