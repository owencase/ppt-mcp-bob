# 3.3.0 변경 내역

## COM 편집 파이프라인

- fail-fast COM 텍스트 수정을 complete-first / validate-after 방식으로 변경
- 1차 편집에서 개별 shape 오류와 overflow를 예외로 중단하지 않고 기록
- 전체 슬라이드 수정 및 Save 완료 후 post-QA 실행
- 문제 shape만 대상으로 최대 0~3회(`max_post_qa_rounds`, 기본 2) 재수정
- 동일 오류 signature 반복 시 cycle breaker로 즉시 중단
- 최종 QA 미통과 시에도 결과 파일과 manifest 반환
- 전체 자동 재생성/자동 재시작 금지 (`automatic_restart_blocked=true`)
- 최종 상태: `completed`, `completed_with_warnings`, `completed_with_unresolved_issues`

## 사용자 화면

- PowerPoint Watch Mode 유지
- 편집 중 PowerPoint modal alert는 억제하고 오류를 post-QA에 수집
- 수정 완료 후 결과 프레젠테이션은 계속 열린 상태 유지

## QA 산출물

- `template-com-validation.json` 추가
- first pass issues, repair rounds, remaining issues, cycle breaker, design audit 기록
- pristine working baseline을 `<output>_qa/template-baseline.*`에 보관

## CLI / MCP

- MCP `edit_template_presentation(..., max_post_qa_rounds=2)` 추가
- CLI `--post-qa-rounds {0,1,2,3}` 추가
