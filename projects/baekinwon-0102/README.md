# PowerPoint 통합 MCP

새 PPT는 `python-pptx` 기반으로 생성하고, 기존 PPT는 Microsoft PowerPoint COM으로 수정하는 MCP 서버입니다. IBM Bob처럼 도구 수에 민감한 클라이언트에서는 생성/수정 서버를 분리해 필요한 도구만 노출합니다.

## 아키텍처

```text
IBM Bob / MCP Client
  -> .bob/skills/powerpoint-routing   요청을 생성 또는 수정으로 분류
  -> powerpoint-generate             생성 도구 7개만 노출
       -> ppt-unified-mcp
       -> ppt-to-python              PPT Master + python-pptx
       -> workspace/*.pptx
  -> powerpoint-edit                 핵심 10개 + 게이트웨이 3개만 노출
       -> ppt-unified-mcp
       -> ppt-to-com                 내부 COM 기능 155개 유지
       -> 기존 .pptx/.pptm/.potx
```

| 경로 | 역할 |
| --- | --- |
| `.bob/` | Bob용 MCP 서버 설정과 생성/수정 라우팅 스킬 |
| `ppt-to-python/` | 새 PPT 생성에 필요한 PPT Master 런타임 |
| `ppt-to-com/` | 기존 PPT 수정에 필요한 PowerPoint COM 런타임 |
| `ppt-unified-mcp/` | 두 엔진을 단일 인터페이스로 연결하고 모드별 도구를 제한하는 서버 |
| `ppt-unified-mcp/workspace/` | 생성 프로젝트, 검증 결과, 내보낸 PPT 저장 위치 |

## 라우팅 원칙

| 요청 | 사용할 서버 | 조건 |
| --- | --- | --- |
| 새 프레젠테이션 생성 | `powerpoint-generate` | PowerPoint 설치 불필요 |
| 기존 파일 열기·수정 | `powerpoint-edit` | Windows와 Microsoft PowerPoint 필요 |

COM 원본의 `ppt_create_presentation`은 의도적으로 제거했습니다. 새 파일은 항상 `ppt-to-python` 경로로 생성됩니다.

## 생성 도구

`PPT_UNIFIED_MODE=generate`에서는 다음 7개만 노출됩니다.

| 도구 | 기능 |
| --- | --- |
| `ppt_generate_read_guide` | 생성 규칙과 캔버스 규격 조회 |
| `ppt_generate_create_project` | 격리된 생성 프로젝트 생성 |
| `ppt_generate_import_sources` | 로컬 파일 또는 URL 자료 등록 |
| `ppt_generate_write_slide_svg` | 슬라이드 SVG 작성 |
| `ppt_generate_get_status` | 프로젝트와 슬라이드 상태 조회 |
| `ppt_generate_validate` | SVG와 최종 산출물 품질 검사 |
| `ppt_generate_export_presentation` | 편집 가능한 PPTX로 변환 |

권장 순서는 `read_guide -> create_project -> import_sources(선택) -> write_slide_svg -> validate -> export_presentation`입니다. 검증 실패 또는 변경 후 재검증하지 않은 SVG는 내보내지 않습니다.

## 수정 도구

`PPT_UNIFIED_MODE=edit`의 기본 compact 프로필은 외부 도구를 13개만 노출하면서 기존 COM 기능 155개를 내부에 유지합니다.

직접 노출되는 핵심 도구는 파일 열기·활성화·정보 조회·슬라이드/텍스트/도형 조회·미리보기·저장·닫기 10개입니다. 나머지 기능은 다음 게이트웨이를 사용합니다.

| 도구 | 기능 |
| --- | --- |
| `ppt_edit_capabilities` | 카테고리나 검색어로 작업을 찾고 정확한 입력 스키마 조회 |
| `ppt_edit_execute` | 기존 COM 작업 하나를 원래 입력 검증과 함께 실행 |
| `ppt_edit_batch` | 최대 30개 작업을 순차 실행해 MCP 왕복 횟수 감소 |

게이트웨이를 통해 제공하는 전체 기능 범위는 다음과 같습니다.

- PowerPoint 연결, 창, 프레젠테이션 열기·활성화·저장·닫기
- 슬라이드 추가·복제·삭제·순서 변경·미리보기
- 도형, 텍스트, 서식, 정렬, 그룹, 레이어, 테마 색상
- 표, 차트, 이미지, 아이콘, 미디어, SmartArt, 애니메이션
- 노트, 섹션, 하이퍼링크, 마스터·레이아웃, PDF·이미지 내보내기

수정 흐름은 `open -> activate -> inspect -> capabilities -> execute/batch -> save_as`입니다. `capabilities`가 반환한 `input_schema`를 확인한 뒤 `arguments`에는 내부 필드만 전달합니다. Batch는 트랜잭션이 아니므로 중간 실패 전에 완료된 수정은 자동 롤백되지 않습니다. 원본 덮어쓰기가 명시되지 않았다면 다른 파일명으로 저장하는 것이 기본 원칙입니다.

기존 155개 도구를 직접 노출해야 하는 클라이언트는 `PPT_EDIT_PROFILE=full`을 사용할 수 있습니다.

## IBM Bob에서 사용

프로젝트의 `.bob/mcp.json`에는 두 서버가 등록되어 있습니다. 한 번에 하나만 활성화해야 Bob이 불필요한 도구 스키마를 읽지 않아 타임아웃 가능성이 줄어듭니다.

- 새 PPT 작업: `powerpoint-generate.disabled = false`, `powerpoint-edit.disabled = true`
- 기존 PPT 수정: 위 값을 반대로 설정
- 설정 변경 후 Bob을 재시작하거나 MCP 연결을 다시 로드
- Bob의 MCP 네트워크 타임아웃은 5분 이상 권장

라우팅 스킬은 `.bob/skills/powerpoint-routing/SKILL.md`에 있습니다. 현재 활성 서버가 요청과 다르면 어떤 서버를 켜야 하는지 안내합니다.

## 설치 및 실행

Python 3.10 이상이 필요합니다. 현재 `.venv`가 없다면 다음 명령으로 설치합니다.

```powershell
cd C:\Users\INWONBAEK\Desktop\ppt-intergration\ppt-unified-mcp
python -m venv .venv
.venv\Scripts\python -m pip install "."
```

직접 실행할 때는 모드를 지정합니다.

```powershell
$env:PPT_UNIFIED_MODE = "generate" # 또는 edit, all
.venv\Scripts\python -m ppt_unified_mcp.server
```

폴더가 현재 구조와 다를 때만 아래 환경변수를 지정합니다.

| 환경변수 | 의미 |
| --- | --- |
| `PPT_TO_PYTHON_ROOT` | `ppt-to-python` 경로 |
| `PPT_TO_COM_ROOT` | `ppt-to-com` 경로 |
| `PPT_UNIFIED_WORKSPACE` | 생성 결과 저장 경로 |
| `PPT_UNIFIED_MODE` | `generate`, `edit`, `all` 중 하나 |
| `PPT_EDIT_PROFILE` | `compact`(기본, 13개 노출) 또는 `full`(155개 노출) |
| `PPT_AUTO_DISMISS_DIALOG` | COM 작업 중 PowerPoint 대화상자 자동 닫기 여부 |

이전 이름인 `PPT_MASTER_ROOT`, `PPT_MCP_ROOT`도 호환을 위해 계속 인식합니다.

## 문제 해결

- Bob 타임아웃: 두 서버를 동시에 켜지 말고 작업 모드 하나만 활성화합니다.
- COM 연결 실패: Windows에서 PowerPoint가 설치·실행 가능한지 확인합니다.
- 생성 스크립트 없음: `PPT_TO_PYTHON_ROOT`가 `ppt-to-python`을 가리키는지 확인합니다.
- 파일 수정 대상 오류: 먼저 `ppt_activate_presentation`으로 대상 파일을 고정합니다.
- 결과 위치: 기본값은 `ppt-unified-mcp/workspace/<project>/exports/`입니다.

## 원본 및 라이선스

- 생성 엔진: [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master), MIT
- COM 엔진: [ykuwai/ppt-mcp](https://github.com/ykuwai/ppt-mcp), MIT
- 통합 고지: `ppt-unified-mcp/THIRD_PARTY_NOTICES.md`
