# IBM Bob PowerPoint MCP

> **Windows COM 자동화를 통한 IBM Bob 전용 안전한 실시간 PowerPoint 편집 서버**

<p align="center">
  <a href="README.md">English</a>
</p>

<p align="center">
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.10%2B-blue.svg" alt="Python 3.10+"></a>
  <img src="https://img.shields.io/badge/Platform-Windows-0078d4.svg" alt="Windows">
  <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="MIT License">
  <img src="https://img.shields.io/badge/IBM%20Bob-ready-052FAD.svg" alt="IBM Bob ready">
  <img src="https://img.shields.io/badge/version-1.7.0-lightgrey.svg" alt="v1.7.0">
  <img src="https://img.shields.io/badge/MCP-165%20tools-blueviolet.svg" alt="165 tools">
</p>

IBM Bob이 실행 중인 Microsoft PowerPoint를 직접 확인하고 편집할 수 있도록 하는 Model Context Protocol(MCP) 서버입니다. Windows COM을 통해 데스크톱 PowerPoint 애플리케이션과 직접 통신하므로 변경 사항이 즉시 화면에 반영되며, PowerPoint 네이티브 오브젝트, 레이아웃, 테마, 애니메이션, 미디어를 그대로 유지합니다.

`python-pptx`와 같은 파일 기반 라이브러리와 달리, 이 서버는 PowerPoint에서 이미 열려 있는 프레젠테이션을 직접 제어합니다. 대상 잠금, 도형 상태 사전 조건, 유효성 검사, 재시도 가능 여부를 포함한 오류를 통해 에이전트 기반 편집을 더 안전하게 만듭니다.

또한 이 프로젝트는 [owencase/ppt-skills](https://github.com/owencase/ppt-skills)의 Bob 스킬을 활용하여 IBM Bob이 템플릿을 인식하고 일관된 디자인으로 프레젠테이션을 작성·편집할 수 있도록 안내합니다.

## 주요 특징

- **IBM Bob 전용 설계** — 도구 지침과 가드레일이 Bob을 확인 → 편집 → 차이 비교 → 검증 워크플로우로 안내합니다.
- **실시간 COM 자동화** — 파일 모델을 재구성하지 않고 실행 중인 PowerPoint를 직접 제어합니다.
- **28개 카테고리, 165개 도구** — 프레젠테이션, 슬라이드, 도형, 텍스트, 표, 차트, 테마, 애니메이션, 미디어, 디자인 참조, 내보내기 등을 지원합니다.
- **페일-클로즈드 대상 관리** — 대상이 닫히거나 없을 경우, 다른 프레젠테이션으로 자동 전환되지 않습니다.
- **사전 조건 편집** — 변환, 삭제, 시각 교체 도구가 변경 전 확인된 도형 상태를 검증합니다.
- **재시도 인식 오류** — MCP 오류에 `retryable` 플래그와 수정 `hint`가 포함되어 위험한 맹목적 재시도를 줄입니다.
- **출력 경로 제한** — 다른 이름으로 저장 및 내보내기 작업은 명시된 신뢰 디렉터리 내에서만 수행됩니다.
- **네이티브 유효성 검사** — Bob은 저장 전에 도형 스냅샷을 비교하고 결과를 검증할 수 있습니다.

## 요구 사항

| 구성 요소 | 최소 버전 |
|---|---|
| Windows | 10 또는 11 |
| Microsoft PowerPoint | 데스크톱 버전 (최신 버전 권장) |
| Python | 3.10 |
| [`uv`](https://docs.astral.sh/uv/getting-started/installation/) | 최신 안정 버전 |

PowerPoint COM 자동화는 Windows 전용입니다. Python 테스트 스위트는 PowerPoint를 실행하지 않고도 대부분의 유효성 검사 및 정책 로직을 실행할 수 있지만, 실제 프레젠테이션 편집에는 Windows에 설치된 PowerPoint가 필요합니다.

## 설치

`projects/owencase`에서 다음을 실행합니다:

```powershell
uv sync --frozen
uv run ppt-mcp-bob
```

서버는 MCP 통신에 stdio를 사용합니다. 직접 실행하면 MCP 클라이언트를 기다리는 상태가 됩니다. 일반적으로 IBM Bob이 MCP 설정을 통해 서버를 시작합니다.

## IBM Bob에 등록

Bob이 이 로컬 소스 트리를 실행하도록 구성합니다. 예시 경로를 Windows 머신의 절대 경로로 교체하세요:

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "C:\\Users\\YOUR_NAME\\.local\\bin\\uv.exe",
      "args": [
        "--directory",
        "C:\\ABSOLUTE\\PATH\\TO\\ppt-mcp-bob\\projects\\owencase",
        "run",
        "ppt-mcp-bob"
      ],
      "env": {
        "PPT_TEMPLATES_DIR": "C:\\ABSOLUTE\\PATH\\TO\\templates",
        "PPT_MCP_OUTPUT_DIR": "C:\\ABSOLUTE\\PATH\\TO\\output"
      }
    }
  }
}
```

> 개발 또는 배포 시에는 `--directory` 구성을 사용하세요. `uvx ppt-mcp`는 별도로 게시된 업스트림 패키지를 해석하므로 이 디렉터리의 코드를 실행하지 않습니다.

기존 설정과의 호환성을 위해 `ppt-mcp` 명령도 유지되지만, **이 프로젝트의 정식 명령은 `ppt-mcp-bob`입니다**.

## Bob 스킬

이 프로젝트는 **[owencase/ppt-skills](https://github.com/owencase/ppt-skills)**의 Bob 스킬과 함께 사용하도록 설계되었습니다. 이 스킬은 MCP 도구만으로는 강제할 수 없는 템플릿 인식 슬라이드 생성, 레이아웃 선택, 디자인 일관성 편집에 대한 상세 지침을 IBM Bob에 제공합니다.

해당 리포지터리의 안내에 따라 스킬을 설치하면, Bob이 프레젠테이션 작업 시 자동으로 적용합니다.

## Bob의 안전한 편집 워크플로우

기존 프레젠테이션 편집 시, Bob은 다음 순서로 작업합니다:

1. `ppt_activate_presentation`을 호출하고 `ppt_get_presentation_info`, `ppt_list_shapes`로 대상을 확인합니다.
2. 확인된 전체 경로와 슬라이드 수로 `ppt_set_work_mode`를 호출합니다.
3. 이동, 크기 조정, 편집, 삭제, 교체 요청 시 `allow_create=false`를 유지합니다.
4. 도형 스냅샷을 캡처하고 안정적인 `shape_id` 값을 보관합니다.
5. 확인된 상태 사전 조건과 함께 `ppt_transform_shapes`, `ppt_delete_shapes`, 또는 `ppt_replace_visual`을 사용합니다.
6. `ppt_proofread_text`를 호출하고, 반환된 모든 텍스트를 문맥에 맞게 검토하며 확인된 오류를 수정한 후 재실행합니다.
7. 저장 전에 `ppt_diff_shape_snapshot`과 `ppt_validate_presentation`을 호출합니다.
8. 교정 결과가 비어 있고, 차이 비교와 유효성 검사 결과가 요청된 변경 사항과 일치할 때만 저장합니다.

`ppt_proofread_text`는 읽기 전용입니다. 신뢰도 높은 한국어·영어 오타, 사용자 정의 교체, 반복 단어, 구두점, 괄호, 제어 문자, 글자 깨짐을 확인합니다. 또한 도형, 그룹, 표, SmartArt, 차트에서 위치 정보가 포함된 텍스트를 반환합니다. 제품명에는 `allowed_terms`, 조직별 전문 용어에는 `custom_replacements`를 활용하세요. 발표자 노트는 선택 사항입니다.

MCP 서버는 Bob에게 조용한 핸드오프 정책을 전달합니다. 프레젠테이션 작업 중에는 계획, 도구 호출, 진행 상황 등 중간 컨텍스트를 노출하지 않습니다. 저장 및 검증이 완료된 후, 사용자의 언어로 결과, 출력 경로 및 범위, 검증 결과를 정확히 세 줄로 응답합니다.

### 기본 작업 모드

| 설정 | 기본값 | 효과 |
|---|---:|---|
| `allow_create` | `false` | 생성이 명시적으로 허용될 때까지 추가, 복사, 복제 도구를 차단합니다. |
| `require_preconditions` | `true` | 변경 시 대상 프레젠테이션 경로와 도형 상태 사전 조건을 요구합니다. |

## 환경 변수

| 변수 | 필수 | 설명 |
|---|---:|---|
| `PPT_TEMPLATES_DIR` | 아니오 | 디렉터리 인수가 없을 때 `ppt_list_templates`가 검색하는 위치입니다. |
| `PPT_MCP_OUTPUT_DIR` | 권장 | 다른 이름으로 저장 및 내보내기 경로의 신뢰 경계입니다. 미지정 시, 잠긴 프레젠테이션의 로컬 디렉터리를 사용합니다. |
| `PPT_DOWNLOAD_TIMEOUT_SECONDS` | 아니오 | 원격 이미지 및 아이콘 메타데이터 타임아웃입니다. 기본값: `15`. |
| `PPT_MAX_DOWNLOAD_BYTES` | 아니오 | 원격 다운로드 최대 크기(바이트)입니다. 기본값: `20971520` (20 MiB). |
| `PPT_AUTO_DISMISS_DIALOG` | 아니오 | PowerPoint가 바쁜 상태로 COM 호출을 거부할 때 Escape를 보냅니다. 기본적으로 비활성화되어 있습니다. |

> `PPT_AUTO_DISMISS_DIALOG=true`는 무인 실행에 유용하지만 사용자가 열어둔 대화 상자를 닫을 수 있습니다. 자동 닫기를 의도한 경우가 아니라면 대화형 세션에서는 비활성화 상태로 유지하세요.

## 도구 카테고리

| 카테고리 | 도구 수 | 지원 범위 |
|---|---:|---|
| 앱 | 5 | 연결, 애플리케이션 상태, 활성 창, 열려 있는 프레젠테이션 |
| 프레젠테이션 | 8 | 만들기, 열기, 저장, 닫기, 활성화, 정보 확인, 템플릿 목록 |
| 슬라이드 | 10 | 추가, 삭제, 복제, 이동, 복사, 정보 확인, 노트, 탐색 |
| 도형 | 10 | 안정적인 ID를 사용한 도형, 텍스트 상자, 그림, 선 추가 및 확인 |
| 안전한 편집 | 8 | 작업 모드, 안전한 변환/삭제/교체, 스냅샷, 차이 비교, 유효성 검사 |
| 텍스트 | 11 | 텍스트 내용, 범위, 단락, 글머리 기호, 검색, 추출, 타이포그래피, 교정 |
| 자리 표시자 | 6 | 자리 표시자 내용 확인 및 업데이트 |
| 서식 | 3 | 채우기, 선, 그림자 |
| 표 | 13 | 데이터, 셀, 행, 열, 병합/분할, 스타일, 레이아웃, 테두리 |
| 내보내기 | 3 | PDF, 이미지, 클립보드 복사 |
| 슬라이드쇼 | 6 | 시작, 정지, 탐색, 슬라이드쇼 상태 확인 |
| 차트 | 7 | 만들기, 확인, 데이터 채우기, 서식, 차트 종류 변경 |
| 애니메이션 | 6 | 전환 효과 및 애니메이션 라이프사이클 작업 |
| 테마 | 4 | 테마 적용, 테마 색상, 머리글/바닥글 |
| 그룹 | 3 | 그룹화, 그룹 해제, 그룹 항목 확인 |
| 커넥터 | 2 | 커넥터 추가 및 서식 |
| 하이퍼링크 | 3 | 추가, 확인, 제거 |
| 섹션 | 3 | 추가, 목록, 관리 |
| 속성 | 2 | 프레젠테이션 메타데이터 읽기 및 업데이트 |
| 미디어 | 3 | 비디오, 오디오, 미디어 설정 |
| SmartArt | 3 | 추가, 수정, SmartArt 레이아웃 목록 |
| 편집 작업 | 6 | 실행 취소, 다시 실행, 도형 또는 서식 복사 |
| 레이아웃 | 7 | 정렬, 균등 배분, 크기, 배경, 대칭, 병합 |
| 효과 | 3 | 광선, 반사, 부드러운 가장자리 효과 |
| 댓글 | 3 | 추가, 목록, 삭제 |
| 고급 | 19 | 태그, 글꼴, 자르기, 그림 작업, 선택, 아이콘, URL, 일괄 적용 |
| 자유형 | 7 | 경로 작성 및 자유형 노드 확인·수정 |
| 디자인 참조 | 1 | React Bits에서 영감을 받은 PowerPoint 네이티브 디자인 레시피 및 공개 참조 링크 |
| **합계** | **165** | |

## 개발

`projects/owencase`에서 실행합니다:

```powershell
uv sync --group dev
uv run pytest
```

테스트 스위트는 스키마 엄격성, 대상 잠금, 재시도 동작, 경로 제약, 안정적인 도형 ID, 슬라이드 작업, 유효성 검사 로직을 검증합니다.

## 버전

현재 버전: **1.7.0** ([`pyproject.toml`](pyproject.toml)의 `ppt-mcp-bob` 패키지)

## 라이선스 및 크레딧

[MIT License](https://opensource.org/licenses/MIT)로 공개되어 있습니다.

이 IBM Bob 통합은 [owencase](https://github.com/owencase)가 관리하며, [ykuwai/ppt-mcp](https://github.com/ykuwai/ppt-mcp)의 PowerPoint MCP 작업을 기반으로 합니다. [FastMCP](https://github.com/jlowin/fastmcp), [pywin32](https://github.com/mhammond/pywin32), [Model Context Protocol](https://modelcontextprotocol.io/)을 사용합니다.
