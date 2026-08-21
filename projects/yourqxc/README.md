# ppt-mcp

PowerPoint(.pptx) 덱을 만들고 고치는 MCP 서버. IBM BOB을 비롯한 MCP 클라이언트에서
도구로 연결해 쓴다.

- 덱 한 벌을 JSON 스펙 한 번으로 생성 (`create_deck`)
- 열어 놓고 슬라이드 단위로 추가·교체·삭제·이동 후 저장 (`open_deck` → … → `save_deck`)
- 슬라이드 13종: 표지 · 목차 · 간지 · 불릿 · 2단 · 비교 · 표 · 차트 · 이미지 · 인용 · KPI · 타임라인 · 자유배치
- 내장 테마 4종, 또는 사내 표준 템플릿(.potx/.pptx)의 레이아웃·마스터를 그대로 사용
- 한글 폰트를 `<a:ea>`까지 지정해 PowerPoint에서 대체 폰트로 새지 않게 처리
- 내용 길이에 맞춰 글자 크기를 자동으로 줄여 텍스트 넘침 방지

## 설치

```bash
python3 -m venv .venv && .venv/bin/pip install .
```

Python 3.11 이상이 필요하다. 직접 선언한 의존성은 `mcp`, `python-pptx`, `pydantic`뿐이고,
HTTP 전송에 쓰는 uvicorn/starlette는 `mcp` 2.0이 함께 가져온다.

코드를 고쳐 가며 쓸 거면 `pip install -e .` 대신 `PYTHONPATH=src`로 실행해도 된다.

## 실행

```bash
.venv/bin/ppt-mcp --output-dir ~/Documents/decks
```

원격 배포(사내 서버에 올려 여러 사람이 공유)할 때는 Streamable HTTP로 띄운다.

```bash
.venv/bin/ppt-mcp --transport http --host 0.0.0.0 --port 8080 --output-dir /srv/decks
```

| 옵션 | 설명 |
| --- | --- |
| `--transport` | `stdio`(기본) 또는 `http` |
| `--host`, `--port`, `--path` | http 모드 바인딩 정보 (기본 `127.0.0.1:8080/mcp`) |
| `--output-dir` | 생성 파일 저장 위치. 쓰기는 이 안으로 제한된다 |
| `--template-dir` | 템플릿을 상대 경로로 찾을 디렉터리 |
| `--default-template` | `template`을 지정하지 않았을 때 쓸 기본 템플릿 |
| `--theme` | 기본 테마 (`carbon_light` `carbon_dark` `minimal` `vivid`) |
| `--allow-remote-images` | 이미지 URL 다운로드 허용 (기본 차단) |

같은 설정을 환경변수로도 줄 수 있다: `PPT_MCP_OUTPUT_DIR`, `PPT_MCP_TEMPLATE_DIR`,
`PPT_MCP_DEFAULT_TEMPLATE`, `PPT_MCP_DEFAULT_THEME`, `PPT_MCP_ALLOW_REMOTE_IMAGES`,
`PPT_MCP_ALLOW_ANY_PATH`, `PPT_MCP_MAX_SLIDES`.

## IBM BOB에 연결

BOB의 MCP 서버 설정에 stdio 서버로 등록한다. 설정 파일 형식은 Claude Desktop 계열과
같은 `mcpServers` 구조를 쓴다.

```json
{
  "mcpServers": {
    "ppt": {
      "command": "/절대경로/PPT_MCP/.venv/bin/ppt-mcp",
      "args": ["--output-dir", "/Users/<이름>/Documents/decks"],
      "env": { "PPT_MCP_DEFAULT_THEME": "carbon_light" }
    }
  }
}
```

`command`는 반드시 절대 경로로 준다. 사내 표준 서식이 있으면 `--default-template`으로
지정해 두면 모든 덱이 그 서식으로 만들어진다. 바로 복사해 쓸 수 있는 설정은
`examples/bob-mcp-config.json`에 있다.

원격(HTTP)으로 붙일 때는 BOB의 원격 MCP 서버 등록에 `http://<호스트>:8080/mcp`를 넣는다.
이 경우 파일은 **서버 쪽** `--output-dir`에 쌓이므로, 공유 스토리지를 물리거나 별도
다운로드 경로를 마련해야 한다.

## 도구

| 도구 | 하는 일 |
| --- | --- |
| `describe_options` | 테마·슬라이드 종류·(템플릿을 주면) 레이아웃 목록 |
| `create_deck` | 덱 전체 스펙을 받아 .pptx 생성·저장 |
| `read_deck` | 기존 .pptx의 슬라이드별 제목·텍스트·노트 읽기 |
| `open_deck` | 새 덱 또는 기존 파일을 편집용으로 열기 |
| `add_slide` / `update_slide` / `delete_slide` / `move_slide` | 슬라이드 단위 편집 |
| `save_deck` | 편집 중인 덱을 파일로 저장 |
| `close_deck` / `list_open_decks` | 세션 정리·조회 |

편집 도구는 메모리에서만 동작한다. **`save_deck`을 불러야 디스크에 쓰인다.**

`create_deck`과 편집 도구는 결과에 `warnings`를 함께 돌려준다. 한 장에 내용이 너무 많아
글자를 줄여야 했던 슬라이드가 여기 나오므로, 에이전트가 그 슬라이드를 두 장으로 나누는
판단에 쓸 수 있다.

## 슬라이드 스펙

`examples/weekly_report.json`이 12장짜리 실제 예시다. 최소 형태는 이렇다.

```json
{
  "title": "분기 리뷰",
  "theme": "carbon_light",
  "footer": "IBM Korea · 내부용",
  "slides": [
    {"type": "title", "title": "2026 3분기 리뷰", "eyebrow": "분기 보고", "date": "2026-09-30"},
    {"type": "bullets", "title": "요약", "bullets": ["매출 목표 달성", "- 신규 고객 12곳", "이탈률 개선"]},
    {"type": "kpi", "title": "핵심 지표",
     "items": [{"value": "128%", "label": "목표 달성률", "delta": "+18%p", "tone": "positive"}]},
    {"type": "chart", "title": "월별 추이", "chart_type": "line",
     "categories": ["7월", "8월", "9월"], "series": [{"name": "매출", "values": [82, 95, 128]}]}
  ]
}
```

불릿은 문자열 앞의 `-` 개수로 들여쓰기 단계를 정한다(`"- 하위"`, `"-- 더 하위"`).
색이나 굵기를 직접 주려면 `{"text": "...", "bold": true, "color": "accent"}` 형태를 쓴다.

## 테마와 템플릿

내장 테마 4종은 **가독성 위주의 중립 팔레트**이지 공식 브랜드 자산이 아니다.
`carbon_light`/`carbon_dark`는 IBM Plex 계열 폰트 이름을 지정하므로, 해당 폰트가 설치돼
있지 않으면 PowerPoint가 임의로 대체한다. 사내 브랜드 규정을 지켜야 한다면 테마 대신
실제 템플릿 파일을 넘기는 편이 확실하다.

```json
{"title": "…", "template": "/서식/IBM_표준.potx", "slides": [...]}
```

템플릿을 주면 그 파일의 레이아웃·마스터를 쓰고, 슬라이드 종류에 맞는 레이아웃을 이름으로
찾아 제목·본문 플레이스홀더를 채운다. 표·차트·KPI처럼 플레이스홀더로 표현할 수 없는
것만 도형으로 얹는다. 어떤 레이아웃이 있는지는 `describe_options`에 `template`을 주면 된다.

## 보안 관련 기본값

도구 인자는 LLM이 채우므로, 프롬프트 인젝션으로 파일 시스템이나 네트워크를 건드리지
못하게 기본값을 좁게 잡았다.

- **쓰기는 `--output-dir` 안으로만** 제한된다. 밖으로 쓰려면 `PPT_MCP_ALLOW_ANY_PATH=1`.
- **이미지 URL 다운로드는 차단**이 기본. 열려면 `--allow-remote-images`.
- 읽기는 `.pptx`/`.potx`/이미지 확장자만 허용한다.
- 슬라이드 수 상한은 200장(`PPT_MCP_MAX_SLIDES`).

## 지금 없는 것

- PDF 내보내기 (LibreOffice 같은 외부 변환기가 필요해 넣지 않았다)
- 기존 슬라이드의 개별 도형 편집 — `update_slide`는 슬라이드를 통째로 교체한다
- SmartArt, 애니메이션, 화면 전환

## 개발

```bash
.venv/bin/pip install -e ".[dev]" && .venv/bin/python -m pytest
```

테스트는 `tests/layout_audit.py`의 레이아웃 불변식 검사를 함께 돌린다. 렌더링 없이
도형 좌표만으로 "슬라이드 밖으로 나감 / 글자 넘침 / 텍스트 겹침"을 잡아내므로, 슬라이드
레이아웃을 손볼 때 눈으로 확인하지 않아도 회귀를 막을 수 있다.
