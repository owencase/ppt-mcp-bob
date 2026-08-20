# PPT MCP Server

> LLM 클라이언트(Claude / Copilot / ChatGPT)에서 자연어로 PowerPoint 파일을 자동 생성·편집할 수 있는 MCP 서버입니다.

---

## 아키텍처

```
LLM Client (Claude / Copilot / ChatGPT)
        │  MCP Protocol (stdio JSON-RPC)
        ▼
packages/mcp-server    ← TypeScript / Node.js · Tool 정의
        │  stdin/stdout JSON (1 tool call = 1 프로세스)
        ▼
packages/ppt-bridge    ← Python / python-pptx · 실제 파일 조작
        │  파일 I/O
        ▼
.pptx 파일
```

> `.pptx` 파일을 **직접 읽고 쓰는** 방식입니다. PowerPoint COM 자동화가 아니므로
> Windows나 PowerPoint 설치가 없어도 동작합니다. (열어서 보려면 뷰어는 필요합니다.)

---

## 폴더 구조

```
ppt-mcp-bob/
├── .github/
│   ├── workflows/ci.yml          # PR 마다 도는 자동 검사
│   └── pull_request_template.md
├── config/
│   └── mcp.json                  # MCP 클라이언트 등록 설정 (샘플)
├── examples/
│   └── build_ibm_quantum.py      # 브릿지 직접 호출 예시 스크립트
└── packages/
    ├── mcp-server/               # Node.js MCP 서버
    │   ├── src/index.ts          # 모든 Tool 정의
    │   ├── test/tools.test.mjs   # tool 노출 + 브릿지 계약 테스트
    │   ├── package.json
    │   └── tsconfig.json
    └── ppt-bridge/               # Python 브릿지
        ├── bridge.py             # python-pptx 기반 액션 핸들러
        ├── tests/test_bridge.py  # 프로토콜 · 액션 테스트
        ├── requirements.txt
        └── requirements-dev.txt  # pytest (테스트용)
```

새 컴포넌트를 추가할 때는 `packages/<이름>/` 아래에 만들어 주세요.
`node_modules/`, `build/`, 생성된 `.pptx` 는 **커밋하지 않습니다** (`.gitignore` 참고).

---

## 패키지

각 패키지는 **독립적으로 설치·실행**됩니다. 하나를 쓰려고 전체를 설치할 필요는 없습니다.
자세한 사용법은 각 패키지의 `README.md` 를 보세요.

| 패키지 | 하는 일 | 담당 | 설치 · 실행 | 상태 |
|---|---|---|---|---|
| `packages/mcp-server` | MCP 서버. tool 9종을 LLM 클라이언트에 노출 | (레퍼런스) | `npm ci && npm run build` | `main` |
| `packages/ppt-bridge` | python-pptx 로 `.pptx` 를 직접 조작 | (레퍼런스) | `pip install -r requirements.txt` | `main` |
| `packages/design-library` | PPT 템플릿·테마 메타데이터 | `@github-id` | — | `feat/bridge-extended` 작업 중 |
| `packages/html-render-pptx` | HTML/CSS 렌더링 → PPTX 내보내기 | `@github-id` | `npm install && npm test` | `feat/html-render-pptx` 작업 중 |
| `packages/html-ppt-mcp` | HTML 기반 프레젠테이션 생성 MCP 서버 | `@github-id` | `npm install && npm run build` | `feat/html-ppt-mcp` 작업 중 |

> 작업 중인 패키지는 해당 브랜치에만 있습니다. `main` 에 머지되면 상태가 `main` 으로 바뀝니다.
>
> **담당** 은 그 패키지를 만든 사람입니다. 질문·리뷰 요청은 담당자에게 보내세요.
> 이 표는 멘토가 관리합니다 — 여러분이 고칠 일은 없습니다.
> `(레퍼런스)` 는 모두가 참고하는 공용 구현이라 바꾸려면 이슈를 먼저 열어 주세요.

> 비슷한 걸 두 사람이 각자 만드는 것은 **문제가 아닙니다.** 접근이 다르면 비교해서
> 배울 게 생깁니다. 다만 남의 패키지를 지우거나 덮어쓰지는 마세요 (CONTRIBUTING.md 4절).

### 어느 걸 쓰면 되나요

- **PPT 를 자연어로 만들고 싶다** → `mcp-server` + `ppt-bridge` 를 설치하고 MCP 클라이언트에 등록하세요. 아래 "설치 및 빌드" 참고.
- **직접 스크립트로 PPT 를 만들고 싶다** → `ppt-bridge` 만 설치하고 `examples/build_ibm_quantum.py` 를 참고하세요.
- **HTML 로 슬라이드를 디자인하고 싶다** → `html-render-pptx` / `html-ppt-mcp` (아직 작업 중)

---

## 사전 요구사항

| 항목 | 버전 |
|------|------|
| Node.js | 20 이상 |
| Python | 3.10 이상 |

---

## 설치 및 빌드

### 1. Python 의존성 설치

```bash
cd packages/ppt-bridge
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt -r requirements-dev.txt
```

> venv 경로(`packages/ppt-bridge/.venv`)를 그대로 쓰는 걸 권장합니다.
> 아래 MCP 설정에서 이 경로를 그대로 가리키게 되어 있습니다.

### 2. Node.js 의존성 설치 및 빌드

```bash
cd packages/mcp-server
npm ci
npm run build
```

빌드 결과물: `packages/mcp-server/build/index.js`

### 3. 테스트

```bash
cd packages/ppt-bridge && .venv/bin/python -m pytest tests -v
cd packages/mcp-server && npm test
```

PR 을 열면 이 검사가 [CI](.github/workflows/ci.yml) 에서 자동으로 돕니다.
빨간불이 뜨면 Actions 탭의 요약(Summary)에 고치는 방법이 적혀 있습니다.

`mcp-server` 의 계약 테스트는 `index.ts` 의 tool 목록과 `bridge.py` 의
`HANDLERS` 가 어긋났는지 확인합니다. 액션을 추가할 때 한쪽만 고치면 여기서 잡힙니다.

---

## MCP 클라이언트에 등록하기

`config/mcp.json` 을 클라이언트의 MCP 설정 파일에 병합하고,
`args` 의 경로를 **본인 환경의 절대경로**로 바꿔 주세요.

```json
{
  "mcpServers": {
    "ppt-mcp-server": {
      "command": "node",
      "args": ["/ABSOLUTE/PATH/TO/ppt-mcp-bob/packages/mcp-server/build/index.js"],
      "env": {
        "PYTHON_BIN": "/ABSOLUTE/PATH/TO/ppt-mcp-bob/packages/ppt-bridge/.venv/bin/python"
      }
    }
  }
}
```

MCP 클라이언트는 임의의 작업 디렉터리에서 서버를 실행하므로 **상대경로를 쓰면 안 됩니다.**

> ⚠️ `PYTHON_BIN` 을 그냥 `python3` 로 두면 안 됩니다. MCP 클라이언트는 venv 가
> 활성화되지 않은 상태로 서버를 띄우기 때문에, `python3` 는 `python-pptx` 가 없는
> 시스템 파이썬을 가리키게 되고 **모든 tool 호출이 실패합니다.**
> venv 안의 python 실행 파일을 절대경로로 직접 가리켜 주세요.

### 환경변수

| 변수 | 기본값 | 설명 |
|------|--------|------|
| `PYTHON_BIN` | `python3` | Python 실행 파일 경로. **venv 안의 python 을 절대경로로 지정하세요.** 기본값은 venv 를 못 찾습니다. |
| `BRIDGE_SCRIPT` | `packages/ppt-bridge/bridge.py` | 브릿지 스크립트 경로 직접 지정 (경로에 한글이 섞여 문제가 될 때 사용) |

---

## 사용 가능한 Tools

| Tool | 설명 |
|------|------|
| `create_presentation` | 새 .pptx 파일 생성 |
| `add_slide` | 슬라이드 추가 |
| `add_text_box` | 텍스트 박스 추가 (폰트, 색상, 정렬 지원) |
| `add_image` | 이미지 삽입 |
| `set_background_color` | 슬라이드 배경색 변경 |
| `add_shape` | 사각형/둥근사각형 도형 추가 |
| `apply_theme` | 전체 슬라이드에 배경색 일괄 적용 + 액센트 팔레트 반환 |
| `save_presentation` | 파일 저장 |
| `get_presentation_info` | 슬라이드/도형 메타데이터 조회 |

모든 위치·크기 단위는 **cm**, 색상은 `RRGGBB` **hex 문자열**(`#` 없이)입니다.

### 내장 테마

| 테마명 | 특징 |
|--------|------|
| `minimal_dark` | 어두운 배경, 흰 텍스트, 보라 계열 액센트 |
| `minimal_light` | 흰 배경, 어두운 텍스트, 파랑 계열 액센트 |
| `tech_blue` | 네이비 배경, 하늘색 액센트 |
| `marketing_warm` | 따뜻한 크림 배경, 오렌지 계열 액센트 |

> `apply_theme` 은 **배경색만** 적용하고 나머지 팔레트(`text`/`accent1`/`accent2`)는
> 응답으로 돌려줍니다. 그 값을 이어지는 `add_text_box` / `add_shape` 호출에 넣어 주세요.

---

## 사용 예시

LLM 클라이언트에서 다음과 같이 요청하면 됩니다:

```
"AI 기술 소개 PPT를 5장짜리로 만들어줘. tech_blue 테마로 해주고,
 첫 번째 슬라이드에 제목은 'AI 혁신의 시대'로 크게 넣어줘."
```

브릿지를 직접 호출하는 스크립트 예시는 `examples/build_ibm_quantum.py` 를 참고하세요.

```bash
python3 -X utf8 examples/build_ibm_quantum.py
```

---

## 보안 유의사항

- `file_path`, `image_path` 는 **절대경로**를 사용하세요. 상대경로는 클라이언트의
  작업 디렉터리 기준으로 해석되어 예측할 수 없는 위치에 파일이 생깁니다.
- 서버는 stdio로만 통신하며 네트워크 포트를 열지 않습니다.
- API 키나 비밀 정보는 코드에 하드코딩하지 말고 환경변수를 사용하세요.

---

## 기여

브랜치 규칙과 PR 절차는 [CONTRIBUTING.md](CONTRIBUTING.md) 를 읽어 주세요.
