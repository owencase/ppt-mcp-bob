# PPT MCP Server

> LLM 클라이언트(Claude / Copilot / ChatGPT)에서 자연어로 PowerPoint 파일을 자동 생성·편집할 수 있는 MCP 서버입니다.

---

## 아키텍처

```
LLM Client (Claude / Copilot / ChatGPT)
        │  MCP Protocol (stdio JSON-RPC)
        ▼
reference/mcp-server   ← TypeScript / Node.js · Tool 정의
        │  stdin/stdout JSON (1 tool call = 1 프로세스)
        ▼
reference/ppt-bridge   ← Python / python-pptx · 실제 파일 조작
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
│   ├── workflows/ci.yml      # PR 마다 도는 자동 검사
│   ├── CODEOWNERS
│   └── pull_request_template.md
├── config/mcp.json           # MCP 클라이언트 등록 설정 (샘플)
├── examples/                 # 실행 가능한 예시 스크립트
│
├── reference/                # ← 공용. 모두가 읽는 레퍼런스 구현
│   ├── mcp-server/           #   TypeScript MCP 서버
│   └── ppt-bridge/           #   Python 브릿지 (python-pptx)
│
└── projects/                 # ← 사람별. 폴더명 = GitHub 아이디
    ├── SeoJHeasdw/           #   멘토의 ppt-mcp
    ├── design-library/       #   (담당 확인 대기)
    ├── html-ppt-mcp/         #   (담당 확인 대기)
    └── html-render-pptx/     #   (담당 확인 대기)
```

**`reference/` 는 읽는 것, `projects/` 는 만드는 것입니다.** 성격이 달라서
디렉터리를 나눴습니다. 내 작업은 `projects/<내-GitHub-아이디>/` 아래에 만듭니다.
`node_modules/`, `build/`, 생성된 `.pptx` 는 **커밋하지 않습니다** (`.gitignore` 참고).

---

## 패키지

각 패키지는 **독립적으로 설치·실행**됩니다. 하나를 쓰려고 전체를 설치할 필요는 없습니다.
자세한 사용법은 각 패키지의 `README.md` 를 보세요.

| 경로 | 하는 일 | 담당 |
|---|---|---|
| `reference/mcp-server` | 레퍼런스 MCP 서버 (TypeScript) | 공용 |
| `reference/ppt-bridge` | 레퍼런스 브릿지 (Python · python-pptx) | 공용 |
| `projects/SeoJHeasdw` | ppt-mcp — 의도 높이의 tool + resources/prompts | `@SeoJHeasdw` |
| `projects/design-library` | 템플릿·테마 메타데이터 | *미확인* |
| `projects/html-render-pptx` | HTML/CSS → PPTX | *미확인* |
| `projects/html-ppt-mcp` | HTML 기반 MCP 서버 | *미확인* |

> **`projects/` 폴더는 GitHub 아이디로 만듭니다.** 각자 자기 ppt-mcp 를 통째로
> 만들고 서로 비교하는 구조라, 기능 이름으로 나누면 이름이 겹칩니다.
> 자세한 이유는 [CONTRIBUTING.md](CONTRIBUTING.md) 3절.
>
> 아래 세 개는 아이디 규칙 이전에 만들어진 것이라 이름이 기능 기준입니다.
> 담당자가 정해지면 아이디 폴더로 옮깁니다.
> `공용` 은 모두가 참고하는 레퍼런스라 바꾸려면 이슈를 먼저 열어 주세요.

### 어느 걸 쓰면 되나요

- **PPT 를 자연어로 만들고 싶다** → `reference/mcp-server` + `reference/ppt-bridge` 를 설치하고 MCP 클라이언트에 등록하세요. 아래 "설치 및 빌드" 참고.
- **직접 스크립트로 PPT 를 만들고 싶다** → `reference/ppt-bridge` 만 설치하고 `examples/build_ibm_quantum.py` 를 참고하세요.
- **HTML 로 슬라이드를 디자인하고 싶다** → `projects/html-render-pptx` / `projects/html-ppt-mcp` (아직 작업 중)

---

## 사전 요구사항

| 항목 | 버전 |
|------|------|
| Node.js | 20 이상 |
| Python | 3.10 이상 (CI 는 3.11 로 검사) |

> 개별 패키지가 더 높은 버전을 요구할 수 있습니다. 각 패키지의
> `pyproject.toml` / `package.json` 을 확인하세요.

---

## 설치 및 빌드

### 1. Python 의존성 설치

```bash
cd reference/ppt-bridge
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt -r requirements-dev.txt
```

> venv 경로(`reference/ppt-bridge/.venv`)를 그대로 쓰는 걸 권장합니다.
> 아래 MCP 설정에서 이 경로를 그대로 가리키게 되어 있습니다.

### 2. Node.js 의존성 설치 및 빌드

```bash
cd reference/mcp-server
npm ci
npm run build
```

빌드 결과물: `reference/mcp-server/build/index.js`

### 3. 테스트

```bash
cd reference/ppt-bridge && .venv/bin/python -m pytest tests -v
cd reference/mcp-server && npm test
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
      "args": ["/ABSOLUTE/PATH/TO/ppt-mcp-bob/reference/mcp-server/build/index.js"],
      "env": {
        "PYTHON_BIN": "/ABSOLUTE/PATH/TO/ppt-mcp-bob/reference/ppt-bridge/.venv/bin/python"
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
| `BRIDGE_SCRIPT` | `reference/ppt-bridge/bridge.py` | 브릿지 스크립트 경로 직접 지정 (경로에 한글이 섞여 문제가 될 때 사용) |

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
