# ppt-mcp-bob

> **각자 PowerPoint MCP 서버를 만들고, 서로의 구현을 비교하며 배우는 레포입니다.**

**IBM BoB**(코딩 에이전트)에 붙여서 자연어로 `.pptx` 를 만드는 MCP 서버를
각자 하나씩 만듭니다.

같은 문제를 서로 다르게 푼 구현들이 `projects/` 아래에 나란히 놓이고, 그걸
읽고 이야기하면서 더 나은 방향을 찾는 것이 이 레포의 목적입니다.
**비슷한 걸 여러 명이 만드는 것은 중복이 아니라 의도입니다.**

> 붙일 대상은 **IBM BoB** 입니다. MCP 는 표준 프로토콜이라 Claude Code ·
> Copilot 같은 다른 클라이언트에도 그대로 붙지만, 동작 확인은 BoB 기준으로
> 하세요.

---

## 처음 오셨다면

### 1. 레퍼런스를 읽습니다

`reference/` 에 동작하는 구현이 있습니다. 작고 단순해서 한 번에 읽힙니다.

```
reference/mcp-server/src/index.ts       tool 9개 정의 (TypeScript)
reference/ppt-bridge/bridge.py          python-pptx 로 실제 파일 조작
```

### 2. 직접 돌려 봅니다

아래 "설치 및 빌드" 를 따라 하면 BoB 에서 실제로 PPT 가 만들어집니다.
남의 코드를 읽기 전에 **돌아가는 걸 한 번 보는 게** 이해가 빠릅니다.

### 3. 내 프로젝트를 만듭니다

```
projects/<내-GitHub-아이디>/
```

레퍼런스를 그대로 베껴도 되고, 완전히 다르게 만들어도 됩니다.
설계를 어떤 축으로 판단하면 되는지는 [MCP-DESIGN.md](MCP-DESIGN.md) 에 정리했습니다.
남의 구현을 리뷰할 때도 같은 문서를 씁니다.

작업 절차(브랜치·PR·CI)는 [CONTRIBUTING.md](CONTRIBUTING.md) 0절이면 충분합니다.

---

## 폴더 구조

```
ppt-mcp-bob/
├── .github/
│   ├── workflows/ci.yml      # PR 마다 도는 자동 검사
│   ├── CODEOWNERS
│   └── pull_request_template.md
├── config/mcp.json           # MCP 클라이언트 등록 설정 (샘플)
├── MCP-DESIGN.md             # 설계 판단 기준 · 리뷰 체크리스트
│
├── reference/                # ← 공용. 모두가 읽는 것
│   ├── mcp-server/           #   TypeScript MCP 서버
│   ├── ppt-bridge/           #   Python 브릿지 (python-pptx) + 최소 예제
│   └── design-library/       #   템플릿·테마 메타데이터 규약
│
└── projects/                 # ← 사람별. 폴더명 = GitHub 아이디
    └── SeoJHeasdw/           #   멘토의 ppt-mcp
```

**`reference/` 는 읽는 것, `projects/` 는 만드는 것입니다.** 성격이 달라서
디렉터리를 나눴습니다.

`node_modules/`, `build/`, 생성된 `.pptx` 는 **커밋하지 않습니다** (`.gitignore` 참고).

---

## 지금 있는 것들

| 경로 | 하는 일 | 담당 |
|---|---|---|
| `reference/mcp-server` | 레퍼런스 MCP 서버 (TypeScript) | 공용 |
| `reference/ppt-bridge` | 레퍼런스 브릿지 (Python · python-pptx) | 공용 |
| `reference/design-library` | 템플릿·테마 메타데이터 규약 | 공용 |
| `projects/SeoJHeasdw` | ppt-mcp — 의도 높이의 tool + resources/prompts | `@SeoJHeasdw` |

**여기가 비어 보이는 게 정상입니다.** `projects/` 는 여러분이 채우는 자리입니다.

각 항목은 **독립적으로 설치·실행**됩니다. 하나를 보려고 전체를 설치할 필요는
없습니다. 자세한 사용법은 각 폴더의 `README.md` 를 보세요.

> `projects/` 폴더명은 **GitHub 아이디**입니다.
> `공용` 은 모두가 참고하는 코드라 바꾸려면 이슈를 먼저 열어 주세요.

---

## 레퍼런스 구현의 구조

아래 "설치 및 빌드" 부터는 **`reference/` 구현** 에 대한 설명입니다.
여러분의 프로젝트는 이 구조를 따르지 않아도 됩니다.

```
IBM BoB  ← 주 대상 (코딩 에이전트)
  다른 MCP 클라이언트(Claude Code, Copilot 등)도 동일하게 동작
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

> `.pptx` 를 **직접 읽고 쓰는** 방식입니다. PowerPoint COM 자동화가 아니므로
> Windows 나 PowerPoint 설치가 없어도 동작합니다. (열어서 보려면 뷰어는 필요합니다.)

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

`config/mcp.json` 을 **IBM BoB** 의 MCP 설정에 병합하고, `args` 의 경로를
**본인 환경의 절대경로**로 바꿔 주세요.

`mcpServers` 형식은 MCP 클라이언트 공통이라 내용은 그대로 쓰면 되지만,
**설정 파일이 어디 있는지는 클라이언트마다 다릅니다.** BoB 쪽 경로는 BoB
문서를 확인하세요.

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

## 레퍼런스 구현의 Tools

> `reference/` 구현이 노출하는 tool 입니다. 여러분의 프로젝트는 다른 tool 을
> 노출해도 됩니다 — 오히려 그게 비교할 거리가 됩니다.


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

BoB 에서 다음과 같이 요청하면 됩니다:

```
"AI 기술 소개 PPT를 5장짜리로 만들어줘. tech_blue 테마로 해주고,
 첫 번째 슬라이드에 제목은 'AI 혁신의 시대'로 크게 넣어줘."
```

MCP 서버 없이 브릿지만 직접 호출하는 46줄짜리 예제도 있습니다.

```bash
cd reference/ppt-bridge && .venv/bin/python example_deck.py
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
