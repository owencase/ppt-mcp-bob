# ppt-mcp (SeoJHeasdw)

의도 높이의 tool 로 PowerPoint 덱을 만드는 MCP 서버.

같은 재료(python-pptx)로 같은 결과물(.pptx)을 만들지만, **tool 을 어느 높이에
두느냐**만 바꿨습니다. 이 문서는 그 선택의 이유를 적은 것입니다.

---

## 1. 핵심: tool 의 고도(altitude)

MCP tool 은 *라이브러리가 할 수 있는 것* 이 아니라 *모델이 하려는 것* 의
높이에 맞춰야 합니다.

레퍼런스 구현(`reference/mcp-server`)은 python-pptx 의 API 를 거의 그대로
노출합니다.

```
add_text_box(file_path, slide_index, text, left_cm, top_cm, width_cm, height_cm,
             font_size_pt, bold, color_hex, align)
```

이러면 **좌표 계산이 LLM 의 일이 됩니다.** 5장짜리 덱을 만들려면 30번쯤
호출해야 하고, 계산이 틀리면 글자가 슬라이드 밖으로 나가거나 상자를 넘칩니다.
`examples/build_ibm_quantum.py` 가 214줄인 이유가 그겁니다 — 좌표를 사람이
전부 손으로 넣었습니다.

여기서는 **서버가 레이아웃을 책임집니다.**

```
create_deck(spec, output_path)
```

호출자는 슬라이드 종류와 내용만 말합니다.

```json
{
  "theme": "tech_blue",
  "slides": [
    {"kind": "title",   "title": "AI 혁신의 시대", "subtitle": "2026 기술 전망"},
    {"kind": "section", "title": "핵심 변화"},
    {"kind": "bullets", "title": "세 가지 축", "points": ["자동화", "개인화", "예측"]},
    {"kind": "chart",   "title": "성장률", "series": {"2024": 12, "2025": 31, "2026": 58}}
  ]
}
```

**호출 1회.** 좌표도, 폰트 크기도 나오지 않습니다. 그건 서버가 압니다.

> 이 주장은 테스트로 고정되어 있습니다 —
> `tests/test_server.py::test_create_deck_takes_a_whole_spec_not_coordinates`.
> 누가 `left_cm` 같은 인자를 tool 표면에 추가하면 테스트가 깨집니다.

---

## 2. MCP 의 세 요소를 다 씁니다

대부분 tool 만 씁니다. 나머지 둘이 있어야 서버가 제 몫을 합니다.

| | 뜻 | 여기서는 |
|---|---|---|
| **tool** | 모델이 *실행* 하는 것 (부작용 있음) | `create_deck` · `validate_deck` · `describe_deck` |
| **resource** | 모델이 *읽는* 것 (부작용 없음) | `theme://list` · `theme://{name}` |
| **prompt** | 사용자가 *고르는* 잘 만든 시작점 | `deck_from_topic` |

테마 팔레트를 예로 들면 이렇습니다. tool 로 만들면 모델이 "색을 알아내려고"
동작을 실행해야 합니다. resource 로 두면 그냥 읽습니다. **읽기와 쓰기를
구분하는 것** 이 설계의 요점입니다.

tool 이 셋뿐인 것도 의도입니다. 적을수록 모델이 덜 헷갈립니다.

---

## 3. 에러가 고치는 법을 알려줍니다

모델은 에러를 읽고 다시 시도합니다. 그래서 **에러 메시지의 품질이 곧 다음
시도의 성공률** 입니다.

```
✗  invalid input
✓  불릿은 한 장에 최대 6개입니다 (9개를 보냈습니다). 슬라이드를 나누세요.

✗  Unknown theme
✓  '없는테마' 테마는 없습니다. 사용 가능: marketing_warm, minimal_dark, ...

✗  ValueError: text too long
✓  2번째 슬라이드(section): 14pt 로 줄여도 상자에 안 들어갑니다
   (상자 28.8×3.8cm, 글자 400자). 텍스트를 줄이거나 슬라이드를 나누세요.
```

몇 번째 슬라이드인지 붙이는 것까지가 한 세트입니다. 그래야 그 장만 고칩니다.

---

## 4. 넘치는 텍스트는 조용히 잘리지 않습니다

`layout.fit_font_size()` 가 상자에 맞을 때까지 폰트를 줄이고, **하한(본문
14pt)까지 줄여도 안 들어가면 실패시킵니다.**

잘린 PPT 는 인쇄하고 나서야 발견됩니다. 그때는 늦습니다. 만들다 실패하는 게
낫습니다.

한글은 라틴 문자보다 넓어서 줄 수 추정에 반영했습니다. 같게 계산하면
과소추정해서 넘칩니다.

---

## 5. 저장이 원자적입니다

임시 파일에 다 쓴 뒤 `os.replace` 로 바꿔칩니다. 렌더 도중 죽어도 반쯤 쓰인
`.pptx` 가 남지 않습니다.

레퍼런스 구현은 tool 호출마다 대상 파일을 바로 덮어써서, 20번 호출짜리 덱이
8번째에서 실패하면 반쯤 만들어진 파일이 디스크에 남습니다.

---

## 6. 쓰기가 한 군데로 모입니다

파일을 만드는 지점이 `render.py` 하나뿐이라, 경로 검사도 한 곳에서 합니다.

```bash
export PPT_MCP_OUTPUT_DIR=~/Documents/decks   # 이 밖으로는 못 씀
```

MCP 서버는 LLM 이 준 경로를 그대로 받습니다. 어디든 쓸 수 있게 두면 안 됩니다.
`..` 로 빠져나가는 것도 막습니다 (테스트 있음).

---

## 구조

```
src/ppt_mcp/
├── server.py   MCP 표면 — tool 3 · resource 2 · prompt 1
├── models.py   덱 스펙 (pydantic) — tool 의 입력 계약
├── layout.py   좌표·폰트 계산 ← 핵심. python-pptx 를 모르는 순수 함수
├── render.py   python-pptx 로 그리기 + 원자적 저장 + 경로 게이트
└── theme.py    팔레트
tests/          layout · spec · render · server(계약)
```

`layout.py` 가 python-pptx 를 import 하지 않는 게 중요합니다. 파일을 만들지
않고도 레이아웃 규칙을 테스트할 수 있고, 그래서 촘촘히 검사할 수 있습니다.
레이아웃을 렌더링 코드에 섞으면 그게 안 됩니다.

---

## 설치 · 실행

```bash
cd projects/SeoJHeasdw
python3 -m venv .venv
.venv/bin/pip install -e '.[dev]'
.venv/bin/python -m pytest tests -v
```

테스트는 `pytest` 하나만 있으면 돕니다. MCP 목록 API 가 async 지만
`asyncio.run` 으로 감싸서 `pytest-asyncio` 를 쓰지 않습니다 — 감싸면 되는
일에 의존성을 늘릴 이유가 없습니다.

MCP 클라이언트 등록:

```json
{
  "mcpServers": {
    "ppt-mcp-seojh": {
      "command": "/ABSOLUTE/PATH/TO/projects/SeoJHeasdw/.venv/bin/ppt-mcp-seojh",
      "env": { "PPT_MCP_OUTPUT_DIR": "/ABSOLUTE/PATH/TO/decks" }
    }
  }
}
```

---

## 안 한 것

의도적으로 뺐습니다. 작고 읽을 수 있는 상태를 유지하는 게 이 패키지의 목적입니다.

- 이미지 삽입 — 넣는다면 SSRF 방어가 먼저입니다 (`projects/inwon` 의
  `_public_image_url` 이 좋은 예시입니다)
- 기존 템플릿 편집 — 완전히 다른 문제입니다
- 임의 도형 배치 — 이걸 넣는 순간 1절의 주장이 무너집니다
