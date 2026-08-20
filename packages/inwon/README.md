# Canva-style PowerPoint MCP 3.0


## 3.0 Hybrid Engine: python-pptx 생성 + PowerPoint COM 템플릿 수정

3.0부터 PowerPoint 작업을 두 가지 엔진으로 완전히 분리합니다.

| 사용자 목적 | 실행 모드 | 엔진 | 디자인 처리 |
|---|---|---|---|
| 처음부터 새 PPT 생성 | `generate` | `python-pptx` | 기존 자동 디자인 시스템으로 새로 생성 |
| `/template`의 기존 PPT 내용만 수정 | `template_com` | Windows Microsoft PowerPoint COM (`pywin32`) | 원본 디자인 구조를 유지하고 텍스트 내용만 변경 |

### 반드시 먼저 모드를 확인

MCP는 PPT 요청을 받자마자 파일을 만들지 않습니다. 다음 순서를 강제합니다.

1. `prepare_presentation_task(user_request)` — 문장을 분석해 추천 모드와 사용자 확인 질문을 반환
2. AI가 사용자에게 **새로 생성 / 기존 템플릿 내용만 수정** 중 하나를 질문
3. 사용자 답변 이후 `confirm_presentation_mode(...)` 호출
4. 서버가 1회용 `execution_token` 발급
5. 해당 토큰으로 `create_presentation` 또는 `edit_template_presentation` 중 정확히 하나만 실행

추천 의도가 명확해도 2번 질문을 생략하면 안 됩니다. 실행 토큰이 없거나 다른 모드의 토큰이면 서버가 작업을 거부합니다.

### 문장별 추천 규칙

- `IBM 소개 PPT 만들어줘`, `AI 전략 발표자료 생성해줘` → `generate` 추천 → 사용자 확인 후 python-pptx 생성
- `QBR 템플릿의 내용만 IBM에 맞게 수정해줘`, `이 디자인 그대로 내용만 바꿔줘` → `template_com` 추천 → 사용자 확인 후 COM 수정

### `/template` 폴더

프로젝트 루트의 `template/` 폴더에 `.pptx`, `.pptm`, `.potx`, `.potm` 파일을 넣습니다. COM 모드는 기본적으로 이 폴더 밖의 파일을 템플릿으로 받지 않습니다. 원본은 절대 덮어쓰지 않고 별도 `output_path`로 저장합니다.

```text
project/
├─ template/
│  ├─ company_qbr.pptx
│  └─ sales_template.pptx
├─ output/
└─ src/canva_ppt_mcp/
```

### COM 수정 안전장치

`template_com`은 ykuwai/ppt-mcp의 COM 설계를 참고해 다음 원칙을 적용합니다.

- 실행 중 PowerPoint가 있으면 연결하고, 없으면 새 인스턴스를 시작
- 활성 창을 믿지 않고 **명시적으로 연 output 프레젠테이션 객체만** 수정
- `RPC_E_CALL_REJECTED` / `RPC_E_SERVERCALL_RETRYLATER`는 제한 횟수 재시도
- 텍스트 상자의 문자열만 교체하고 shape geometry를 직접 수정하지 않음
- 수정 전/후에 텍스트를 제외한 디자인 지문을 비교
- 슬라이드 수, 도형 수, 위치/크기/회전, fill/line, 기본 텍스트 스타일 등이 달라지면 `design_preserved=false`로 성공 처리하지 않음
- Windows, PowerPoint, pywin32가 없으면 python-pptx 템플릿 모드로 몰래 폴백하지 않고 실패

벤치마크: https://github.com/ykuwai/ppt-mcp

### Windows COM 설치

```bash
pip install -e ".[windows]"
```

Microsoft PowerPoint가 설치된 Windows에서만 `template_com`을 사용할 수 있습니다.

### CLI 예시

직접 CLI를 쓸 때도 기본적으로 모드를 다시 묻습니다. 자동화에서 사용자가 이미 모드를 명시적으로 선택한 경우에만 `--yes-mode-confirmed`를 사용합니다.

```bash
# 추천 모드만 확인 (파일 변경 없음)
canva-ppt route --request "IBM 소개 PPT 만들어줘"

# 처음부터 생성
canva-ppt create --topic "IBM의 생성형 AI 전략" --slides 8 \
  --output ./output/ibm-ai.pptx

# /template/company_qbr.pptx의 디자인은 그대로 두고 내용만 COM으로 수정
canva-ppt edit-template --topic "2026 IBM QBR" --template company_qbr.pptx \
  --output ./output/ibm-qbr.pptx

canva-ppt list-templates
```

### Skill

`skills/powerpoint-routing/SKILL.md`와 `.codex/skills/powerpoint-routing/SKILL.md`를 함께 제공합니다. 이 스킬은 파일 변경 도구를 호출하기 전에 사용자에게 두 모드 중 하나를 반드시 확인하도록 규정합니다. OpenAI Skills의 일반적인 `SKILL.md` 워크플로 형식을 따릅니다.

---

주제에 맞는 공개 자료를 먼저 조사하고 디자인 시스템을 확정한 뒤 PowerPoint를 만드는 로컬 MCP입니다. OpenAI 패키지와 API 키가 없어도 Wikipedia 또는 사용자가 제공한 조사 문서로 동작합니다.

## 2.2 제목 한 줄 고정

- 본문 슬라이드, 표지, 마무리마다 실제 제목 상자에 맞는 별도 한 줄 너비 예산 적용
- 긴 주제도 핵심어를 보존해 한 줄용 제목으로 의미 중심 재요약
- PowerPoint 제목 텍스트 상자의 자동 줄바꿈 비활성화
- Semantic QA와 PPTX 구조 QA에서 각각 `SLIDE_TITLE_WRAP_RISK`, `TITLE_WRAP` 검사
- 글자 크기를 44pt 아래로 줄이지 않고 제목 문구를 먼저 압축

## 2.1 제목·소제목 요약 개선

- 슬라이드 제목을 원문 앞뒤 절단이 아닌 30자 안팎의 결론형 문장으로 재작성
- 항목 소제목을 본문 첫 구절이 아닌 18자 안팎의 의미 중심 명사구로 요약
- 한국어 조사 선택과 숫자·연도·순위·금액·고유명사 보존 규칙 추가
- 34자를 넘는 제목, 22자를 넘는 소제목, 본문 앞부분을 그대로 복사한 소제목을 Semantic QA에서 차단
- 오프라인 모드에서도 정의·효과·운영·시장·역사·통계 문장을 발표용 문구로 변환

## 2.0 주요 개선

- `Claim → Source → Slide` 근거 매핑과 숫자 일치 검사
- 여러 개의 `{title, url, text}` 조사 문서 입력 및 자료별 출처 유지
- 청중·목적 키워드를 반영한 섹션 우선순위와 takeaway형 제목
- 주제·대상·목적에 따라 7개 디자인 프리셋을 결정적으로 선택하거나 `--style`로 지정
- Wikipedia 대표 이미지를 안전하게 내려받아 풀블리드 이미지 슬라이드에 사용
- `bar`, `column`, `line`, `pie` 차트 스키마와 네이티브 PowerPoint 차트
- 템플릿의 테마 색·폰트·레이아웃 목록 추출, 보호 푸터 제외, 이미지 슬롯 교체
- 예시 슬라이드가 없는 템플릿은 오류 대신 자동 디자인 모드로 안전하게 전환
- 그룹 내부 텍스트, 텍스트 위 시각 요소, 반복 글자 손실, 마지막 수정 후 재검사
- 파일 잠금, 입력 크기 제한, 색상 전체 검증, 출력 루트 제한
- 한국어 덱의 고정 UI 문구 현지화

## 1.2.0 실제 주제 내용 Grounding

- `Research → DeckPlan → Semantic QA → Render QA` 파이프라인
- OpenAI API 키가 없어도 한국어/영어 Wikipedia 공개 API로 주제 조사
- 인터넷 없이 사용할 때는 `research_text`와 `source_urls`로 사내 자료나 사용자가 제공한 내용을 입력
- 모든 외부 사실의 speaker notes에 `[Sources]` 블록 기록
- 범용 폴백 문구를 찾는 `GENERIC_FILLER`, 주제 누락을 찾는 `TOPIC_NOT_PRESENT`, 출처 누락을 찾는 `SOURCE_NOTES_MISSING` 검사
- 조사 실패 시 범용 슬라이드를 만들지 않고 명확한 오류 반환

## 1.1.0 화려함 강화

- 제목 44~60pt, 핵심 통계 80~120pt의 극단적인 스케일 대비
- 강조 슬라이드에 미리 생성한 래스터 그라데이션 PNG와 미묘한 노이즈 텍스처 적용
- 전체 40~50% 다크 슬라이드, 화려한 장과 차분한 장의 교차 리듬
- 반투명 대형 도형, 비대칭 배치, 회전 요소, 깨진 2×2 그리드, 50/50 컬러 블로킹
- 이미지 슬라이드의 풀블리드 크롭과 가독성 오버레이
- 표지와 마무리에 동일 장식/동일 구도를 반복하지 않음
- QA에 `DECOR_TEXT_CONTRAST`와 `DARK_RHYTHM` 검사 추가

## 제공 도구

- `prepare_presentation_task`: PPT 요청의 추천 모드를 판별하고 **필수 사용자 확인 질문** + `confirmation_id` 반환. 파일 변경 없음
- `confirm_presentation_mode`: 사용자가 고른 `generate` / `template_com` 모드를 등록하고 1회용 `execution_token` 발급
- `create_presentation`: 확인된 `generate` 토큰으로만 python-pptx 새 PPT 생성 + Semantic/Render QA
- `list_templates`: `/template` 폴더의 사용 가능한 템플릿 목록 조회
- `edit_template_presentation`: 확인된 `template_com` 토큰으로만 PowerPoint COM 텍스트 수정 + 디자인 지문 검증
- `inspect_template`: PPTX/POTX의 전체 슬라이드 유형, 색상, 폰트, 슬롯 분석
- `qa_presentation`: 기존 PPTX를 렌더링하고 오버플로·겹침·대비·placeholder 검사

## 설치

Python 3.11+와 LibreOffice(`soffice`)가 필요합니다. Linux에서 한글 덱을 만들 때는 `fonts-noto-cjk` 같은 CJK 글꼴도 설치해야 렌더 QA를 통과합니다. Windows에서는 맑은 고딕을 동아시아 fallback으로 사용합니다.

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -e .
```

콘텐츠 계획에 OpenAI API를 사용하려는 경우에만 선택 의존성을 설치하고 환경 변수를 설정합니다.

```bash
pip install -e ".[ai]"
set OPENAI_API_KEY=...
set PPT_MCP_MODEL=gpt-5.6
```

## 로컬 MCP 연결

Codex CLI:

```bash
codex mcp add canva-ppt -- canva-ppt-mcp
```

프로젝트 `.codex/config.toml` 예시:

```toml
[mcp_servers.canva-ppt]
command = "C:/path/to/project/.venv/Scripts/canva-ppt-mcp.exe"
startup_timeout_sec = 20
tool_timeout_sec = 600
```

ChatGPT 데스크톱 또는 MCP 지원 클라이언트에서는 STDIO 서버 명령으로 `canva-ppt-mcp`를 등록합니다.

## 직접 실행

```bash
canva-ppt create --topic "IBM의 생성형 AI 전략" --slides 8 --output ./output/ibm-ai.pptx
canva-ppt create --topic "API 없이 테스트" --content-json ./examples/demo_plan.json --output ./output/demo.pptx
canva-ppt inspect-template ./template.pptx
canva-ppt qa ./output/ibm-ai.pptx
```

API 없이 공개 자료를 자동 조사하는 기본 사용법:

```bash
canva-ppt create --topic "IBM" --slides 8 --language ko --output ./output/ibm.pptx
canva-ppt create --topic "IBM" --slides 8 --style editorial --output ./output/ibm-editorial.pptx
```

인터넷 없이 직접 제공한 자료로 생성:

```bash
canva-ppt create --topic "사내 AI 플랫폼" --research-text ./research.txt \
  --source-url "https://intranet.example/research" --output ./output/internal-ai.pptx
```

여러 자료를 교차 사용하려면 JSON 배열을 전달합니다.

```json
[
  {"title": "공식 문서", "url": "https://example.com/official", "text": "..."},
  {"title": "운영 보고서", "url": "https://example.com/report", "text": "..."}
]
```

```bash
canva-ppt create --topic "사내 AI 플랫폼" --research-documents ./sources.json \
  --output ./output/internal-ai.pptx
```

MCP가 쓸 수 있는 출력 위치를 제한하려면 `PPT_MCP_OUTPUT_ROOT`를 설정합니다.

`create_presentation` 입력의 `content_json`에 직접 슬라이드 내용을 넘기면 API 호출 없이 그대로 렌더링합니다. 이미지에는 로컬 파일 또는 공개 HTTP(S) URL을 쓸 수 있습니다. 원격 이미지는 사설 IP 차단, 크기 제한, 실제 이미지 형식 검사를 통과한 경우에만 사용합니다.

## QA 동작

1. PPTX를 LibreOffice로 PDF 변환
2. 각 페이지를 PNG로 렌더링
3. PPTX 구조와 PDF 텍스트 위치를 함께 검사
4. 실패한 슬라이드만 최소 글자 크기 범위 안에서 수정
5. 수정된 페이지만 다시 PNG로 추출
6. 설정된 횟수 내 통과하지 못하면 성공으로 보고하지 않고 오류 반환

검사 보고서는 PPTX 옆의 `<파일명>_qa/qa-report.json`에 저장됩니다.

## 테스트

`pytest`가 없어도 핵심 회귀 테스트를 실행할 수 있습니다.

```bash
PYTHONPATH=src:. python scripts/run_tests.py
PYTHONPATH=src python scripts/smoke_test.py
```

자동 공개 조사는 현재 Wikipedia를 기본 자료로 사용합니다. 단일 자료만 확보된 경우 결과는 생성하되 `SINGLE_SOURCE` 경고를 남깁니다. 중요한 의사결정용 덱에는 `research_documents`로 공식 문서와 최신 자료를 함께 제공하는 것이 권장됩니다.
