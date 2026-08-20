# 기여 가이드

`main` 은 보호 브랜치입니다. 직접 push하지 말고 아래 절차를 따라 주세요.

---

## 0. 이것만 하면 됩니다

바쁘면 이 절만 읽으세요. 나머지는 필요할 때 찾아보면 됩니다.

```bash
git checkout feat/내-브랜치      # 본인 브랜치는 이미 만들어져 있습니다
git pull

# ... 자기 프로젝트(projects/<내-아이디>/) 안에서 작업 ...

git add .
git commit -m "feat: 무엇을 왜 바꿨는지"
git push
```

그 다음 GitHub 에서 **Pull Request 를 엽니다.**

- PR 을 열면 CI 가 자동으로 검사합니다. **초록불이면 끝입니다.**
- 빨간불이면 Actions 탭 Summary 에 고치는 방법이 적혀 있습니다. 그대로 하면 됩니다.
- `main` 을 직접 건드리거나 남의 `projects/` 를 고칠 일은 없습니다.

작업은 **자기 패키지 안에서만** 하면 됩니다. 그러면 남과 부딪힐 일이 없어서
충돌을 해결할 필요도 없습니다.

---

## 1. 절대 커밋하지 않는 것

| 대상 | 이유 |
|------|------|
| `node_modules/` | `package-lock.json` + `npm ci` 로 누구나 동일하게 재현됩니다. 커밋하면 레포가 수십 MB로 불고, diff·리뷰가 불가능해집니다. |
| `build/`, `dist/` | 소스에서 생성되는 산출물입니다. 소스와 어긋나면 디버깅이 지옥이 됩니다. |
| 생성된 `.pptx` | 바이너리라 diff가 안 되고 히스토리에 영구히 남습니다. |
| `.env`, API 키 | 한 번 push되면 지워도 히스토리에 남습니다. |

> ⚠️ **GitHub 웹 UI의 "Add files via upload" 로 코드를 올리지 마세요.**
> 실행 권한(exec bit)이 전부 날아가서, 받는 사람은 `npm run build` 조차 실패합니다.
> 실제로 이 레포에서 그렇게 됐습니다. 반드시 `git add` / `git commit` / `git push` 로 올려 주세요.

커밋 전 확인:

```bash
git status          # 올라갈 파일 목록을 눈으로 확인
git diff --cached   # 실제로 무엇이 바뀌는지 확인
```

실수로 이미 추적 중이라면 (파일은 남기고 추적만 해제):

```bash
git rm -r --cached node_modules
```

---

## 2. 폴더 규칙

```
config/      MCP 클라이언트 등록 설정
examples/    실행 가능한 예시 스크립트 (일회성 데모는 여기로)
reference/   공용 레퍼런스 구현 (모두가 읽는 것)
projects/    각자의 ppt-mcp — 폴더명은 자기 GitHub 아이디
```

내 작업은 `projects/<내-GitHub-아이디>/` 아래에 만듭니다.
**레포 루트에 소스 파일을 직접 두지 마세요.**

---

## 3. 브랜치 / PR

### 브랜치 이름

이름(`inwon`, `kunwoo`)이 아니라 **작업 단위**로 만듭니다.

```
feat/html-to-pptx       새 기능
fix/theme-font-color    버그 수정
chore/repo-hygiene      설정·정리 작업
docs/readme-setup       문서
```

### 왜 개인 이름 브랜치를 안 쓰나요

이 레포는 실제로 개인 이름 브랜치로 운영하다가 한 번 무너졌습니다. 무슨 일이
있었는지 그대로 적어 둡니다.

| 브랜치 | 결과 |
|--------|------|
| `geonhee` | 수백 KB짜리 `.thmx` 바이너리와 `build/` 산출물이 커밋됨. 커밋 메시지는 `setting`, `새로운 함수 추가`, `transcom` |
| `inwon` | `config/`, `mcp-server/`, `ppt-bridge/` 디렉터리를 **통째로 삭제**하고 자기 구조로 다시 씀 (83만 줄 삭제) |
| `kunwoo` | `inwon` 의 삭제 커밋을 그대로 물려받아, 되돌릴 방법이 없어짐. 커밋 메시지 `k` |

개인 브랜치의 문제는 실력이 아니라 **구조**입니다.

- **끝이 없습니다.** `feat/html-to-pptx` 는 머지되면 죽지만 `inwon` 은 영원히 삽니다.
  머지 시점이 없으니 main 과 매일 조금씩 벌어지고, 두 달 뒤엔 합칠 수 없게 됩니다.
- **이름이 아무것도 말해주지 않습니다.** `inwon` 브랜치가 무엇을 하는지 알려면
  체크아웃해서 읽어야 합니다. 리뷰가 안 됩니다.
- **"내 브랜치니까" 가 됩니다.** 남의 디렉터리를 지워도 내 브랜치라 괜찮다고 느낍니다.
  위 표의 세 번째 줄이 그렇게 나왔습니다.

### 그럼 내 영역은 어디인가요

**`projects/<내-GitHub-아이디>/` 가 여러분의 영역입니다.**

```
reference/                   공용 — 모두가 읽는 레퍼런스 구현
├── mcp-server/
└── ppt-bridge/

projects/                    사람별 — 각자의 ppt-mcp
├── baekinwon-0102/
├── kunwoo1016/
└── SeoJHeasdw/
```

디렉터리가 둘로 나뉜 이유는 성격이 다르기 때문입니다. `reference/` 는 **읽는
것**이고 `projects/` 는 **만드는 것**입니다. 한 바구니에 섞여 있으면 처음 온
사람이 무엇부터 봐야 할지 모릅니다.

이 레포는 **각자 자기 ppt-mcp 를 통째로 만드는** 구조입니다. 서로의 구현을
보면서 비교하고 이야기하다가 더 나은 방향을 찾는 게 목적입니다. 그래서 폴더는
*기능* 이 아니라 *사람* 으로 나눕니다.

- 기능 이름(`html-ppt-mcp`)으로 나누면 두 사람이 비슷한 걸 만들 때 이름이
  겹칩니다. 그런데 **비슷한 걸 각자 만드는 게 이 레포의 목적입니다.**
- 아이디로 나누면 절대 안 겹치고, 폴더만 봐도 누구 것인지 압니다.

소유를 *시간*(브랜치)이 아니라 *공간*(디렉터리)으로 나눕니다. 디렉터리로 나누면
매일 main 을 머지해도 서로 안 부딪힙니다. 브랜치로 나누면 안 부딪히는 대신
영원히 안 만납니다.

- 내 패키지 안에서는 마음대로 해도 됩니다. 갈아엎어도 남에게 영향이 없습니다.
- 남의 프로젝트와 `reference/` 아래(레퍼런스)는
  이슈를 먼저 열어 주세요.
- **비슷한 걸 두 사람이 각자 만드는 것은 문제가 아닙니다.** 접근이 다르면
  비교해서 배울 게 생깁니다. 합치라고 요구하지 않습니다.

리뷰는 [.github/CODEOWNERS](.github/CODEOWNERS) 에 따라 자동 지정됩니다.

작업 브랜치는 얼마든지 만들어도 됩니다. 브랜치는 싸요. 조건은 두 개뿐입니다 —
**이름이 무엇을 하는지 말해줄 것**, 그리고 **끝나면 머지되고 사라질 것**.

### 절차

0절을 보세요. 브랜치를 새로 만들 때만 아래가 추가됩니다.

```bash
git checkout main
git pull                          # 항상 최신 main에서 시작
git checkout -b feat/my-feature
...
git push -u origin feat/my-feature
```

main 에 직접 push 는 하지 않습니다.

### 커밋 메시지

`k`, `kkk`, `asdf`, `setting`, `새로운 함수 추가` 같은 메시지는 3일 뒤의 나에게도 아무 정보가 안 됩니다.

```
feat: HTML 슬라이드를 PPTX로 변환하는 tool 추가
fix: apply_theme이 두 번째 슬라이드부터 적용되지 않던 문제 수정
chore: node_modules 추적 해제 및 .gitignore 추가
```

---

## 4. 기존 코드를 지우고 다시 쓰고 싶을 때

**`Delete <directory>` 커밋을 연달아 올리지 마세요.** 다른 사람 작업까지 지워지고,
main과 병합이 사실상 불가능해집니다.

방향 전환이 필요하면 먼저 이슈나 PR에서 이야기해 주세요.
기존 컴포넌트를 대체하는 거라면, 지우는 게 아니라 `projects/` 아래에 **새로 추가**하고
동작을 확인한 뒤 정리하는 순서로 갑니다.

---

## 5. 올리기 전 체크리스트

```bash
cd reference/ppt-bridge && .venv/bin/python -m pytest tests   # 브릿지 테스트
cd reference/mcp-server && npm ci && npm test                 # 빌드 + 테스트
git status                                                   # 의도한 파일만 있는지
```

- [ ] `node_modules/`, `build/`, `.pptx` 가 목록에 없다
- [ ] 커밋 메시지가 무엇을 왜 바꿨는지 말해 준다
- [ ] `main`이 아니라 작업 브랜치에 있다

---

## 6. CI

PR 을 열면 [.github/workflows/ci.yml](.github/workflows/ci.yml) 이 자동으로 돕니다.

| 검사 | 하는 일 |
|------|---------|
| 커밋 위생 | `node_modules/`, `build/`, `.pptx` 가 추적되는지 / 루트에 소스가 흩어졌는지 / 파일 삭제가 있는지 |
| Node 패키지 | `reference/*` · `projects/*` 를 훑어서 있는 것마다 빌드·테스트 |
| Python 브릿지 | 의존성 설치 + `pytest` |

빨간불이 뜨면 **Actions 탭의 Summary** 를 먼저 보세요. 무엇이 왜 걸렸고
어떻게 고치는지 적혀 있습니다. 사람한테 물어보기 전에 거기부터 읽으면 대부분 풀립니다.

새 프로젝트를 만들면 CI 를 고칠 필요 없습니다. `package.json` 에 `build` / `test`
스크립트만 넣어 두면 알아서 실행됩니다. **테스트가 없으면 경고가 뜹니다** —
빌드는 통과시켜 주지만, 리뷰에서 물어볼 겁니다.

---

## 7. 테스트를 어떻게 쓰나요

레퍼런스 구현에 예시가 있습니다. 새로 짤 때 이걸 보고 따라 하면 됩니다.

- [reference/ppt-bridge/tests/test_bridge.py](reference/ppt-bridge/tests/test_bridge.py) — 순수 함수 / 프로토콜 계약 / 실제 동작
- [reference/mcp-server/test/tools.test.mjs](reference/mcp-server/test/tools.test.mjs) — 진짜 MCP 클라이언트로 서버에 붙어서 tool 확인

각 테스트마다 **왜 이 테스트가 있는지** 주석을 달아 뒀습니다. 테스트를 추가할 때도
같이 적어 주세요. 3개월 뒤에 그 테스트가 깨졌을 때, 고쳐야 할지 지워야 할지를
판단할 수 있는 건 그 주석뿐입니다.

---

## 8. 설계는 어떤 기준으로 판단하나요

[MCP-DESIGN.md](MCP-DESIGN.md) 에 여덟 개 축과 리뷰 체크리스트가 있습니다.
**남의 `projects/` 를 리뷰할 때도 같은 문서를 씁니다.** 기준이 없으면 리뷰가
취향 싸움이 되기 때문입니다.

리뷰는 점수 매기기가 아니라 질문하기입니다. 나와 다르게 한 게 있으면
"왜 그렇게 했어?" 를 물어보세요.
