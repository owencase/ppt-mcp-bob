# 기여 가이드

`main` 은 보호 브랜치입니다. 직접 push하지 말고 아래 절차를 따라 주세요.

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
packages/    독립적으로 설치·빌드되는 컴포넌트
```

새 컴포넌트는 `packages/<이름>/` 아래에 만듭니다.
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

**`packages/<내-패키지>/` 가 여러분의 영역입니다.**

소유를 *시간*(브랜치)이 아니라 *공간*(디렉터리)으로 나눕니다. 디렉터리로 나누면
매일 main 을 머지해도 서로 안 부딪힙니다. 브랜치로 나누면 안 부딪히는 대신
영원히 안 만납니다.

[.github/CODEOWNERS](.github/CODEOWNERS) 에 경로별 담당자가 적혀 있습니다.
자기 패키지를 건드리는 PR 은 자동으로 자기에게 리뷰가 옵니다.

- 내 패키지 안에서는 마음대로 해도 됩니다. 갈아엎어도 남에게 영향이 없습니다.
- 남의 패키지와 `packages/mcp-server` · `packages/ppt-bridge`(레퍼런스)는
  이슈를 먼저 열어 주세요.
- **비슷한 걸 두 사람이 각자 만드는 것은 문제가 아닙니다.** 접근이 다르면
  비교해서 배울 게 생깁니다. 합치라고 요구하지 않습니다.

작업 브랜치는 얼마든지 만들어도 됩니다. 브랜치는 싸요. 조건은 두 개뿐입니다 —
**이름이 무엇을 하는지 말해줄 것**, 그리고 **끝나면 머지되고 사라질 것**.

### 절차

```bash
git checkout main
git pull                          # 항상 최신 main에서 시작
git checkout -b feat/my-feature

# ... 작업 ...

git add -p                        # 필요한 것만 골라서 스테이징
git commit -m "feat: 무엇을 왜 바꿨는지"
git push -u origin feat/my-feature
```

그 다음 GitHub에서 **PR을 열고 리뷰를 요청**합니다. main에 직접 push 금지.

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
기존 컴포넌트를 대체하는 거라면, 지우는 게 아니라 `packages/` 아래에 **새 컴포넌트로 추가**하고
동작을 확인한 뒤 정리하는 순서로 갑니다.

---

## 5. 올리기 전 체크리스트

```bash
cd packages/ppt-bridge && .venv/bin/python -m pytest tests   # 브릿지 테스트
cd packages/mcp-server && npm ci && npm test                 # 빌드 + 테스트
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
| Node 패키지 | `packages/*` 를 훑어서 있는 것마다 빌드·테스트 |
| Python 브릿지 | 의존성 설치 + `pytest` |

빨간불이 뜨면 **Actions 탭의 Summary** 를 먼저 보세요. 무엇이 왜 걸렸고
어떻게 고치는지 적혀 있습니다. 사람한테 물어보기 전에 거기부터 읽으면 대부분 풀립니다.

새 패키지를 만들면 CI 를 고칠 필요 없습니다. `package.json` 에 `build` / `test`
스크립트만 넣어 두면 알아서 실행됩니다. **테스트가 없으면 경고가 뜹니다** —
빌드는 통과시켜 주지만, 리뷰에서 물어볼 겁니다.

---

## 7. 테스트를 어떻게 쓰나요

레퍼런스 구현에 예시가 있습니다. 새로 짤 때 이걸 보고 따라 하면 됩니다.

- [packages/ppt-bridge/tests/test_bridge.py](packages/ppt-bridge/tests/test_bridge.py) — 순수 함수 / 프로토콜 계약 / 실제 동작
- [packages/mcp-server/test/tools.test.mjs](packages/mcp-server/test/tools.test.mjs) — 진짜 MCP 클라이언트로 서버에 붙어서 tool 확인

각 테스트마다 **왜 이 테스트가 있는지** 주석을 달아 뒀습니다. 테스트를 추가할 때도
같이 적어 주세요. 3개월 뒤에 그 테스트가 깨졌을 때, 고쳐야 할지 지워야 할지를
판단할 수 있는 건 그 주석뿐입니다.
