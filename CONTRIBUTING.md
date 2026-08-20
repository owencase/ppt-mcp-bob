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
개인 이름 브랜치는 오래 살아남아 main과 걷잡을 수 없이 벌어집니다.

```
feat/html-to-pptx       새 기능
fix/theme-font-color    버그 수정
chore/repo-hygiene      설정·정리 작업
docs/readme-setup       문서
```

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
cd packages/mcp-server && npm ci && npm run build   # 빌드 통과
git status                                          # 의도한 파일만 있는지
```

- [ ] `node_modules/`, `build/`, `.pptx` 가 목록에 없다
- [ ] 커밋 메시지가 무엇을 왜 바꿨는지 말해 준다
- [ ] `main`이 아니라 작업 브랜치에 있다
