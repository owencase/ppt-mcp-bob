# html-render-pptx

HTML/CSS로 슬라이드를 렌더링해 PPTX로 내보내는 MCP 서버.

## 실행

```bash
cd projects/html-render-pptx
npm install
npm test
```

## 구조

```
src/themes.mjs   테마(색상) 정의
src/render.mjs   슬라이드 → HTML 렌더링
src/server.mjs   MCP 서버 (아직 없음 — 여기부터 만드시면 됩니다)
test/            node --test 로 실행
```

`npm start` 는 `src/server.mjs` 를 실행합니다. 아직 그 파일이 없어서 지금은
실패합니다. `reference/mcp-server/src/index.ts` 의 부팅 부분을 참고하세요.

## 규칙

- 색상은 `#` 없는 `RRGGBB` hex 문자열
- 사용자 입력은 반드시 `escapeHtml()` 을 거칩니다
- `node_modules/`, 생성된 `.pptx` 는 커밋하지 않습니다
