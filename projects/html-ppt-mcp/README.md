# html-ppt-mcp

HTML 기반으로 PowerPoint 프레젠테이션을 만드는 MCP 서버.

## 실행

```bash
cd projects/html-ppt-mcp
npm install
npm run build
npm start
```

## 구조

```
src/index.ts        tool 정의 + 서버 부팅
test/               node --test 로 실행 (npm test)
build/              tsc 산출물 (커밋하지 않음)
```

`npm test` 는 빌드를 먼저 돌린 뒤, 진짜 MCP 클라이언트로 서버에 붙어서
tool 이 실제로 노출되는지 확인합니다. tool 을 추가하면 테스트도 같이
추가해 주세요. 더 자세한 예시는 `reference/mcp-server/test/tools.test.mjs`.

## 규칙

- 색상은 `#` 없는 `RRGGBB` hex 문자열
- 사용자 입력은 반드시 `escapeHtml()` 을 거칩니다
- 로그는 `console.error` (stderr)로. stdout은 MCP 프로토콜 전용입니다
- `node_modules/`, `build/` 는 커밋하지 않습니다
