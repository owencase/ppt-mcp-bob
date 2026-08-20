# html-ppt-mcp

HTML 기반으로 PowerPoint 프레젠테이션을 만드는 MCP 서버.

## 실행

```bash
cd packages/html-ppt-mcp
npm install
npm run build
npm start
```

## 구조

```
src/index.ts   tool 정의 + 서버 부팅
build/         tsc 산출물 (커밋하지 않음)
```

## 규칙

- 색상은 `#` 없는 `RRGGBB` hex 문자열
- 사용자 입력은 반드시 `escapeHtml()` 을 거칩니다
- 로그는 `console.error` (stderr)로. stdout은 MCP 프로토콜 전용입니다
- `node_modules/`, `build/` 는 커밋하지 않습니다
