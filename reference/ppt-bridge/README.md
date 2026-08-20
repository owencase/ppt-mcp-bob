# ppt-bridge

python-pptx 로 `.pptx` 파일을 직접 읽고 쓰는 브릿지.
`reference/mcp-server` 가 tool 호출마다 이 스크립트를 프로세스로 띄웁니다.

## 프로토콜

stdin/stdout 으로 JSON 한 줄씩 주고받습니다.

```
stdin  → {"action": "<이름>", "params": {...}}
stdout → {"success": true,  "message": "...", "data": {...}}
         {"success": false, "error": "..."}
```

직접 호출해서 테스트할 수 있습니다:

```bash
echo '{"action":"create_presentation","params":{"file_path":"test.pptx"}}' \
  | python3 bridge.py
```

## 액션 추가하는 법

1. 핸들러 함수를 만듭니다. 성공하면 `ok()`, 실패하면 `err()` 를 반환합니다.

```python
def handle_add_table(params: dict) -> dict:
    file_path: str = params["file_path"]
    slide_index: int = int(params["slide_index"])

    prs = load_prs(file_path)
    slide = prs.slides[slide_index]
    # ... 실제 작업 ...
    save_prs(prs, file_path)
    return ok("Table added")
```

2. `HANDLERS` 딕셔너리에 등록합니다.

```python
HANDLERS = {
    ...
    "add_table": handle_add_table,
}
```

3. `reference/mcp-server/src/index.ts` 에 대응하는 tool 을 등록합니다.
   여기까지 해야 LLM 클라이언트에서 보입니다.

## 규칙

- 위치·크기 단위는 **cm** 입니다. 내부에서 `cm()` 으로 EMU 변환합니다 (1cm = 360,000 EMU)
- 색상은 `#` 없는 **RRGGBB** hex 문자열입니다
- 도형 종류는 매직넘버 대신 `MSO_SHAPE` 열거형을 씁니다
- 예외는 `main()` 에서 잡아 JSON 으로 돌려줍니다. 스택트레이스를 stdout 에 흘리면 프로토콜이 깨집니다
- 생성된 `.pptx` 는 커밋하지 않습니다

## 설치

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```
