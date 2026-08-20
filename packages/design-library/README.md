# design-library

PPT 템플릿 / 테마 메타데이터.

## 중요: 바이너리는 커밋하지 않습니다

`.pptx`, `.thmx`, `.potx` 같은 원본 파일은 **이 레포에 넣지 않습니다.**
개당 수 MB ~ 수십 MB 라 히스토리에 한 번 들어가면 영구히 남고, 레포 클론이 무거워집니다.
(이전에 이것 때문에 레포가 96MB 까지 커졌습니다.)

커밋하는 것은 **메타데이터 `.json` 만** 입니다.

```
templates/back-to-school.json      ← 커밋 O
templates/back-to-school.pptx      ← 커밋 X (.gitignore 로 막혀 있음)
```

## 메타데이터 형식

```json
{
  "name": "back-to-school",
  "kind": "template",
  "source": "원본 파일을 어디서 받는지 (URL 또는 공유 드라이브 경로)",
  "slide_size": "16:9",
  "palette": { "bg": "0F172A", "text": "E2E8F0", "accent": "38BDF8" },
  "tags": ["education", "bright"]
}
```

원본 파일 보관 위치는 아직 정하지 않았습니다. 공유 드라이브 / Git LFS 중
어느 쪽으로 갈지 정해지면 `source` 에 그 경로를 적습니다.
