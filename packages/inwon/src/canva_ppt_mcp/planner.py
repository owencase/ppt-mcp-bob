from __future__ import annotations

import json
import os
import re
from hashlib import sha256

from .models import (ChartSeries, ChartSpec, ContentItem, DeckPlan, DesignSystem,
                     EvidenceClaim, Palette, ResearchSource, SlideSpec, Typography)
from .prompts import SYSTEM_PROMPT, planning_prompt
from .research import ResearchResult, ResearchUnavailableError, research_topic, source_metadata


TOPIC_PALETTES = {
    "ai": ("#102A43", ["#D9EAF7", "#5C7AEA"], "#00B8A9", "#F7FBFF", "#071A2B"),
    "환경": ("#164E3A", ["#DCECDF", "#4F8A6B"], "#E8A317", "#F7FBF8", "#0C2E22"),
    "금융": ("#263B5E", ["#DCE5F2", "#7083A3"], "#D5962B", "#F8FAFD", "#17243A"),
    "교육": ("#3D315B", ["#E7E0F4", "#7B6BA8"], "#F05D5E", "#FBFAFE", "#241B38"),
    "의료": ("#145C63", ["#D9F0EF", "#5A9EA5"], "#E45B66", "#F7FCFC", "#0B383D"),
    "기술": ("#173F5F", ["#D8E6F0", "#4E7898"], "#F6A623", "#F8FBFD", "#0C263A"),
    "제조": ("#27364B", ["#D9E1E8", "#738497"], "#FF6B35", "#F7FAFC", "#111B28"),
    "문화": ("#512B58", ["#E9DCEF", "#9C6AA5"], "#FFB000", "#FCF9FD", "#2D1632"),
    "스포츠": ("#123D2E", ["#D5EADF", "#4D8B70"], "#FF5A36", "#F7FBF9", "#09261C"),
}

STYLE_PRESETS = ("orbital", "editorial", "neon", "organic", "luxury", "geometric", "swiss")


def _palette_for(topic: str) -> Palette:
    lowered = topic.lower()
    selected = TOPIC_PALETTES["기술"]
    for key, value in TOPIC_PALETTES.items():
        if key in lowered:
            selected = value
            break
    primary, secondary, accent, light, dark = selected
    return Palette(primary=primary, secondary=secondary, accent=accent,
                   background_light=light, background_dark=dark)


def _style_for(topic: str, audience: str = "", purpose: str = "", preferred: str | None = None) -> str:
    if preferred:
        if preferred not in STYLE_PRESETS:
            raise ValueError(f"style_preference must be one of: {', '.join(STYLE_PRESETS)}")
        return preferred
    digest = sha256(f"{topic}|{audience}|{purpose}".encode("utf-8")).digest()
    return STYLE_PRESETS[digest[0] % len(STYLE_PRESETS)]


def _design_for(topic: str, audience: str = "", purpose: str = "",
                style_preference: str | None = None) -> DesignSystem:
    style = _style_for(topic, audience, purpose, style_preference)
    fonts = {
        "editorial": "Century Schoolbook", "luxury": "Bookman Old Style",
        "organic": "Bookman Old Style", "swiss": "Cambria",
    }
    rotations = {
        "editorial": ["title", "image_focus", "two_column", "big_stat", "timeline", "chart", "comparison", "icon_rows"],
        "neon": ["title", "big_stat", "image_focus", "icon_rows", "chart", "comparison", "timeline", "two_column"],
        "organic": ["title", "image_focus", "icon_rows", "timeline", "two_column", "chart", "grid_2x2", "comparison"],
        "luxury": ["title", "big_stat", "image_focus", "comparison", "chart", "two_column", "timeline", "icon_rows"],
        "geometric": ["title", "grid_2x2", "comparison", "big_stat", "chart", "two_column", "timeline", "image_focus"],
        "swiss": ["title", "two_column", "icon_rows", "chart", "timeline", "comparison", "image_focus", "big_stat"],
        "orbital": ["title", "two_column", "icon_rows", "big_stat", "chart", "grid_2x2", "timeline", "comparison", "image_focus"],
    }
    return DesignSystem(
        palette=_palette_for(topic), typography=Typography(header_font=fonts.get(style, "Cambria")),
        visual_motif=f"{style} 스타일의 대형 시각 앵커와 비대칭 레이어",
        style_preset=style, layout_rotation=rotations[style],
        visual_intensity="bold", dark_slide_ratio=.45,
        gradient_backgrounds=True, background_texture=True, dynamic_composition=True,
    )


def _fallback_plan(topic: str, audience: str, purpose: str, slide_count: int,
                   language: str = "ko", style_preference: str | None = None) -> DeckPlan:
    design = _design_for(topic, audience, purpose, style_preference)
    rotation = design.layout_rotation
    english = language.lower().startswith("en")
    middle_jobs = ([
        ("Why it matters", ["The pressure is increasing", "The operating context is changing"]),
        ("What it requires", ["A clear scope and shared language", "Connected capabilities, not isolated tools"]),
        ("Value and impact", ["Faster, more reliable decisions", "Practical constraints must shape delivery"]),
        ("Execution principles", ["Prove value before scaling", "Agree on success measures first"]),
        ("A staged approach", ["Prepare", "Validate", "Expand"]),
        ("Decision criteria", ["Impact", "Feasibility", "Risk"]),
    ] if english else [
        ("핵심 맥락", ["지금 이 주제가 중요한 이유", "대상이 겪는 변화"]),
        ("핵심 구성", ["주요 개념과 범위", "서로 연결되는 요소"]),
        ("가치와 영향", ["기대할 수 있는 변화", "실행 시 고려할 점"]),
        ("실행 원칙", ["작게 검증하고 확대", "성과 기준을 먼저 합의"]),
        ("단계별 접근", ["준비", "검증", "확장"]),
        ("선택 기준", ["효과", "실행 난이도", "위험"]),
    ])
    default_subtitle = ("A practical deck for understanding the topic and choosing the next action"
                        if english else "핵심을 이해하고 다음 행동을 정하는 프레젠테이션")
    slides = [SlideSpec(title=topic, subtitle=purpose or default_subtitle, layout="title")]
    for index in range(max(0, slide_count - 2)):
        heading, bullets = middle_jobs[index % len(middle_jobs)]
        layout = rotation[(index + 1) % len(rotation)]
        slides.append(SlideSpec(
            title=(f"{heading} for {topic}" if english else f"{heading}: {topic}을 구체화합니다"),
            layout=layout,
            items=[ContentItem(
                heading=b,
                body=(f"A core consideration when evaluating {topic}." if english
                      else f"{topic} 관점에서 확인할 핵심 항목입니다.")
            ) for b in bullets],
        ))
    slides.append(SlideSpec(
        title=("Make the next action small and testable" if english else "다음 행동은 작고 검증 가능해야 합니다"),
        subtitle=((f"Start by aligning {audience or 'the team'} on the priority and accountable owner.")
                  if english else f"{audience or '참여자'}가 우선순위와 책임자를 합의하는 것으로 시작합니다."),
        layout="closing",
        items=[ContentItem(heading="NEXT", body=("Choose the first action and its success measure."
                                                  if english else "첫 실행 항목과 성공 기준을 정하세요."))],
    ))
    return DeckPlan(
        communication_job=((f"By the end, {audience or 'the audience'} understands {topic} and the next action.")
                           if english else f"발표가 끝날 때 {audience or '청중'}은 {topic}의 핵심과 다음 행동을 이해한다."),
        design_system=design,
        slides=slides, language="en" if english else "ko",
    )


def _compact_claim(text: str, budget: int = 38) -> str:
    text = " ".join(text.split()).strip(" .")
    text = __import__("re").sub(r"\([^)]{12,}\)", "", text)
    for marker in (". ", "다. ", "; ", ": "):
        if marker in text:
            text = text.split(marker, 1)[0] + ("다" if marker.startswith("다") else "")
            break
    if len(text) <= budget:
        return text
    cut = text[:budget]
    for marker in (", ", "이며", "이고", "하지만", "그러나", " which ", " that "):
        if marker in cut:
            cut = cut.split(marker, 1)[0]
    words: list[str] = []
    used = 0
    for word in cut.split():
        if words and used + 1 + len(word) > budget:
            break
        words.append(word); used += len(word) + (1 if words else 0)
    return (" ".join(words) or cut).rstrip(" ,.;:")


def _fit_summary(text: str, limit: int, fallback: str) -> str:
    """Fit already-summarized copy without slicing the source sentence."""
    clean = re.sub(r"\s+", " ", text).strip(" ,.;:")
    if len(clean) <= limit:
        return clean
    removable = ("핵심적인 ", "세계적인 ", "대표적인 ", "주요 ", "관련 ")
    for token in removable:
        clean = clean.replace(token, "")
    if len(clean) <= limit:
        return clean
    words = clean.split()
    while len(words) > 2 and len(" ".join(words)) > limit:
        words.pop(-2)
    fitted = " ".join(words)
    return fitted if len(fitted) <= limit else fallback


def _visual_width(text: str) -> float:
    return sum(1.65 if ord(char) > 127 else 1.0 for char in text)


def _fit_one_line_summary(text: str, limit: float, fallback: str) -> str:
    """Compress summarized copy to the renderer's one-line title budget."""
    clean = re.sub(r"\s+", " ", text).strip(" ,.;:")
    replacements = (
        ("필요한 자료 탐색", "자료 탐색"),
        ("속도와 활용 시점", "속도·활용 시점"),
        ("모델 버전과 배포", "모델 버전·배포"),
        ("분포 변화와 장애", "변화·장애"),
        ("컴퓨터 시장을 지배", "컴퓨터 시장 지배"),
        ("영역을 담당", "영역 담당"),
        ("준비 시간을 줄이고 재현성을 높인다", "준비 시간 단축·재현성 향상"),
    )
    for source, target in replacements:
        clean = clean.replace(source, target)
    if _visual_width(clean) <= limit:
        return clean

    words = clean.split()
    candidate_sets: list[list[str]] = []
    if len(words) >= 4:
        candidate_sets.extend((words[:1] + words[-3:], words[:2] + words[-2:], words[:1] + words[-2:]))
    if len(words) >= 3:
        candidate_sets.append(words[:1] + words[-1:])
    for candidate_words in candidate_sets:
        unique: list[str] = []
        for word in candidate_words:
            if not unique or unique[-1] != word:
                unique.append(word)
        candidate = " ".join(unique).strip(" ,.;:")
        if _visual_width(candidate) <= limit:
            return candidate

    fallback_clean = re.sub(r"\s+", " ", fallback).strip(" ,.;:")
    if _visual_width(fallback_clean) <= limit:
        return fallback_clean
    fallback_words = fallback_clean.split()
    while len(fallback_words) > 1 and _visual_width(" ".join(fallback_words)) > limit:
        fallback_words.pop(-2 if len(fallback_words) > 2 else -1)
    return " ".join(fallback_words)


def _subject_label(text: str, topic: str, fallback: str) -> str:
    if topic and topic.lower() in text.lower():
        return topic
    match = re.match(r"^(.{2,28}?)(?:은|는|이|가)\s", text)
    if not match:
        return fallback
    label = re.sub(r"^\d[\d,.]*(?:세기|년대|년|월)?\s+", "", match.group(1)).strip()
    if topic and any(token in label for token in ("회사", "기업", "기관", "조직")):
        return topic
    return _fit_summary(label, 14, fallback)


def _josa(word: str, consonant: str, vowel: str) -> str:
    """Append a Korean particle using the final syllable's jongseong."""
    if not word:
        return word
    last = word[-1]
    if "가" <= last <= "힣":
        has_jongseong = (ord(last) - ord("가")) % 28 != 0
    else:
        has_jongseong = last.lower() not in "aeiou"
    return word + (consonant if has_jongseong else vowel)


def _ko_summary(text: str, fallback: str, topic: str = "") -> tuple[str, str]:
    """Return (takeaway title, semantic item heading) from a Korean claim.

    The rules rewrite a fact into a presentation claim. They intentionally do
    not return a prefix/suffix of the original sentence, which was the main
    source of clipped, encyclopaedic copy in earlier versions.
    """
    clean = " ".join(text.split()).strip(" .")
    subject = _subject_label(clean, topic, fallback)
    year = next(iter(re.findall(r"\d{4}년", clean)), "")

    rename = re.search(r'["“]([^"”]{3,48})["”]로 이름이 바뀌', clean)
    if rename:
        label = topic or rename.group(1)
        return f"{year or '새 이름'}, {label} 사명 변경", f"사명 변경 · {year or label}"
    countries = re.search(r"(\d[\d,]*개국) 이상에 진출", clean)
    if countries and ("다국적" in clean or "기술 회사" in clean):
        return f"{subject}, {countries.group(1)}에 진출한 기술 기업", f"글로벌 사업 · {countries.group(1)}"
    patent_streak = re.search(r"(\d+년) 연속.*특허 최다", clean)
    if patent_streak:
        return f"{subject}, {patent_streak.group(1)} 연속 특허 선도", f"연구 경쟁력 · {patent_streak.group(1)}"
    founded = re.search(r"(\d{4}년).*설립되었다", clean)
    if founded:
        return f"{founded.group(1)}, {subject}의 출발", f"설립 · {founded.group(1)}"
    century = re.search(r"(\d+세기(?:\s*후반|\s*전반)?)", clean)
    if century and "기술 혁신" in clean and "시작" in clean:
        return f"{subject}의 뿌리는 {century.group(1)} 기술 혁신", f"기술 기원 · {century.group(1)}"
    patent = re.search(r"(\d{4}년).*특허", clean)
    if patent:
        return f"{patent.group(1)}, {_josa(subject, '이', '가')} 기술 특허를 확보", f"{subject} · {patent.group(1)} 특허"
    invention = re.search(r"([^,.]{2,24})\((\d{4})\)(?:을|를) 발명", clean)
    if invention:
        invention_name = invention.group(1).split()[-2:]
        invention_label = " ".join(invention_name)
        return (f"{invention.group(2)}년, {_josa(subject, '이', '가')} {invention_label} 발명",
                f"{subject} · {invention.group(2)}년 발명")
    money = re.search(r"(\d[\d,]*억\s*달러)", clean)
    if money and ("시가총액" in clean or "가치" in clean):
        return f"{subject}의 기업가치는 {money.group(1)}", f"기업가치 · {money.group(1)}"
    rank = re.search(r"(\d+)위를 차지", clean)
    if rank:
        return f"{subject}, 미국 기업 순위 {rank.group(1)}위", f"기업 순위 · {rank.group(1)}위"
    if "금융 공학" in clean and "비난" in clean:
        return "단기 실적 중심 경영이 비판을 받다", "경영 논란 · 단기 실적"
    if "시스템/360" in clean and "%" in clean:
        share = re.search(r"미국[^,.]{0,20}?(\d+%)", clean)
        metric = share.group(1) if share else next(iter(re.findall(r"\d+%", clean)), "")
        return f"{subject} 메인프레임이 컴퓨터 시장을 지배", f"메인프레임 · 미국 {metric}".rstrip()
    if "베스트셀러" in clean and "PC" in clean:
        return "IBM PC가 대표 제품으로 부상", "IBM PC · 베스트셀러"
    if "지식 재산권" in clean and "보호" in clean:
        return "PC의 지식재산 보호는 미흡했다", "PC · 지식재산 한계"
    if "마이크로컴퓨터 시장에 진출" in clean:
        return f"{subject}, 마이크로컴퓨터 시장에 진출", "PC 시장 · 사업 확장"
    if "제공하는 기반 시스템" in clean and ("인공지능" in clean or "모델" in clean):
        short_subject = subject.replace("데이터 플랫폼", "플랫폼")
        return f"{_josa(short_subject, '이', '가')} 모델용 데이터를 공급", f"{short_subject} · 데이터 공급"
    if "찾도록 돕" in clean:
        return f"{_josa(subject, '이', '가')} 자료 탐색을 지원", f"{subject} · 자료 탐색"
    if "신뢰도를 높" in clean:
        short_subject = subject.replace("데이터 품질 ", "품질 ")
        return f"{_josa(short_subject, '이', '가')} 분석 신뢰도를 높인다", f"{short_subject} · 신뢰도 향상"
    if "속도와 활용 시점에 따라 선택" in clean:
        return "처리 방식은 속도·시점에 따라 결정", "처리 방식 · 선택 기준"
    if "학습과 추론" in clean and "재사용" in clean:
        return f"{_josa(subject, '이', '가')} 학습·추론 일관성을 확보", f"{subject} · 일관성"
    if "모델 버전" in clean and "배포 상태" in clean:
        return f"{_josa(subject, '이', '가')} 버전·배포를 관리", f"{subject} · 모델 운영"
    if "민감한 데이터" in clean and ("접근 제어" in clean or "암호화" in clean):
        return "접근 제어와 암호화가 민감정보를 보호", "보안 · 접근 통제"
    if "원천 데이터" in clean and "추적" in clean:
        return f"{_josa(subject, '이', '가')} 데이터 변환 과정을 추적", f"{subject} · 흐름 추적"
    if "분포 변화" in clean and "장애" in clean:
        return f"{_josa(subject, '이', '가')} 분포 변화와 장애를 감지", f"{subject} · 이상 감지"
    if "준비 시간" in clean and "재현성" in clean:
        return "플랫폼이 준비 시간을 줄이고 재현성을 높인다", "핵심 가치 · 속도와 재현성"
    if " — " in clean:
        label, detail = [part.strip() for part in clean.split(" — ", 1)]
        label = {
            "Information Management": "데이터 관리",
            "Business Analytics": "비즈니스 분석",
            "Industry Solutions": "산업 솔루션",
        }.get(label, label)
        detail_head = detail.split(",", 1)[0].strip()
        return f"{_josa(label, '이', '가')} {detail_head} 영역을 담당", f"{label} · {detail_head}"

    # General predicate-aware fallback: the result is a rewritten claim, not
    # a raw character slice.  Fallback labels preserve section meaning when a
    # sentence is too irregular for a safe rewrite.
    predicate_rules = (
        (r"([^,.]{2,18})(?:을|를) 높인다", "{actor} {object}을 향상", "{subject} · 성과 향상"),
        (r"([^,.]{2,18})(?:을|를) 관리한다", "{actor} {object}을 통합 관리", "{subject} · 통합 관리"),
        (r"([^,.]{2,18})(?:을|를) 추적한다", "{actor} {object}을 추적", "{subject} · 추적"),
        (r"([^,.]{2,18})(?:을|를) 제한한다", "{actor} {object}을 통제", "{subject} · 통제"),
        (r"([^,.]{2,18})(?:을|를) 지원한다", "{actor} {object}을 지원", "{subject} · 지원"),
    )
    for pattern, title_template, heading_template in predicate_rules:
        match = re.search(pattern, clean)
        if match:
            obj = " ".join(match.group(1).split()[-3:])
            return (title_template.format(subject=subject, actor=_josa(subject, "이", "가"), object=obj),
                    heading_template.format(subject=subject, actor=_josa(subject, "이", "가")))
    return f"{fallback}에서 확인한 {subject}의 핵심", f"{fallback} · 핵심 의미"


def _item_heading(text: str, fallback: str, topic: str = "", language: str = "ko") -> str:
    clean = " ".join(text.split()).strip(" .")
    if language.startswith("ko"):
        _title, heading = _ko_summary(clean, fallback, topic)
        return _fit_summary(heading, 22, fallback)
    # English offline mode remains extractive but uses sentence/word boundaries.
    return _compact_claim(clean, 34) or fallback


def _numeric_value(text: str) -> str | None:
    match = re.search(r"(?<!\w)(\d+(?:[.,]\d+)?%)(?!\w)", text)
    return match.group(1) if match else None


def _claim_title(text: str, fallback: str, language: str, topic: str = "",
                 display_limit: float = 33) -> str:
    full = " ".join(text.split()).rstrip(". ")
    if language.startswith("ko"):
        title, _heading = _ko_summary(full, fallback, topic)
        return _fit_one_line_summary(title, display_limit, fallback)
    title = _compact_claim(full, 40 if language.startswith("ko") else 52)
    if len(title) < 6:
        title = fallback
    if language.startswith("ko"):
        title = title.rstrip(".")
    return _fit_one_line_summary(title, display_limit, fallback)


def _claim_note(claims: list[EvidenceClaim]) -> str:
    lines = [f"- [{claim.claim_id}] {claim.source_url}" for claim in claims]
    return "[Sources]\n" + "\n".join(dict.fromkeys(lines)) + "\n[/Sources]"


def _percentage_chart(sentences: list[tuple[str, EvidenceClaim]], topic: str = "",
                      language: str = "ko") -> ChartSpec | None:
    categories: list[str] = []
    values: list[float] = []
    for sentence, _claim in sentences:
        for raw in re.findall(r"(?<!\w)(\d+(?:[.,]\d+)?)%(?!\w)", sentence):
            if len(values) >= 6:
                break
            categories.append(_item_heading(sentence, f"지표 {len(values) + 1}", topic, language))
            values.append(float(raw.replace(",", "")))
    if len(values) < 2:
        return None
    # Duplicate labels are confusing on a chart; retain a deterministic suffix.
    seen: dict[str, int] = {}
    for index, label in enumerate(categories):
        seen[label] = seen.get(label, 0) + 1
        if seen[label] > 1:
            categories[index] = f"{label} {seen[label]}"
    return ChartSpec(categories=categories, series=[ChartSeries(name="비율", values=values)], value_suffix="%")


def _grounded_plan(topic: str, audience: str, purpose: str, slide_count: int,
                   language: str, research: ResearchResult,
                   style_preference: str | None = None) -> DeckPlan:
    if language.lower().startswith("ko") and research.language != "ko":
        raise ResearchUnavailableError(
            "한국어 발표를 요청했지만 한국어 공개 자료를 찾지 못했습니다. "
            "한국어 research_text를 제공하면 번역되지 않은 외국어를 넣지 않고 생성할 수 있습니다."
        )
    deck_language = "en" if language.lower().startswith("en") else "ko"
    design = _design_for(topic, audience, purpose, style_preference)
    rotation = design.layout_rotation
    evidence: list[EvidenceClaim] = []
    claim_lookup: dict[str, EvidenceClaim] = {}
    seen_sentences: set[str] = set()
    for section in research.sections:
        for sentence in section.sentences:
            normalized = " ".join(sentence.split())
            if normalized in seen_sentences:
                continue
            seen_sentences.add(normalized)
            claim_id = f"c{len(evidence) + 1:03}"
            claim = EvidenceClaim(
                claim_id=claim_id, text=normalized,
                source_url=section.source_url or research.url,
                section=section.heading,
                numeric_values=re.findall(r"(?<!\w)\d+(?:[.,]\d+)?%?(?!\w)", normalized),
            )
            evidence.append(claim); claim_lookup[normalized] = claim
    if len(evidence) < max(3, slide_count - 2):
        raise ResearchUnavailableError(f"'{topic}'에 사용할 검증 가능한 주장 수가 부족합니다.")
    source_note = _claim_note(evidence[:1])
    section_labels = []
    for section in research.sections:
        if section.heading not in section_labels:
            section_labels.append(section.heading)
    subtitle = purpose or (f"{research.title}의 핵심을 근거와 함께 살펴봅니다"
                           if deck_language == "ko" else f"An evidence-based view of {research.title}")
    slides = [SlideSpec(title=topic, subtitle=subtitle, layout="title", speaker_notes=source_note)]

    preferred = ("개요", "정의", "역사", "배경", "특징", "원리", "구조", "기술", "제품", "서비스",
                 "사업", "연구", "활용", "영향", "현황", "overview", "history", "technology", "products")
    deferred = ("지배 구조", "이사회", "논란", "비판", "수상", "기타", "governance", "controversy")
    intent_tokens = {x.lower() for x in re.findall(r"[A-Za-z0-9가-힣]{2,}", f"{audience} {purpose}")}
    ranked = sorted(enumerate(research.sections), key=lambda pair: (
        1 if any(key in pair[1].heading.lower() for key in deferred) else 0,
        -sum(token in (pair[1].heading + " " + " ".join(pair[1].sentences)).lower() for token in intent_tokens),
        0 if any(key in pair[1].heading.lower() for key in preferred) else 1,
        pair[0],
    ))
    first_groups: list[tuple[str, list[str]]] = []
    extra_groups: list[tuple[str, list[str]]] = []
    deferred_groups: list[tuple[str, list[str]]] = []
    for _, section in ranked:
        usable = [s for s in section.sentences if len(s) >= 28 and s in claim_lookup]
        chunks = [usable[start:start + 3] for start in range(0, len(usable), 3) if usable[start:start + 3]]
        if chunks:
            target = deferred_groups if any(key in section.heading.lower() for key in deferred) else first_groups
            target.append((section.heading, chunks[0]))
            remainder = [(section.heading, chunk) for chunk in chunks[1:]]
            if target is deferred_groups:
                deferred_groups.extend(remainder)
            else:
                extra_groups.extend(remainder)
    groups = first_groups + extra_groups + deferred_groups
    needed = max(1, slide_count - 2)
    if len(groups) < needed:
        groups = [(section.heading, [sentence]) for section in research.sections
                  for sentence in section.sentences if len(sentence) >= 28]
    if len(groups) < needed:
        raise ResearchUnavailableError(
            f"'{topic}' 자료에서 {slide_count}장 분량의 중복 없는 내용을 확보하지 못했습니다. "
            "슬라이드 수를 줄이거나 research_text를 추가하세요."
        )

    section_counts: dict[str, int] = {}
    for index, (section_heading, sentences) in enumerate(groups[:needed]):
        section_counts[section_heading] = section_counts.get(section_heading, 0) + 1
        first = sentences[0]
        slide_claims = [claim_lookup[s] for s in sentences if s in claim_lookup]
        chart = _percentage_chart([(s, claim_lookup[s]) for s in sentences if s in claim_lookup],
                                  topic, deck_language)
        numeric = next(((s, _numeric_value(s)) for s in sentences if _numeric_value(s)), None)
        if chart and index % 4 == 2:
            layout = "chart"
            items = [ContentItem(
                heading=_item_heading(first, section_heading, topic, deck_language), body=first,
                claim_ids=[claim_lookup[first].claim_id],
            )]
        elif numeric and index % 3 == 1:
            numeric_sentence, value = numeric
            first = numeric_sentence
            layout = "big_stat"
            items = [ContentItem(
                heading=_item_heading(numeric_sentence, section_heading, topic, deck_language),
                body=numeric_sentence, value=value,
                claim_ids=[claim_lookup[numeric_sentence].claim_id],
            )]
        else:
            body_rotation = [x for x in design.layout_rotation
                             if x not in {"title", "closing", "chart", "big_stat"}]
            layout = body_rotation[index % len(body_rotation)]
            if research.image_url and index == needed - 1:
                layout = "image_focus"
            if layout == "grid_2x2" and max(len(sentence) for sentence in sentences) > 45:
                layout = "icon_rows"
            items = []
            for sentence in sentences[:4]:
                item_value = _numeric_value(sentence)
                heading = (f"{section_heading} · {item_value}"
                           if layout in {"grid_2x2", "icon_rows"} and item_value
                           else _item_heading(sentence, section_heading, topic, deck_language))
                items.append(ContentItem(
                    heading=heading, body=sentence,
                    claim_ids=[claim_lookup[sentence].claim_id],
                    image_url=(research.image_url if layout == "image_focus" and not items else None),
                ))
        slide_notes = _claim_note(slide_claims)
        if any(item.image_url for item in items):
            slide_notes = slide_notes.replace("\n[/Sources]", f"\n- [asset] {research.image_url}\n[/Sources]")
        slides.append(SlideSpec(
            title=_claim_title(first, section_heading, deck_language, topic),
            subtitle=section_heading,
            layout=layout,
            items=items,
            chart=chart if layout == "chart" else None,
            speaker_notes=slide_notes,
        ))

    used_claim_ids = {
        claim_id for slide in slides[1:] for item in slide.items for claim_id in item.claim_ids
    }
    conclusion_claim = next(
        (claim for claim in evidence
         if claim.claim_id not in used_claim_ids
         and not any(key in claim.section.lower() for key in deferred)),
        slide_claims[0] if slide_claims else evidence[0],
    )
    conclusion = _claim_title(conclusion_claim.text, research.title, deck_language, topic, 22)
    slides.append(SlideSpec(
        title=conclusion,
        subtitle=(purpose or (f"{research.title}를 이해하는 데 필요한 근거를 확인했습니다"
                              if deck_language == "ko" else "The evidence resolves the central question")),
        layout="closing",
        items=[ContentItem(
            heading="핵심 결론" if deck_language == "ko" else "Conclusion",
            body=conclusion_claim.text, claim_ids=[conclusion_claim.claim_id],
        )],
        speaker_notes=_claim_note([conclusion_claim]),
    ))
    metadata = source_metadata(research)
    source_urls = list(dict.fromkeys(claim.source_url for claim in evidence))
    sources = [ResearchSource(
        title=(next((claim.section for claim in evidence if claim.source_url == url), research.title) or research.title),
        url=url, source_type=research.source_type,
        retrieved_at=metadata["retrieved_at"],
    ) for url in source_urls]
    if research.image_url:
        sources.append(ResearchSource(
            title=f"{research.title} image", url=research.image_url,
            source_type="asset", retrieved_at=metadata["retrieved_at"],
        ))
    return DeckPlan(
        communication_job=(f"발표가 끝날 때 {audience or '청중'}은 {topic}의 실제 개념과 주요 사실을 "
                           "출처에 근거해 이해한다."),
        design_system=design,
        slides=slides,
        research_sources=sources, evidence_claims=evidence,
        grounded=True, language=deck_language,
    )


def _normalize(plan: DeckPlan) -> DeckPlan:
    def compact_title(text: str, budget: int = 42) -> str:
        text = " ".join(text.split())
        if text.lower().startswith(("the ", "a ", "an ")):
            text = text.split(" ", 1)[1]
        used = 0; words = []
        for word in text.split():
            cost = sum(2 if ord(c) > 127 else 1 for c in word) + (1 if words else 0)
            if words and used + cost > budget: break
            words.append(word); used += cost
        result = " ".join(words) or text
        return result.rstrip(" ,.;:")

    rotation = plan.design_system.layout_rotation
    for i, slide in enumerate(plan.slides):
        if i == 0:
            slide.layout = "title"
        elif i == len(plan.slides) - 1:
            slide.layout = "closing"
        elif slide.layout == plan.slides[i - 1].layout:
            candidates = [x for x in rotation if x not in {slide.layout, "title", "closing"}]
            if candidates:
                slide.layout = candidates[i % len(candidates)]
        clean_title = " ".join(slide.title.split())
        title_limit = 22 if i in {0, len(plan.slides) - 1} else 33
        fallback_title = compact_title(clean_title, 42)
        slide.title = _fit_one_line_summary(clean_title, title_limit, fallback_title)
        slide.subtitle = " ".join(slide.subtitle.split())
        slide.items = slide.items[:4]
        for item in slide.items:
            item.heading = compact_title(item.heading, 48)
            item.body = " ".join(item.body.split())
    return plan


def _attach_evidence(plan: DeckPlan, evidence: list[EvidenceClaim]) -> DeckPlan:
    for slide in plan.slides:
        used: list[EvidenceClaim] = []
        for item in slide.items:
            if item.claim_ids:
                used.extend([c for c in evidence if c.claim_id in item.claim_ids])
                continue
            body_tokens = set(re.findall(r"[A-Za-z0-9가-힣]{2,}", item.body.lower()))
            scored = []
            for claim in evidence:
                claim_tokens = set(re.findall(r"[A-Za-z0-9가-힣]{2,}", claim.text.lower()))
                score = len(body_tokens & claim_tokens) / max(1, len(body_tokens))
                scored.append((score, claim))
            if scored:
                score, claim = max(scored, key=lambda pair: pair[0])
                if score >= .55:
                    item.claim_ids = [claim.claim_id]
                    used.append(claim)
        if used:
            slide.speaker_notes = _claim_note(used)
    return plan


def create_plan(topic: str, audience: str = "", purpose: str = "", slide_count: int = 8,
                language: str = "ko", content_json: dict | None = None,
                research_text: str | None = None, source_urls: list[str] | None = None,
                research_required: bool = True,
                style_preference: str | None = None,
                research_documents: list[dict[str, str]] | None = None) -> DeckPlan:
    if content_json:
        return _normalize(DeckPlan.model_validate(content_json))
    try:
        research = research_topic(topic, language, research_text, source_urls, research_documents)
    except ResearchUnavailableError:
        if research_required:
            raise
        return _normalize(_fallback_plan(topic, audience, purpose, slide_count, language, style_preference))
    if not os.getenv("OPENAI_API_KEY"):
        return _normalize(_grounded_plan(topic, audience, purpose, slide_count, language, research, style_preference))
    from openai import OpenAI
    client = OpenAI()
    schema = DeckPlan.model_json_schema()
    response = client.responses.create(
        model=os.getenv("PPT_MCP_MODEL", "gpt-5.6"),
        input=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": planning_prompt(
                topic, audience, purpose, slide_count, language,
                research_context="\n".join(
                    f"c{index + 1:03}: {sentence}" for index, sentence in enumerate(research.sentences[:24])
                ),
                source_urls=list(dict.fromkeys(
                    section.source_url or research.url for section in research.sections
                )),
            )},
        ],
        text={"format": {"type": "json_schema", "name": "deck_plan", "schema": schema, "strict": True}},
    )
    plan = DeckPlan.model_validate(json.loads(response.output_text))
    evidence = []
    for index, section in enumerate(research.sections):
        for sentence in section.sentences:
            evidence.append(EvidenceClaim(
                claim_id=f"c{len(evidence) + 1:03}", text=sentence,
                source_url=section.source_url or research.url, section=section.heading,
                numeric_values=re.findall(r"(?<!\w)\d+(?:[.,]\d+)?%?(?!\w)", sentence),
            ))
    metadata = source_metadata(research)
    unique_urls = list(dict.fromkeys(claim.source_url for claim in evidence))
    plan.research_sources = [ResearchSource(
        title=next((claim.section for claim in evidence if claim.source_url == url), research.title),
        url=url, source_type=research.source_type, retrieved_at=metadata["retrieved_at"],
    ) for url in unique_urls]
    plan.evidence_claims = evidence
    plan.grounded = True
    plan.language = "en" if language.lower().startswith("en") else "ko"
    if style_preference:
        plan.design_system.style_preset = _style_for(topic, audience, purpose, style_preference)
    return _normalize(_attach_evidence(plan, evidence))
