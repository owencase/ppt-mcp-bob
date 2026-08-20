from __future__ import annotations

import re

from .models import DeckPlan, QAIssue


GENERIC_FILLER = (
    "관점에서 확인할 핵심 항목입니다",
    "지금 이 주제가 중요한 이유",
    "대상이 겪는 변화",
    "주요 개념과 범위",
    "서로 연결되는 요소",
    "기대할 수 있는 변화",
    "실행 시 고려할 점",
    "a core consideration when evaluating",
    "the operating context is changing",
    "practical constraints must shape delivery",
)


def inspect_grounding(plan: DeckPlan, topic: str, research_required: bool = True) -> list[QAIssue]:
    issues: list[QAIssue] = []
    if research_required and not plan.grounded and not plan.research_sources:
        issues.append(QAIssue(
            slide=1, code="UNGROUNDED_CONTENT",
            message="실제 조사 자료나 사용자가 제공한 콘텐츠 없이 덱을 생성하려고 했습니다.",
        ))
    full_text = "\n".join(
        part for slide in plan.slides
        for part in [slide.title, slide.subtitle, *(f"{i.heading} {i.body}" for i in slide.items)]
    ).lower()
    for phrase in GENERIC_FILLER:
        if phrase.lower() in full_text:
            issues.append(QAIssue(
                slide=1, code="GENERIC_FILLER",
                message=f"주제와 무관한 범용 문구가 남아 있습니다: {phrase}",
            ))
    topic_tokens = [token.lower() for token in re.findall(r"[A-Za-z0-9가-힣]{2,}", topic)]
    if research_required and topic_tokens and not any(token in full_text for token in topic_tokens):
        issues.append(QAIssue(
            slide=1, code="TOPIC_NOT_PRESENT",
            message="생성된 내용에서 입력 주제를 확인할 수 없습니다.",
        ))
    for index, slide in enumerate(plan.slides, 1):
        title_width = sum(1.65 if ord(char) > 127 else 1 for char in slide.title)
        one_line_limit = 22 if index in {1, len(plan.slides)} else 33
        if title_width > one_line_limit:
            issues.append(QAIssue(
                slide=index, code="SLIDE_TITLE_WRAP_RISK",
                message="슬라이드 제목이 한 줄 너비를 넘습니다. 의미를 유지하며 더 짧게 요약해야 합니다.",
            ))
    if plan.research_sources:
        claim_map = {claim.claim_id: claim for claim in plan.evidence_claims}
        for index, slide in enumerate(plan.slides, 1):
            if "[Sources]" not in slide.speaker_notes or "[/Sources]" not in slide.speaker_notes:
                issues.append(QAIssue(
                    slide=index, code="SOURCE_NOTES_MISSING",
                    message="조사된 사실을 사용했지만 speaker notes에 출처 블록이 없습니다.",
                ))
            used_ids = [claim_id for item in slide.items for claim_id in item.claim_ids]
            if index not in {1, len(plan.slides)} and slide.items and not used_ids:
                issues.append(QAIssue(
                    slide=index, code="CLAIM_SOURCE_MISSING",
                    message="본문 주장과 연결된 claim_id가 없습니다.",
                ))
            for item in slide.items:
                if plan.language == "ko" and len(item.heading) > 22:
                    issues.append(QAIssue(
                        slide=index, code="ITEM_HEADING_TOO_LONG",
                        message="항목 소제목이 22자를 넘습니다. 본문의 의미를 더 짧게 요약해야 합니다.",
                    ))
                if (plan.language == "ko" and len(item.body) >= 32 and len(item.heading) >= 8
                        and item.body.startswith(item.heading)):
                    issues.append(QAIssue(
                        slide=index, code="EXTRACTIVE_ITEM_HEADING",
                        message="항목 소제목이 본문 앞부분을 그대로 잘라 사용했습니다.",
                    ))
                if plan.language == "ko" and item.heading.endswith(
                        (" 대한", " 위한", " 있는", " 하는", " 되고", "이며", "이고", "으로")):
                    issues.append(QAIssue(
                        slide=index, code="INCOMPLETE_ITEM_HEADING",
                        message="항목 소제목이 문장 중간에서 끊긴 형태입니다.",
                    ))
                if not item.body or not item.claim_ids:
                    continue
                body_tokens = set(re.findall(r"[A-Za-z0-9가-힣]{2,}", item.body.lower()))
                supported = False
                for claim_id in item.claim_ids:
                    claim = claim_map.get(claim_id)
                    if not claim:
                        issues.append(QAIssue(
                            slide=index, code="UNKNOWN_CLAIM_ID",
                            message=f"존재하지 않는 근거 ID를 참조합니다: {claim_id}",
                        ))
                        continue
                    claim_tokens = set(re.findall(r"[A-Za-z0-9가-힣]{2,}", claim.text.lower()))
                    overlap = len(body_tokens & claim_tokens) / max(1, len(body_tokens))
                    if item.body in claim.text or claim.text in item.body or overlap >= .55:
                        body_numbers = set(re.findall(r"\d+(?:[.,]\d+)?%?", item.body))
                        claim_numbers = set(re.findall(r"\d+(?:[.,]\d+)?%?", claim.text))
                        if body_numbers <= claim_numbers:
                            supported = True
                    if claim.source_url not in slide.speaker_notes:
                        issues.append(QAIssue(
                            slide=index, code="CLAIM_URL_MISSING",
                            message=f"근거 {claim_id}의 URL이 speaker notes에 없습니다.",
                        ))
                if not supported:
                    issues.append(QAIssue(
                        slide=index, code="UNSUPPORTED_CLAIM",
                        message="본문과 연결된 출처 문장 간 의미·수치 일치도가 부족합니다.",
                    ))
            if index not in {1}:
                for item in slide.items:
                    if (len(item.body) >= 40 and len(slide.title) >= 12
                            and item.body.startswith(slide.title)):
                        issues.append(QAIssue(
                            slide=index, code="EXTRACTIVE_SLIDE_TITLE",
                            message="슬라이드 제목이 본문 일부를 그대로 잘라 사용했습니다.",
                        ))
                        break
        if len({source.url for source in plan.research_sources if source.source_type != "asset"}) < 2:
            issues.append(QAIssue(
                slide=1, code="SINGLE_SOURCE", severity="warning",
                message="현재 덱은 단일 자료에 기반합니다. 중요한 의사결정용이면 추가 출처를 권장합니다.",
            ))
    if plan.language == "ko":
        forbidden = ("why it matters", "source", "orbital edition", "premium edition", "system / 01")
        for index, slide in enumerate(plan.slides, 1):
            visible = " ".join([slide.title, slide.subtitle, *(f"{x.heading} {x.body}" for x in slide.items)]).lower()
            if any(token in visible for token in forbidden):
                issues.append(QAIssue(
                    slide=index, code="UNLOCALIZED_LABEL",
                    message="한국어 덱에 영문 고정 UI 문구가 남아 있습니다.",
                ))
    return issues
