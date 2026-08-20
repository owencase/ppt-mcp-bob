from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from dataclasses import dataclass, field
from html import unescape
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


WIKIPEDIA_APIS = {
    "ko": "https://ko.wikipedia.org/w/api.php",
    "en": "https://en.wikipedia.org/w/api.php",
}


class ResearchUnavailableError(RuntimeError):
    pass


@dataclass
class ResearchSection:
    heading: str
    sentences: list[str] = field(default_factory=list)
    source_url: str = ""


@dataclass
class ResearchResult:
    query: str
    title: str
    url: str
    summary: str
    sections: list[ResearchSection]
    language: str
    image_url: str | None = None
    source_type: str = "web"

    @property
    def sentences(self) -> list[str]:
        values: list[str] = []
        for section in self.sections:
            values.extend(section.sentences)
        return values


def _request_json(url: str, params: dict[str, str | int], timeout: int = 15) -> dict:
    target = f"{url}?{urlencode(params)}"
    request = Request(target, headers={"User-Agent": "canva-ppt-mcp/2.2 (+local PowerPoint research)"})
    try:
        with urlopen(request, timeout=timeout) as response:
            return json.loads(response.read().decode("utf-8"))
    except (HTTPError, URLError, TimeoutError, json.JSONDecodeError) as exc:
        raise ResearchUnavailableError(f"공개 자료 조회 실패: {exc}") from exc


def _clean(text: str) -> str:
    text = unescape(text)
    text = re.sub(r"\[[^]]+\]", "", text)
    text = re.sub(r"\([^)]*IPA[^)]*\)", "", text, flags=re.I)
    # Long foreign-language aliases from encyclopedic leads are poor slide
    # copy. In offline Korean mode remove the alias rather than leak it into
    # audience-facing text without a translation engine.
    text = re.sub(r"\((?=[^)]*(?:[一-龥]|[A-Za-z]{3}))[^)]*\)", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" \n\t-–—")


def _sentences(text: str) -> list[str]:
    parts = re.split(r"(?<=[.!?。])\s+|\n+", text)
    result: list[str] = []
    for part in parts:
        value = _clean(part)
        lowered = value.lower()
        meta = ("위키미디어 공용", "문서를 참고", "공식 웹사이트", "category:", "분류:")
        if (28 <= len(value) <= 260 and not any(token in lowered for token in meta)
                and not (value.startswith("(") and value.endswith(")"))):
            result.append(value)
    return result


def _sections_from_extract(extract: str, default_heading: str, source_url: str = "") -> list[ResearchSection]:
    heading = default_heading
    collected: dict[str, list[str]] = {heading: []}
    for raw in extract.splitlines():
        line = raw.strip()
        match = re.match(r"^=+\s*(.*?)\s*=+$", line)
        if match:
            heading = _clean(match.group(1)) or default_heading
            collected.setdefault(heading, [])
            continue
        if line:
            collected.setdefault(heading, []).extend(_sentences(line))
    ignored = re.compile(r"^(각주|참고 문헌|참고문헌|외부 링크|같이 보기|관련 항목|주해|references?|external links?|see also|further reading)$", re.I)
    sections = [ResearchSection(heading=k, sentences=v[:8], source_url=source_url) for k, v in collected.items()
                if v and not ignored.match(k.strip())]
    return sections[:10]


def research_from_text(topic: str, text: str, source_url: str = "user-provided") -> ResearchResult:
    sentences = _sentences(text)
    if len(sentences) < 3:
        raise ResearchUnavailableError("제공된 조사 텍스트에서 슬라이드에 사용할 문장을 충분히 찾지 못했습니다.")
    return ResearchResult(
        query=topic,
        title=topic,
        url=source_url,
        summary=sentences[0],
        sections=[ResearchSection(heading="제공 자료", sentences=sentences, source_url=source_url)],
        language="ko" if any("가" <= c <= "힣" for c in text) else "en",
        source_type="user",
    )


def research_from_documents(topic: str, documents: list[dict[str, str]]) -> ResearchResult:
    sections: list[ResearchSection] = []
    all_sentences: list[str] = []
    for index, document in enumerate(documents, 1):
        text = document.get("text", "")
        url = document.get("url", f"user-provided-{index}")
        title = document.get("title", f"제공 자료 {index}")
        sentences = _sentences(text)
        if len(sentences) < 2:
            raise ResearchUnavailableError(f"'{title}'에서 사용할 문장을 충분히 찾지 못했습니다.")
        sections.append(ResearchSection(heading=title, sentences=sentences[:12], source_url=url))
        all_sentences.extend(sentences)
    if len(all_sentences) < 3:
        raise ResearchUnavailableError("제공된 조사 문서에서 슬라이드에 사용할 문장이 부족합니다.")
    return ResearchResult(
        query=topic, title=topic, url=sections[0].source_url,
        summary=all_sentences[0], sections=sections,
        language="ko" if any("가" <= c <= "힣" for c in " ".join(all_sentences)) else "en",
        source_type="user",
    )


def _search_title(topic: str, language: str) -> str:
    api = WIKIPEDIA_APIS[language]
    data = _request_json(api, {
        "action": "query", "list": "search", "srsearch": topic,
        "srlimit": 5, "utf8": 1, "format": "json",
    })
    results = data.get("query", {}).get("search", [])
    if not results:
        raise ResearchUnavailableError(f"'{topic}'과 관련된 공개 백과사전 문서를 찾지 못했습니다.")
    normalized = re.sub(r"\W+", "", topic).lower()
    for result in results:
        candidate = result.get("title", "")
        if normalized and normalized in re.sub(r"\W+", "", candidate).lower():
            return candidate
    return results[0].get("title", topic)


def research_wikipedia(topic: str, language: str = "ko") -> ResearchResult:
    requested = "en" if language.lower().startswith("en") else "ko"
    errors: list[str] = []
    for lang in (requested, "en") if requested != "en" else ("en",):
        try:
            title = _search_title(topic, lang)
            api = WIKIPEDIA_APIS[lang]
            data = _request_json(api, {
                "action": "query", "prop": "extracts|info|pageimages", "titles": title,
                "redirects": 1, "explaintext": 1, "inprop": "url",
                "pithumbsize": 1600, "format": "json", "utf8": 1,
            })
            pages = data.get("query", {}).get("pages", {})
            page = next((p for p in pages.values() if "missing" not in p), None)
            if not page:
                raise ResearchUnavailableError(f"'{title}' 문서 본문을 가져오지 못했습니다.")
            extract = page.get("extract", "")
            page_url = page.get("fullurl", f"https://{lang}.wikipedia.org/wiki/{title.replace(' ', '_')}")
            sections = _sections_from_extract(extract, "개요", page_url)
            sentences = [sentence for section in sections for sentence in section.sentences]
            if len(sentences) < 5:
                raise ResearchUnavailableError(f"'{title}' 문서의 유효한 본문이 부족합니다.")
            return ResearchResult(
                query=topic,
                title=page.get("title", title),
                url=page_url,
                summary=sentences[0],
                sections=sections,
                language=lang,
                image_url=(page.get("thumbnail") or {}).get("source"),
                source_type="wikipedia",
            )
        except ResearchUnavailableError as exc:
            errors.append(str(exc))
    raise ResearchUnavailableError(" / ".join(errors))


def research_topic(topic: str, language: str = "ko", research_text: str | None = None,
                   source_urls: list[str] | None = None,
                   research_documents: list[dict[str, str]] | None = None) -> ResearchResult:
    if research_documents:
        documents = list(research_documents)
        if research_text:
            documents.insert(0, {"title": "제공 자료", "text": research_text,
                                 "url": (source_urls or ["user-provided"])[0]})
        return research_from_documents(topic, documents)
    if research_text:
        return research_from_text(topic, research_text, (source_urls or ["user-provided"])[0])
    return research_wikipedia(topic, language)


def source_metadata(result: ResearchResult) -> dict[str, str | None]:
    return {
        "title": result.title,
        "url": result.url,
        "source_type": result.source_type,
        "retrieved_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
