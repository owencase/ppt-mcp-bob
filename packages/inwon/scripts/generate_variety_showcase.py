from __future__ import annotations

import os
from pathlib import Path

from canva_ppt_mcp.models import ContentItem, DeckPlan, DesignSystem, Palette, SlideSpec, Typography
from canva_ppt_mcp.pipeline import create_presentation


STYLES = [
    {
        "slug": "editorial", "name": "Editorial Red", "preset": "editorial",
        "palette": ("#191919", ["#E7EBF0", "#A7B0BD"], "#E23D28", "#FFFFFF", "#191919"),
        "font": "Bookman Old Style", "layout": "big_stat",
        "title": "Ideas Need a Clear Point of View",
        "subtitle": "Editorial hierarchy, bold contrast, and disciplined whitespace",
        "items": [ContentItem(heading="Signal over noise", body="One clear claim should anchor every slide.", value="01")],
    },
    {
        "slug": "neon", "name": "Neon Future", "preset": "neon",
        "palette": ("#0B1026", ["#6C4DFF", "#14213D"], "#00F5D4", "#F5F7FF", "#070A18"),
        "font": "Cambria", "layout": "comparison",
        "title": "Build for the Next Interface",
        "subtitle": "Electric accents and dark depth for emerging technology stories",
        "items": [ContentItem(heading="Current", body="Disconnected tools slow the experience."), ContentItem(heading="Next", body="A unified layer makes capability feel immediate.")],
    },
    {
        "slug": "organic", "name": "Organic Growth", "preset": "organic",
        "palette": ("#164E3A", ["#B8D8BA", "#DDEEDB"], "#FF7A59", "#F6FBF7", "#0B2F23"),
        "font": "Century Schoolbook", "layout": "timeline",
        "title": "Growth Works in Living Systems",
        "subtitle": "Soft forms and grounded color for sustainability and people topics",
        "items": [ContentItem(heading="Seed", body="Start with one useful behavior."), ContentItem(heading="Root", body="Build routines that make it repeatable."), ContentItem(heading="Scale", body="Expand only after the system is healthy.")],
    },
    {
        "slug": "luxury", "name": "Dark Luxury", "preset": "luxury",
        "palette": ("#151515", ["#2C2C2C", "#6C5A3D"], "#D6AE61", "#FBFBFA", "#0D0D0D"),
        "font": "Bookman Old Style", "layout": "two_column",
        "title": "Precision Creates Premium Value",
        "subtitle": "Restrained gold, deep black, and a slower visual rhythm",
        "items": [ContentItem(heading="Craft", body="Details communicate intent before words do."), ContentItem(heading="Restraint", body="Fewer elements make each decision feel deliberate.")],
    },
    {
        "slug": "geometric", "name": "Bold Geometric", "preset": "geometric",
        "palette": ("#243B53", ["#D9E2EC", "#486581"], "#FF6B35", "#F7FAFC", "#102A43"),
        "font": "Cambria", "layout": "grid_2x2",
        "title": "Structure Makes Complexity Usable",
        "subtitle": "Strong blocks and modular rhythm for product and strategy decks",
        "items": [ContentItem(heading="Frame", body="Define the decision."), ContentItem(heading="Focus", body="Prioritize the signal."), ContentItem(heading="Build", body="Connect the capabilities."), ContentItem(heading="Measure", body="Track the outcome.")],
    },
    {
        "slug": "swiss", "name": "Swiss Minimal", "preset": "swiss",
        "palette": ("#111111", ["#E8E8E8", "#B7B7B7"], "#E31B23", "#FFFFFF", "#111111"),
        "font": "Century Schoolbook", "layout": "icon_rows",
        "title": "Clarity Is the Design System",
        "subtitle": "High contrast, strict alignment, and direct communication",
        "items": [ContentItem(heading="Align", body="Use one visible grid."), ContentItem(heading="Reduce", body="Remove decoration without meaning."), ContentItem(heading="Emphasize", body="Let scale create the hierarchy.")],
    },
]


def build(root: Path) -> None:
    os.environ.pop("OPENAI_API_KEY", None)
    root.mkdir(parents=True, exist_ok=True)
    for style in STYLES:
        primary, secondary, accent, light, dark = style["palette"]
        design = DesignSystem(
            palette=Palette(primary=primary, secondary=secondary, accent=accent,
                            background_light=light, background_dark=dark),
            typography=Typography(header_font=style["font"], body_font="Arial"),
            visual_motif=style["name"], style_preset=style["preset"],
            layout_rotation=["title", style["layout"], "comparison", "timeline"],
        )
        plan = DeckPlan(
            communication_job="Show a distinct offline presentation design direction.",
            design_system=design,
            slides=[
                SlideSpec(title=style["title"], subtitle=style["subtitle"], layout="title"),
                SlideSpec(title=f"{style['name']} in practice",
                          layout=style["layout"], items=style["items"]),
                SlideSpec(title="Choose the style that matches the message",
                          subtitle="The next version can apply this direction to a complete topic deck.",
                          layout="closing", items=[ContentItem(heading="NEXT", body=f"Use the {style['name']} direction.")]),
            ],
        )
        result = create_presentation(topic=style["name"], output_path=str(root / f"{style['slug']}.pptx"),
                                     slide_count=3, language="en", content_json=plan.model_dump())
        if not result["qa"]["passed"]:
            raise RuntimeError(f"QA failed for {style['name']}")
        print(style["slug"], "passed", result["qa"]["rounds"])


if __name__ == "__main__":
    build(Path("output/design-variety"))
