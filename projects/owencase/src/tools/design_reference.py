"""PowerPoint-native design recipes inspired by public React Bits references.

This module deliberately does not install, execute, scrape, or translate React
components. It gives the presentation agent a small, stable design contract
that can be implemented with the existing PowerPoint COM tools.
"""

from __future__ import annotations

import json
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


StyleName = Literal[
    "list",
    "agency_dark",
    "aurora_gradient",
    "magic_bento",
    "spotlight_cards",
    "glass_surface",
    "editorial_motion",
    "pixel_tech",
]
SlidePurpose = Literal[
    "cover", "section", "content", "metrics", "comparison", "closing"
]
MotionLevel = Literal["none", "subtle", "expressive"]


class GetDesignReferenceInput(BaseModel):
    """Select a React Bits reference and translate it into a PPT design recipe."""

    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    style: StyleName = Field(
        default="list",
        description=(
            "Reference style. Use 'list' to return the available PowerPoint-native "
            "reinterpretations and their public React Bits documentation URLs."
        ),
    )
    slide_purpose: SlidePurpose = Field(
        default="content",
        description="Slide role used to select an appropriate fixed 16:9 layout.",
    )
    motion: MotionLevel = Field(
        default="subtle",
        description=(
            "PowerPoint-native motion level. This never attempts to reproduce "
            "React, CSS, canvas, WebGL, hover, cursor, or scroll behavior."
        ),
    )


_STYLES: dict[str, dict] = {
    "agency_dark": {
        "name": "Agency Dark",
        "reference_url": "https://pro.reactbits.dev/docs/templates/agency-site",
        "reference_focus": "bold agency hierarchy, dark canvas, restrained luminous accents",
        "palette": {
            "background": "#09090B",
            "surface": "#18181B",
            "text": "#FAFAFA",
            "muted": "#A1A1AA",
            "accent": "#A78BFA",
            "accent_secondary": "#22D3EE",
        },
        "typography": {
            "latin_heading": "Aptos Display",
            "latin_body": "Aptos",
            "east_asian": "Malgun Gothic",
            "heading_weight": "bold",
            "case_rule": "sentence case; reserve all-caps for short labels",
        },
        "native_elements": [
            "oversized left-aligned headline",
            "one thin accent rule or compact status pill",
            "asymmetric image or card cluster on the right",
            "large unused dark area to preserve hierarchy",
        ],
        "avoid": ["shader imitation", "decorative cursor", "more than two accent colors"],
    },
    "aurora_gradient": {
        "name": "Aurora Gradient",
        "reference_url": "https://reactbits.dev/backgrounds/aurora",
        "reference_focus": "soft luminous color fields on a deep background",
        "palette": {
            "background": "#070A12",
            "surface": "#111827",
            "text": "#F8FAFC",
            "muted": "#CBD5E1",
            "accent": "#34D399",
            "accent_secondary": "#38BDF8",
        },
        "typography": {
            "latin_heading": "Aptos Display",
            "latin_body": "Aptos",
            "east_asian": "Malgun Gothic",
            "heading_weight": "semibold",
            "case_rule": "short calm headlines; avoid condensed all-caps",
        },
        "native_elements": [
            "two or three oversized translucent gradient shapes",
            "soft glow with generous overlap outside the slide bounds",
            "solid high-contrast text placed away from the brightest region",
            "one crisp foreground card to anchor the composition",
        ],
        "avoid": ["frame-by-frame aurora imitation", "rainbow palette", "glow behind body text"],
    },
    "magic_bento": {
        "name": "Magic Bento",
        "reference_url": "https://reactbits.dev/components/magic-bento",
        "reference_focus": "modular bento hierarchy with one emphasized card",
        "palette": {
            "background": "#0B0B10",
            "surface": "#17171F",
            "text": "#F4F4F5",
            "muted": "#A1A1AA",
            "accent": "#8B5CF6",
            "accent_secondary": "#F59E0B",
        },
        "typography": {
            "latin_heading": "Aptos Display",
            "latin_body": "Aptos",
            "east_asian": "Malgun Gothic",
            "heading_weight": "bold",
            "case_rule": "compact labels and short card titles",
        },
        "native_elements": [
            "two-column or three-column modular card grid",
            "one card spanning two grid tracks",
            "consistent 12-18 pt corner radius and internal padding",
            "small accent icon or metric per card instead of ornamental effects",
        ],
        "avoid": ["hover glow", "random card sizes", "more than six cards on one slide"],
    },
    "spotlight_cards": {
        "name": "Spotlight Cards",
        "reference_url": "https://reactbits.dev/components/spotlight-card",
        "reference_focus": "dark cards with a single high-contrast focal card",
        "palette": {
            "background": "#0A0A0F",
            "surface": "#16161D",
            "text": "#FAFAFA",
            "muted": "#9CA3AF",
            "accent": "#22D3EE",
            "accent_secondary": "#818CF8",
        },
        "typography": {
            "latin_heading": "Aptos Display",
            "latin_body": "Aptos",
            "east_asian": "Malgun Gothic",
            "heading_weight": "semibold",
            "case_rule": "short declarative headings",
        },
        "native_elements": [
            "three aligned cards with identical geometry",
            "accent border and low-radius glow on only the focal card",
            "muted borders on supporting cards",
            "large numeric or icon anchor above concise copy",
        ],
        "avoid": ["mouse-following spotlight", "glow on every card", "long paragraphs"],
    },
    "glass_surface": {
        "name": "Glass Surface",
        "reference_url": "https://reactbits.dev/components/glass-surface",
        "reference_focus": "translucent panels, fine borders, and layered depth",
        "palette": {
            "background": "#111827",
            "surface": "#FFFFFF",
            "text": "#F8FAFC",
            "muted": "#CBD5E1",
            "accent": "#60A5FA",
            "accent_secondary": "#C084FC",
        },
        "typography": {
            "latin_heading": "Aptos Display",
            "latin_body": "Aptos",
            "east_asian": "Malgun Gothic",
            "heading_weight": "semibold",
            "case_rule": "clean sentence case",
        },
        "native_elements": [
            "rounded panels using 12-22 percent white transparency",
            "0.75-1 pt translucent white border",
            "soft shadow below panels rather than blur inside them",
            "one saturated accent shape behind the glass layer",
        ],
        "avoid": ["fake refraction", "low-contrast text", "stacking more than two glass layers"],
    },
    "editorial_motion": {
        "name": "Editorial Motion",
        "reference_url": "https://reactbits.dev/animations/fade-content",
        "reference_focus": "editorial spacing with restrained reveal choreography",
        "palette": {
            "background": "#F5F2EA",
            "surface": "#FFFFFF",
            "text": "#18181B",
            "muted": "#71717A",
            "accent": "#EA580C",
            "accent_secondary": "#0F766E",
        },
        "typography": {
            "latin_heading": "Georgia",
            "latin_body": "Aptos",
            "east_asian": "Malgun Gothic",
            "heading_weight": "regular",
            "case_rule": "sentence case with deliberate line breaks",
        },
        "native_elements": [
            "strong baseline grid and wide page margins",
            "one serif display statement paired with neutral sans-serif body",
            "thin rules and small folio labels",
            "one image with a strict rectangular crop",
        ],
        "avoid": ["multiple display fonts", "heavy card chrome", "animation on every line"],
    },
    "pixel_tech": {
        "name": "Pixel Tech",
        "reference_url": "https://reactbits.dev/backgrounds/faulty-terminal",
        "reference_focus": "terminal-inspired rhythm without animated glitch simulation",
        "palette": {
            "background": "#050806",
            "surface": "#0D1510",
            "text": "#E7FBEA",
            "muted": "#7FA889",
            "accent": "#4ADE80",
            "accent_secondary": "#FACC15",
        },
        "typography": {
            "latin_heading": "Consolas",
            "latin_body": "Aptos",
            "east_asian": "Malgun Gothic",
            "heading_weight": "bold",
            "case_rule": "short labels may use all-caps; body remains sentence case",
        },
        "native_elements": [
            "fine grid or scan-line motif at very low contrast",
            "monospace labels paired with a readable sans-serif body",
            "square or lightly rounded modules",
            "single green accent plus one warning color",
        ],
        "avoid": ["animated glitch", "dense fake code", "green text for long paragraphs"],
    },
}


_PURPOSE_LAYOUTS: dict[str, dict] = {
    "cover": {
        "safe_margin_pt": 54,
        "structure": "headline block 52-58% wide; supporting visual 30-36% wide",
        "title_size_pt": "38-48",
        "body_size_pt": "18-22",
        "content_limit": "one headline, one supporting sentence, one label",
    },
    "section": {
        "safe_margin_pt": 58,
        "structure": "oversized section number or label plus one short statement",
        "title_size_pt": "34-44",
        "body_size_pt": "18-20",
        "content_limit": "no more than two text groups",
    },
    "content": {
        "safe_margin_pt": 44,
        "structure": "top title band plus a 42/58 or 50/50 content split",
        "title_size_pt": "30-38",
        "body_size_pt": "17-20",
        "content_limit": "one message and at most four supporting items",
    },
    "metrics": {
        "safe_margin_pt": 42,
        "structure": "one headline followed by three or four aligned metric modules",
        "title_size_pt": "28-36",
        "body_size_pt": "16-19",
        "content_limit": "one metric, one label, and one qualifier per module",
    },
    "comparison": {
        "safe_margin_pt": 42,
        "structure": "two equal columns with a centered divider or shared baseline",
        "title_size_pt": "28-36",
        "body_size_pt": "16-19",
        "content_limit": "three comparable points per side",
    },
    "closing": {
        "safe_margin_pt": 60,
        "structure": "centered or left-weighted closing statement with one action",
        "title_size_pt": "36-46",
        "body_size_pt": "18-22",
        "content_limit": "one closing statement and one contact/action line",
    },
}


_MOTION_PLANS: dict[str, dict] = {
    "none": {
        "principle": "Use no object animation; preserve hierarchy entirely through layout.",
        "effects": [],
        "transition": {"effect": "none"},
    },
    "subtle": {
        "principle": "Use a single reveal family and keep total build time below 1.2 seconds.",
        "effects": [
            {"target": "title", "effect": "fade", "trigger": "on_click", "duration": 0.35},
            {"target": "supporting content", "effect": "fade", "trigger": "after_previous", "duration": 0.25, "delay": 0.08},
            {"target": "accent rule or focal card", "effect": "wipe", "trigger": "with_previous", "duration": 0.3},
        ],
        "transition": {"effect": "fade", "duration": 0.3},
    },
    "expressive": {
        "principle": "Animate the hierarchy, not decoration; use at most three effect families.",
        "effects": [
            {"target": "title", "effect": "float", "trigger": "on_click", "duration": 0.45},
            {"target": "primary visual", "effect": "zoom", "trigger": "with_previous", "duration": 0.4},
            {"target": "cards or supporting items", "effect": "fade", "trigger": "after_previous", "duration": 0.25, "stagger_seconds": 0.08},
        ],
        "transition": {"effect": "fade", "duration": 0.35},
    },
}


def _style_catalog() -> list[dict]:
    return [
        {
            "style": key,
            "name": value["name"],
            "reference_url": value["reference_url"],
            "reference_focus": value["reference_focus"],
        }
        for key, value in _STYLES.items()
    ]


def _build_recipe(style: str, slide_purpose: str, motion: str) -> dict:
    selected = _STYLES[style]
    return {
        "success": True,
        "mode": "powerpoint_native_redesign",
        "reference_only": True,
        "source": "React Bits public design reference",
        "style": style,
        "slide_purpose": slide_purpose,
        "reference": {
            "name": selected["name"],
            "url": selected["reference_url"],
            "focus": selected["reference_focus"],
            "connection": "public documentation link only; no source code or runtime is fetched",
        },
        "design_contract": {
            "canvas": {"ratio": "16:9", "width_pt": 960, "height_pt": 540},
            "palette": selected["palette"],
            "typography": selected["typography"],
            "layout": _PURPOSE_LAYOUTS[slide_purpose],
            "native_elements": selected["native_elements"],
            "motion": _MOTION_PLANS[motion],
            "avoid": selected["avoid"],
        },
        "build_order": [
            "Set the fixed 16:9 slide background and theme colors.",
            "Build the layout grid with native PowerPoint shapes only.",
            "Apply typography and verify actual line breaks in PowerPoint.",
            "Add only the selected native animation plan.",
            "Review the editable slide structure and refine the visual hierarchy.",
        ],
        "recommended_tools": [
            "ppt_set_slide_background",
            "ppt_add_shape",
            "ppt_add_textbox",
            "ppt_set_fill",
            "ppt_set_line",
            "ppt_set_shadow",
            "ppt_set_glow",
            "ppt_format_text",
            "ppt_align_shapes",
            "ppt_distribute_shapes",
            "ppt_add_animation",
            "ppt_set_slide_transition",
        ],
        "guardrails": [
            "Do not import or execute React Bits code inside PowerPoint.",
            "Do not claim pixel-identical reproduction of CSS, canvas, WebGL, hover, cursor, or scroll effects.",
            "Treat the reference as visual direction, then make a coherent PowerPoint-native composition.",
            "Keep text, charts, and core content editable unless the user explicitly requests rasterization.",
        ],
    }


def get_design_reference(params: GetDesignReferenceInput) -> str:
    """Return curated React Bits references or one PowerPoint-native recipe."""

    if params.style == "list":
        result = {
            "success": True,
            "mode": "powerpoint_native_redesign",
            "reference_only": True,
            "styles": _style_catalog(),
            "usage": "Call again with one style, slide_purpose, and motion level.",
        }
    else:
        result = _build_recipe(params.style, params.slide_purpose, params.motion)

    return json.dumps(result, ensure_ascii=False)
