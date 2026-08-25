"""PowerPoint COM constants used by the MCP server."""

# MsoTriState
msoTrue = -1
msoFalse = 0

# PpSlideLayout
ppLayoutBlank = 12

# PpSaveAsFileType
ppSaveAsDefault = 11
ppSaveAsJPG = 17
ppSaveAsPNG = 18
ppSaveAsOpenXMLPresentation = 24
ppSaveAsPDF = 32

# MsoShapeType
msoChart = 3
msoFreeform = 5
msoGroup = 6
msoLinkedPicture = 11
msoPicture = 13
msoPlaceholder = 14
msoTextBox = 17
msoTable = 19
msoSmartArt = 24

# MsoAutoShapeType (commonly used)
msoShapeRectangle = 1

# PpPlaceholderType
ppPlaceholderTitle = 1
ppPlaceholderBody = 2
ppPlaceholderCenterTitle = 3
ppPlaceholderSubtitle = 4

# PpParagraphAlignment
ppAlignLeft = 1
ppAlignCenter = 2
ppAlignRight = 3
ppAlignJustify = 4
ppAlignDistribute = 5

# MsoTextOrientation
msoTextOrientationHorizontal = 1
msoTextOrientationUpward = 2
msoTextOrientationDownward = 3
msoTextOrientationVertical = 5

# PpAutoSize
ppAutoSizeNone = 0
ppAutoSizeShapeToFitText = 1
ppAutoSizeTextToFitShape = 2  # MsoAutoSize (TextFrame2 only) — shrink text on overflow

# MsoZOrderCmd
msoBringToFront = 0
msoSendToBack = 1
msoBringForward = 2
msoSendBackward = 3

# MsoGradientStyle
msoGradientHorizontal = 1
msoGradientVertical = 2
msoGradientDiagonalUp = 3
msoGradientDiagonalDown = 4
msoGradientFromCorner = 5
msoGradientFromCenter = 7

# MsoLineDashStyle
msoLineSolid = 1
msoLineRoundDot = 2
msoLineDot = 3
msoLineDash = 4
msoLineDashDot = 5
msoLineDashDotDot = 6
msoLineLongDash = 7
msoLineLongDashDot = 8

# MsoArrowheadStyle
msoArrowheadNone = 1
msoArrowheadTriangle = 2
msoArrowheadOpen = 3
msoArrowheadStealth = 4
msoArrowheadDiamond = 5
msoArrowheadOval = 6

# PpBulletType
ppBulletNone = 0
ppBulletUnnumbered = 1
ppBulletNumbered = 2

# PpNumberedBulletStyle (subset commonly used outside CJK/Hindi/Hebrew/Thai)
ppBulletAlphaLCPeriod = 0
ppBulletAlphaUCPeriod = 1
ppBulletArabicParenRight = 2
ppBulletArabicPeriod = 3
ppBulletRomanLCParenBoth = 4
ppBulletRomanLCParenRight = 5
ppBulletRomanLCPeriod = 6
ppBulletRomanUCPeriod = 7
ppBulletAlphaLCParenBoth = 8
ppBulletAlphaLCParenRight = 9
ppBulletAlphaUCParenBoth = 10
ppBulletAlphaUCParenRight = 11
ppBulletArabicParenBoth = 12
ppBulletArabicPlain = 13
ppBulletRomanUCParenBoth = 14
ppBulletRomanUCParenRight = 15
ppBulletKanjiKoreanPlain = 26
ppBulletKanjiKoreanPeriod = 27
ppBulletArabicDBPlain = 28
ppBulletArabicDBPeriod = 29

# PpSelectionType
ppSelectionNone = 0
ppSelectionSlides = 1
ppSelectionShapes = 2
ppSelectionText = 3

# PpSlideShowType
ppShowTypeSpeaker = 1
ppShowTypeWindow = 2
ppShowTypeKiosk = 3

# PpSlideShowRangeType
ppShowAll = 1
ppShowSlideRange = 2

# PpFixedFormatType
ppFixedFormatTypePDF = 2

# Shape type name mapping for human-readable output
SHAPE_TYPE_NAMES = {
    1: "AutoShape",
    2: "Callout",
    3: "Chart",
    4: "Comment",
    5: "Freeform",
    6: "Group",
    7: "EmbeddedOLEObject",
    8: "FormControl",
    9: "Line",
    10: "LinkedOLEObject",
    11: "LinkedPicture",
    12: "OLEControlObject",
    13: "Picture",
    14: "Placeholder",
    15: "TextEffect",
    16: "Media",
    17: "TextBox",
    19: "Table",
    24: "SmartArt",
}
PLACEHOLDER_TYPE_NAMES = {
    1: "Title",
    2: "Body",
    3: "CenterTitle",
    4: "Subtitle",
    5: "VerticalTitle",
    6: "VerticalBody",
    7: "Object",
    8: "Chart",
    9: "Bitmap",
    10: "MediaClip",
    11: "OrgChart",
    12: "Table",
    13: "SlideNumber",
    14: "Header",
    15: "Footer",
    16: "Date",
    17: "VerticalObject",
    18: "Picture",
}
WINDOW_STATE_NAMES = {
    1: "normal",
    2: "minimized",
    3: "maximized",
}
SLIDESHOW_STATE_NAMES = {
    1: "running",
    2: "paused",
    3: "black_screen",
    4: "white_screen",
    5: "done",
}
SHOW_TYPE_NAMES = {
    1: "speaker",
    2: "window",
    3: "kiosk",
}

# PpActionType
ppActionNone = 0
ppActionHyperlink = 7

# PpMouseActivation
ppMouseClick = 1
ppMouseOver = 2

# XlAxisType
xlCategory = 1
xlValue = 2

# Name lookup maps for Phase 3
ANIMATION_EFFECT_NAMES = {
    # Entrance effects (1-53)
    1: "appear", 2: "fly", 3: "blinds", 4: "box",
    5: "checkerboard", 6: "circle", 8: "diamond",
    9: "dissolve", 10: "fade", 16: "split", 22: "wipe",
    23: "zoom", 26: "bounce", 30: "float", 31: "grow_and_turn",
    # Emphasis effects (54-82)
    54: "change_fill_color", 55: "change_font", 56: "change_font_color",
    57: "change_font_size", 59: "grow_shrink", 61: "spin", 62: "transparency",
    63: "bold_flash", 69: "color_wave", 73: "darken", 74: "desaturate",
    75: "flash_bulb", 78: "lighten", 80: "teeter", 82: "wave",
    # Motion path effects (86-149)
    86: "path_circle", 88: "path_diamond", 90: "path_star",
    92: "path_square", 94: "path_heart", 109: "path_loop",
    120: "path_left", 122: "path_arc_down", 123: "path_zigzag",
    125: "path_sine_wave", 126: "path_bounce_left", 127: "path_down",
    129: "path_arc_up", 131: "path_spiral_right", 132: "path_wave",
    134: "path_diagonal_down_right", 136: "path_arc_left",
    137: "path_funnel", 138: "path_spring", 139: "path_bounce_right",
    141: "path_diagonal_up_right", 143: "path_arc_right",
    148: "path_up", 149: "path_right",
}
ANIMATION_TRIGGER_NAMES = {
    0: "none", 1: "on_click", 2: "with_previous",
    3: "after_previous", 4: "on_shape_click",
}

# MsoAnimDirection
ANIM_DIRECTION_MAP = {
    "none": 0, "up": 1, "right": 2, "down": 3, "left": 4,
    "up_left": 6, "up_right": 7, "down_right": 8, "down_left": 9,
    "top": 10, "bottom": 11, "top_left": 12, "top_right": 13,
    "bottom_right": 14, "bottom_left": 15,
    "horizontal": 16, "vertical": 17, "across": 18,
    "in": 19, "out": 20,
    "clockwise": 21, "counterclockwise": 22,
    "horizontal_in": 23, "horizontal_out": 24,
    "vertical_in": 25, "vertical_out": 26,
}
ANIM_DIRECTION_NAMES = {v: k for k, v in ANIM_DIRECTION_MAP.items()}

# PpSelectionType
ppSelectionNone = 0
ppSelectionSlides = 1
ppSelectionShapes = 2
ppSelectionText = 3

# MsoGradientStyle (for backgrounds)
msoGradientHorizontal = 1
msoGradientVertical = 2
msoGradientDiagonalUp = 3
msoGradientDiagonalDown = 4
msoGradientFromCorner = 5
msoGradientFromCenter = 7

# Friendly name maps for Phase 4
ALIGN_CMD_MAP = {
    "left": 0, "center": 1, "right": 2,
    "top": 3, "middle": 4, "bottom": 5,
}
DISTRIBUTE_CMD_MAP = {
    "horizontal": 0, "vertical": 1,
}
FLIP_CMD_MAP = {
    "horizontal": 0, "vertical": 1,
}
MERGE_CMD_MAP = {
    "union": 1, "combine": 2, "intersect": 3,
    "subtract": 4, "fragment": 5,
}
SLIDE_SIZE_MAP = {
    "4:3": 1, "letter": 2, "a4": 3, "35mm": 4,
    "overhead": 5, "banner": 6, "custom": 7,
    "a3": 8, "16:9": 9, "16:10": 10, "widescreen": 9,
}
GRADIENT_STYLE_MAP = {
    "horizontal": 1, "vertical": 2,
    "diagonal_up": 3, "diagonal_down": 4,
    "from_corner": 5, "from_title": 6, "from_center": 7,
}
SHAPE_FORMAT_MAP = {
    "gif": 0, "jpg": 1, "png": 2, "bmp": 3,
    "wmf": 4, "emf": 5,
}
VIEW_TYPE_MAP = {
    "normal": 1, "slide_master": 2, "notes_page": 3,
    "handout_master": 4, "notes_master": 5,
    "outline": 6, "slide_sorter": 7,
    "title_master": 8, "reading": 10,
}
VIEW_TYPE_NAMES = {v: k for k, v in VIEW_TYPE_MAP.items()}

# PpBorderType
ppBorderTop          = 1
ppBorderLeft         = 2
ppBorderBottom       = 3
ppBorderRight        = 4
ppBorderDiagonalDown = 5
ppBorderDiagonalUp   = 6

# MsoVerticalAnchor
msoAnchorTop            = 1
msoAnchorMiddle         = 3
msoAnchorBottom         = 4

# MsoEditingType (for FreeformBuilder / ShapeNodes)
msoEditingAuto      = 0  # Auto-select type based on connected segments

# MsoSegmentType (for FreeformBuilder / ShapeNodes)
msoSegmentLine  = 0  # Straight line segment

# Lookup maps for freeform tools
EDITING_TYPE_MAP = {
    "auto": 0, "corner": 1, "smooth": 2, "symmetric": 3,
}
EDITING_TYPE_NAMES = {
    0: "auto", 1: "corner", 2: "smooth", 3: "symmetric",
}
SEGMENT_TYPE_MAP = {
    "line": 0, "curve": 1,
}
SEGMENT_TYPE_NAMES = {
    0: "line", 1: "curve",
}

# MsoPictureColorType
PICTURE_COLOR_TYPE_MAP = {
    "automatic": 1,
    "grayscale": 2,
    "black_and_white": 3,
    "watermark": 4,
}
PICTURE_COLOR_TYPE_NAMES = {v: k for k, v in PICTURE_COLOR_TYPE_MAP.items()}

# MsoAnimAfterEffect
AFTER_EFFECT_MAP = {
    "none": 0, "dim": 1, "hide": 2, "hide_on_next_click": 3,
}
AFTER_EFFECT_NAMES = {v: k for k, v in AFTER_EFFECT_MAP.items()}

# MsoAnimateByLevel (text build level)
BUILD_LEVEL_MAP = {
    "none": 0, "all_levels": 1, "first_level": 2, "second_level": 3,
    "third_level": 4, "fourth_level": 5, "fifth_level": 6,
}
BUILD_LEVEL_NAMES = {v: k for k, v in BUILD_LEVEL_MAP.items()}

# MsoAnimTextUnitEffect
TEXT_UNIT_EFFECT_MAP = {
    "by_paragraph": 0, "by_character": 1, "by_word": 2,
}
TEXT_UNIT_EFFECT_NAMES = {v: k for k, v in TEXT_UNIT_EFFECT_MAP.items()}
