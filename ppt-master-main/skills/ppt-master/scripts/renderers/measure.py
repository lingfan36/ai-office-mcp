"""
Pre-layout text dimension estimation.
Uses character-width heuristics — no external font files or PIL required.
96 DPI assumed (matches PPTX canvas: 914400 EMU/inch ÷ 9525 EMU/px).
"""

_PT_TO_PX = 96.0 / 72.0  # 1pt = 1.333... px


def _is_wide(ch: str) -> bool:
    """True for CJK / full-width characters (square cell)."""
    cp = ord(ch)
    return (
        0x1100 <= cp <= 0x11FF or   # Hangul Jamo
        0x2E80 <= cp <= 0x2EFF or   # CJK Radicals Supplement
        0x2F00 <= cp <= 0x2FDF or   # Kangxi Radicals
        0x3000 <= cp <= 0x9FFF or   # CJK misc + Unified Ideographs
        0xAC00 <= cp <= 0xD7AF or   # Hangul Syllables
        0xF900 <= cp <= 0xFAFF or   # CJK Compatibility Ideographs
        0xFF00 <= cp <= 0xFFEF       # Fullwidth / Halfwidth Forms
    )


def _char_w(ch: str, px_size: float) -> float:
    """Approximate glyph advance width in pixels."""
    if _is_wide(ch):
        return px_size              # CJK / full-width: square cell
    if ch in (' ', '　', '\t'):
        return px_size * 0.30
    if ch in ('.', ',', ':', ';', '!', '?', '|', "'", '"'):
        return px_size * 0.35
    if ch in ('W', 'M', 'm', 'w'):
        return px_size * 0.72
    if ch.isupper():
        return px_size * 0.65
    if ch.isdigit():
        return px_size * 0.55
    return px_size * 0.52           # lower-case / misc ASCII


def _para_lines(para: str, px_size: float, max_w: int) -> int:
    """How many display lines a single paragraph needs after word-wrap."""
    if not para:
        return 1                    # blank line still occupies a row
    # Greedy character-level wrap (conservative — treats every char as atomic)
    cur_w = 0.0
    lines = 1
    for ch in para:
        cw = _char_w(ch, px_size)
        if cur_w + cw > max_w and cur_w > 0:
            lines += 1
            cur_w = cw
        else:
            cur_w += cw
    return lines


def text_height(
    text: str,
    font_size_pt: float,
    max_width_px: int,
    line_spacing: float = 1.15,
) -> int:
    """
    Estimate rendered text height in pixels (no box padding).

    Parameters
    ----------
    text          : the string to measure (may contain \\n line breaks)
    font_size_pt  : font size in points
    max_width_px  : available width in pixels (for wrap calculation)
    line_spacing  : multiplier per line (default 1.15 ≈ PPTX auto)

    Returns
    -------
    Integer pixel height with a 20 % safety margin for font-metric variance.
    """
    if not text:
        return 0
    px_size = font_size_pt * _PT_TO_PX
    n_lines = sum(_para_lines(p, px_size, max_width_px) for p in text.split('\n'))
    return int(n_lines * px_size * line_spacing * 1.20)


def text_height_box(
    text: str,
    font_size_pt: float,
    max_width_px: int,
    line_spacing: float = 1.15,
    v_pad: int = 8,
) -> int:
    """text_height + vertical padding (top + bottom combined)."""
    return text_height(text, font_size_pt, max_width_px, line_spacing) + v_pad * 2
