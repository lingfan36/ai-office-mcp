from pptx.dml.color import RGBColor

# ── Palette ───────────────────────────────────────────────────────────────────
DARK_BG    = "1E293B"
WHITE_BG   = "FFFFFF"
CARD_BG    = "F1F5F9"
ACCENT     = "0EA5E9"
ACCENT2    = "38BDF8"
GREEN      = "22C55E"
INDIGO     = "6366F1"
AMBER      = "F59E0B"
RED        = "EF4444"
T_PRIMARY  = "1E293B"
T_SECOND   = "64748B"
T_MUTED    = "94A3B8"
BORDER     = "CBD5E1"
BORDER_DK  = "334155"

# ── Fonts ─────────────────────────────────────────────────────────────────────
FONT_CJK   = "Microsoft YaHei UI"   # UI variant: better hinting at small sizes
FONT_NUM   = "Segoe UI"             # clean modern numerals
FONT_SANS  = "Segoe UI"

# ── Gradient pairs (c1=start, c2=end) ────────────────────────────────────────
DARKEST     = "080F1A"
GRAD_DARK   = (DARK_BG, "080F1A")     # cover / fundraising background
GRAD_ACCENT = (ACCENT,  INDIGO)       # accent bars, hub, buttons
GRAD_CARD   = ("FFFFFF", "EFF6FF")    # light card background

# ── Canvas: 1280 × 720 px (96 DPI → 13.33" × 7.5") ──────────────────────────
SLIDE_W_PX = 1280
SLIDE_H_PX = 720
PX_TO_EMU  = 9525          # 914400 EMU/inch ÷ 96 DPI


def rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# Light tint map for icon background circles
TINT = {
    ACCENT: "E0F2FE",
    GREEN:  "DCFCE7",
    INDIGO: "EEF2FF",
    AMBER:  "FEF3C7",
    RED:    "FEE2E2",
}


# ── Decoration flags (can be overridden by theme presets) ────────────────────
DECO_GLOW  = True   # ambient radial glow orbs
DECO_NOISE = True   # grain / noise overlay texture
DECO_RINGS = True   # decorative concentric ring overlays


def is_dark_bg() -> bool:
    """True when WHITE_BG has been overridden to a dark colour."""
    h = WHITE_BG.lstrip("#")
    lum = int(h[:2], 16) * 0.299 + int(h[2:4], 16) * 0.587 + int(h[4:6], 16) * 0.114
    return lum < 128


def tint(color: str) -> str:
    """Return a tinted version of color: dark variant on dark bg, light on light bg."""
    h = color.lstrip("#")
    r, g, b = int(h[:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    if is_dark_bg():
        dr = min(255, int(r * 0.28))
        dg = min(255, int(g * 0.28))
        db = min(255, int(b * 0.28))
        return f"{dr:02X}{dg:02X}{db:02X}"
    if color in TINT:
        return TINT[color]
    lr = min(255, int(r + (255 - r) * 0.85))
    lg = min(255, int(g + (255 - g) * 0.85))
    lb = min(255, int(b + (255 - b) * 0.85))
    return f"{lr:02X}{lg:02X}{lb:02X}"
