"""
Built-in theme presets.
Each preset is a dict of overrides applied to theme.py module constants.
Keys must match attribute names in theme.py exactly.
"""

THEMES = {

    # ── 1. tech_dark (default) — dark slate + sky-blue ───────────────────────
    "tech_dark": {},   # no overrides; theme.py defaults apply

    # ── 2. corporate — navy + clean blue, no glow ────────────────────────────
    "corporate": {
        "WHITE_BG":    "FFFFFF",
        "CARD_BG":     "F1F5F9",
        "DARK_BG":     "1E3A5F",
        "DARKEST":     "0A1628",
        "ACCENT":      "1D4ED8",
        "ACCENT2":     "3B82F6",
        "GREEN":       "16A34A",
        "INDIGO":      "4338CA",
        "AMBER":       "D97706",
        "RED":         "DC2626",
        "T_PRIMARY":   "0F172A",
        "T_SECOND":    "334155",
        "T_MUTED":     "64748B",
        "BORDER":      "CBD5E1",
        "BORDER_DK":   "1E3A5F",
        "GRAD_DARK":   ("1E3A5F", "0A1628"),
        "GRAD_ACCENT": ("1D4ED8", "4338CA"),
        "GRAD_CARD":   ("FFFFFF", "EFF6FF"),
        "DECO_GLOW":   False,
        "DECO_RINGS":  False,
    },

    # ── 3. warm — cream + amber, energetic / startup ──────────────────────────
    "warm": {
        "WHITE_BG":    "FFFBF5",
        "CARD_BG":     "FEF3C7",
        "DARK_BG":     "1C1408",
        "DARKEST":     "0D0A04",
        "ACCENT":      "D97706",
        "ACCENT2":     "F59E0B",
        "GREEN":       "15803D",
        "INDIGO":      "7C3AED",
        "AMBER":       "B45309",
        "RED":         "DC2626",
        "T_PRIMARY":   "1C1408",
        "T_SECOND":    "78350F",
        "T_MUTED":     "92400E",
        "BORDER":      "FDE68A",
        "BORDER_DK":   "92400E",
        "GRAD_DARK":   ("1C1408", "0D0A04"),
        "GRAD_ACCENT": ("D97706", "F59E0B"),
        "GRAD_CARD":   ("FFFBF5", "FEF3C7"),
    },

    # ── 4. purple — deep indigo + violet, creative / AI ──────────────────────
    "purple": {
        "WHITE_BG":    "FAF5FF",
        "CARD_BG":     "EDE9FE",
        "DARK_BG":     "1E1B4B",
        "DARKEST":     "0F0A2E",
        "ACCENT":      "7C3AED",
        "ACCENT2":     "A78BFA",
        "GREEN":       "059669",
        "INDIGO":      "4F46E5",
        "AMBER":       "D97706",
        "RED":         "E11D48",
        "T_PRIMARY":   "1E1B4B",
        "T_SECOND":    "4C1D95",
        "T_MUTED":     "7C3AED",
        "BORDER":      "DDD6FE",
        "BORDER_DK":   "3730A3",
        "GRAD_DARK":   ("1E1B4B", "0F0A2E"),
        "GRAD_ACCENT": ("7C3AED", "A78BFA"),
        "GRAD_CARD":   ("FAF5FF", "EDE9FE"),
    },

    # ── 5. forest — deep green + emerald, sustainability / health ─────────────
    "forest": {
        "WHITE_BG":    "F0FDF4",
        "CARD_BG":     "DCFCE7",
        "DARK_BG":     "052E16",
        "DARKEST":     "022C22",
        "ACCENT":      "16A34A",
        "ACCENT2":     "22C55E",
        "GREEN":       "15803D",
        "INDIGO":      "0D9488",
        "AMBER":       "CA8A04",
        "RED":         "B91C1C",
        "T_PRIMARY":   "052E16",
        "T_SECOND":    "166534",
        "T_MUTED":     "4ADE80",
        "BORDER":      "BBF7D0",
        "BORDER_DK":   "14532D",
        "GRAD_DARK":   ("052E16", "022C22"),
        "GRAD_ACCENT": ("16A34A", "22C55E"),
        "GRAD_CARD":   ("F0FDF4", "DCFCE7"),
    },

    # ── 6. minimal — pure white + charcoal, no deco ───────────────────────────
    "minimal": {
        "WHITE_BG":    "FFFFFF",
        "CARD_BG":     "FAFAFA",
        "DARK_BG":     "18181B",
        "DARKEST":     "09090B",
        "ACCENT":      "18181B",
        "ACCENT2":     "3F3F46",
        "GREEN":       "16A34A",
        "INDIGO":      "4338CA",
        "AMBER":       "D97706",
        "RED":         "DC2626",
        "T_PRIMARY":   "09090B",
        "T_SECOND":    "52525B",
        "T_MUTED":     "A1A1AA",
        "BORDER":      "E4E4E7",
        "BORDER_DK":   "27272A",
        "GRAD_DARK":   ("18181B", "09090B"),
        "GRAD_ACCENT": ("18181B", "3F3F46"),
        "GRAD_CARD":   ("FFFFFF", "F4F4F5"),
        "DECO_GLOW":   False,
        "DECO_NOISE":  False,
        "DECO_RINGS":  False,
    },

    # ── 7. sunset — dark warm + orange-red, bold / creative ───────────────────
    "sunset": {
        "WHITE_BG":    "FFF7ED",
        "CARD_BG":     "FED7AA",
        "DARK_BG":     "1C0A00",
        "DARKEST":     "0D0500",
        "ACCENT":      "EA580C",
        "ACCENT2":     "FB923C",
        "GREEN":       "65A30D",
        "INDIGO":      "C026D3",
        "AMBER":       "EAB308",
        "RED":         "E11D48",
        "T_PRIMARY":   "1C0A00",
        "T_SECOND":    "9A3412",
        "T_MUTED":     "C2410C",
        "BORDER":      "FED7AA",
        "BORDER_DK":   "7C2D12",
        "GRAD_DARK":   ("1C0A00", "0D0500"),
        "GRAD_ACCENT": ("EA580C", "EAB308"),
        "GRAD_CARD":   ("FFF7ED", "FED7AA"),
    },

    # ── 8. ocean — deep teal + cyan, data / analytics ─────────────────────────
    "ocean": {
        "WHITE_BG":    "F0FDFA",
        "CARD_BG":     "CCFBF1",
        "DARK_BG":     "042F2E",
        "DARKEST":     "021A19",
        "ACCENT":      "0D9488",
        "ACCENT2":     "2DD4BF",
        "GREEN":       "0EA5E9",
        "INDIGO":      "0891B2",
        "AMBER":       "CA8A04",
        "RED":         "B91C1C",
        "T_PRIMARY":   "042F2E",
        "T_SECOND":    "0F766E",
        "T_MUTED":     "5EEAD4",
        "BORDER":      "99F6E4",
        "BORDER_DK":   "134E4A",
        "GRAD_DARK":   ("042F2E", "021A19"),
        "GRAD_ACCENT": ("0D9488", "2DD4BF"),
        "GRAD_CARD":   ("F0FDFA", "CCFBF1"),
    },
}


def get_theme(name: str) -> dict:
    """Return preset dict; defaults to tech_dark for unknown names."""
    return THEMES.get(name, THEMES["tech_dark"])


THEME_NAMES = list(THEMES.keys())
