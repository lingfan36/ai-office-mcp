"""
Slide layout renderers.
Each function signature: render_<name>(slide, spec: dict) -> None
All coordinates are in pixels on a 1280 × 720 canvas.
"""

import os
from utils import (px, set_bg, add_rect, add_oval, add_line,
                   add_text, title_bar, page_num,
                   set_grad_bg, apply_grad, apply_shadow,
                   add_rect_alpha, set_bg_image,
                   add_glow_orb, set_multi_grad_bg,
                   apply_glass, apply_glass_border, add_noise_texture)
import theme
import images as _images
from measure import text_height as _th
from layout_engine import LayoutFlow


# ─────────────────────────────────────────────────────────────────────────────
# 1. cover — dark anchor page
# ─────────────────────────────────────────────────────────────────────────────

def render_cover(slide, spec):
    img_path = _images.fetch_pexels(spec.get("image_bg", "dark abstract technology"))
    if img_path:
        set_bg(slide, "000000")
        set_bg_image(slide, img_path)
        add_rect_alpha(slide, 0, 0, 1280, 720, theme.DARK_BG, 75)
    else:
        set_multi_grad_bg(slide, [
            (0,   theme.DARKEST),
            (45,  theme.DARK_BG),
            (100, theme.DARKEST),
        ], angle=135)

    # Ambient glow orbs
    add_glow_orb(slide, 1160, 80,  340, theme.ACCENT,  22)
    add_glow_orb(slide, 160,  660, 280, theme.INDIGO,  18)
    add_glow_orb(slide, 640,  360, 180, theme.ACCENT2, 6)

    # Noise grain
    add_noise_texture(slide, alpha=5)

    # Left accent bar (gradient)
    bar = add_rect(slide, 0, 80, 5, 560, fill=theme.ACCENT)
    apply_grad(bar, *theme.GRAD_ACCENT, angle=90)

    # Decorative rings (top-right corner)
    if theme.DECO_RINGS:
        add_oval(slide, 1240,  60, 260, line=theme.ACCENT,  line_w=0.5)
        add_oval(slide, 1240,  60, 175, line=theme.ACCENT2, line_w=1.0)
        add_oval(slide, 1240,  60,  85, fill=theme.ACCENT)

    # Content
    # ── Dynamic content stack (prevents title/subtitle overlap) ─────────────
    _CX   = 80
    _CW   = 880
    title = spec["title"]
    sub   = spec.get("subtitle",  "")
    sub2  = spec.get("subtitle2", "")
    badge = spec.get("badge",     "")

    title_h = max(110, _th(title, 88, _CW, line_spacing=0.9))
    sub_h   = max(36,  _th(sub,   20, 860)) + 8  if sub   else 0
    sub2_h  = max(28,  _th(sub2,  16, 860)) + 8  if sub2  else 0

    flow = LayoutFlow(136, 700, gap=0)
    title_y, _ = flow.next(title_h, gap=10)
    div_y,   _ = flow.next(4,       gap=14)
    sub_y,   _ = flow.next(sub_h,   gap=8)
    sub2_y,  _ = flow.next(sub2_h,  gap=16)
    badge_y    = flow.y

    add_text(slide, spec.get("tag", ""), _CX, 88, 720, 28,
             size=11, bold=True, color=theme.ACCENT2)
    add_text(slide, title, _CX, title_y, _CW, title_h,
             font=theme.FONT_CJK, size=88, bold=True, color="FFFFFF",
             line_spacing=0.9)

    divider = add_rect(slide, _CX, div_y, 360, 4, fill=theme.ACCENT)
    apply_grad(divider, *theme.GRAD_ACCENT, angle=0)

    if sub:
        add_text(slide, sub, _CX, sub_y, 860, sub_h,
                 size=20, color="94A3B8")
    if sub2:
        add_text(slide, sub2, _CX, sub2_y, 860, sub2_h,
                 size=16, color="64748B")

    if badge:
        br = add_rect(slide, _CX, badge_y, 300, 38, fill=theme.DARKEST,
                      line=theme.ACCENT, line_w=1.0, radius=19)
        apply_glass(br, 30)
        add_text(slide, badge, _CX, badge_y + 1, 300, 36, size=12,
                 color=theme.ACCENT, align="center")

    page_num(slide, spec.get("page", ""))


# ─────────────────────────────────────────────────────────────────────────────
# 2. hero_stat — breathing page with giant number + 3 columns
# ─────────────────────────────────────────────────────────────────────────────

def render_hero_stat(slide, spec):
    dark = theme.is_dark_bg()
    # Background
    set_multi_grad_bg(slide, [
        (0,   theme.DARKEST if dark else "F8FAFC"),
        (50,  theme.DARK_BG if dark else "FFFFFF"),
        (100, theme.DARKEST if dark else "EEF2FF"),
    ], angle=135)
    add_glow_orb(slide, 1200, 90, 320, theme.ACCENT,  20 if dark else 8)
    add_glow_orb(slide, 100,  660, 200, theme.INDIGO, 12 if dark else 5)
    add_noise_texture(slide, alpha=5 if dark else 3)

    tc  = "F1F5F9" if dark else theme.T_PRIMARY
    tm  = "94A3B8" if dark else "64748B"

    if spec.get("title"):
        add_text(slide, spec["title"], 60, 38, 1100, 52,
                 size=32, bold=True, color=tc)
        add_rect(slide, 60, 92, 80, 3, fill=theme.ACCENT)

    # Giant hero stat — wide fixed box, no n_w hack
    hero_n = spec.get("hero_number", "")
    hero_u = spec.get("hero_unit",   "")
    add_text(slide, hero_n, 60, 110, 700, 204,
             font=theme.FONT_NUM, size=130, bold=True, color=theme.ACCENT,
             line_spacing=0.85)

    # Unit  +  label on the line below the number (no side-by-side overlap)
    hero_label = spec.get("hero_label", "")
    if hero_u:
        add_text(slide, hero_u, 60, 318, 90, 44,
                 font=theme.FONT_NUM, size=30, bold=True, color=tc)
        add_text(slide, hero_label, 158, 322, 740, 38, size=15, color=tm)
    else:
        add_text(slide, hero_label, 60, 322, 820, 38, size=15, color=tm)

    # Three column glass cards
    col_colors = [theme.ACCENT, theme.GREEN, theme.INDIGO]
    card_fill  = "FFFFFF"
    card_line  = "FFFFFF" if dark else "CBD5E1"
    card_alpha = 9 if dark else 100

    for i, col in enumerate(spec.get("columns", [])[:3]):
        cx, cw = 60 + i * 400, 360
        c = col_colors[i]
        add_rect(slide, cx, 375, cw, 4, fill=c)
        card = add_rect(slide, cx, 379, cw, 278, fill=card_fill,
                        line=card_line, line_w=0.5, radius=10)
        if dark:
            apply_glass(card, fill_alpha=card_alpha)
            apply_glass_border(card, border_alpha=18)
        apply_shadow(card, blur=70000, dist=28000, alpha=28 if dark else 15)
        add_text(slide, col.get("title", ""), cx + 22, 400, cw - 44, 34,
                 size=16, bold=True, color=tc)
        add_text(slide, col.get("body", ""), cx + 22, 440, cw - 44, 200,
                 size=14, color=tm, wrap=True)

    page_num(slide, spec.get("page", ""))


# ─────────────────────────────────────────────────────────────────────────────
# 3. solution — dark-left / light-right split
# ─────────────────────────────────────────────────────────────────────────────

def render_solution(slide, spec):
    set_bg(slide, theme.WHITE_BG)
    add_rect(slide, 0, 0, 520, 720, fill=theme.DARK_BG)
    add_rect(slide, 0, 0, 4, 720, fill=theme.ACCENT)
    add_noise_texture(slide, alpha=4)

    hero_txt = spec.get("hero", "")
    hero_h   = max(100, _th(hero_txt, 28, 440)) + 12

    flow_l = LayoutFlow(60, 700, gap=0)
    tag_y,  _ = flow_l.next(28,     gap=20)
    hero_y, _ = flow_l.next(hero_h, gap=10)
    div_y,  _ = flow_l.next(3,      gap=18)

    add_text(slide, spec.get("tag", ""), 42, tag_y, 440, 28,
             size=12, bold=True, color=theme.ACCENT2)
    add_text(slide, hero_txt, 42, hero_y, 440, hero_h,
             size=28, bold=True, color="FFFFFF", wrap=True)
    add_rect(slide, 42, div_y, 100, 3, fill=theme.ACCENT)

    bullet_colors = [theme.ACCENT, theme.ACCENT2, "7DD3FC"]
    bullets = spec.get("bullets", [])[:3]
    bflow = LayoutFlow(flow_l.y, 700, gap=0)
    for i, b in enumerate(bullets):
        b_title = b.get("title", "")
        b_body  = b.get("body",  "")
        title_h = max(28, _th(b_title, 15, 420)) + 4
        body_h  = max(32, _th(b_body,  13, 420)) + 4
        slot_h  = title_h + body_h + 4
        by, _   = bflow.next(slot_h, gap=12)
        c = bullet_colors[i]
        add_rect(slide, 42, by, 3, slot_h, fill=c)
        add_text(slide, b_title, 60, by + 2, 420, title_h,
                 size=15, bold=True, color="FFFFFF")
        add_text(slide, b_body, 60, by + title_h + 4, 420, body_h,
                 size=13, color="94A3B8")

    card_colors = [theme.ACCENT, theme.GREEN, theme.INDIGO]
    for i, card in enumerate(spec.get("cards", [])[:3]):
        cy, c = 80 + i * 192, card_colors[i]
        add_rect(slide, 540, cy, 720, 172, fill="FFFFFF",
                 line="CBD5E1", line_w=1, radius=10)
        add_rect(slide, 540, cy, 5, 172, fill=c, radius=2)
        add_oval(slide, 602, cy + 86, 32, fill=theme.tint(c))
        add_text(slide, card.get("icon_char", "◆"), 576, cy + 68,
                 52, 38, size=20, color=c, align="center")
        add_text(slide, card.get("title", ""), 652, cy + 30, 580, 34,
                 size=18, bold=True, color=theme.T_PRIMARY)
        add_text(slide, card.get("body", ""), 652, cy + 68, 580, 88,
                 size=14, color="64748B", wrap=True)

    page_num(slide, spec.get("page", ""))


# ─────────────────────────────────────────────────────────────────────────────
# 4. pipeline — horizontal N-stage flow
# ─────────────────────────────────────────────────────────────────────────────

def render_pipeline(slide, spec):
    set_bg(slide, theme.WHITE_BG)
    add_noise_texture(slide, alpha=3)
    if spec.get("title"):
        title_bar(slide, 60, 28, spec["title"], size=30)

    stages = spec.get("stages", [])
    n = min(len(stages), 5)
    if n == 0:
        return
    default_colors = [theme.INDIGO, theme.ACCENT, theme.GREEN, theme.AMBER, theme.DARK_BG]
    stage_w = 218
    total_w = 1160
    gap = (total_w - n * stage_w) // max(n - 1, 1)

    for i, stage in enumerate(stages[:n]):
        sx = 60 + i * (stage_w + gap)
        c = stage.get("color", default_colors[i % len(default_colors)])
        opt = stage.get("optional", False)

        add_rect(slide, sx, 100, stage_w, 546,
                 fill="F0F9FF" if opt else "FAFBFC",
                 line=c if opt else "CBD5E1", line_w=1.5 if opt else 1, radius=12)
        add_rect(slide, sx, 100, stage_w, 8, fill=c, radius=4)

        add_text(slide, str(i + 1), sx, 118, stage_w, 30,
                 font=theme.FONT_NUM, size=22, bold=True, color=c, align="center")
        stage_title = stage.get("title", "")
        st_h = max(40, _th(stage_title, 15, stage_w - 24)) + 8
        add_text(slide, stage_title, sx + 12, 156, stage_w - 24, st_h,
                 size=15, bold=True, color=theme.T_PRIMARY, align="center", wrap=True)

        bflow = LayoutFlow(156 + st_h + 8, 590, gap=4)
        for bullet in stage.get("bullets", [])[:5]:
            bh = max(28, _th(f"· {bullet}", 12, stage_w - 28)) + 4
            by, _ = bflow.next(bh)
            add_text(slide, f"· {bullet}", sx + 14, by, stage_w - 28, bh,
                     size=12, color="475569", wrap=True)

        if opt:
            add_rect(slide, sx + stage_w // 2 - 36, 558, 72, 24, fill=c, radius=12)
            add_text(slide, "可选", sx + stage_w // 2 - 36, 559, 72, 22,
                     size=10, bold=True, color="FFFFFF", align="center")

        if i < n - 1:
            ax = sx + stage_w + gap // 2 - 8
            add_text(slide, "▶", ax, 348, 20, 28, size=14, color=theme.T_MUTED, align="center")

    page_num(slide, spec.get("page", ""))


# ─────────────────────────────────────────────────────────────────────────────
# 5. hub_spoke — center hub + up to 6 spoke cards
# ─────────────────────────────────────────────────────────────────────────────

def render_hub_spoke(slide, spec):
    set_bg(slide, theme.WHITE_BG)
    add_noise_texture(slide, alpha=3)
    if spec.get("title"):
        title_bar(slide, 60, 28, spec["title"], size=30)

    cx, cy = 640, 400
    add_oval(slide, cx, cy, 160, line=theme.ACCENT, line_w=0.5)
    add_oval(slide, cx, cy, 250, line="CBD5E1", line_w=0.5)
    add_oval(slide, cx, cy, 72, fill=theme.ACCENT)
    hub = spec.get("hub", {})
    add_text(slide, hub.get("title", "AI\n引擎"), cx - 65, cy - 26,
             130, 52, size=14, bold=True, color="FFFFFF", align="center")

    positions = [
        (cx, 115), (cx + 312, 230), (cx + 312, 570),
        (cx, 655), (cx - 312, 570), (cx - 312, 230),
    ]
    spoke_colors = [theme.ACCENT, theme.ACCENT2, theme.GREEN,
                    theme.INDIGO, theme.AMBER, "EC4899"]
    sw, sh = 200, 108

    for i, (px_, py_) in enumerate(positions[:len(spec.get("spokes", []))]):
        spoke = spec["spokes"][i]
        c = spoke.get("color", spoke_colors[i % len(spoke_colors)])
        add_line(slide, cx, cy, px_, py_, color=c, width=1.2)
        cxs = px_ - sw // 2
        cys = max(5, min(py_ - sh // 2, theme.SLIDE_H_PX - sh - 5))
        add_rect(slide, cxs, cys, sw, sh, fill="FFFFFF",
                 line=c, line_w=1.5, radius=8)
        add_rect(slide, cxs, cys, sw, 4, fill=c, radius=2)
        add_text(slide, spoke.get("title", ""), cxs + 10, cys + 12,
                 sw - 20, 28, size=13, bold=True, color=theme.T_PRIMARY)
        add_text(slide, spoke.get("body", ""), cxs + 10, cys + 42,
                 sw - 20, 60, size=11, color="64748B", wrap=True)

    page_num(slide, spec.get("page", ""))


# ─────────────────────────────────────────────────────────────────────────────
# 6. kpi_grid — 2 × 2 metric cards
# ─────────────────────────────────────────────────────────────────────────────

def render_kpi_grid(slide, spec):
    dark = theme.is_dark_bg()
    set_multi_grad_bg(slide, [
        (0,   theme.DARKEST if dark else "F8FAFC"),
        (55,  theme.DARK_BG if dark else "FFFFFF"),
        (100, theme.DARKEST if dark else "EEF2FF"),
    ], angle=160)
    add_glow_orb(slide, 1180, 100, 360, theme.ACCENT,  18 if dark else 7)
    add_glow_orb(slide, 120,  650, 300, theme.INDIGO,  15 if dark else 5)
    add_noise_texture(slide, alpha=5 if dark else 3)

    if spec.get("title"):
        title_bar(slide, 60, 26, spec["title"], size=30)

    positions     = [(60, 106), (660, 106), (60, 406), (660, 406)]
    default_colors = [theme.ACCENT, theme.GREEN, theme.AMBER, theme.INDIGO]
    cw, ch = 560, 272

    for i, card in enumerate(spec.get("cards", [])[:4]):
        cx, cy = positions[i]
        c    = card.get("color", default_colors[i])
        card_dark = card.get("dark", False)

        bg   = theme.DARKEST if card_dark else "FFFFFF"
        kcard = add_rect(slide, cx, cy, cw, ch, fill=bg,
                         line="FFFFFF" if dark else "CBD5E1",
                         line_w=0.5 if dark else 1, radius=14)

        if dark and not card_dark:
            apply_glass(kcard, fill_alpha=9)
            apply_glass_border(kcard, border_alpha=16)
        elif not dark and not card_dark:
            apply_grad(kcard, "FFFFFF", "F0F9FF", angle=90)

        apply_shadow(kcard, blur=80000, dist=35000, alpha=30 if dark else 12)

        # Coloured top accent bar (replaces per-card orbs which caused overlaps)
        add_rect(slide, cx, cy, cw, 5, fill=c, radius=3)

        tc = "FFFFFF" if (dark or card_dark) else theme.T_PRIMARY
        tm = "94A3B8" if (dark or card_dark) else "64748B"
        vc = c if not card_dark else "FFFFFF"

        add_text(slide, card.get("value", ""), cx + 28, cy + 24, 420, 108,
                 font=theme.FONT_NUM, size=72, bold=True, color=vc)
        add_text(slide, card.get("unit", ""), cx + 28, cy + 138, cw - 56, 34,
                 size=19, bold=True, color=tc)
        add_text(slide, card.get("label", ""), cx + 28, cy + 176, cw - 56, 30,
                 size=13, color=tm)
        if card.get("note"):
            add_text(slide, card["note"], cx + 28, cy + 210, cw - 56, 48,
                     size=12, color=tm, wrap=True)

    page_num(slide, spec.get("page", ""))


# ─────────────────────────────────────────────────────────────────────────────
# 7. radar — matplotlib radar image + score table
# ─────────────────────────────────────────────────────────────────────────────

def render_radar(slide, spec):
    set_bg(slide, theme.WHITE_BG)
    add_noise_texture(slide, alpha=3)
    if spec.get("title"):
        title_bar(slide, 60, 28, spec["title"], size=30)

    try:
        img_buf = _radar_image(spec)
        slide.shapes.add_picture(img_buf, px(55), px(88), px(580), px(580))
    except Exception:
        add_rect(slide, 60, 100, 580, 568, fill="F1F5F9", line="CBD5E1", line_w=1, radius=8)
        add_text(slide, "Radar Chart\n(pip install matplotlib)", 60, 360,
                 580, 56, size=14, color="94A3B8", align="center")

    # Legend
    add_oval(slide, 716, 110, 7, fill=theme.ACCENT)
    add_text(slide, spec.get("legend_us", "我们"), 728, 102, 120, 22,
             size=13, bold=True, color=theme.ACCENT)
    add_oval(slide, 840, 110, 7, fill="94A3B8")
    add_text(slide, spec.get("legend_avg", "行业均值"), 852, 102, 120, 22,
             size=13, color="94A3B8")

    dims = spec.get("dimensions", [])
    row_h = 62
    for i, dim in enumerate(dims[:8]):
        ry = 140 + i * row_h
        if i % 2 == 0:
            add_rect(slide, 680, ry, 560, row_h, fill="F8FAFC", radius=4)
        add_text(slide, dim.get("name", ""), 698, ry + 10, 240, 24,
                 size=14, color=theme.T_PRIMARY)
        bar_x, bw = 950, 240
        us, avg = dim.get("us", 9), dim.get("avg", 5)
        add_rect(slide, bar_x, ry + 22, int(bw * avg / 10), 10,
                 fill="CBD5E1", radius=5)
        add_rect(slide, bar_x, ry + 38, int(bw * us / 10), 10,
                 fill=theme.ACCENT, radius=5)
        add_text(slide, str(us), 1200, ry + 28, 38, 22,
                 font=theme.FONT_NUM, size=14, bold=True, color=theme.ACCENT, align="right")

    page_num(slide, spec.get("page", ""))


def _radar_image(spec):
    import io
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.font_manager as fm
    import numpy as np

    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei",
                                       "PingFang SC", "Arial Unicode MS", "DejaVu Sans"]
    plt.rcParams["font.family"]     = "sans-serif"
    plt.rcParams["axes.unicode_minus"] = False

    dims = spec.get("dimensions", [])
    labels = [d.get("name", "") for d in dims]
    us     = [d.get("us",  9) for d in dims]
    avg    = [d.get("avg", 5) for d in dims]
    n = len(labels)
    if n < 3:
        raise ValueError("need ≥ 3 dimensions")

    angles = np.linspace(0, 2 * np.pi, n, endpoint=False).tolist() + [0]
    us  += us[:1]
    avg += avg[:1]

    fig, ax = plt.subplots(figsize=(5.5, 5.5), subplot_kw=dict(polar=True))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    ax.plot(angles, us,  "o-",  lw=2,   color="#0EA5E9", label=spec.get("legend_us",  "我们"))
    ax.fill(angles, us,  alpha=0.18, color="#0EA5E9")
    ax.plot(angles, avg, "s--", lw=1.5, color="#94A3B8", label=spec.get("legend_avg", "行业均值"))
    ax.fill(angles, avg, alpha=0.10, color="#94A3B8")
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, fontsize=10)
    ax.set_ylim(0, 10)
    ax.set_yticks([2, 4, 6, 8, 10])
    ax.set_yticklabels(["2","4","6","8","10"], fontsize=7, color="#94A3B8")
    ax.grid(color="#E2E8F0", lw=0.8)
    ax.spines["polar"].set_color("#E2E8F0")

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────────────────────
# 8. comparison — competitor feature-comparison table
# ─────────────────────────────────────────────────────────────────────────────

def render_comparison(slide, spec):
    set_bg(slide, theme.WHITE_BG)
    add_noise_texture(slide, alpha=3)
    if spec.get("title"):
        title_bar(slide, 60, 28, spec["title"], size=30)

    comps    = spec.get("competitors", [])   # [{"name":..., "note":...}, ...]
    features = spec.get("features", [])       # [{"name":..., "values":[T/F/str]}, ...]

    label_w = 200
    us_w    = 240
    oth_w   = 200
    gap     = 10
    tx, ty  = 60, 98
    rh, hh  = 60, 54

    # Compute column left-edges
    col_xs = [tx + label_w + gap]
    for i in range(1, len(comps)):
        col_xs.append(col_xs[-1] + (us_w if i == 1 else oth_w) + gap)
    col_ws = [us_w] + [oth_w] * (len(comps) - 1)

    # Header
    add_rect(slide, tx, ty, label_w, hh, fill="F1F5F9", line="CBD5E1", line_w=1)
    for i, (cx, cw, comp) in enumerate(zip(col_xs, col_ws, comps)):
        bg = theme.DARK_BG if i == 0 else "F8FAFC"
        add_rect(slide, cx, ty, cw, hh, fill=bg, line="CBD5E1", line_w=1)
        add_text(slide, comp.get("name", ""), cx + 8, ty + 8, cw - 16, 24,
                 size=15, bold=True,
                 color="FFFFFF" if i == 0 else theme.T_PRIMARY, align="center")
        if comp.get("note"):
            add_text(slide, comp["note"], cx + 8, ty + 32, cw - 16, 18,
                     size=11, color=theme.ACCENT2 if i == 0 else "94A3B8", align="center")

    # Feature rows
    for r, feat in enumerate(features[:8]):
        ry = ty + hh + r * rh
        row_bg = "FAFBFC" if r % 2 == 0 else "FFFFFF"
        add_rect(slide, tx, ry, label_w, rh, fill=row_bg, line="E2E8F0", line_w=0.5)
        add_text(slide, feat.get("name", ""), tx + 10, ry + 16, label_w - 20, 30,
                 size=13, color=theme.T_PRIMARY)
        for i, (cx, cw, val) in enumerate(zip(col_xs, col_ws, feat.get("values", []))):
            cell_bg = "EFF8FF" if i == 0 else row_bg
            add_rect(slide, cx, ry, cw, rh, fill=cell_bg, line="E2E8F0", line_w=0.5)
            if isinstance(val, bool):
                sym, col = ("✓", theme.GREEN) if val else ("✗", theme.RED)
                add_text(slide, sym, cx, ry + 12, cw, 34,
                         size=18, bold=True, color=col, align="center")
            else:
                add_text(slide, str(val), cx + 8, ry + 14, cw - 16, 32,
                         size=12, color=theme.T_PRIMARY if i == 0 else "64748B",
                         align="center")

    page_num(slide, spec.get("page", ""))


# ─────────────────────────────────────────────────────────────────────────────
# 9. business — vertical list with price tags (business model)
# ─────────────────────────────────────────────────────────────────────────────

def render_business(slide, spec):
    dark = theme.is_dark_bg()
    set_multi_grad_bg(slide, [
        (0,   theme.DARKEST if dark else "F8FAFC"),
        (50,  theme.DARK_BG if dark else "FAFBFC"),
        (100, theme.DARKEST if dark else "F1F5F9"),
    ], angle=150)
    add_glow_orb(slide, 1200, 360, 280, theme.ACCENT,  12 if dark else 5)
    add_noise_texture(slide, alpha=5 if dark else 3)

    tc = "F1F5F9" if dark else theme.T_PRIMARY
    tm = "94A3B8" if dark else "64748B"

    if spec.get("title"):
        title_bar(slide, 60, 26, spec["title"], size=30)
    if spec.get("subtitle"):
        add_text(slide, spec["subtitle"], 60, 94, 920, 32, size=15, color=tm)

    default_colors = [theme.ACCENT, theme.GREEN, theme.INDIGO]
    track_x = 160

    # Timeline track
    track_bar = add_rect(slide, track_x - 2, 152, 4, 504, fill=theme.ACCENT)
    apply_grad(track_bar, *theme.GRAD_ACCENT, angle=90)
    add_text(slide, "▼", track_x - 8, 658, 18, 22, size=12, color=theme.INDIGO, align="center")

    # Pre-compute card heights so stacking doesn't overlap
    items = spec.get("items", [])[:3]
    _title_hs = [max(30, _th(it.get("title",""), 18, 640)) + 4 for it in items]
    _body_hs  = [max(26, _th(it.get("body", ""), 13, 640)) + 4 for it in items]
    _card_hs  = [max(80, th + bh + 24) for th, bh in zip(_title_hs, _body_hs)]
    _spacing  = max(180, max(_card_hs, default=[80]) + 60) if items else 180

    for i, item in enumerate(items):
        c  = item.get("color", default_colors[i])
        ch = _card_hs[i]
        iy = 152 + _spacing // 2 + i * _spacing   # node centre

        # Node
        add_oval(slide, track_x, iy, 30, fill=theme.tint(c))
        add_oval(slide, track_x, iy, 20, fill=c)
        add_text(slide, f"0{i+1}", track_x - 12, iy - 11, 24, 22,
                 font=theme.FONT_NUM, size=13, bold=True, color="FFFFFF", align="center")
        add_line(slide, track_x + 20, iy, 244, iy, color=c, width=2)

        # Card centred on node
        card_fill = "FFFFFF"
        card_line = "FFFFFF" if dark else "CBD5E1"
        cy = iy - ch // 2
        bcard = add_rect(slide, 244, cy, 984, ch, fill=card_fill,
                         line=card_line, line_w=0.5, radius=12)
        if dark:
            apply_glass(bcard, fill_alpha=9)
            apply_glass_border(bcard, border_alpha=16)
        apply_shadow(bcard, blur=70000, dist=28000, alpha=30 if dark else 12)

        add_rect(slide, 244, cy, 5, ch, fill=c, radius=2)
        add_oval(slide, 298, iy, 24, fill=theme.tint(c))

        th_ = _title_hs[i]
        bh_ = _body_hs[i]
        add_text(slide, item.get("title", ""), 338, cy + 12, 640, th_,
                 size=18, bold=True, color=tc)
        add_text(slide, item.get("body",  ""), 338, cy + 12 + th_ + 2, 640, bh_,
                 size=13, color=tm)

        if item.get("price"):
            pw = 148
            price_card = add_rect(slide, 1078, iy - 36, pw, 72, fill=c, radius=10)
            apply_shadow(price_card, blur=50000, dist=20000, alpha=30)
            add_text(slide, item["price"], 1078, iy - 30, pw, 34,
                     font=theme.FONT_NUM, size=22, bold=True, color="FFFFFF", align="center")
            if item.get("price_label"):
                add_text(slide, item["price_label"], 1078, iy + 8, pw, 22,
                         size=11, color="FFFFFF", align="center")

    page_num(slide, spec.get("page", ""))


# ─────────────────────────────────────────────────────────────────────────────
# 10. timeline — horizontal milestone track
# ─────────────────────────────────────────────────────────────────────────────

def render_timeline(slide, spec):
    set_bg(slide, theme.WHITE_BG)
    add_noise_texture(slide, alpha=3)
    if spec.get("title"):
        title_bar(slide, 60, 28, spec["title"], size=30)

    mss = spec.get("milestones", [])
    n = min(len(mss), 5)
    if n == 0:
        return

    ty = 380
    tx1, tx2 = 120, 1160
    add_rect(slide, tx1, ty - 4, tx2 - tx1, 8, fill="E2E8F0", radius=4)

    step  = (tx2 - tx1) / max(n - 1, 1)
    node_xs = [int(tx1 + i * step) for i in range(n)]

    cur_idx = next((i for i, m in enumerate(mss[:n]) if m.get("state") == "current"), -1)
    if cur_idx >= 0:
        add_rect(slide, tx1, ty - 4, node_xs[cur_idx] - tx1 + 30, 8,
                 fill=theme.ACCENT, radius=4)

    cw, ch = 180, 125
    above_y, below_y = 90, 415

    for i, ms in enumerate(mss[:n]):
        nx    = node_xs[i]
        state = ms.get("state", "planned")
        above = (i % 2 == 0)
        card_y = above_y if above else below_y
        cxc   = nx - cw // 2

        # Node
        r = 20 if state == "current" else 14
        if state == "current":
            add_oval(slide, nx, ty, 28, fill="CEE9F8")
            add_oval(slide, nx, ty, r,  fill="FFFFFF", line=theme.ACCENT,  line_w=3)
            add_oval(slide, nx, ty,  7, fill=theme.ACCENT)
        elif state == "target":
            add_oval(slide, nx, ty, r,  fill="FFFFFF", line=theme.INDIGO,  line_w=3)
        else:
            add_oval(slide, nx, ty, r,  fill="FFFFFF", line=theme.T_MUTED, line_w=2.5)

        # Connector
        lc = theme.ACCENT if state == "current" else (
             theme.INDIGO if state == "target" else theme.T_MUTED)
        if above:
            add_line(slide, nx, card_y + ch, nx, ty - r, color=lc, width=2)
        else:
            add_line(slide, nx, ty + r, nx, card_y, color=lc, width=2)

        # Card
        bc = theme.ACCENT if state == "current" else (
             theme.INDIGO if state == "target" else "CBD5E1")
        bw = 2.0 if state != "planned" else 1.5
        add_rect(slide, cxc, card_y, cw, ch, fill="FFFFFF", line=bc, line_w=bw, radius=12)
        if state in ("current",):
            add_rect(slide, cxc, card_y, cw, 6, fill=theme.ACCENT, radius=3)
        elif state == "target":
            add_rect(slide, cxc, card_y, cw, 6, fill=theme.INDIGO, radius=3)

        dc = theme.ACCENT if state == "current" else (
             theme.INDIGO if state == "target" else "64748B")
        add_text(slide, ms.get("date", ""), cxc, card_y + 14, cw, 22,
                 size=13, bold=True, color=dc, align="center")
        add_text(slide, ms.get("title", ""), cxc + 10, card_y + 38, cw - 20, 30,
                 size=14, bold=True, color=theme.T_PRIMARY, align="center", wrap=True)
        add_text(slide, ms.get("body", ""), cxc + 10, card_y + 70, cw - 20, 44,
                 size=12, color="94A3B8", align="center", wrap=True)

        if state == "current":
            add_rect(slide, nx - 36, card_y + ch - 16, 72, 22, fill=theme.ACCENT, radius=11)
            add_text(slide, "进行中", nx - 36, card_y + ch - 15, 72, 18,
                     size=10, bold=True, color="FFFFFF", align="center")
        elif state == "target":
            add_rect(slide, nx - 36, card_y + ch - 16, 72, 22, fill=theme.INDIGO, radius=11)
            add_text(slide, "融资目标", nx - 36, card_y + ch - 15, 72, 18,
                     size=10, bold=True, color="FFFFFF", align="center")

    if spec.get("footer"):
        add_line(slide, 60, 630, 1220, 630, color="E2E8F0")
        add_text(slide, spec["footer"], 320, 636, 640, 28,
                 size=13, color="94A3B8", align="center")

    page_num(slide, spec.get("page", ""))


# ─────────────────────────────────────────────────────────────────────────────
# 11. fundraising — dark split with breakdown cards
# ─────────────────────────────────────────────────────────────────────────────

def render_fundraising(slide, spec):
    set_multi_grad_bg(slide, [
        (0,   theme.DARKEST),
        (40,  theme.DARK_BG),
        (70,  theme.DARKEST),
        (100, "010409"),
    ], angle=150)

    # Ambient orbs
    add_glow_orb(slide, 1240,  60, 360, theme.ACCENT,  20)
    add_glow_orb(slide, 80,   660, 260, theme.INDIGO,  16)
    add_glow_orb(slide, 540,  360, 200, theme.ACCENT2,  7)

    add_noise_texture(slide, alpha=5)

    # Decorative rings
    if theme.DECO_RINGS:
        add_oval(slide, 1240,  60, 250, line=theme.ACCENT,  line_w=0.5)
        add_oval(slide, 1240,  60, 160, line=theme.ACCENT2, line_w=0.8)

    # Vertical divider
    add_line(slide, 540, 76, 540, 648, color="1E3A5F", width=1)

    # Left panel
    add_text(slide, spec.get("tag", "FUNDRAISING · PRE-A"), 80, 88, 420, 28,
             size=11, bold=True, color=theme.ACCENT2)
    add_text(slide, spec.get("title", ""), 80, 144, 428, 80,
             size=30, bold=True, color="FFFFFF", wrap=True)

    divider = add_rect(slide, 80, 232, 180, 3, fill=theme.ACCENT)
    apply_grad(divider, *theme.GRAD_ACCENT, angle=0)

    # Amount
    add_text(slide, spec.get("amount", "¥3000"), 80, 246, 380, 122,
             font=theme.FONT_NUM, size=104, bold=True, color=theme.ACCENT,
             line_spacing=0.85)
    add_text(slide, spec.get("amount_unit", "万"), 80, 374, 200, 54,
             size=38, bold=True, color="FFFFFF")

    # Round badge (glass)
    br = add_rect(slide, 80, 410, 148, 34, fill=theme.DARKEST,
                  line=theme.ACCENT, line_w=1, radius=17)
    apply_glass(br, 35)
    add_text(slide, spec.get("round_label", "Pre-A 轮"), 80, 411, 148, 32,
             size=14, color=theme.ACCENT, align="center")

    add_text(slide, spec.get("target_time", ""), 80, 468, 432, 28, size=15, color="94A3B8")
    add_text(slide, spec.get("target_arr",  ""), 80, 498, 432, 26, size=14, color="64748B")

    # Right panel: breakdown glass cards — dynamic heights
    bk_colors = [theme.ACCENT, theme.ACCENT2, theme.GREEN]
    bk_flow   = LayoutFlow(106, 610, gap=12)
    for i, bk in enumerate(spec.get("breakdowns", [])[:3]):
        c = bk.get("color", bk_colors[i])
        title_h = max(30, _th(bk.get("title",    ""), 18, 508)) + 4
        sub_h   = max(22, _th(bk.get("subtitle", ""), 13, 508)) + 4
        body_h  = max(22, _th(bk.get("body",     ""), 13, 508)) + 4
        ch      = max(130, 24 + title_h + sub_h + body_h + 16)
        by, _   = bk_flow.next(ch)

        card = add_rect(slide, 562, by, 650, ch, fill="FFFFFF",
                        line="FFFFFF", line_w=0.5, radius=12)
        apply_glass(card, fill_alpha=8)
        apply_glass_border(card, border_alpha=16)
        apply_shadow(card, blur=80000, dist=30000, alpha=35)

        add_glow_orb(slide, 1210, by + ch // 2, 110, c, 15)
        add_rect(slide, 562, by, 4, ch, fill=c, radius=2)

        add_text(slide, bk.get("pct", ""), 578, by + (ch - 84) // 2, 86, 84,
                 font=theme.FONT_NUM, size=52, bold=True, color=c, align="center")

        add_line(slide, 668, by + 18, 668, by + ch - 18, color="334155", width=1)

        ty = by + 20
        add_text(slide, bk.get("title",    ""), 684, ty, 508, title_h,
                 size=18, bold=True, color="FFFFFF")
        add_text(slide, bk.get("subtitle", ""), 684, ty + title_h + 2, 508, sub_h,
                 size=13, color="94A3B8")
        add_text(slide, bk.get("body",     ""), 684, ty + title_h + sub_h + 6, 508, body_h,
                 size=13, color="64748B", wrap=True)

    # Footer
    add_line(slide, 60, 620, 1220, 620, color="1E3A5F", width=1)
    add_text(slide, "联系方式", 540, 632, 200, 22, size=13, color="64748B", align="center")
    add_text(slide, spec.get("contact", ""), 440, 654, 400, 28,
             size=16, bold=True, color=theme.ACCENT, align="center")

    page_num(slide, spec.get("page", ""))

