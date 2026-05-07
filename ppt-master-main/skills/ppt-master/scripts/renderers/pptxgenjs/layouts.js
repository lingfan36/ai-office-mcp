// PptxGenJS layout renderers
// Rules: no accent lines under titles, no full-width decorative bars,
// every slide has a visual element, strong typography contrast.
'use strict';

const { W, H, M, CW, sh, cardOpts, addDot, addDecorCircles,
        iconPng, pageNum, addTitle, PALETTE } = require('./utils');

// ─────────────────────────────────────────────────────────────────────────────
// 1. COVER
// ─────────────────────────────────────────────────────────────────────────────
async function renderCover(slide, pres, spec, t) {
  slide.background = { color: t.bgDeep };

  // Decorative concentric ring outlines (top-right)
  addDecorCircles(slide, pres, t.accent);

  // Subtle mid-slide glow patch
  slide.addShape(pres.shapes.OVAL, {
    x: 5.5, y: 1.8, w: 5.0, h: 3.0,
    fill: { color: t.accent, transparency: 92 },
    line: { type: 'none' },
  });

  // Tag label (above title — no underline)
  if (spec.tag) {
    slide.addText(spec.tag.toUpperCase(), {
      x: M, y: 0.48, w: 6, h: 0.28,
      fontSize: 10, fontFace: t.fontBody,
      bold: true, color: t.accent2, align: 'left', margin: 0,
      charSpacing: 3,
    });
  }

  // Main title
  const titleLines = spec.title || '';
  slide.addText(titleLines, {
    x: M, y: 0.88, w: 6.8, h: 2.5,
    fontSize: 46, fontFace: t.fontTitle,
    bold: true, color: 'FFFFFF',
    align: 'left', valign: 'top', margin: 0,
    lineSpacingMultiple: 1.08,
  });

  // Subtitle
  let curY = 3.2;
  if (spec.subtitle) {
    slide.addText(spec.subtitle, {
      x: M, y: curY, w: 7.2, h: 0.6,
      fontSize: 17, fontFace: t.fontBody,
      color: t.textSecond, align: 'left', margin: 0,
    });
    curY += 0.52;
  }
  if (spec.subtitle2) {
    slide.addText(spec.subtitle2, {
      x: M, y: curY, w: 7.2, h: 0.4,
      fontSize: 13, fontFace: t.fontBody,
      color: t.textMuted, align: 'left', margin: 0,
    });
    curY += 0.42;
  }

  // Badge pill
  if (spec.badge) {
    curY += 0.18;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: M, y: curY, w: 1.8, h: 0.34,
      fill: { color: t.bgDeep },
      line: { color: t.accent, width: 1.0 },
      rectRadius: 0.17,
    });
    slide.addText(spec.badge, {
      x: M, y: curY, w: 1.8, h: 0.34,
      fontSize: 11, fontFace: t.fontBody,
      bold: true, color: t.accent, align: 'center', margin: 0,
    });
  }

  pageNum(slide, spec.page, t.textMuted);
}

// ─────────────────────────────────────────────────────────────────────────────
// 2. HERO_STAT
// ─────────────────────────────────────────────────────────────────────────────
async function renderHeroStat(slide, pres, spec, t) {
  slide.background = { color: t.bg };

  // Title (no underline bar)
  if (spec.title) addTitle(slide, spec.title, t, M, 0.35, CW, 28);

  // Giant number — left area
  const heroN = spec.hero_number || '';
  slide.addText(heroN, {
    x: M, y: 0.88, w: 5.2, h: 2.6,
    fontSize: 96, fontFace: t.fontNum,
    bold: true, color: t.accent,
    align: 'left', valign: 'middle', margin: 0,
    lineSpacingMultiple: 0.9,
  });

  // Unit (beside or below number)
  const heroU = spec.hero_unit || '';
  const heroL = spec.hero_label || '';
  if (heroU) {
    slide.addText(heroU, {
      x: M, y: 3.52, w: 1.2, h: 0.44,
      fontSize: 26, fontFace: t.fontNum,
      bold: true, color: t.textPrimary, align: 'left', margin: 0,
    });
  }
  if (heroL) {
    slide.addText(heroL, {
      x: heroU ? M + 1.3 : M, y: 3.58, w: heroU ? 4.5 : 5.8, h: 0.36,
      fontSize: 13, fontFace: t.fontBody,
      color: t.textSecond, align: 'left', margin: 0,
    });
  }

  // Three column cards (bottom strip)
  const cols = (spec.columns || []).slice(0, 3);
  const colColors = PALETTE(t);
  const cardW = 2.85, cardH = 1.38, gap = 0.225, startY = 4.0;
  cols.forEach((col, i) => {
    const cx = M + i * (cardW + gap);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: startY, w: cardW, h: cardH,
      fill: { color: t.cardBg },
      line: { color: t.cardBorder, width: 0.5 },
      shadow: sh(6, 2, 0.10),
    });
    // Color accent strip (top of card, inside — not full-slide width)
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: startY, w: cardW, h: 0.06,
      fill: { color: colColors[i] },
      line: { type: 'none' },
    });
    slide.addText(col.title || '', {
      x: cx + 0.18, y: startY + 0.14, w: cardW - 0.36, h: 0.34,
      fontSize: 14, fontFace: t.fontTitle,
      bold: true, color: t.textPrimary, margin: 0,
    });
    slide.addText(col.body || '', {
      x: cx + 0.18, y: startY + 0.5, w: cardW - 0.36, h: 0.78,
      fontSize: 12, fontFace: t.fontBody,
      color: t.textSecond, margin: 0,
      align: 'left',
    });
  });

  pageNum(slide, spec.page, t.textMuted);
}

// ─────────────────────────────────────────────────────────────────────────────
// 3. SOLUTION (dark-left / light-right split)
// ─────────────────────────────────────────────────────────────────────────────
async function renderSolution(slide, pres, spec, t) {
  slide.background = { color: t.bg };

  // Left dark panel
  const panelW = 4.4;
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: panelW, h: H,
    fill: { color: t.bgDeep },
    line: { type: 'none' },
  });

  // Tag
  if (spec.tag) {
    slide.addText(spec.tag.toUpperCase(), {
      x: M, y: 0.45, w: panelW - M, h: 0.26,
      fontSize: 10, fontFace: t.fontBody,
      bold: true, color: t.accent2, align: 'left', margin: 0, charSpacing: 2,
    });
  }

  // Hero text (left panel)
  if (spec.hero) {
    slide.addText(spec.hero, {
      x: M, y: 0.82, w: panelW - M - 0.2, h: 1.6,
      fontSize: 24, fontFace: t.fontTitle,
      bold: true, color: 'FFFFFF',
      align: 'left', valign: 'top', margin: 0,
      lineSpacingMultiple: 1.2,
    });
  }

  // Bullets (left panel) — colored dot instead of border bar
  const bullets = (spec.bullets || []).slice(0, 3);
  const dotColors = [t.accent, t.accent2, 'FFFFFF'];
  let bY = 2.55;
  bullets.forEach((b, i) => {
    // Colored dot bullet
    slide.addShape(pres.shapes.OVAL, {
      x: M, y: bY + 0.07, w: 0.12, h: 0.12,
      fill: { color: dotColors[i] },
      line: { type: 'none' },
    });
    slide.addText(b.title || '', {
      x: M + 0.22, y: bY, w: panelW - M - 0.3, h: 0.28,
      fontSize: 13, fontFace: t.fontTitle,
      bold: true, color: 'FFFFFF', align: 'left', margin: 0,
    });
    slide.addText(b.body || '', {
      x: M + 0.22, y: bY + 0.28, w: panelW - M - 0.3, h: 0.34,
      fontSize: 11, fontFace: t.fontBody,
      color: t.textSecond, align: 'left', margin: 0,
    });
    bY += 0.76;
  });

  // Right cards
  const cards = (spec.cards || []).slice(0, 3);
  const cardColors = PALETTE(t);
  const cX = panelW + 0.3, cW = W - panelW - 0.6;
  cards.forEach((card, i) => {
    const cY = 0.35 + i * 1.68;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cX, y: cY, w: cW, h: 1.55,
      fill: { color: t.cardBg },
      line: { color: t.cardBorder, width: 0.5 },
      shadow: sh(5, 2, 0.09),
    });

    // Icon circle
    const ic = cardColors[i];
    slide.addShape(pres.shapes.OVAL, {
      x: cX + 0.22, y: cY + 0.44, w: 0.58, h: 0.58,
      fill: { color: ic, transparency: 82 },
      line: { type: 'none' },
    });
    slide.addText(card.icon_char || '◆', {
      x: cX + 0.22, y: cY + 0.44, w: 0.58, h: 0.58,
      fontSize: 18, fontFace: t.fontBody,
      color: ic, align: 'center', valign: 'middle', margin: 0,
    });

    slide.addText(card.title || '', {
      x: cX + 0.96, y: cY + 0.2, w: cW - 1.1, h: 0.36,
      fontSize: 16, fontFace: t.fontTitle,
      bold: true, color: t.textPrimary, align: 'left', margin: 0,
    });
    slide.addText(card.body || '', {
      x: cX + 0.96, y: cY + 0.6, w: cW - 1.1, h: 0.7,
      fontSize: 12, fontFace: t.fontBody,
      color: t.textSecond, align: 'left', margin: 0,
    });
  });

  pageNum(slide, spec.page, t.textMuted);
}

// ─────────────────────────────────────────────────────────────────────────────
// 4. PIPELINE (horizontal N-stage flow)
// ─────────────────────────────────────────────────────────────────────────────
async function renderPipeline(slide, pres, spec, t) {
  slide.background = { color: t.bg };
  if (spec.title) addTitle(slide, spec.title, t, M, 0.3, CW, 28);

  const stages = (spec.stages || []).slice(0, 5);
  const n = stages.length;
  if (n === 0) return;

  const stageW = (CW - (n - 1) * 0.18) / n;
  const stageH = 4.5, stageY = 0.88;
  const colPalette = PALETTE(t);

  stages.forEach((stage, i) => {
    const sx = M + i * (stageW + 0.18);
    const c = stage.color || colPalette[i % colPalette.length];

    // Card
    slide.addShape(pres.shapes.RECTANGLE, {
      x: sx, y: stageY, w: stageW, h: stageH,
      fill: { color: stage.optional ? t.bg : t.cardBg },
      line: { color: stage.optional ? c : t.cardBorder, width: stage.optional ? 1.0 : 0.5 },
      shadow: sh(5, 2, 0.09),
    });

    // Step number circle
    slide.addShape(pres.shapes.OVAL, {
      x: sx + stageW / 2 - 0.28, y: stageY + 0.15,
      w: 0.56, h: 0.56,
      fill: { color: c },
      line: { type: 'none' },
    });
    slide.addText(String(i + 1), {
      x: sx + stageW / 2 - 0.28, y: stageY + 0.15, w: 0.56, h: 0.56,
      fontSize: 18, fontFace: t.fontNum,
      bold: true, color: 'FFFFFF', align: 'center', valign: 'middle', margin: 0,
    });

    // Stage title
    slide.addText(stage.title || '', {
      x: sx + 0.1, y: stageY + 0.82, w: stageW - 0.2, h: 0.6,
      fontSize: 13, fontFace: t.fontTitle,
      bold: true, color: t.textPrimary, align: 'center', margin: 0,
    });

    // Bullets
    let bY = stageY + 1.52;
    (stage.bullets || []).slice(0, 5).forEach(b => {
      slide.addText('· ' + b, {
        x: sx + 0.12, y: bY, w: stageW - 0.24, h: 0.38,
        fontSize: 11, fontFace: t.fontBody,
        color: t.textSecond, align: 'left', margin: 0,
      });
      bY += 0.38;
    });

    // Optional badge
    if (stage.optional) {
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: sx + stageW / 2 - 0.5, y: stageY + stageH - 0.46,
        w: 1.0, h: 0.3,
        fill: { color: c },
        line: { type: 'none' },
        rectRadius: 0.15,
      });
      slide.addText('可选', {
        x: sx + stageW / 2 - 0.5, y: stageY + stageH - 0.46, w: 1.0, h: 0.3,
        fontSize: 10, fontFace: t.fontBody,
        bold: true, color: 'FFFFFF', align: 'center', margin: 0,
      });
    }

    // Arrow
    if (i < n - 1) {
      slide.addText('▶', {
        x: sx + stageW + 0.02, y: stageY + stageH / 2 - 0.15, w: 0.16, h: 0.3,
        fontSize: 10, color: t.textMuted, align: 'center', margin: 0,
      });
    }
  });

  pageNum(slide, spec.page, t.textMuted);
}

// ─────────────────────────────────────────────────────────────────────────────
// 5. HUB_SPOKE
// ─────────────────────────────────────────────────────────────────────────────
async function renderHubSpoke(slide, pres, spec, t) {
  slide.background = { color: t.bg };
  if (spec.title) addTitle(slide, spec.title, t, M, 0.3, CW, 28);

  const hub = spec.hub || {};
  const hx = W / 2, hy = H / 2 + 0.2;

  // Outer ring
  slide.addShape(pres.shapes.OVAL, {
    x: hx - 1.4, y: hy - 1.4, w: 2.8, h: 2.8,
    fill: { type: 'none' },
    line: { color: t.accent, width: 0.8, transparency: 60 },
  });
  // Hub circle
  slide.addShape(pres.shapes.OVAL, {
    x: hx - 0.6, y: hy - 0.6, w: 1.2, h: 1.2,
    fill: { color: t.accent },
    line: { type: 'none' },
  });
  slide.addText(hub.title || 'AI\n核心', {
    x: hx - 0.6, y: hy - 0.6, w: 1.2, h: 1.2,
    fontSize: 13, fontFace: t.fontTitle,
    bold: true, color: 'FFFFFF', align: 'center', valign: 'middle', margin: 0,
  });

  const positions = [
    [hx,       hy - 2.2], [hx + 2.5, hy - 1.2], [hx + 2.5, hy + 0.8],
    [hx,       hy + 2.1], [hx - 2.5, hy + 0.8], [hx - 2.5, hy - 1.2],
  ];
  const colPalette = PALETTE(t);
  const spokes = (spec.spokes || []).slice(0, 6);
  const cardW = 2.0, cardH = 1.0;

  spokes.forEach((spoke, i) => {
    const [px, py] = positions[i];
    const c = spoke.color || colPalette[i % colPalette.length];

    // Spoke line
    slide.addShape(pres.shapes.LINE, {
      x: hx, y: hy, w: px - hx, h: py - hy,
      line: { color: c, width: 1, transparency: 40 },
    });

    // Spoke card
    const cx = Math.max(M, Math.min(W - M - cardW, px - cardW / 2));
    const cy = Math.max(0.4, Math.min(H - 0.2 - cardH, py - cardH / 2));

    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cardW, h: cardH,
      fill: { color: t.cardBg },
      line: { color: c, width: 1.0 },
      shadow: sh(4, 2, 0.09),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cardW, h: 0.05,
      fill: { color: c },
      line: { type: 'none' },
    });
    slide.addText(spoke.title || '', {
      x: cx + 0.1, y: cy + 0.1, w: cardW - 0.2, h: 0.3,
      fontSize: 12, fontFace: t.fontTitle,
      bold: true, color: t.textPrimary, margin: 0,
    });
    slide.addText(spoke.body || '', {
      x: cx + 0.1, y: cy + 0.44, w: cardW - 0.2, h: 0.5,
      fontSize: 10, fontFace: t.fontBody,
      color: t.textSecond, margin: 0,
    });
  });

  pageNum(slide, spec.page, t.textMuted);
}

// ─────────────────────────────────────────────────────────────────────────────
// 6. KPI_GRID (2 × 2 metric cards)
// ─────────────────────────────────────────────────────────────────────────────
async function renderKpiGrid(slide, pres, spec, t) {
  slide.background = { color: t.bg };
  if (spec.title) addTitle(slide, spec.title, t, M, 0.28, CW, 28);

  const cards = (spec.cards || []).slice(0, 4);
  const colPalette = PALETTE(t);
  const cW = 4.3, cH = 2.22, gX = 0.1, gY = 0.14;
  const positions = [
    [M, 0.82],          [M + cW + gX, 0.82],
    [M, 0.82 + cH + gY], [M + cW + gX, 0.82 + cH + gY],
  ];

  cards.forEach((card, i) => {
    const [cx, cy] = positions[i];
    const c = card.color || colPalette[i];
    const isDark = card.dark;

    // Card background
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: cW, h: cH,
      fill: { color: isDark ? t.bgDeep : t.cardBg },
      line: { color: isDark ? t.cardBorder : t.cardBorder, width: 0.5 },
      shadow: sh(8, 3, 0.11),
    });

    // Color accent (small left-side strip inside card)
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: cy, w: 0.06, h: cH,
      fill: { color: c },
      line: { type: 'none' },
    });

    // Value
    slide.addText(card.value || '', {
      x: cx + 0.22, y: cy + 0.14, w: cW - 0.36, h: 1.05,
      fontSize: 64, fontFace: t.fontNum,
      bold: true, color: isDark ? 'FFFFFF' : c,
      align: 'left', valign: 'middle', margin: 0,
    });

    // Unit
    if (card.unit) {
      slide.addText(card.unit, {
        x: cx + 0.22, y: cy + 1.24, w: cW - 0.36, h: 0.36,
        fontSize: 17, fontFace: t.fontTitle,
        bold: true, color: isDark ? t.textPrimary : t.textPrimary,
        align: 'left', margin: 0,
      });
    }

    // Label
    slide.addText(card.label || '', {
      x: cx + 0.22, y: cy + 1.62, w: cW - 0.36, h: 0.3,
      fontSize: 11, fontFace: t.fontBody,
      color: isDark ? t.textSecond : t.textSecond,
      align: 'left', margin: 0,
    });

    // Note
    if (card.note) {
      slide.addText(card.note, {
        x: cx + 0.22, y: cy + 1.88, w: cW - 0.36, h: 0.26,
        fontSize: 10, fontFace: t.fontBody,
        color: t.textMuted, align: 'left', margin: 0,
      });
    }
  });

  pageNum(slide, spec.page, t.textMuted);
}

// ─────────────────────────────────────────────────────────────────────────────
// 7. RADAR (matplotlib chart + score table — image placeholder)
// ─────────────────────────────────────────────────────────────────────────────
async function renderRadar(slide, pres, spec, t) {
  slide.background = { color: t.bg };
  if (spec.title) addTitle(slide, spec.title, t, M, 0.28, CW, 28);

  // Chart area placeholder (python radar is passed as image path if present)
  if (spec._radar_image) {
    slide.addImage({ path: spec._radar_image, x: M, y: 0.82, w: 5.5, h: 4.6 });
  } else {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: M, y: 0.82, w: 5.5, h: 4.6,
      fill: { color: t.cardBg },
      line: { color: t.cardBorder, width: 0.5 },
    });
    slide.addText('Radar Chart', {
      x: M, y: 2.8, w: 5.5, h: 0.5,
      fontSize: 14, color: t.textMuted, align: 'center', margin: 0,
    });
  }

  // Legend
  slide.addShape(pres.shapes.OVAL, {
    x: 6.2, y: 0.88, w: 0.14, h: 0.14,
    fill: { color: t.accent }, line: { type: 'none' },
  });
  slide.addText(spec.legend_us || '我们', {
    x: 6.4, y: 0.82, w: 1.2, h: 0.26,
    fontSize: 12, fontFace: t.fontBody,
    bold: true, color: t.accent, margin: 0,
  });
  slide.addShape(pres.shapes.OVAL, {
    x: 7.8, y: 0.88, w: 0.14, h: 0.14,
    fill: { color: t.textMuted }, line: { type: 'none' },
  });
  slide.addText(spec.legend_avg || '行业均值', {
    x: 8.0, y: 0.82, w: 1.6, h: 0.26,
    fontSize: 12, fontFace: t.fontBody,
    color: t.textSecond, margin: 0,
  });

  // Score rows
  const dims = (spec.dimensions || []).slice(0, 8);
  const rowH = 0.54;
  dims.forEach((d, i) => {
    const ry = 1.22 + i * rowH;
    if (i % 2 === 0) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: 6.0, y: ry, w: 3.7, h: rowH,
        fill: { color: t.cardBg }, line: { type: 'none' },
      });
    }
    slide.addText(d.name || '', {
      x: 6.12, y: ry + 0.06, w: 1.8, h: 0.24,
      fontSize: 12, fontFace: t.fontBody, color: t.textPrimary, margin: 0,
    });
    const barX = 8.0, barW = 1.5;
    const us = d.us || 9, avg = d.avg || 5;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: barX, y: ry + 0.2, w: barW * avg / 10, h: 0.1,
      fill: { color: t.textMuted }, line: { type: 'none' },
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: barX, y: ry + 0.34, w: barW * us / 10, h: 0.1,
      fill: { color: t.accent }, line: { type: 'none' },
    });
    slide.addText(String(us), {
      x: 9.55, y: ry + 0.12, w: 0.3, h: 0.3,
      fontSize: 12, fontFace: t.fontNum,
      bold: true, color: t.accent, align: 'right', margin: 0,
    });
  });

  pageNum(slide, spec.page, t.textMuted);
}

// ─────────────────────────────────────────────────────────────────────────────
// 8. COMPARISON (feature table)
// ─────────────────────────────────────────────────────────────────────────────
async function renderComparison(slide, pres, spec, t) {
  slide.background = { color: t.bg };
  if (spec.title) addTitle(slide, spec.title, t, M, 0.28, CW, 28);

  const comps    = (spec.competitors || []).slice(0, 5);
  const features = (spec.features    || []).slice(0, 8);
  if (!comps.length) return;

  const labelW = 1.8, usW = 2.2, othW = 1.6, gap = 0.06;
  const tx = M, ty = 0.92, rh = 0.58, hh = 0.6;

  const colXs = [tx + labelW + gap];
  const colWs = [usW];
  for (let i = 1; i < comps.length; i++) {
    colXs.push(colXs[i - 1] + colWs[i - 1] + gap);
    colWs.push(othW);
  }

  // Header
  slide.addShape(pres.shapes.RECTANGLE, {
    x: tx, y: ty, w: labelW, h: hh,
    fill: { color: t.cardBg }, line: { color: t.cardBorder, width: 0.5 },
  });
  comps.forEach((comp, i) => {
    const bg = i === 0 ? t.bgDeep : t.cardBg;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: colXs[i], y: ty, w: colWs[i], h: hh,
      fill: { color: bg }, line: { color: t.cardBorder, width: 0.5 },
    });
    slide.addText(comp.name || '', {
      x: colXs[i] + 0.06, y: ty + 0.06, w: colWs[i] - 0.12, h: 0.3,
      fontSize: 13, fontFace: t.fontTitle, bold: true,
      color: i === 0 ? 'FFFFFF' : t.textPrimary, align: 'center', margin: 0,
    });
    if (comp.note) {
      slide.addText(comp.note, {
        x: colXs[i] + 0.06, y: ty + 0.36, w: colWs[i] - 0.12, h: 0.2,
        fontSize: 9, fontFace: t.fontBody,
        color: i === 0 ? t.accent2 : t.textMuted, align: 'center', margin: 0,
      });
    }
  });

  // Feature rows
  features.forEach((feat, r) => {
    const ry = ty + hh + r * rh;
    const rowBg = r % 2 === 0 ? t.cardBg : t.bg;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: tx, y: ry, w: labelW, h: rh,
      fill: { color: rowBg }, line: { color: t.cardBorder, width: 0.3 },
    });
    slide.addText(feat.name || '', {
      x: tx + 0.12, y: ry + 0.12, w: labelW - 0.24, h: 0.36,
      fontSize: 12, fontFace: t.fontBody, color: t.textPrimary, margin: 0,
    });

    const vals = feat.values || [];
    comps.forEach((_, i) => {
      const val = vals[i];
      const cellBg = i === 0 ? (r % 2 === 0 ? 'EFF8FF' : 'F8FBFF') : rowBg;
      slide.addShape(pres.shapes.RECTANGLE, {
        x: colXs[i], y: ry, w: colWs[i], h: rh,
        fill: { color: cellBg }, line: { color: t.cardBorder, width: 0.3 },
      });
      if (val === true) {
        slide.addText('✓', {
          x: colXs[i], y: ry + 0.06, w: colWs[i], h: rh - 0.12,
          fontSize: 18, fontFace: t.fontBody, bold: true,
          color: t.green, align: 'center', valign: 'middle', margin: 0,
        });
      } else if (val === false) {
        slide.addText('✗', {
          x: colXs[i], y: ry + 0.06, w: colWs[i], h: rh - 0.12,
          fontSize: 18, fontFace: t.fontBody, bold: true,
          color: t.red, align: 'center', valign: 'middle', margin: 0,
        });
      } else if (val !== undefined) {
        slide.addText(String(val), {
          x: colXs[i] + 0.06, y: ry + 0.1, w: colWs[i] - 0.12, h: rh - 0.2,
          fontSize: 11, fontFace: t.fontBody,
          color: i === 0 ? t.textPrimary : t.textSecond, align: 'center', margin: 0,
        });
      }
    });
  });

  pageNum(slide, spec.page, t.textMuted);
}

// ─────────────────────────────────────────────────────────────────────────────
// 9. BUSINESS (vertical timeline with price cards)
// ─────────────────────────────────────────────────────────────────────────────
async function renderBusiness(slide, pres, spec, t) {
  slide.background = { color: t.bg };
  if (spec.title) addTitle(slide, spec.title, t, M, 0.28, CW, 28);
  if (spec.subtitle) {
    slide.addText(spec.subtitle, {
      x: M, y: 0.78, w: CW, h: 0.28,
      fontSize: 13, fontFace: t.fontBody, color: t.textSecond, margin: 0,
    });
  }

  const items = (spec.items || []).slice(0, 3);
  const colPalette = PALETTE(t);
  const trackX = 1.35;

  // Track line
  slide.addShape(pres.shapes.LINE, {
    x: trackX, y: 1.18, w: 0, h: 3.9,
    line: { color: t.accent, width: 1.5 },
  });

  items.forEach((item, i) => {
    const c = item.color || colPalette[i];
    const iy = 1.55 + i * 1.44;

    // Node circle
    slide.addShape(pres.shapes.OVAL, {
      x: trackX - 0.24, y: iy - 0.24, w: 0.48, h: 0.48,
      fill: { color: c, transparency: 78 }, line: { type: 'none' },
    });
    slide.addShape(pres.shapes.OVAL, {
      x: trackX - 0.16, y: iy - 0.16, w: 0.32, h: 0.32,
      fill: { color: c }, line: { type: 'none' },
    });
    slide.addText(`0${i + 1}`, {
      x: trackX - 0.16, y: iy - 0.16, w: 0.32, h: 0.32,
      fontSize: 10, fontFace: t.fontNum, bold: true,
      color: 'FFFFFF', align: 'center', valign: 'middle', margin: 0,
    });

    // Card
    const cx = trackX + 0.36, cW = 5.8;
    const titleH = 0.36, bodyH = 0.3;
    const ch = titleH + bodyH + 0.28;
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: iy - ch / 2, w: cW, h: ch,
      fill: { color: t.cardBg },
      line: { color: t.cardBorder, width: 0.5 },
      shadow: sh(5, 2, 0.09),
    });
    slide.addShape(pres.shapes.RECTANGLE, {
      x: cx, y: iy - ch / 2, w: 0.05, h: ch,
      fill: { color: c }, line: { type: 'none' },
    });
    slide.addText(item.title || '', {
      x: cx + 0.18, y: iy - ch / 2 + 0.08, w: cW - 0.28, h: titleH,
      fontSize: 15, fontFace: t.fontTitle,
      bold: true, color: t.textPrimary, margin: 0,
    });
    slide.addText(item.body || '', {
      x: cx + 0.18, y: iy - ch / 2 + titleH + 0.08, w: cW - 0.28, h: bodyH,
      fontSize: 11, fontFace: t.fontBody, color: t.textSecond, margin: 0,
    });

    // Price tag (right)
    if (item.price) {
      const pW = 1.55;
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
        x: W - M - pW, y: iy - 0.42, w: pW, h: 0.84,
        fill: { color: c }, line: { type: 'none' },
        rectRadius: 0.1, shadow: sh(5, 2, 0.18),
      });
      slide.addText(item.price, {
        x: W - M - pW, y: iy - 0.35, w: pW, h: 0.38,
        fontSize: 18, fontFace: t.fontNum, bold: true,
        color: 'FFFFFF', align: 'center', margin: 0,
      });
      if (item.price_label) {
        slide.addText(item.price_label, {
          x: W - M - pW, y: iy + 0.04, w: pW, h: 0.24,
          fontSize: 10, fontFace: t.fontBody,
          color: 'FFFFFF', align: 'center', margin: 0,
        });
      }
    }
  });

  pageNum(slide, spec.page, t.textMuted);
}

// ─────────────────────────────────────────────────────────────────────────────
// 10. TIMELINE (horizontal milestone track)
// ─────────────────────────────────────────────────────────────────────────────
async function renderTimeline(slide, pres, spec, t) {
  slide.background = { color: t.bg };
  if (spec.title) addTitle(slide, spec.title, t, M, 0.28, CW, 28);

  const mss = (spec.milestones || []).slice(0, 5);
  const n = mss.length;
  if (n === 0) return;

  const trackY = 3.0, tx1 = M + 0.3, tx2 = W - M - 0.3;
  const step = (tx2 - tx1) / Math.max(n - 1, 1);

  // Track
  slide.addShape(pres.shapes.LINE, {
    x: tx1, y: trackY, w: tx2 - tx1, h: 0,
    line: { color: t.cardBorder, width: 2 },
  });

  // Progress
  const curIdx = mss.findIndex(m => m.state === 'current');
  if (curIdx >= 0) {
    slide.addShape(pres.shapes.LINE, {
      x: tx1, y: trackY, w: step * curIdx + 0.18, h: 0,
      line: { color: t.accent, width: 2 },
    });
  }

  const cW = 1.7, cH = 1.1;

  mss.forEach((ms, i) => {
    const nx = tx1 + i * step;
    const state = ms.state || 'planned';
    const above = i % 2 === 0;
    const cardY = above ? trackY - cH - 0.32 : trackY + 0.38;

    // Connector
    slide.addShape(pres.shapes.LINE, {
      x: nx, y: above ? cardY + cH : trackY,
      w: 0, h: above ? trackY - cardY - cH : cardY - trackY,
      line: { color: state === 'current' ? t.accent : t.cardBorder, width: 1 },
    });

    // Card
    const bc = state === 'current' ? t.accent : (state === 'target' ? t.indigo : t.cardBorder);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: nx - cW / 2, y: cardY, w: cW, h: cH,
      fill: { color: t.cardBg },
      line: { color: bc, width: state !== 'planned' ? 1.2 : 0.5 },
      shadow: sh(4, 2, 0.08),
    });
    if (state === 'current') {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: nx - cW / 2, y: cardY, w: cW, h: 0.05,
        fill: { color: t.accent }, line: { type: 'none' },
      });
    }
    slide.addText(ms.date || '', {
      x: nx - cW / 2 + 0.1, y: cardY + 0.1, w: cW - 0.2, h: 0.24,
      fontSize: 11, fontFace: t.fontBody, bold: true,
      color: state === 'current' ? t.accent : t.textSecond, align: 'center', margin: 0,
    });
    slide.addText(ms.title || '', {
      x: nx - cW / 2 + 0.08, y: cardY + 0.36, w: cW - 0.16, h: 0.42,
      fontSize: 12, fontFace: t.fontTitle, bold: true,
      color: t.textPrimary, align: 'center', margin: 0,
    });
    slide.addText(ms.body || '', {
      x: nx - cW / 2 + 0.08, y: cardY + 0.78, w: cW - 0.16, h: 0.3,
      fontSize: 10, fontFace: t.fontBody,
      color: t.textSecond, align: 'center', margin: 0,
    });

    // Node
    const nr = state === 'current' ? 0.22 : 0.16;
    if (state === 'current') {
      slide.addShape(pres.shapes.OVAL, {
        x: nx - 0.24, y: trackY - 0.24, w: 0.48, h: 0.48,
        fill: { color: t.accent, transparency: 80 }, line: { type: 'none' },
      });
    }
    slide.addShape(pres.shapes.OVAL, {
      x: nx - nr, y: trackY - nr, w: nr * 2, h: nr * 2,
      fill: { color: t.bg },
      line: { color: state === 'current' ? t.accent : (state === 'target' ? t.indigo : t.textMuted), width: 2 },
    });
    if (state === 'current') {
      slide.addShape(pres.shapes.OVAL, {
        x: nx - 0.08, y: trackY - 0.08, w: 0.16, h: 0.16,
        fill: { color: t.accent }, line: { type: 'none' },
      });
    }
  });

  if (spec.footer) {
    slide.addText(spec.footer, {
      x: M, y: H - 0.42, w: CW, h: 0.28,
      fontSize: 11, fontFace: t.fontBody,
      color: t.textMuted, align: 'center', margin: 0,
    });
  }

  pageNum(slide, spec.page, t.textMuted);
}

// ─────────────────────────────────────────────────────────────────────────────
// 11. FUNDRAISING (dark premium)
// ─────────────────────────────────────────────────────────────────────────────
async function renderFundraising(slide, pres, spec, t) {
  slide.background = { color: t.bgDeep };

  // Soft accent glow (top-right)
  slide.addShape(pres.shapes.OVAL, {
    x: 8.0, y: -0.8, w: 3.5, h: 3.5,
    fill: { color: t.accent, transparency: 90 }, line: { type: 'none' },
  });

  // Divider line
  slide.addShape(pres.shapes.LINE, {
    x: 4.5, y: 0.55, w: 0, h: H - 1.0,
    line: { color: t.cardBorder, width: 0.5 },
  });

  // Left panel
  if (spec.tag) {
    slide.addText(spec.tag.toUpperCase(), {
      x: M, y: 0.52, w: 3.8, h: 0.26,
      fontSize: 10, fontFace: t.fontBody, bold: true,
      color: t.accent2, align: 'left', margin: 0, charSpacing: 2,
    });
  }

  if (spec.title) {
    slide.addText(spec.title, {
      x: M, y: 0.9, w: 3.7, h: 0.9,
      fontSize: 22, fontFace: t.fontTitle, bold: true,
      color: 'FFFFFF', align: 'left', margin: 0, lineSpacingMultiple: 1.2,
    });
  }

  // Amount
  slide.addText(spec.amount || '', {
    x: M, y: 1.95, w: 3.5, h: 1.3,
    fontSize: 88, fontFace: t.fontNum, bold: true,
    color: t.accent, align: 'left', valign: 'middle', margin: 0, lineSpacingMultiple: 0.9,
  });
  slide.addText(spec.amount_unit || '万', {
    x: M, y: 3.28, w: 1.2, h: 0.48,
    fontSize: 28, fontFace: t.fontNum, bold: true, color: 'FFFFFF', margin: 0,
  });

  // Round badge
  if (spec.round_label) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: M, y: 3.9, w: 1.8, h: 0.34,
      fill: { color: t.bgDeep },
      line: { color: t.accent, width: 1.0 },
      rectRadius: 0.17,
    });
    slide.addText(spec.round_label, {
      x: M, y: 3.9, w: 1.8, h: 0.34,
      fontSize: 12, fontFace: t.fontBody, bold: true,
      color: t.accent, align: 'center', margin: 0,
    });
  }

  if (spec.target_time) {
    slide.addText(spec.target_time, {
      x: M, y: 4.35, w: 3.7, h: 0.3,
      fontSize: 13, fontFace: t.fontBody, color: t.textSecond, margin: 0,
    });
  }
  if (spec.target_arr) {
    slide.addText(spec.target_arr, {
      x: M, y: 4.65, w: 3.7, h: 0.28,
      fontSize: 12, fontFace: t.fontBody, color: t.textMuted, margin: 0,
    });
  }

  // Right panel: breakdown cards
  const bkColors = PALETTE(t);
  const bks = (spec.breakdowns || []).slice(0, 3);
  const cW = 5.1, startX = 4.75;
  let curY = 0.52;
  const availH = H - 1.22;
  const ch = (availH - (bks.length - 1) * 0.12) / Math.max(bks.length, 1);

  bks.forEach((bk, i) => {
    const c = bk.color || bkColors[i];
    const by = curY;

    slide.addShape(pres.shapes.RECTANGLE, {
      x: startX, y: by, w: cW, h: ch,
      fill: { color: t.cardBg, transparency: 88 },
      line: { color: t.cardBorder, width: 0.5 },
      shadow: sh(6, 2, 0.18),
    });

    // Left accent strip
    slide.addShape(pres.shapes.RECTANGLE, {
      x: startX, y: by, w: 0.05, h: ch,
      fill: { color: c }, line: { type: 'none' },
    });

    // Pct
    slide.addText(bk.pct || '', {
      x: startX + 0.15, y: by + (ch - 0.9) / 2, w: 0.9, h: 0.9,
      fontSize: 36, fontFace: t.fontNum, bold: true,
      color: c, align: 'center', valign: 'middle', margin: 0,
    });

    // Divider
    slide.addShape(pres.shapes.LINE, {
      x: startX + 1.15, y: by + ch * 0.15, w: 0, h: ch * 0.7,
      line: { color: t.cardBorder, width: 0.5 },
    });

    const tx = startX + 1.3, tW = cW - 1.44;
    const titleH = Math.min(0.36, ch * 0.35), subH = Math.min(0.28, ch * 0.25);
    const bodyH = Math.min(0.28, ch * 0.25);
    const innerY = by + (ch - titleH - subH - bodyH) / 2;

    slide.addText(bk.title || '', {
      x: tx, y: innerY, w: tW, h: titleH,
      fontSize: 15, fontFace: t.fontTitle, bold: true,
      color: 'FFFFFF', margin: 0,
    });
    slide.addText(bk.subtitle || '', {
      x: tx, y: innerY + titleH, w: tW, h: subH,
      fontSize: 11, fontFace: t.fontBody, color: t.textSecond, margin: 0,
    });
    slide.addText(bk.body || '', {
      x: tx, y: innerY + titleH + subH, w: tW, h: bodyH,
      fontSize: 11, fontFace: t.fontBody, color: t.textMuted, margin: 0,
    });

    curY += ch + 0.12;
  });

  // Footer
  slide.addShape(pres.shapes.LINE, {
    x: M, y: H - 0.5, w: CW, h: 0,
    line: { color: t.cardBorder, width: 0.5 },
  });
  slide.addText(spec.contact || '', {
    x: M, y: H - 0.44, w: CW, h: 0.3,
    fontSize: 13, fontFace: t.fontBody, bold: true,
    color: t.accent, align: 'center', margin: 0,
  });
}

// ── Dispatch map ──────────────────────────────────────────────────────────────
const DISPATCH = {
  cover:       renderCover,
  hero_stat:   renderHeroStat,
  solution:    renderSolution,
  pipeline:    renderPipeline,
  hub_spoke:   renderHubSpoke,
  kpi_grid:    renderKpiGrid,
  radar:       renderRadar,
  comparison:  renderComparison,
  business:    renderBusiness,
  timeline:    renderTimeline,
  fundraising: renderFundraising,
};

module.exports = { DISPATCH };
