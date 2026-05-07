// Shared helpers for PptxGenJS layouts
// Coordinates are in inches; slide = 10" × 5.625" (16:9)

const W = 10;      // slide width
const H = 5.625;   // slide height
const M = 0.5;     // standard margin
const CW = W - 2 * M;  // content width = 9"

// ── Shadow factory (never reuse objects — PptxGenJS mutates in-place) ─────────
function sh(blur = 8, offset = 3, opacity = 0.14) {
  return { type: 'outer', blur, offset, angle: 135, color: '000000', opacity };
}

// ── Card helper ──────────────────────────────────────────────────────────────
// Returns options for addShape(RECTANGLE, ...) with shadow + optional accent line
function cardOpts(t, fill, border, radius = 0.1) {
  return {
    fill: { color: fill },
    line: border ? { color: border, width: 0.5 } : { type: 'none' },
    rectRadius: radius,
    shadow: sh(),
  };
}

// ── Color accent circle (dot) ────────────────────────────────────────────────
function addDot(slide, pres, cx, cy, r, color) {
  slide.addShape(pres.shapes.OVAL, {
    x: cx - r, y: cy - r, w: r * 2, h: r * 2,
    fill: { color },
    line: { type: 'none' },
  });
}

// ── Subtle background panel ──────────────────────────────────────────────────
function addPanel(slide, pres, x, y, w, h, color, alpha = 100) {
  const opts = {
    x, y, w, h,
    fill: { color, transparency: 100 - alpha },
    line: { type: 'none' },
  };
  slide.addShape(pres.shapes.RECTANGLE, opts);
}

// ── Abstract decorative circles (top-right corner of cover) ──────────────────
function addDecorCircles(slide, pres, accent) {
  const circles = [
    { x: 7.8, y: -0.9, w: 3.2, h: 3.2, t: 85 },
    { x: 8.3, y: -0.4, w: 2.0, h: 2.0, t: 75 },
    { x: 8.7, y:  0.1, w: 1.1, h: 1.1, t: 60 },
  ];
  circles.forEach(({ x, y, w, h, t }) => {
    slide.addShape(pres.shapes.OVAL, {
      x, y, w, h,
      fill: { type: 'none' },
      line: { color: accent, width: 1.2, transparency: t },
    });
  });
}

// ── Icon using react-icons → PNG ─────────────────────────────────────────────
async function iconPng(IconComp, color, size = 256) {
  const React = require('react');
  const ReactDOM = require('react-dom/server');
  const sharp = require('sharp');
  const svg = ReactDOM.renderToStaticMarkup(
    React.createElement(IconComp, { color: '#' + color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

// ── Page number ───────────────────────────────────────────────────────────────
function pageNum(slide, text, color) {
  if (!text) return;
  slide.addText(text, {
    x: W - M - 0.6, y: H - 0.38, w: 0.6, h: 0.25,
    fontSize: 10, fontFace: 'Segoe UI', color,
    align: 'right', margin: 0,
  });
}

// ── Title block (no underline — PPT_SKILL.md rule) ────────────────────────────
function addTitle(slide, text, t, x = M, y = 0.35, w = CW, fontSize = 30) {
  slide.addText(text, {
    x, y, w, h: 0.7,
    fontSize, fontFace: t.fontTitle,
    bold: true, color: t.textPrimary,
    align: 'left', margin: 0,
  });
}

// ── Multicolor row of small accent dots for section headers ──────────────────
const PALETTE = (t) => [t.accent, t.green, t.indigo, t.amber, t.red, t.accent2];

module.exports = { W, H, M, CW, sh, cardOpts, addDot, addPanel,
                   addDecorCircles, iconPng, pageNum, addTitle, PALETTE };
