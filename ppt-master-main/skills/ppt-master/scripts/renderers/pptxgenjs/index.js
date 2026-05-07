#!/usr/bin/env node
'use strict';
/**
 * PptxGenJS CLI renderer
 * Usage: node index.js <spec.json> <output.pptx>
 */

const fs   = require('fs');
const path = require('path');
const PptxGenJS = require('pptxgenjs');
const { getTheme } = require('./theme');
const { DISPATCH }  = require('./layouts');

async function main() {
  const [,, specPath, outPath] = process.argv;
  if (!specPath || !outPath) {
    console.error('Usage: node index.js <spec.json> <output.pptx>');
    process.exit(1);
  }

  const deck = JSON.parse(fs.readFileSync(specPath, 'utf8'));
  const meta = deck.meta || {};

  // Resolve theme
  const baseTheme = getTheme(meta.theme || 'tech_dark');
  const t = Object.assign({}, baseTheme, meta.colors || {});

  // Create presentation
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';   // 10" × 5.625"
  pres.author  = meta.author  || 'PPT Master';
  pres.company = meta.company || '';
  pres.subject = meta.subject || '';
  pres.title   = meta.title   || (deck.slides[0] && deck.slides[0].title) || '';

  let slideIndex = 0;
  for (const spec of deck.slides) {
    slideIndex++;
    const fn = DISPATCH[spec.layout];
    if (!fn) {
      console.warn(`[SKIP] slide ${slideIndex}: unknown layout "${spec.layout}"`);
      continue;
    }
    const slide = pres.addSlide();
    try {
      await fn(slide, pres, spec, t);
      console.log(`[OK]   ${String(slideIndex).padStart(2, '0')}  ${spec.id || spec.layout}`);
    } catch (err) {
      console.error(`[ERR]  slide ${slideIndex} (${spec.layout}):`, err.message);
    }
  }

  // Ensure output directory exists
  const outDir = path.dirname(path.resolve(outPath));
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  await pres.writeFile({ fileName: outPath });
  console.log(`\nSaved → ${outPath}`);
}

main().catch(err => {
  console.error('Fatal:', err);
  process.exit(1);
});
