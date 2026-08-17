#!/usr/bin/env node
/**
 * inspect-pptx.mjs — Look inside a .pptx without a PowerPoint renderer.
 *
 * Answers "does this deck actually look right?" from the emitted OOXML rather than by
 * opening it in an office suite. Two outputs, both agent-friendly:
 *
 *   1. A geometry report + lint  — every shape's position, size, font and text, plus
 *      automatic checks for the failure modes that eyeballing is used to catch:
 *      text pushed off the slide, boxes overlapping each other, absurd font sizes.
 *   2. A rendered preview        — each slide rebuilt as HTML from the .pptx's own
 *      coordinates and screenshotted in headless Chromium. Because it is built from
 *      the written package, not from the converter's memory, it independently proves
 *      what was emitted.
 *
 * Usage:
 *   node inspect-pptx.mjs <file.pptx> [options]
 *
 * Options:
 *   --preview <dir>       Render each slide to <dir>/slide-NNN.png
 *   --slides <list>       Only these slides: "1,10,14" or "1-5" (default: all)
 *   --json                Machine-readable report on stdout
 *   --quiet               Lint findings only, no per-shape listing
 *   --overlap <pct>       Overlap area, as % of the smaller box, that counts (default 25)
 *   --min-font <pt>       Font size below which to warn (default 8)
 *   --min-contrast <n>    Text/background contrast below which to warn (default 3)
 *   --fail-contrast <n>   Text/background contrast below which to fail (default 2)
 *   --no-contrast         Skip text/background contrast sampling
 *
 * No office suite is involved: everything here comes from the OOXML and from the same
 * headless Chromium the converter already depends on.
 */

import { readFile, mkdir } from 'fs/promises';
import { existsSync } from 'fs';
import { resolve, join } from 'path';
import { createRequire } from 'module';

let JSZip, chromium;
try {
  const depsRoot = process.env.SLIDES_PPTX_DEPS;
  if (depsRoot) {
    const req = createRequire(join(depsRoot, 'index.js'));
    JSZip = req('jszip');
    ({ chromium } = req('playwright'));
  } else {
    JSZip = (await import('jszip')).default;
    ({ chromium } = await import('playwright'));
  }
} catch (e) {
  console.error(`✗ Missing dependencies: ${e.message}`);
  console.error('  Install with: npm install jszip playwright');
  process.exit(2);
}

/* ── CLI ────────────────────────────────────────────────── */

const argv = process.argv.slice(2);
const opts = {
  preview: null, slides: null, json: false, quiet: false, overlap: 25,
  minFont: 8, minContrast: 3, failContrast: 2, contrast: true,
};
const positional = [];
for (let i = 0; i < argv.length; i++) {
  const a = argv[i];
  if (a === '--preview') opts.preview = argv[++i];
  else if (a === '--slides') opts.slides = argv[++i];
  else if (a === '--json') opts.json = true;
  else if (a === '--quiet') opts.quiet = true;
  else if (a === '--overlap') opts.overlap = Number(argv[++i]);
  else if (a === '--min-font') opts.minFont = Number(argv[++i]);
  else if (a === '--min-contrast') opts.minContrast = Number(argv[++i]);
  else if (a === '--fail-contrast') opts.failContrast = Number(argv[++i]);
  else if (a === '--no-contrast') opts.contrast = false;
  else if (a.startsWith('--')) { console.error(`Unknown option: ${a}`); process.exit(2); }
  else positional.push(a);
}
if (!positional.length) {
  console.error('Usage: node inspect-pptx.mjs <file.pptx> [--preview <dir>] [--slides 1,10] [--json]');
  process.exit(2);
}
const FILE = resolve(positional[0]);
if (!existsSync(FILE)) { console.error(`✗ File not found: ${FILE}`); process.exit(2); }

/** "1,10,14" / "1-5" / "1-3,8" → Set of 1-based indices, or null for all. */
function parseSlideList(spec) {
  if (!spec) return null;
  const out = new Set();
  for (const part of spec.split(',')) {
    const range = part.trim().match(/^(\d+)\s*-\s*(\d+)$/);
    if (range) for (let n = Number(range[1]); n <= Number(range[2]); n++) out.add(n);
    else if (/^\d+$/.test(part.trim())) out.add(Number(part.trim()));
  }
  return out.size ? out : null;
}
const WANTED = parseSlideList(opts.slides);

/* ── OOXML parsing ──────────────────────────────────────── */

const EMU_PER_IN = 914400;
const EMU_PER_PT = 12700;

const zip = await JSZip.loadAsync(await readFile(FILE));
const read = async (p) => (zip.file(p) ? zip.file(p).async('string') : null);

const presXml = await read('ppt/presentation.xml');
if (!presXml) { console.error('✗ ppt/presentation.xml missing — not a PowerPoint package'); process.exit(1); }
const sz = presXml.match(/<p:sldSz\s+cx="(\d+)"\s+cy="(\d+)"/);
const DECK = { cx: sz ? Number(sz[1]) : 12192000, cy: sz ? Number(sz[2]) : 6858000 };

const unescapeXml = (s) => s
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, d) => String.fromCharCode(Number(d)))
  .replace(/&amp;/g, '&');

/** <a:off x/> + <a:ext cx/> from a shape's own <a:xfrm>, in EMU. */
function xfrmOf(xml) {
  const off = xml.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"/);
  const ext = xml.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
  if (!off || !ext) return null;
  return { x: Number(off[1]), y: Number(off[2]), w: Number(ext[1]), h: Number(ext[2]) };
}

/** Every <tag>…</tag> block at this nesting, non-greedy but nesting-aware for our shapes. */
function blocks(xml, tag) {
  const out = [];
  const open = new RegExp(`<${tag}(?:\\s[^>]*)?>`, 'g');
  let m;
  while ((m = open.exec(xml))) {
    // Walk forward counting opens/closes so nested same-name tags do not truncate.
    let depth = 1, i = open.lastIndex;
    const scan = new RegExp(`<${tag}(?:\\s[^>]*)?>|</${tag}>`, 'g');
    scan.lastIndex = i;
    let s;
    while (depth > 0 && (s = scan.exec(xml))) {
      depth += s[0].startsWith('</') ? -1 : 1;
      i = scan.lastIndex;
    }
    out.push(xml.slice(m.index, i));
    open.lastIndex = i;
  }
  return out;
}

/** Parse the runs of one <p:txBody> into paragraphs of styled runs. */
function parseTextBody(xml) {
  const paragraphs = [];
  for (const para of blocks(xml, 'a:p')) {
    const runs = [];
    // <a:r> carries text; <a:br> is an explicit line break.
    for (const piece of para.match(/<a:r>[\s\S]*?<\/a:r>|<a:br\s*\/>|<a:br>[\s\S]*?<\/a:br>/g) || []) {
      if (piece.startsWith('<a:br')) { runs.push({ break: true }); continue; }
      const t = piece.match(/<a:t>([\s\S]*?)<\/a:t>/);
      const rPr = piece.match(/<a:rPr\b[^>]*>|<a:rPr\b[^>]*\/>/);
      const attrs = rPr ? rPr[0] : '';
      const num = (name) => {
        const m = attrs.match(new RegExp(`${name}="(-?\\d+)"`));
        return m ? Number(m[1]) : null;
      };
      const sz = num('sz');
      runs.push({
        text: t ? unescapeXml(t[1]) : '',
        pt: sz ? sz / 100 : null,
        bold: /\bb="1"/.test(attrs),
        italic: /\bi="1"/.test(attrs),
        spacing: num('spc'),
        color: (piece.match(/<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"/) || [])[1] || null,
        // <a:alpha> is opacity in thousandths of a percent, and absent means opaque —
        // so a missing element is 100, not 0. Searchable mode's hidden layer is 0.
        opacity: Number((piece.match(/<a:alpha val="(\d+)"/) || [])[1] ?? 100000) / 1000,
        font: (piece.match(/<a:latin[^>]*typeface="([^"]*)"/) || [])[1] || null,
      });
    }
    const align = (para.match(/<a:pPr[^>]*\balgn="([^"]*)"/) || [])[1] || 'l';
    const lnSpc = (para.match(/<a:lnSpc><a:spcPts val="(\d+)"/) || [])[1];
    paragraphs.push({ align, lineSpacingPt: lnSpc ? Number(lnSpc) / 100 : null, runs });
  }
  return paragraphs;
}

const slideNames = Object.keys(zip.files)
  .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
  .sort((a, b) => Number(a.match(/(\d+)/)[1]) - Number(b.match(/(\d+)/)[1]));

const slides = [];
for (const [i, name] of slideNames.entries()) {
  const num = i + 1;
  if (WANTED && !WANTED.has(num)) continue;
  const xml = await read(name);
  const rels = (await read(name.replace('slides/', 'slides/_rels/') + '.rels')) || '';
  const relMap = new Map();
  for (const m of rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)) relMap.set(m[1], m[2]);

  const pictures = [];
  for (const pic of blocks(xml, 'p:pic')) {
    const geo = xfrmOf(pic);
    const embed = (pic.match(/<a:blip[^>]*r:embed="([^"]+)"/) || [])[1];
    const target = embed ? relMap.get(embed) : null;
    pictures.push({
      geo,
      part: target ? `ppt/${target.replace(/^\.\.\//, '')}` : null,
      alt: (pic.match(/<p:cNvPr[^>]*descr="([^"]*)"/) || [])[1] || '',
    });
  }

  const shapes = [];
  for (const sp of blocks(xml, 'p:sp')) {
    const geo = xfrmOf(sp);
    const txBody = blocks(sp, 'p:txBody')[0] || '';
    const paragraphs = txBody ? parseTextBody(txBody) : [];
    const wrap = /wrap="none"/.test(txBody) ? 'none' : 'square';
    const plain = paragraphs
      .map((p) => p.runs.map((r) => (r.break ? '\n' : r.text)).join(''))
      .join('\n');
    shapes.push({
      index: shapes.length + 1,
      geo, wrap, paragraphs, text: plain,
      hasInk: plain.trim().length > 0,
      // A fully transparent run is a searchable-mode text layer, not visible content.
      visible: paragraphs.some((p) => p.runs.some((r) => r.text.trim() && r.opacity > 0)),
      hyperlink: (sp.match(/<a:hlinkClick[^>]*r:id="([^"]+)"/) || [])[1]
        ? relMap.get((sp.match(/<a:hlinkClick[^>]*r:id="([^"]+)"/) || [])[1]) : null,
    });
  }

  slides.push({ number: num, part: name, pictures, shapes });
}

/* ── Lint ───────────────────────────────────────────────── */

const findings = [];
const inches = (emu) => Math.round((emu / EMU_PER_IN) * 100) / 100;
const pctOf = (a, b) => Math.round((a / b) * 100);

function normalizeHex(hex) {
  if (!hex) return '000000';
  return hex.replace(/^#/, '').toUpperCase().padStart(6, '0').slice(0, 6);
}

function hexToRgb(hex) {
  const value = normalizeHex(hex);
  return [0, 2, 4].map((i) => parseInt(value.slice(i, i + 2), 16));
}

function rgbToHex(rgb) {
  return rgb.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase();
}

function luminance(rgb) {
  const [r, g, b] = rgb.map((v) => {
    const c = v / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

function contrastRatio(fgHex, bgRgb) {
  const fg = luminance(hexToRgb(fgHex));
  const bg = luminance(bgRgb);
  const light = Math.max(fg, bg);
  const dark = Math.min(fg, bg);
  return (light + 0.05) / (dark + 0.05);
}

function visibleTextColors(shape) {
  const colors = new Set();
  for (const p of shape.paragraphs) {
    for (const r of p.runs) {
      if (!r.text?.trim() || r.opacity < 50) continue;
      colors.add(normalizeHex(r.color));
    }
  }
  return [...colors];
}

/** Overlap area of two EMU rects. */
function overlapArea(a, b) {
  const w = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
  const h = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
  return w > 0 && h > 0 ? w * h : 0;
}

for (const s of slides) {
  const visible = s.shapes.filter((sh) => sh.geo && sh.visible);

  for (const sh of visible) {
    const g = sh.geo;
    const snippet = sh.text.replace(/\s+/g, ' ').trim().slice(0, 60);

    // Off-slide: the single most common symptom of a substituted font reflowing text.
    const over = {
      left: -g.x, top: -g.y,
      right: g.x + g.w - DECK.cx, bottom: g.y + g.h - DECK.cy,
    };
    for (const [edge, amount] of Object.entries(over)) {
      // A tenth of an inch of slop is rounding, not a layout error.
      if (amount > EMU_PER_IN / 10) {
        findings.push({
          level: 'warn', slide: s.number, kind: 'off-slide',
          message: `text box runs ${inches(amount)}in past the ${edge} edge: "${snippet}"`,
        });
      }
    }

    for (const p of sh.paragraphs) {
      for (const r of p.runs) {
        if (r.text.trim() && r.pt != null && r.pt < opts.minFont) {
          findings.push({
            level: 'warn', slide: s.number, kind: 'tiny-font',
            message: `${r.pt}pt text (below ${opts.minFont}pt): "${r.text.trim().slice(0, 40)}"`,
          });
        }
      }
    }
  }

  // Collisions between visible text boxes. Boxes are positioned from the HTML layout,
  // where nothing overlapped — so a substantial overlap here means text reflowed.
  for (let i = 0; i < visible.length; i++) {
    for (let j = i + 1; j < visible.length; j++) {
      const a = visible[i], b = visible[j];
      const area = overlapArea(a.geo, b.geo);
      if (!area) continue;
      const smaller = Math.min(a.geo.w * a.geo.h, b.geo.w * b.geo.h);
      const pct = pctOf(area, smaller);
      if (pct < opts.overlap) continue;
      findings.push({
        level: 'warn', slide: s.number, kind: 'overlap',
        message: `two text boxes overlap by ${pct}% — "${a.text.replace(/\s+/g, ' ').trim().slice(0, 32)}" / ` +
          `"${b.text.replace(/\s+/g, ' ').trim().slice(0, 32)}"`,
      });
    }
  }

  if (!s.pictures.length && !s.shapes.some((sh) => sh.hasInk)) {
    findings.push({ level: 'fail', slide: s.number, kind: 'empty', message: 'slide has no pictures and no text' });
  }
}

async function runContrastInspection() {
  if (!opts.contrast || !slides.some((s) => s.shapes.some((sh) => sh.geo && sh.visible && visibleTextColors(sh).length))) return;

  const PX_W = 1920;
  const PX_PER_EMU = PX_W / DECK.cx;
  const PX_H = Math.round(DECK.cy * PX_PER_EMU);
  const px = (emu) => emu * PX_PER_EMU;
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();

  try {
    for (const s of slides) {
      const checks = s.shapes
        .filter((sh) => sh.geo && sh.visible)
        .map((sh) => ({
          index: sh.index,
          text: sh.text.replace(/\s+/g, ' ').trim().slice(0, 60),
          x: px(sh.geo.x), y: px(sh.geo.y), w: px(sh.geo.w), h: px(sh.geo.h),
          colors: visibleTextColors(sh),
        }))
        .filter((sh) => sh.colors.length);
      if (!checks.length) continue;

      const pictures = [];
      for (const pic of s.pictures) {
        if (!pic.geo || !pic.part || !zip.file(pic.part)) continue;
        const buf = await zip.file(pic.part).async('nodebuffer');
        const ext = pic.part.split('.').pop().toLowerCase();
        const mime = ext === 'png' ? 'image/png' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`;
        pictures.push({
          src: `data:${mime};base64,${buf.toString('base64')}`,
          x: px(pic.geo.x), y: px(pic.geo.y), w: px(pic.geo.w), h: px(pic.geo.h),
        });
      }

      const sampled = await page.evaluate(async ({ width, height, pictures, checks }) => {
        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d', { willReadFrequently: true });
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, width, height);
        const loadImage = (src) => new Promise((resolve, reject) => {
          const img = new Image();
          img.onload = () => resolve(img);
          img.onerror = reject;
          img.src = src;
        });
        for (const pic of pictures) {
          const img = await loadImage(pic.src);
          ctx.drawImage(img, pic.x, pic.y, pic.w, pic.h);
        }
        const samplePoints = (box) => {
          const xs = [0.18, 0.5, 0.82];
          const ys = [0.3, 0.5, 0.7];
          const out = [];
          for (const xf of xs) {
            for (const yf of ys) {
              const x = Math.max(0, Math.min(width - 1, Math.round(box.x + box.w * xf)));
              const y = Math.max(0, Math.min(height - 1, Math.round(box.y + box.h * yf)));
              out.push([...ctx.getImageData(x, y, 1, 1).data.slice(0, 3)]);
            }
          }
          return out;
        };
        return checks.map((check) => ({ ...check, backgrounds: samplePoints(check) }));
      }, { width: PX_W, height: PX_H, pictures, checks });

      for (const check of sampled) {
        for (const color of check.colors) {
          let worst = { ratio: Infinity, bg: [255, 255, 255] };
          for (const bg of check.backgrounds) {
            const ratio = contrastRatio(color, bg);
            if (ratio < worst.ratio) worst = { ratio, bg };
          }
          if (worst.ratio >= opts.minContrast) continue;
          findings.push({
            level: worst.ratio < opts.failContrast ? 'fail' : 'warn',
            slide: s.number,
            kind: 'low-contrast',
            message: `${worst.ratio.toFixed(2)}:1 contrast for #${color} text on #${rgbToHex(worst.bg)} background: "${check.text}"`,
          });
        }
      }
    }
  } finally {
    await browser.close();
  }
}

await runContrastInspection();

/* ── Preview render ─────────────────────────────────────── */

let previewFiles = [];
if (opts.preview) {
  const dir = resolve(opts.preview);
  await mkdir(dir, { recursive: true });

  // Render at the deck's own aspect, 1920 wide, so a preview PNG lines up 1:1 with the
  // screenshots the converter captured (`--keep-shots`) for direct comparison.
  const PX_W = 1920;
  const PX_PER_EMU = PX_W / DECK.cx;
  const PX_H = Math.round(DECK.cy * PX_PER_EMU);
  const px = (emu) => Math.round(emu * PX_PER_EMU * 100) / 100;
  // 1pt = 1/72in, and the deck is DECK.cx EMU wide, so points scale with the same ratio.
  const ptToPx = (p) => Math.round(p * EMU_PER_PT * PX_PER_EMU * 100) / 100;

  const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  // Give each family a fallback of the same class, so a font missing locally degrades
  // the way PowerPoint would rather than falling back to a proportional default.
  const MONO = /consolas|courier|mono|menlo/i;
  const SERIF = /georgia|times|garamond|cambria|serif/i;
  const fontStack = (f) => {
    if (!f) return 'Arial, sans-serif';
    const generic = MONO.test(f) ? 'monospace' : SERIF.test(f) ? 'serif' : 'sans-serif';
    // Single quotes: this goes inside style="…", and a double quote would close the
    // attribute and silently discard every declaration after it.
    return `'${f.replace(/'/g, '')}', ${generic}`;
  };

  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: PX_W, height: PX_H } });

  for (const s of slides) {
    const parts = [];
    for (const pic of s.pictures) {
      if (!pic.geo || !pic.part || !zip.file(pic.part)) continue;
      const buf = await zip.file(pic.part).async('nodebuffer');
      const ext = pic.part.split('.').pop().toLowerCase();
      const mime = ext === 'png' ? 'image/png' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`;
      parts.push(
        `<img src="data:${mime};base64,${buf.toString('base64')}" style="position:absolute;` +
        `left:${px(pic.geo.x)}px;top:${px(pic.geo.y)}px;width:${px(pic.geo.w)}px;height:${px(pic.geo.h)}px">`
      );
    }
    for (const sh of s.shapes) {
      if (!sh.geo || !sh.visible) continue;
      const ALIGN = { l: 'left', ctr: 'center', r: 'right', just: 'justify' };
      const paras = sh.paragraphs.map((p) => {
        const runs = p.runs.map((r) => {
          if (r.break) return '<br>';
          if (!r.text) return '';
          const st = [
            `font-family:${fontStack(r.font)}`,
            r.pt ? `font-size:${ptToPx(r.pt)}px` : '',
            r.bold ? 'font-weight:700' : 'font-weight:400',
            r.italic ? 'font-style:italic' : '',
            r.color ? `color:#${r.color}` : '',
            r.opacity < 100 ? `opacity:${(r.opacity / 100).toFixed(3)}` : '',
            r.spacing ? `letter-spacing:${ptToPx(r.spacing / 100)}px` : '',
          ].filter(Boolean).join(';');
          return `<span style="${st}">${esc(r.text)}</span>`;
        }).join('');
        const pst = [
          `text-align:${ALIGN[p.align] || 'left'}`,
          p.lineSpacingPt ? `line-height:${ptToPx(p.lineSpacingPt)}px` : '',
        ].filter(Boolean).join(';');
        return `<p style="${pst}">${runs || '<br>'}</p>`;
      }).join('');
      parts.push(
        `<div style="position:absolute;left:${px(sh.geo.x)}px;top:${px(sh.geo.y)}px;` +
        `width:${px(sh.geo.w)}px;height:${px(sh.geo.h)}px;` +
        // wrap="none" boxes must not re-wrap here either, or the preview lies.
        `white-space:${sh.wrap === 'none' ? 'pre' : 'pre-wrap'};overflow:visible">${paras}</div>`
      );
    }

    const html =
      `<!doctype html><meta charset="utf-8"><style>` +
      `*{margin:0;padding:0;box-sizing:border-box}` +
      `html,body{width:${PX_W}px;height:${PX_H}px;overflow:hidden;background:#fff}` +
      `</style><body>${parts.join('')}</body>`;
    await page.setContent(html, { waitUntil: 'load' });
    await page.evaluate(() => document.fonts.ready);
    const out = join(dir, `slide-${String(s.number).padStart(3, '0')}.png`);
    await page.screenshot({ path: out });
    previewFiles.push(out);
  }
  await browser.close();
}

/* ── Report ─────────────────────────────────────────────── */

const fails = findings.filter((f) => f.level === 'fail');
const warns = findings.filter((f) => f.level === 'warn');

if (opts.json) {
  console.log(JSON.stringify({
    file: FILE,
    deck: { emu: DECK, inches: { w: inches(DECK.cx), h: inches(DECK.cy) } },
    slides: slides.map((s) => ({
      number: s.number,
      pictures: s.pictures.map((p) => ({ ...p, inches: p.geo ? { x: inches(p.geo.x), y: inches(p.geo.y), w: inches(p.geo.w), h: inches(p.geo.h) } : null })),
      shapes: s.shapes.map((sh) => ({
        inches: sh.geo ? { x: inches(sh.geo.x), y: inches(sh.geo.y), w: inches(sh.geo.w), h: inches(sh.geo.h) } : null,
        wrap: sh.wrap, visible: sh.visible, hyperlink: sh.hyperlink, text: sh.text,
        fonts: [...new Set(sh.paragraphs.flatMap((p) => p.runs.map((r) => r.font).filter(Boolean)))],
        sizes: [...new Set(sh.paragraphs.flatMap((p) => p.runs.map((r) => r.pt).filter(Boolean)))],
        colors: visibleTextColors(sh),
      })),
    })),
    findings, previews: previewFiles, ok: fails.length === 0,
  }, null, 2));
} else {
  console.log('');
  console.log(`  ${FILE}`);
  console.log(`  ${slideNames.length} slides · ${inches(DECK.cx)}×${inches(DECK.cy)}in` +
    (WANTED ? `  (inspecting ${slides.length})` : ''));
  console.log('');

  if (!opts.quiet) {
    for (const s of slides) {
      const vis = s.shapes.filter((sh) => sh.visible).length;
      const hidden = s.shapes.filter((sh) => sh.hasInk && !sh.visible).length;
      console.log(`  ── slide ${s.number} — ${s.pictures.length} picture(s), ${vis} visible text box(es)` +
        (hidden ? `, ${hidden} invisible` : ''));
      for (const sh of s.shapes) {
        if (!sh.geo || !sh.hasInk) continue;
        const g = sh.geo;
        const fonts = [...new Set(sh.paragraphs.flatMap((p) => p.runs.map((r) => r.font).filter(Boolean)))];
        const sizes = [...new Set(sh.paragraphs.flatMap((p) => p.runs.map((r) => r.pt).filter(Boolean)))].sort((a, b) => b - a);
        const at = `${inches(g.x)},${inches(g.y)} ${inches(g.w)}×${inches(g.h)}in`;
        const style = `${fonts.join('/') || '?'} ${sizes.join('/')}pt${sh.wrap === 'none' ? ' nowrap' : ''}${sh.visible ? '' : ' hidden'}`;
        console.log(`     ${at.padEnd(28)} ${style.padEnd(30)} ${JSON.stringify(sh.text.replace(/\s+/g, ' ').trim().slice(0, 54))}`);
      }
    }
    console.log('');
  }

  for (const f of warns) console.log(`  ⚠ slide ${f.slide} [${f.kind}] ${f.message}`);
  for (const f of fails) console.log(`  ✗ slide ${f.slide} [${f.kind}] ${f.message}`);
  if (!findings.length) console.log('  no layout problems found');
  console.log('');
  if (previewFiles.length) {
    console.log(`  ${previewFiles.length} preview PNG(s) in ${resolve(opts.preview)}`);
    console.log('');
  }
}

process.exit(fails.length ? 1 : 0);
