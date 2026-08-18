#!/usr/bin/env node
/**
 * html2pptx.mjs — Convert a frontend-slides HTML deck into a PowerPoint .pptx
 *
 * Usage:
 *   node html2pptx.mjs <input.html> [output.pptx] [options]
 *
 * Options:
 *   --mode <image|searchable|editable>   Conversion mode (default: image)
 *   --selector <css>                     Slide selector (default: .slide)
 *   --scale <n>                          Screenshot device pixel ratio (default: 2)
 *   --jpeg [quality]                     Encode slide art as JPEG (default: PNG)
 *   --native-images                      editable mode: lift <img> out as real pictures
 *   --no-background-art                  editable mode: omit full-slide screenshot artwork
 *   --keep-chrome                        Do not hide nav/counter/edit UI outside the stage
 *   --font-map "A=B,C=D"                 Substitute webfont A with Office font B
 *   --no-notes                           Skip speaker notes
 *   --no-links                           Skip hyperlink hotspots
 *   --no-shrink                          editable mode: do not shrink text that a
 *                                        substituted font would overflow out of its box
 *   --title/--author/--company/--subject Document metadata
 *   --keep-shots <dir>                   Keep the source screenshots (for validation)
 *   --json                               Print a machine-readable summary to stdout
 *
 * Modes:
 *   image       Each slide is one full-bleed picture. Perfect fidelity, no editable text.
 *   searchable  image + an invisible native text layer (searchable / copyable / screen-readable).
 *   editable    Background art is captured with the text hidden, then text is rebuilt as real
 *               PowerPoint text boxes. Editable, but webfonts are substituted.
 *
 * The pipeline: serve the deck over HTTP -> drive it in headless Chromium -> capture each
 * slide at the deck's own authored stage size -> emit a .pptx at PowerPoint's canonical
 * 13.3333333 x 7.5in widescreen size -> repair pptxgenjs's OOXML/PowerPoint quirks.
 */

import { createServer } from 'http';
import { readFileSync, mkdirSync, existsSync, rmSync } from 'fs';
import { readFile, writeFile } from 'fs/promises';
import { join, extname, dirname, basename, resolve } from 'path';
import { createRequire } from 'module';
import { tmpdir } from 'os';

/* ═══════════════════════════════════════════════════════════
   DEPENDENCY LOADING
   ESM resolves imports relative to this file, not the cwd, so a shared
   dependency cache has to be resolved explicitly. SLIDES_PPTX_DEPS points at
   the directory containing node_modules; without it we fall back to normal
   resolution so the script also runs from a project that already has the deps.
   ═══════════════════════════════════════════════════════════ */

let chromium, PptxGenJS, JSZip;
try {
  const depsRoot = process.env.SLIDES_PPTX_DEPS;
  if (depsRoot) {
    const req = createRequire(join(depsRoot, 'index.js'));
    ({ chromium } = req('playwright'));
    PptxGenJS = req('pptxgenjs');
    JSZip = req('jszip');
  } else {
    ({ chromium } = await import('playwright'));
    PptxGenJS = (await import('pptxgenjs')).default;
    JSZip = (await import('jszip')).default;
  }
} catch (e) {
  console.error(`✗ Missing dependencies: ${e.message}`);
  console.error('  Install with: npm install playwright pptxgenjs jszip && npx playwright install chromium');
  console.error('  Or run through scripts/slides-to-pptx.sh, which manages them for you.');
  process.exit(1);
}

/* ═══════════════════════════════════════════════════════════
   CLI PARSING
   ═══════════════════════════════════════════════════════════ */

const argv = process.argv.slice(2);
const opts = {
  mode: 'image',
  selector: '.slide',
  scale: 2,
  jpeg: false,
  jpegQuality: 88,
  nativeImages: false,
  backgroundArt: true,
  keepChrome: false,
  fontMap: {},
  notes: true,
  links: true,
  meta: {},
  keepShots: null,
  json: false,
  shrink: true,
};
const positional = [];

for (let i = 0; i < argv.length; i++) {
  const a = argv[i];
  const next = () => argv[++i];
  switch (a) {
    case '--mode': opts.mode = next(); break;
    case '--selector': opts.selector = next(); break;
    case '--scale': opts.scale = Number(next()); break;
    case '--jpeg':
      opts.jpeg = true;
      // Optional numeric quality argument
      if (argv[i + 1] && /^\d+$/.test(argv[i + 1])) opts.jpegQuality = Number(next());
      break;
    case '--native-images': opts.nativeImages = true; break;
    case '--no-background-art': opts.backgroundArt = false; break;
    case '--keep-chrome': opts.keepChrome = true; break;
    case '--font-map':
      for (const pair of next().split(',')) {
        const [from, to] = pair.split('=').map((s) => s && s.trim());
        if (from && to) opts.fontMap[from.toLowerCase()] = to;
      }
      break;
    case '--no-notes': opts.notes = false; break;
    case '--no-links': opts.links = false; break;
    case '--no-shrink': opts.shrink = false; break;
    case '--title': opts.meta.title = next(); break;
    case '--author': opts.meta.author = next(); break;
    case '--company': opts.meta.company = next(); break;
    case '--subject': opts.meta.subject = next(); break;
    case '--keep-shots': opts.keepShots = next(); break;
    case '--json': opts.json = true; break;
    default:
      if (a.startsWith('--')) fail(`Unknown option: ${a}`);
      positional.push(a);
  }
}

if (!positional.length) {
  fail('Usage: node html2pptx.mjs <input.html> [output.pptx] [--mode image|searchable|editable]');
}
if (!['image', 'searchable', 'editable'].includes(opts.mode)) {
  fail(`--mode must be image, searchable, or editable (got "${opts.mode}")`);
}
if (!opts.backgroundArt && opts.mode !== 'editable') {
  fail('--no-background-art is only supported with --mode editable');
}

const INPUT = resolve(positional[0]);
if (!existsSync(INPUT)) fail(`File not found: ${INPUT}`);
const OUTPUT = resolve(positional[1] || INPUT.replace(/\.html?$/i, '') + '.pptx');

function fail(msg) {
  console.error(`✗ ${msg}`);
  process.exit(1);
}
// Progress goes to stderr so --json keeps stdout clean for machine consumption.
function log(msg) {
  if (opts.json) console.error(msg);
  else console.log(msg);
}

/* ═══════════════════════════════════════════════════════════
   STATIC FILE SERVER
   Webfonts, relative image paths and CSS all need a real origin —
   file:// URLs break font loading and fetch().
   ═══════════════════════════════════════════════════════════ */

const MIME = {
  '.html': 'text/html', '.css': 'text/css', '.js': 'application/javascript',
  '.json': 'application/json', '.png': 'image/png', '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.svg': 'image/svg+xml',
  '.webp': 'image/webp', '.avif': 'image/avif', '.woff': 'font/woff',
  '.woff2': 'font/woff2', '.ttf': 'font/ttf', '.otf': 'font/otf',
  '.mp4': 'video/mp4', '.webm': 'video/webm',
};

const SERVE_DIR = dirname(INPUT);
const HTML_FILE = basename(INPUT);

const server = createServer((req, res) => {
  const decoded = decodeURIComponent(req.url.split('?')[0]);
  const filePath = join(SERVE_DIR, decoded === '/' ? HTML_FILE : decoded);
  // Refuse to serve outside the deck's own directory.
  if (!resolve(filePath).startsWith(resolve(SERVE_DIR))) {
    res.writeHead(403); res.end('Forbidden'); return;
  }
  try {
    const body = readFileSync(filePath);
    res.writeHead(200, { 'Content-Type': MIME[extname(filePath).toLowerCase()] || 'application/octet-stream' });
    res.end(body);
  } catch {
    res.writeHead(404); res.end('Not found');
  }
});
const port = await new Promise((r) => server.listen(0, () => r(server.address().port)));

/* ═══════════════════════════════════════════════════════════
   FONT SUBSTITUTION
   Webfonts do not exist on the viewer's machine. Map them to fonts that ship with
   Office on both macOS and Windows so substitution is predictable, not arbitrary.
   Prefer substitutes with similar *widths* — a wider stand-in reflows every text box
   and is the main source of overflow in editable mode.
   ═══════════════════════════════════════════════════════════ */

const DEFAULT_FONT_MAP = {
  'clash display': 'Impact', 'satoshi': 'Arial', 'space grotesk': 'Arial',
  'general sans': 'Arial', 'cabinet grotesk': 'Trebuchet MS', 'switzer': 'Arial',
  'chillax': 'Trebuchet MS', 'zodiak': 'Georgia', 'sentient': 'Georgia',
  'gambetta': 'Georgia', 'bespoke serif': 'Georgia', 'playfair display': 'Georgia',
  'dm serif display': 'Georgia', 'libre baskerville': 'Georgia', 'lora': 'Georgia',
  'inter': 'Arial', 'roboto': 'Arial', 'poppins': 'Trebuchet MS',
  'montserrat': 'Trebuchet MS', 'archivo': 'Arial', 'oswald': 'Impact',
  'bebas neue': 'Impact', 'anton': 'Impact', 'jetbrains mono': 'Consolas',
  'ibm plex mono': 'Consolas', 'space mono': 'Consolas', 'fira code': 'Consolas',
  'courier prime': 'Courier New',
};
// Resolved map actually handed to the page and to the builder, keyed by lowercase family.
const FONT_MAP = { ...DEFAULT_FONT_MAP, ...opts.fontMap };
const substitutions = new Map();
function mapFont(name) {
  if (!name) return 'Arial';
  const mapped = FONT_MAP[name.toLowerCase()];
  if (mapped) { substitutions.set(name, mapped); return mapped; }
  return name; // already an Office-safe or locally installed family
}

/* ═══════════════════════════════════════════════════════════
   BROWSER-SIDE HELPERS
   These functions are serialized and run inside the page, so they
   may only reference their own arguments and browser globals.
   ═══════════════════════════════════════════════════════════ */

/** Measure the authored stage: its pixel size and its on-screen scale factor. */
function measureStage(selector) {
  const slide = document.querySelector(selector);
  if (!slide) return null;
  const stage = document.querySelector('.deck-stage') || slide.parentElement;
  const cs = getComputedStyle(stage);
  // The authored size is the untransformed CSS box; getBoundingClientRect() is post-transform.
  const authoredW = parseFloat(cs.width) || slide.offsetWidth || 1920;
  const authoredH = parseFloat(cs.height) || slide.offsetHeight || 1080;
  const rect = stage.getBoundingClientRect();
  return {
    width: authoredW,
    height: authoredH,
    scale: rect.width / authoredW || 1,
    count: document.querySelectorAll(selector).length,
  };
}

/**
 * Hide presentation chrome so it does not get baked into every slide.
 * Chrome is precisely "positioned elements that live outside the slide stage" —
 * page counters, nav arrows, the inline-edit toggle. Anything inside the stage is
 * part of the design (including a deck's own page numbers) and is left alone.
 */
function hideChrome(selector) {
  const slide = document.querySelector(selector);
  const stage = document.querySelector('.deck-stage') || slide?.parentElement;
  if (!stage) return 0;
  let n = 0;
  for (const el of document.body.querySelectorAll('*')) {
    if (stage.contains(el) || el.contains(stage)) continue;
    const cs = getComputedStyle(el);
    if (!['fixed', 'sticky', 'absolute'].includes(cs.position)) continue;
    if (cs.display === 'none') continue;
    el.style.setProperty('display', 'none', 'important');
    n++;
  }
  return n;
}

/** Show slide `index` using the deck's own mechanism, then settle every animation. */
async function showSlide(args) {
  const { selector, index } = args;
  const slides = [...document.querySelectorAll(selector)];

  // 1. Prefer the deck's own controller if it exposed one.
  for (const ctrl of [window.presentation, window.deck, window.slidePresentation, window.app]) {
    if (!ctrl) continue;
    for (const fn of ['showSlide', 'goToSlide', 'goTo', 'setSlide']) {
      if (typeof ctrl[fn] === 'function') { try { ctrl[fn](index); } catch {} }
    }
  }

  // 2. Drive the class contract that viewport-base.css defines. Never touch `display`:
  //    the base CSS switches slides with visibility/opacity, and forcing display
  //    breaks layout classes like `.slide-content { display: flex }`.
  slides.forEach((s, i) => {
    s.classList.toggle('active', i === index);
    s.classList.toggle('visible', i === index);
  });

  // 3. Let the deck's reveal transitions run, then jump every finite animation to
  //    its end state. This lands on the design's intended final frame instead of
  //    stomping opacity/transform (which would break translate-based centering).
  await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
  for (const anim of document.getAnimations()) {
    const iterations = anim.effect?.getTiming?.().iterations;
    if (iterations === Infinity) continue; // looping ambience — leave it running
    try { anim.finish(); } catch {}
  }
  await new Promise((r) => requestAnimationFrame(r));
}

/**
 * Extract text boxes, hyperlinks, images and speaker notes from the active slide.
 * Coordinates come back in authored stage pixels with the slide's origin at (0,0).
 */
function extractSlide(args) {
  const { selector, index, wantNotes, wantLinks, wantFit, fontMap } = args;
  const slide = document.querySelectorAll(selector)[index];
  if (!slide) return null;

  const slideRect = slide.getBoundingClientRect();
  const stageScale = slideRect.width / (parseFloat(getComputedStyle(slide).width) || slideRect.width) || 1;

  /** Convert a client rect into authored stage px relative to the slide. */
  const norm = (r) => ({
    x: (r.left - slideRect.left) / stageScale,
    y: (r.top - slideRect.top) / stageScale,
    w: r.width / stageScale,
    h: r.height / stageScale,
  });

  const isVisible = (el) => {
    const cs = getComputedStyle(el);
    if (cs.display === 'none' || cs.visibility === 'hidden') return false;
    if (parseFloat(cs.opacity) === 0) return false;
    const r = el.getBoundingClientRect();
    return r.width > 0.5 && r.height > 0.5;
  };
  const isRasterized = (el) => el.closest('[data-pptx-raster], [data-pptx-screenshot]');

  // An element is a "text box root" when it is not laid out in its parent's inline flow
  // and owns text that isn't claimed by a nested box root.
  //
  // inline-block/inline-flex/inline-grid count as inline flow: a mid-sentence badge or
  // chip is positioned *by* the surrounding text, so lifting it into its own box makes
  // the parent's words flow through the space it occupies and the two collide.
  //
  // And once inside such an element, everything below it is absorbed too — an
  // inline-flex's own children are blockified to `block`, but the whole chip still
  // occupies one contiguous span of the parent's line, so its parts cannot be
  // positioned independently either.
  const INLINE_FLOW = new Set(['inline', 'inline-block', 'inline-flex', 'inline-grid',
    'inline-table', 'contents', 'ruby', 'ruby-text', 'ruby-base']);
  const isInlineFlow = (el) => INLINE_FLOW.has(getComputedStyle(el).display);
  const isBox = (el) => !isInlineFlow(el);
  /** Already absorbed into an ancestor's text box, so not a root of its own. */
  const absorbed = (el) => {
    for (let p = el.parentElement; p && p !== slide; p = p.parentElement) {
      if (isInlineFlow(p)) return true;
    }
    return false;
  };

  const rgbToHex = (str) => {
    const m = str.match(/rgba?\(([^)]+)\)/);
    if (!m) return { hex: '000000', alpha: 1 };
    const p = m[1].split(/[,\s/]+/).filter(Boolean).map(Number);
    const hex = p.slice(0, 3).map((v) => Math.round(v).toString(16).padStart(2, '0')).join('');
    return { hex: hex.toUpperCase(), alpha: p.length > 3 ? p[3] : 1 };
  };

  const runStyle = (el) => {
    const cs = getComputedStyle(el);
    const { hex, alpha } = rgbToHex(cs.color);
    return {
      font: (cs.fontFamily || '').split(',')[0].replace(/["']/g, '').trim(),
      sizePx: parseFloat(cs.fontSize) || 16,
      bold: (parseInt(cs.fontWeight, 10) || 400) >= 600,
      italic: cs.fontStyle === 'italic' || cs.fontStyle === 'oblique',
      underline: cs.textDecorationLine.includes('underline'),
      color: hex,
      alpha,
      transform: cs.textTransform,
      letterSpacingPx: cs.letterSpacing === 'normal' ? 0 : parseFloat(cs.letterSpacing) || 0,
      pre: /^pre/.test(cs.whiteSpace),
    };
  };

  const applyCase = (text, transform) => {
    if (transform === 'uppercase') return text.toUpperCase();
    if (transform === 'lowercase') return text.toLowerCase();
    if (transform === 'capitalize') return text.replace(/\b\w/g, (c) => c.toUpperCase());
    return text;
  };

  /** Collect runs from a box root, stopping at nested box roots. */
  const collectRuns = (root) => {
    const runs = [];
    const hasInk = () => runs.some((r) => r.text && r.text.trim());

    // Newlines become explicit break markers rather than "\n" inside a run:
    // pptxgenjs splits runs on "\n" itself and reorders the pieces, which scrambles
    // multi-run <pre> blocks into a single jumbled line.
    const push = (text, style) => {
      const lines = text.split('\n');
      for (let i = 0; i < lines.length; i++) {
        if (i > 0) runs.push({ break: true });
        if (lines[i]) runs.push({ text: lines[i], style });
      }
    };

    const walk = (node, styleEl, inInline) => {
      for (const child of node.childNodes) {
        if (child.nodeType === Node.TEXT_NODE) {
          const st = runStyle(styleEl);
          // CSS whitespace processing: `pre` keeps everything verbatim, otherwise a run
          // of whitespace collapses to a single space. That surviving space is real
          // content — `<b>a</b> <i>b</i>` renders "a b", so dropping whitespace-only
          // text nodes glues words together and swallows the newlines between <pre> lines.
          const raw = st.pre ? child.nodeValue : child.nodeValue.replace(/\s+/g, ' ');
          if (!raw) continue;
          // Leading whitespace before any real text is layout, not content.
          if (!st.pre && !raw.trim() && !hasInk()) continue;
          push(applyCase(raw, st.transform), st);
        } else if (child.nodeType === Node.ELEMENT_NODE) {
          if (child.tagName === 'BR') { runs.push({ break: true }); continue; }
          const cs = getComputedStyle(child);
          if (cs.display === 'none' || cs.visibility === 'hidden') continue;
          const inline = isInlineFlow(child);
          if (!inline && !inInline) continue; // a real block sibling — its own text box
          // Blockified children of an inline chip are separated by flex/grid gaps rather
          // than by whitespace, so keep a space or the words run together.
          if (!inline && hasInk() && !/\s$/.test(runs[runs.length - 1]?.text || '')) {
            runs.push({ text: ' ', style: runStyle(child) });
          }
          walk(child, child, true);
        }
      }
    };
    walk(root, root, false);

    // Trailing whitespace and dangling breaks would add a blank line to the box.
    while (runs.length) {
      const last = runs[runs.length - 1];
      if (last.break || !last.text.trim()) runs.pop();
      else break;
    }
    return runs;
  };

  /**
   * Union of the client rects of the text this element owns *directly* — the glyphs that
   * will end up in its text box, excluding anything claimed by a nested box root.
   * A flex or grid row's border box spans its children, so positioning the owned text by
   * the element rect starts it underneath a sibling box instead of after it.
   */
  const ownedTextRect = (root) => {
    let out = null;
    const rows = new Set();
    const add = (r) => {
      if (!r || (r.width <= 0 && r.height <= 0)) return;
      // Client rects come one per line fragment; bucket by top edge to count lines.
      rows.add(Math.round(r.top));
      out = out
        ? { left: Math.min(out.left, r.left), top: Math.min(out.top, r.top),
            right: Math.max(out.right, r.right), bottom: Math.max(out.bottom, r.bottom) }
        : { left: r.left, top: r.top, right: r.right, bottom: r.bottom };
    };
    const walk = (node, inInline) => {
      for (const child of node.childNodes) {
        if (child.nodeType === Node.TEXT_NODE) {
          if (!child.nodeValue.trim()) continue;
          const range = document.createRange();
          range.selectNodeContents(child);
          for (const r of range.getClientRects()) add(r);
        } else if (child.nodeType === Node.ELEMENT_NODE) {
          const cs = getComputedStyle(child);
          if (cs.display === 'none' || cs.visibility === 'hidden') continue;
          const inline = isInlineFlow(child);
          // Same absorption rule as collectRuns, or the rect would not cover the runs.
          if (!inline && !inInline) continue;
          walk(child, true);
        }
      }
    };
    walk(root, false);
    if (!out) return null;
    return {
      left: out.left, top: out.top, width: out.right - out.left, height: out.bottom - out.top,
      // Distinct line boxes the text occupies, so single-line text can be marked
      // no-wrap and never reflowed by a renderer with different font metrics.
      lines: rows.size,
    };
  };

  /* ── Text boxes ─────────────────────────────────────────── */
  const texts = [];
  const textRoots = [];
  const processedListItems = new WeakSet();
  const listItemDepth = (li) => {
    let depth = 0;
    for (let p = li.parentElement; p && p !== slide; p = p.parentElement) {
      if (p.tagName === 'UL' || p.tagName === 'OL') depth++;
    }
    return Math.max(depth, 1);
  };
  for (const el of slide.querySelectorAll('*')) {
    if (isRasterized(el)) continue;
    if (!isBox(el) || absorbed(el) || !isVisible(el)) continue;
    const isList = el.tagName === 'UL' || el.tagName === 'OL';
    const listItems = isList
      ? [...el.children].filter((child) => child.tagName === 'LI' && isVisible(child))
      : [];
    if (el.tagName === 'LI' && processedListItems.has(el)) continue;

    let runs;
    if (isList && listItems.length) {
      runs = [];
      listItems.forEach((li, idx) => {
        const liRuns = collectRuns(li);
        if (!liRuns.some((r) => r.text && r.text.trim())) return;
        if (runs.length && idx > 0) runs.push({ break: true });
        runs.push(...liRuns);
        processedListItems.add(li);
      });
    } else {
      runs = collectRuns(el);
    }
    if (!runs.some((r) => r.text && r.text.trim())) continue;

    const cs = getComputedStyle(el);
    const box = norm(el.getBoundingClientRect());
    // Subtract padding so the text rect matches the glyph area, not the padded box.
    const pad = {
      l: parseFloat(cs.paddingLeft) || 0, r: parseFloat(cs.paddingRight) || 0,
      t: parseFloat(cs.paddingTop) || 0, b: parseFloat(cs.paddingBottom) || 0,
    };
    const lh = cs.lineHeight === 'normal' ? 0 : parseFloat(cs.lineHeight) || 0;
    const align = { start: 'left', end: 'right', 'justify': 'justify', center: 'center', right: 'right', left: 'left' }[cs.textAlign] || 'left';

    let x = box.x + pad.l, y = box.y + pad.t;
    let w = Math.max(box.w - pad.l - pad.r, 1);
    let h = Math.max(box.h - pad.t - pad.b, 1);
    // Text the browser laid out on one line must stay on one line. Sizing a box to
    // exactly its own measured width means any renderer with slightly wider metrics
    // wraps the last word onto a second line, which then collides with the box below.
    let nowrap = false;

    const ink = ownedTextRect(el);
    if (ink) {
      const i = norm(ink);
      nowrap = ink.lines === 1;
      // Does this element share its box with text that became a box of its own? A flex or
      // grid container's border box spans all its children, so a text box drawn from it
      // lies across its siblings, and its own words flow through the space they occupy.
      // When that is the case the glyphs' own rect is the honest geometry.
      //
      // Everywhere else keep the content box: it is what the browser wrapped the text at,
      // and the ink rect is not a reliable substitute — Chrome's range rects follow the
      // font's ascent and descent, which overshoot the line box whenever a design sets
      // `line-height` below 1, as display type usually does.
      const shares = [...el.querySelectorAll('*')]
        .some((d) => isBox(d) && !absorbed(d) && isVisible(d));
      if (shares) {
        x = Math.max(i.x, 0); y = Math.max(i.y, 0);
        w = Math.max(i.w, 1); h = Math.max(i.h, 1);
      }
    }

    textRoots.push(el);
    const listDepth = isList ? 1 : (el.tagName === 'LI' ? listItemDepth(el) : 0);
    const listType = isList ? (el.tagName === 'OL' ? 'number' : 'bullet')
      : (el.tagName === 'LI' ? (el.parentElement?.tagName === 'OL' ? 'number' : 'bullet') : null);
    texts.push({
      x, y, w, h, nowrap,
      align,
      listDepth,
      listType,
      lineHeightPx: lh,
      runs: runs.map((r) => (r.break ? { break: true } : { text: r.text, style: r.style })),
      plain: runs.map((r) => (r.break ? '\n' : r.text)).join('').trim(),
    });
  }

  /* ── Substituted-font fit ───────────────────────────────── */
  // A stand-in font with different metrics reflows text onto more lines, so the words
  // spill out of the box they were measured in and collide with the next one. Measure
  // each box again with its substitute applied and record the largest font scale that
  // still fits the authored bounds. The probe is an off-screen clone left in its real
  // parent, so it inherits the true cascade without disturbing the visible layout.
  if (wantFit) {
    // Substitutes are Office fonts, which are often not installed on the machine doing
    // the conversion (Consolas ships with Office, not with macOS). Naming the family on
    // its own would fall back to the browser's default *proportional* font and measure
    // something far narrower than a viewer will see, so pin a generic of the same class.
    // Keep this classification in step with fontStack() in inspect-pptx.mjs.
    const MONO = /consolas|courier|mono/i;
    const SERIF = /georgia|times|cambria|garamond|serif/i;
    const stack = (f) => `"${f}", ${MONO.test(f) ? 'monospace' : SERIF.test(f) ? 'serif' : 'sans-serif'}`;
    const subFamily = (el) => {
      const first = (getComputedStyle(el).fontFamily || '').split(',')[0].replace(/["']/g, '').trim();
      return fontMap[first.toLowerCase()] || null;
    };
    /**
     * Width of the widest laid-out line, to sub-pixel precision. `scrollWidth` is rounded
     * to a whole pixel, and a line two tenths of a pixel over the edge still wraps — so
     * the rounded value reports a comfortable fit for text that visibly spills. Client
     * rects arrive one per inline fragment, so rows are rebuilt by their top edge.
     */
    const widestLine = (el) => {
      const range = document.createRange();
      range.selectNodeContents(el);
      const rows = new Map();
      for (const r of range.getClientRects()) {
        if (r.width <= 0) continue;
        const key = Math.round(r.top);
        const row = rows.get(key);
        if (row) { row.l = Math.min(row.l, r.left); row.r = Math.max(row.r, r.right); }
        else rows.set(key, { l: r.left, r: r.right });
      }
      let max = 0;
      for (const row of rows.values()) max = Math.max(max, row.r - row.l);
      return max;
    };

    // What counts as overflow depends on whether the box can reflow.
    //
    // A wrapping box must fit its width exactly: a line a tenth of a pixel over the edge
    // still wraps, and the wrapped remainder lands outside whatever panel framed it.
    // Its height also has to fit, with a tolerance for the ascent/descent and
    // `line-height: normal` differences between the two fonts — big enough to ignore
    // metric noise, small enough that a genuine extra line still registers.
    //
    // A no-wrap box cannot gain a line at all, so sub-pixel overhang is invisible and
    // shrinking for it would be pointless. Only a visible amount of overhang matters, and
    // its height is meaningless here — it would just compare the stand-in font's own line
    // box and flag single glyphs like "│" as overflowing.
    const FITS = (probe, t) => t.nowrap
      ? widestLine(probe) <= t.w * 1.02 + 1
      : widestLine(probe) <= t.w && probe.scrollHeight <= t.h * 1.04 + 1;
    // Stop at 80%: below that the text is small enough to be its own defect, and a
    // reported overflow is more useful than silently illegible type.
    const STEPS = [1, 0.97, 0.94, 0.91, 0.88, 0.85, 0.82, 0.8];

    for (const [i, el] of textRoots.entries()) {
      const t = texts[i];
      t.fit = 1;
      if (t.h < 4 || t.w < 4) continue;
      // Nothing is being substituted in this box, so nothing can reflow.
      if (![el, ...el.querySelectorAll('*')].some(subFamily)) continue;

      const probe = el.cloneNode(true);
      // Take the clone out of flow at exactly the geometry PowerPoint will give it.
      // setProperty() only accepts kebab-case names — camelCase is silently ignored.
      const box = {
        position: 'absolute', left: '-99999px', top: '0', visibility: 'hidden',
        width: t.w + 'px', height: 'auto', 'max-height': 'none', 'min-height': '0',
        margin: '0', padding: '0', transform: 'none', animation: 'none',
      };
      // A no-wrap box overflows sideways instead of reflowing, so measure it that way.
      if (t.nowrap) box['white-space'] = 'pre';
      for (const p in box) probe.style.setProperty(p, box[p], 'important');
      el.parentElement.appendChild(probe);
      // Strip what this text box does not carry: nested box roots became their own
      // boxes, and measuring them here would hide the owned text's own overflow.
      // Absorbed inline-flow descendants stay — they are part of these runs.
      for (const nested of [...probe.querySelectorAll('*')]) {
        if (nested.isConnected && isBox(nested) && !absorbed(nested)) nested.remove();
      }

      // Snapshot the authored metrics before rewriting families, then scale from them.
      const nodes = [probe, ...probe.querySelectorAll('*')];
      const base = nodes.map((n) => {
        const cs = getComputedStyle(n);
        return {
          size: parseFloat(cs.fontSize) || 0,
          // Only fixed line-heights and letter-spacings need scaling by hand;
          // unitless/`normal` values already track the font size.
          line: cs.lineHeight === 'normal' ? null : parseFloat(cs.lineHeight) || null,
          spacing: cs.letterSpacing === 'normal' ? null : parseFloat(cs.letterSpacing) || null,
        };
      });
      for (const n of nodes) {
        const to = subFamily(n);
        if (to) n.style.setProperty('font-family', stack(to), 'important');
      }

      for (const s of STEPS) {
        t.fit = s;
        nodes.forEach((n, k) => {
          const b = base[k];
          if (b.size) n.style.setProperty('font-size', b.size * s + 'px', 'important');
          if (b.line) n.style.setProperty('line-height', b.line * s + 'px', 'important');
          if (b.spacing) n.style.setProperty('letter-spacing', b.spacing * s + 'px', 'important');
        });
        t.fitted = FITS(probe, t);
        if (t.fitted) break;
      }
      probe.remove();
    }
  }

  /* ── Hyperlinks ─────────────────────────────────────────── */
  const links = [];
  if (wantLinks) {
    for (const a of slide.querySelectorAll('a[href]')) {
      const href = a.href;
      if (!href || href.startsWith('javascript:') || href.startsWith('#')) continue;
      if (!isVisible(a)) continue;
      const r = norm(a.getBoundingClientRect());
      if (r.w < 1 || r.h < 1) continue;
      links.push({ ...r, url: href, tooltip: (a.textContent || '').trim().slice(0, 100) });
    }
  }

  /* ── Images (editable mode, --native-images) ────────────── */
  const images = [];
  for (const img of slide.querySelectorAll('img[src]')) {
    if (!isVisible(img)) continue;
    const r = norm(img.getBoundingClientRect());
    const cs = getComputedStyle(img);
    images.push({ ...r, src: img.src, alt: img.alt || '', fit: cs.objectFit || 'fill' });
  }

  /* ── Raster regions (editable mode) ───────────────────────
     Mark complex HTML/CSS figures with data-pptx-raster when they should remain as
     cropped slide artwork instead of becoming fragile native text/shapes. */
  const rasters = [];
  let rasterIndex = 0;
  for (const el of slide.querySelectorAll('[data-pptx-raster], [data-pptx-screenshot]')) {
    if (!isVisible(el)) continue;
    const r = norm(el.getBoundingClientRect());
    if (r.w < 1 || r.h < 1) continue;
    const id = `${index}-${rasterIndex++}`;
    el.setAttribute('data-slides-pptx-raster-id', id);
    rasters.push({ ...r, id, alt: (el.getAttribute('aria-label') || el.textContent || 'Figure').replace(/\s+/g, ' ').trim().slice(0, 300) });
  }

  /* ── Speaker notes ──────────────────────────────────────── */
  let notes = '';
  if (wantNotes) {
    const parts = [];
    if (slide.dataset.notes) parts.push(slide.dataset.notes);
    for (const n of slide.querySelectorAll('.speaker-notes, .notes, aside.notes, [data-role="notes"]')) {
      const t = (n.textContent || '').trim();
      if (t) parts.push(t);
    }
    // frontend-slides stores notes from PPT conversion as HTML comments.
    const walker = document.createTreeWalker(slide, NodeFilter.SHOW_COMMENT);
    while (walker.nextNode()) {
      const t = walker.currentNode.nodeValue.trim();
      // Skip section banners like "=== HERO ===" that are structural, not speaker notes.
      if (t && !/^[=\-*\s]*$/.test(t) && !/^={2,}/.test(t)) parts.push(t);
    }
    notes = [...new Set(parts)].join('\n\n').trim();
  }

  /* ── Slide title, for alt text and shape naming ─────────── */
  const heading = slide.querySelector('h1, h2, h3, .slide-title, [class*="title"]');
  const title = (heading?.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 120);

  return { texts, links, images, rasters, notes, title, allText: texts.map((t) => t.plain).filter(Boolean).join(' · ') };
}

/**
 * Hide text for the editable-mode background capture.
 * Preserve the element boxes themselves: backgrounds, borders and shadows are artwork.
 * Only the glyph fill is masked so text containers such as cards do not disappear from
 * the background screenshot.
 */
function toggleTextVisibility(args) {
  const { selector, index, hide, hideImages } = args;
  const slide = document.querySelectorAll(selector)[index];
  if (!slide) return 0;

  // Anything laid out in its parent's inline flow belongs to the parent's text box.
  // inline-block included: a mid-sentence badge is positioned by the surrounding text,
  // so lifting it into its own box makes the parent's words flow through where it sits
  // and the two collide. (A flex/grid child is blockified to `block`, so real flex items
  // still get their own box, which is right — flex positions them independently.)
  const INLINE_FLOW = new Set(['inline', 'inline-block', 'inline-flex', 'inline-grid',
    'inline-table', 'contents', 'ruby', 'ruby-text', 'ruby-base']);
  const isBox = (el) => !INLINE_FLOW.has(getComputedStyle(el).display);
  const isRasterized = (el) => el.closest('[data-pptx-raster], [data-pptx-screenshot]');
  const ownsText = (el) => {
    for (const c of el.childNodes) {
      if (c.nodeType === Node.TEXT_NODE && c.nodeValue.trim()) return true;
      if (c.nodeType === Node.ELEMENT_NODE && !isBox(c) && ownsText(c)) return true;
    }
    return false;
  };

  const MASK_KEY = '__slidesPptxTextMask';
  window[MASK_KEY] ||= new WeakMap();
  const masked = window[MASK_KEY];
  const STYLE_PROPS = ['-webkit-text-fill-color', 'text-shadow'];
  const MEDIA = new Set(['IMG', 'SVG', 'VIDEO', 'CANVAS', 'PICTURE', 'SOURCE']);
  const maskGlyphs = (root) => {
    const targets = [root, ...root.querySelectorAll('*')]
      .filter((node) => !MEDIA.has(node.tagName) && !isRasterized(node));
    for (const node of targets) {
      if (!masked.has(node)) {
        const saved = {};
        for (const prop of STYLE_PROPS) {
          saved[prop] = {
            value: node.style.getPropertyValue(prop),
            priority: node.style.getPropertyPriority(prop),
          };
        }
        masked.set(node, saved);
      }
      node.style.setProperty('-webkit-text-fill-color', 'transparent', 'important');
      node.style.setProperty('text-shadow', 'none', 'important');
    }
  };
  const restoreGlyphs = (root) => {
    const targets = [root, ...root.querySelectorAll('*')]
      .filter((node) => !MEDIA.has(node.tagName) && !isRasterized(node));
    for (const node of targets) {
      const saved = masked.get(node);
      if (!saved) continue;
      for (const prop of STYLE_PROPS) {
        const prev = saved[prop];
        if (prev.value) node.style.setProperty(prop, prev.value, prev.priority);
        else node.style.removeProperty(prop);
      }
      masked.delete(node);
    }
  };
  const preserveGlyphs = (root) => {
    const targets = [root, ...root.querySelectorAll('*')].filter((node) => !MEDIA.has(node.tagName));
    for (const node of targets) {
      if (!masked.has(node)) {
        const saved = {};
        for (const prop of STYLE_PROPS) {
          saved[prop] = {
            value: node.style.getPropertyValue(prop),
            priority: node.style.getPropertyPriority(prop),
          };
        }
        masked.set(node, saved);
      }
      node.style.setProperty('-webkit-text-fill-color', 'currentColor', 'important');
    }
  };

  let n = 0;
  for (const el of slide.querySelectorAll('*')) {
    if (isRasterized(el)) continue;
    if (!isBox(el) || !ownsText(el)) continue;
    if (hide) maskGlyphs(el);
    else restoreGlyphs(el);
    n++;
  }
  for (const el of slide.querySelectorAll('[data-pptx-raster], [data-pptx-screenshot]')) {
    if (hide) preserveGlyphs(el);
    else restoreGlyphs(el);
  }
  if (hideImages) {
    for (const img of slide.querySelectorAll('img[src]')) {
      if (hide) img.style.setProperty('visibility', 'hidden', 'important');
      else img.style.removeProperty('visibility');
    }
  }
  return n;
}

/** Fetch an image through the page so relative/same-origin paths resolve. */
async function fetchAsDataUrl(src) {
  try {
    const res = await fetch(src);
    if (!res.ok) return null;
    const blob = await res.blob();
    return await new Promise((r) => {
      const fr = new FileReader();
      fr.onload = () => r(fr.result);
      fr.onerror = () => r(null);
      fr.readAsDataURL(blob);
    });
  } catch { return null; }
}

/* ═══════════════════════════════════════════════════════════
   CAPTURE
   ═══════════════════════════════════════════════════════════ */

log(`ℹ Mode: ${opts.mode}  ·  serving ${HTML_FILE} on :${port}`);

const browser = await chromium.launch(process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE
  ? { executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE }
  : undefined);
const context = await browser.newContext({
  viewport: { width: 1920, height: 1080 },
  deviceScaleFactor: opts.scale,
});
const page = await context.newPage();

let stage;
const shotDir = opts.keepShots ? resolve(opts.keepShots) : join(tmpdir(), `h2p-${process.pid}`);
mkdirSync(shotDir, { recursive: true });
const slidesData = [];

try {
  await page.goto(`http://localhost:${port}/`, { waitUntil: 'networkidle', timeout: 60000 });
  await page.evaluate(() => document.fonts.ready);

  stage = await page.evaluate(measureStage, opts.selector);
  if (!stage) fail(`No elements matched selector "${opts.selector}". Use --selector to point at your slides.`);
  if (!stage.count) fail(`Found 0 slides with selector "${opts.selector}".`);

  // Resize the viewport to the authored stage so the deck renders at scale 1 —
  // then DOM rects are already in stage coordinates and screenshots are pixel-exact.
  if (stage.width !== 1920 || stage.height !== 1080) {
    await page.setViewportSize({ width: Math.round(stage.width), height: Math.round(stage.height) });
    await page.waitForTimeout(150);
  }

  if (!opts.keepChrome) {
    const hidden = await page.evaluate(hideChrome, opts.selector);
    if (hidden) log(`ℹ hid ${hidden} chrome element(s) outside the slide stage`);
  }

  log(`ℹ ${stage.count} slides at ${Math.round(stage.width)}×${Math.round(stage.height)}`);

  for (let i = 0; i < stage.count; i++) {
    await page.evaluate(showSlide, { selector: opts.selector, index: i });

    // Extract with everything visible, so measurements reflect the real layout.
    const data = await page.evaluate(extractSlide, {
      selector: opts.selector, index: i, wantNotes: opts.notes, wantLinks: opts.links,
      // Only editable mode paints visible text, so only it can visibly overflow.
      wantFit: opts.mode === 'editable' && opts.shrink,
      fontMap: FONT_MAP,
    });

    // In editable mode, pull image bytes before hiding anything.
    if (opts.mode === 'editable' && opts.nativeImages) {
      for (const img of data.images) {
        img.dataUrl = await page.evaluate(fetchAsDataUrl, img.src);
      }
      data.images = data.images.filter((im) => im.dataUrl);
    }

    // Editable mode captures the artwork with the glyphs hidden.
    if (opts.mode === 'editable') {
      await page.evaluate(toggleTextVisibility, {
        selector: opts.selector, index: i, hide: true,
        hideImages: opts.nativeImages,
      });
      await page.waitForTimeout(60);

      for (const [j, raster] of (data.rasters || []).entries()) {
        const handle = await page.$(`[data-slides-pptx-raster-id="${raster.id}"]`);
        if (!handle) continue;
        const rasterShot = join(shotDir, `slide-${String(i + 1).padStart(3, '0')}-raster-${String(j + 1).padStart(2, '0')}.png`);
        await handle.screenshot({ path: rasterShot, type: 'png' });
        raster.shot = rasterShot;
      }
    }

    const ext = opts.jpeg ? 'jpg' : 'png';
    const shot = join(shotDir, `slide-${String(i + 1).padStart(3, '0')}.${ext}`);
    const target = (await page.$('.deck-stage')) || (await page.$$(opts.selector))[i];
    await target.screenshot({
      path: shot,
      ...(opts.jpeg ? { type: 'jpeg', quality: opts.jpegQuality } : { type: 'png' }),
    });

    if (opts.mode === 'editable') {
      await page.evaluate(toggleTextVisibility, {
        selector: opts.selector, index: i, hide: false, hideImages: opts.nativeImages,
      });
    }

    slidesData.push({ ...data, shot });
    log(`  ✓ slide ${i + 1}/${stage.count}${data.notes ? ' (notes)' : ''}`);
  }
} finally {
  await browser.close();
  server.close();
}

/* ═══════════════════════════════════════════════════════════
   BUILD THE PPTX
   Output is PowerPoint's canonical widescreen deck: 13.3333333in
   maps to exactly 12192000 EMU, so the file merges into other
   16:9 decks without triggering a rescale prompt.
   ═══════════════════════════════════════════════════════════ */

const SLIDE_W_IN = 13.3333333;
const SLIDE_H_IN = round4((SLIDE_W_IN * stage.height) / stage.width);
const PX_TO_IN = SLIDE_W_IN / stage.width;
const PX_TO_PT = PX_TO_IN * 72; // exactly 0.5 for a 1920px stage

function round4(n) { return Math.round(n * 10000) / 10000; }
const inch = (px) => round4(px * PX_TO_IN);
const pt = (px) => Math.max(1, Math.round(px * PX_TO_PT * 10) / 10);

const pptx = new PptxGenJS();
pptx.defineLayout({ name: 'FS_WIDESCREEN', width: SLIDE_W_IN, height: SLIDE_H_IN });
pptx.layout = 'FS_WIDESCREEN';
pptx.title = opts.meta.title || slidesData[0]?.title || basename(OUTPUT, '.pptx');
if (opts.meta.author) pptx.author = opts.meta.author;
if (opts.meta.company) pptx.company = opts.meta.company;
if (opts.meta.subject) pptx.subject = opts.meta.subject;

let textBoxCount = 0, linkCount = 0, notesCount = 0, nativeImageCount = 0, shrunkBoxes = 0;
const overflowing = [];

for (const [i, s] of slidesData.entries()) {
  const slide = pptx.addSlide();

  /* Background artwork — full-bleed, and in image/searchable mode this IS the slide. */
  const altBase = s.title ? `Slide ${i + 1}: ${s.title}` : `Slide ${i + 1}`;
  if (opts.backgroundArt) {
    slide.addImage({
      path: s.shot,
      x: 0, y: 0, w: SLIDE_W_IN, h: SLIDE_H_IN,
      // Screen readers announce alt text; for an image-only slide it is the only
      // machine-readable content, so include the slide's words.
      altText: opts.mode === 'image'
        ? [altBase, s.allText].filter(Boolean).join(' — ').slice(0, 2000)
        : `${altBase} (background)`,
    });
  }

  /* Cropped raster figures — preserves complex HTML/CSS charts without setting a
     manual slide background or duplicating the editable text layer. */
  if (opts.mode === 'editable') {
    for (const raster of s.rasters || []) {
      if (!raster.shot) continue;
      slide.addImage({
        path: raster.shot,
        x: inch(raster.x), y: inch(raster.y), w: inch(raster.w), h: inch(raster.h),
        altText: raster.alt || 'Figure',
      });
    }
  }

  /* Native images lifted out of the artwork (editable mode only). */
  if (opts.mode === 'editable' && opts.nativeImages) {
    for (const im of s.images) {
      slide.addImage({
        data: im.dataUrl,
        x: inch(im.x), y: inch(im.y), w: inch(im.w), h: inch(im.h),
        altText: im.alt || 'Image',
        ...(im.fit === 'cover' ? { sizing: { type: 'cover', w: inch(im.w), h: inch(im.h) } } : {}),
      });
      nativeImageCount++;
    }
  }

  /* Text layer — invisible in searchable mode, visible in editable mode. */
  if (opts.mode !== 'image') {
    const invisible = opts.mode === 'searchable';
    for (const t of s.texts) {
      // The measured scale that keeps the substituted font inside the authored box.
      const fit = t.fit && t.fit < 1 ? t.fit : 1;
      if (fit < 1) shrunkBoxes++;
      // Shrinking bottomed out and the text still does not fit. Naming it is more use
      // than shrinking further, which only trades overflow for unreadable type.
      if (t.fitted === false) overflowing.push({ slide: i + 1, text: t.plain.slice(0, 60) });

      // Group runs into lines first. Setting `breakLine` on "the previous run" as the
      // markers arrive loses blank lines (two breaks in a row would collapse), so lines
      // are materialised and the break is put on each line's last run explicitly.
      const lines = [[]];
      for (const r of t.runs) {
        if (r.break) lines.push([]);
        else lines[lines.length - 1].push(r);
      }

      const runs = [];
      lines.forEach((line, li) => {
        const notLast = li < lines.length - 1;
        const paragraphBullet = t.listType
          ? (t.listType === 'number' ? { type: 'number' } : true)
          : undefined;
        if (!line.length) {
          // An empty line still needs a run to carry the paragraph, or the blank
          // line vanishes and the following text moves up.
          runs.push({ text: '', options: { breakLine: notLast, fontSize: pt(12), bullet: paragraphBullet } });
          return;
        }
        line.forEach((r, ri) => {
          const st = r.style;
          runs.push({
            text: r.text,
            options: {
              fontFace: mapFont(st.font),
              fontSize: pt(st.sizePx * fit),
              bold: st.bold,
              italic: st.italic,
              underline: st.underline ? { style: 'sng' } : undefined,
              color: st.color,
              // Alpha and invisibility both ride on the same transparency channel.
              transparency: invisible ? 100 : (st.alpha < 1 ? Math.round((1 - st.alpha) * 100) : undefined),
              charSpacing: st.letterSpacingPx ? round4(st.letterSpacingPx * fit * PX_TO_PT) : undefined,
              breakLine: notLast && ri === line.length - 1,
              bullet: t.listType && ri === 0 ? paragraphBullet : undefined,
            },
          });
        });
      });
      if (!runs.length) continue;

      slide.addText(runs, {
        x: inch(t.x), y: inch(t.y), w: inch(t.w), h: inch(t.h),
        align: t.align === 'justify' ? 'justify' : t.align,
        valign: 'top',
        margin: t.listType ? 0.04 : 0,          // PowerPoint's default inset would shift every box
        wrap: !t.nowrap,    // single-line text must not reflow on a different renderer
        autoFit: false,     // keep the measured size; never let PowerPoint rescale silently
        lineSpacing: t.lineHeightPx ? round4(t.lineHeightPx * fit * PX_TO_PT) : undefined,
      });
      textBoxCount++;
    }
  }

  /* Hyperlink hotspots — invisible rectangles that keep links clickable. */
  for (const l of s.links) {
    slide.addShape(pptx.ShapeType.rect, {
      x: inch(l.x), y: inch(l.y), w: inch(l.w), h: inch(l.h),
      fill: { color: 'FFFFFF', transparency: 100 },
      line: { width: 0 },
      hyperlink: { url: l.url, tooltip: l.tooltip || l.url },
    });
    linkCount++;
  }

  if (s.notes) { slide.addNotes(s.notes); notesCount++; }
}

await pptx.writeFile({ fileName: OUTPUT });

/* ═══════════════════════════════════════════════════════════
   SCHEMA REPAIR
   pptxgenjs 4.x emits two constructs that ECMA-376 forbids. PowerPoint is
   lenient about both, but strict validators, Keynote and Google Slides
   are not — so fix them in the written package.

   1. CT_Presentation fixes child order as sldMasterIdLst -> notesMasterIdLst
      -> handoutMasterIdLst -> sldIdLst -> sldSz. pptxgenjs writes
      notesMasterIdLst after sldIdLst.
   2. CT_TextParagraph allows at most one <a:pPr>, as the first child.
      pptxgenjs repeats it before every run in a multi-run paragraph.
   3. pptxgenjs points the notes master at the same theme part used by the
      presentation/slide master. PowerPoint repairs this by creating a dedicated
      notes theme part and retargeting the notes-master relationship.
   4. pptxgenjs emits empty text runs in blank notes placeholders. PowerPoint
      removes these during repair/save, so remove them before handoff.
   5. pptxgenjs can leave stale content-type overrides behind for parts that no
      longer exist after master/layout de-duplication. PowerPoint repairs those.
   6. Zip directory entries are harmless to most readers but are noise in a PPTX
      package and PowerPoint rewrites them during repair.
   ═══════════════════════════════════════════════════════════ */

const zip = await JSZip.loadAsync(await readFile(OUTPUT));
const repairs = [];

// --- 1. Presentation element order ---
let presXml = await zip.file('ppt/presentation.xml').async('string');
const notesLst = presXml.match(/<p:notesMasterIdLst>[\s\S]*?<\/p:notesMasterIdLst>/);
if (notesLst && presXml.indexOf(notesLst[0]) > presXml.indexOf('<p:sldIdLst>')) {
  presXml = presXml.replace(notesLst[0], '').replace('<p:sldIdLst>', notesLst[0] + '<p:sldIdLst>');
  zip.file('ppt/presentation.xml', presXml);
  repairs.push('presentation element order');
}

// --- 2. Duplicate paragraph properties ---
// Within one <a:p>, keep the first <a:pPr> and drop the repeats. The duplicates
// pptxgenjs writes are byte-identical, so the paragraph's formatting is preserved.
const PPR = /<a:pPr(?:[^>]*\/>|[^>]*>[\s\S]*?<\/a:pPr>)/g;
let dupParagraphs = 0;
for (const name of Object.keys(zip.files)) {
  if (!/^ppt\/(slides|notesSlides|slideLayouts|slideMasters)\/[^/]+\.xml$/.test(name)) continue;
  const xml = await zip.file(name).async('string');
  let changed = false;
  const fixed = xml.replace(/<a:p>[\s\S]*?<\/a:p>/g, (para) => {
    const matches = para.match(PPR);
    if (!matches || matches.length < 2) return para;
    let seen = false;
    const out = para.replace(PPR, (m) => (seen ? '' : ((seen = true), m)));
    changed = true;
    dupParagraphs++;
    return out;
  });
  if (changed) zip.file(name, fixed);
}
if (dupParagraphs) repairs.push(`${dupParagraphs} duplicate <a:pPr>`);

// --- 3. Dedicated notes-master theme ---
// PowerPoint-created decks keep the notes master on its own theme part. pptxgenjs
// reuses theme1.xml, which validates against ECMA-376 but triggers PowerPoint's
// repair pass on some Office builds. Mirroring PowerPoint's package shape avoids
// the repair prompt while keeping the same visual theme.
let contentTypesXml = await zip.file('[Content_Types].xml')?.async('string');
const THEME_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';
const hasOverride = (xml, partName) => new RegExp(`<Override\\s+PartName="/${partName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"\\s+ContentType="[^"]+"\\s*/>`).test(xml);
const ensureOverride = (xml, partName, contentType) => {
  const escaped = partName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`<Override\\s+PartName="/${escaped}"\\s+ContentType="[^"]+"\\s*/>`);
  const entry = `<Override PartName="/${partName}" ContentType="${contentType}"/>`;
  if (re.test(xml)) return xml.replace(re, entry);
  return xml.replace('</Types>', `  ${entry}\n</Types>`);
};

const notesMasterRelsPath = 'ppt/notesMasters/_rels/notesMaster1.xml.rels';
let notesMasterRelsXml = await zip.file(notesMasterRelsPath)?.async('string');
if (notesMasterRelsXml && zip.file('ppt/theme/theme1.xml')) {
  const notesThemeRel = notesMasterRelsXml.match(/<Relationship\b[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/theme"[^>]*Target="([^"]+)"[^>]*\/>/);
  if (notesThemeRel?.[1] === '../theme/theme1.xml') {
    if (!zip.file('ppt/theme/theme2.xml')) {
      zip.file('ppt/theme/theme2.xml', await zip.file('ppt/theme/theme1.xml').async('string'));
    }
    notesMasterRelsXml = notesMasterRelsXml.replace(notesThemeRel[0], notesThemeRel[0].replace('Target="../theme/theme1.xml"', 'Target="../theme/theme2.xml"'));
    zip.file(notesMasterRelsPath, notesMasterRelsXml);
    if (contentTypesXml && !hasOverride(contentTypesXml, 'ppt/theme/theme2.xml')) {
      contentTypesXml = ensureOverride(contentTypesXml, 'ppt/theme/theme2.xml', THEME_CONTENT_TYPE);
      zip.file('[Content_Types].xml', contentTypesXml);
    }
    repairs.push('dedicated notes-master theme');
  }
}

// --- 4. Empty notes text runs ---
const EMPTY_TEXT_RUN = /<a:r>\s*<a:rPr(?:\s[^>]*)?\/>\s*<a:t(?:\s*\/|><\/a:t>)\s*<\/a:r>/g;
let emptyNotesRuns = 0;
for (const name of Object.keys(zip.files)) {
  if (!/^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name)) continue;
  const xml = await zip.file(name).async('string');
  const matches = xml.match(EMPTY_TEXT_RUN);
  if (!matches) continue;
  zip.file(name, xml.replace(EMPTY_TEXT_RUN, ''));
  emptyNotesRuns += matches.length;
}
if (emptyNotesRuns) repairs.push(`${emptyNotesRuns} empty notes text run${emptyNotesRuns === 1 ? '' : 's'}`);

// --- 5. Stale content-type overrides ---
contentTypesXml = await zip.file('[Content_Types].xml')?.async('string');
if (contentTypesXml) {
  let staleOverrides = 0;
  const fixed = contentTypesXml.replace(/<Override\s+PartName="([^"]+)"\s+ContentType="[^"]+"\s*\/>/g, (entry, partName) => {
    const zipName = partName.replace(/^\//, '');
    if (zip.file(zipName)) return entry;
    staleOverrides++;
    return '';
  });
  if (staleOverrides) {
    zip.file('[Content_Types].xml', fixed);
    repairs.push(`${staleOverrides} stale content-type override${staleOverrides === 1 ? '' : 's'}`);
  }
}

// --- 6. Directory entries ---
let directoryEntries = 0;
for (const [name, entry] of Object.entries(zip.files)) {
  if (!entry.dir) continue;
  // JSZip#remove('ppt/') recursively removes all children. We only want to drop
  // the explicit directory record from the central directory, not the files below it.
  delete zip.files[name];
  directoryEntries++;
}
if (directoryEntries) repairs.push(`${directoryEntries} zip director${directoryEntries === 1 ? 'y' : 'ies'}`);

const repaired = repairs.length > 0;
if (repaired) {
  await writeFile(OUTPUT, await zip.generateAsync({
    type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 },
  }));
}

/* ═══════════════════════════════════════════════════════════
   REPORT
   ═══════════════════════════════════════════════════════════ */

if (!opts.keepShots) rmSync(shotDir, { recursive: true, force: true });

const sizeMB = (readFileSync(OUTPUT).length / 1024 / 1024).toFixed(1);
const summary = {
  output: OUTPUT,
  mode: opts.mode,
  slides: slidesData.length,
  stage: { width: stage.width, height: stage.height },
  slideSizeInches: { w: SLIDE_W_IN, h: SLIDE_H_IN },
  textBoxes: textBoxCount,
  shrunkTextBoxes: shrunkBoxes,
  overflowingTextBoxes: overflowing,
  hyperlinks: linkCount,
  notes: notesCount,
  nativeImages: nativeImageCount,
  fontSubstitutions: Object.fromEntries(substitutions),
  schemaRepairs: repairs,
  sizeMB: Number(sizeMB),
  screenshots: opts.keepShots ? shotDir : null,
};

if (opts.json) {
  console.log(JSON.stringify(summary, null, 2));
} else {
  console.log('');
  console.log(`✓ ${OUTPUT}`);
  console.log(`  ${summary.slides} slides · ${sizeMB} MB · ${SLIDE_W_IN}×${SLIDE_H_IN}in`);
  if (textBoxCount) console.log(`  ${textBoxCount} text boxes${opts.mode === 'searchable' ? ' (invisible search layer)' : ''}`);
  if (shrunkBoxes) console.log(`  ${shrunkBoxes} of them shrunk to fit the substituted font (--no-shrink to keep authored sizes)`);
  if (overflowing.length) {
    console.log(`  ⚠ ${overflowing.length} still overflow at the ${Math.round(0.8 * 100)}% floor — check these:`);
    for (const o of overflowing.slice(0, 8)) console.log(`      slide ${o.slide}: ${JSON.stringify(o.text)}`);
    if (overflowing.length > 8) console.log(`      … and ${overflowing.length - 8} more`);
  }
  if (nativeImageCount) console.log(`  ${nativeImageCount} native images`);
  if (linkCount) console.log(`  ${linkCount} hyperlinks`);
  if (notesCount) console.log(`  ${notesCount} slides with speaker notes`);
  if (repaired) console.log(`  schema: repaired ${repairs.join(', ')}`);
  if (substitutions.size) {
    console.log('  font substitutions (webfonts are not installed on viewers\' machines):');
    for (const [from, to] of substitutions) console.log(`    ${from} → ${to}`);
  }
}
