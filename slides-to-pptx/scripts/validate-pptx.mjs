#!/usr/bin/env node
/**
 * validate-pptx.mjs — Verify a generated .pptx is structurally sound, schema-clean,
 * and actually contains the slides you think it does.
 *
 * Usage:
 *   node validate-pptx.mjs <file.pptx> [options]
 *
 * Options:
 *   --expect-slides <n>    Fail if the slide count differs
 *   --expect-text          Fail if any slide has no text runs (searchable/editable modes)
 *   --no-schema            Skip the OOXML schema pass (which needs npx + network once)
 *   --json                 Emit machine-readable results
 *
 * Checks, in order of how much they actually catch:
 *   1. Package structure   — zip integrity, required parts, relationship targets resolve,
 *                            and PowerPoint-repair-prone notes/theme wiring is absent
 *   2. Deck geometry       — slide size in EMU, 16:9 aspect, canonical widescreen match
 *   3. Slide inventory     — slide parts vs sldIdLst, media per slide, notes, text runs
 *   4. Blank-art heuristic — flags slide images small enough to be a flat blank frame
 *   5. OOXML schema        — ECMA-376 validation via @xarsh/ooxml-validator
 *
 * Layout correctness is a separate question — see scripts/inspect-pptx.mjs.
 */

import { readFile } from 'fs/promises';
import { existsSync } from 'fs';
import { execFileSync } from 'child_process';
import { resolve, join, posix as posixPath } from 'path';
import { createRequire } from 'module';

// See html2pptx.mjs — SLIDES_PPTX_DEPS points at the shared dependency cache.
let JSZip;
try {
  const depsRoot = process.env.SLIDES_PPTX_DEPS;
  JSZip = depsRoot ? createRequire(join(depsRoot, 'index.js'))('jszip') : (await import('jszip')).default;
} catch (e) {
  console.error(`✗ Missing dependency jszip: ${e.message}`);
  console.error('  Install with: npm install jszip');
  process.exit(2);
}

/* ── CLI ────────────────────────────────────────────────── */

const argv = process.argv.slice(2);
const opts = { schema: true, json: false, expectSlides: null, expectText: false };
const positional = [];
for (let i = 0; i < argv.length; i++) {
  const a = argv[i];
  if (a === '--expect-slides') opts.expectSlides = Number(argv[++i]);
  else if (a === '--expect-text') opts.expectText = true;
  else if (a === '--no-schema') opts.schema = false;
  else if (a === '--json') opts.json = true;
  else if (a.startsWith('--')) { console.error(`Unknown option: ${a}`); process.exit(2); }
  else positional.push(a);
}
if (!positional.length) {
  console.error('Usage: node validate-pptx.mjs <file.pptx> [--expect-slides N] [--expect-text] [--no-schema]');
  process.exit(2);
}
const FILE = resolve(positional[0]);
if (!existsSync(FILE)) { console.error(`✗ File not found: ${FILE}`); process.exit(2); }

const problems = [];   // hard failures
const warnings = [];   // worth a look, not fatal
const info = {};
const fail = (m) => problems.push(m);
const warn = (m) => warnings.push(m);

/* ── 1. Package structure ───────────────────────────────── */

let zip;
try {
  zip = await JSZip.loadAsync(await readFile(FILE));
} catch (e) {
  console.error(`✗ Not a readable zip/OOXML package: ${e.message}`);
  process.exit(1);
}

const names = Object.keys(zip.files);
const REQUIRED = ['[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels'];
for (const part of REQUIRED) if (!names.includes(part)) fail(`Missing required part: ${part}`);

const read = async (p) => (zip.file(p) ? zip.file(p).async('string') : null);
const presXml = await read('ppt/presentation.xml');
if (!presXml) { console.error('✗ ppt/presentation.xml missing — not a PowerPoint package'); process.exit(1); }

const presRels = (await read('ppt/_rels/presentation.xml.rels')) || '';
// Every r:id referenced by the presentation must resolve to a declared relationship.
const relIds = new Set([...presRels.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]));
for (const m of presXml.matchAll(/r:id="([^"]+)"/g)) {
  if (!relIds.has(m[1])) fail(`presentation.xml references undeclared relationship ${m[1]}`);
}
// Presentation relationship targets must exist in the package.
for (const m of presRels.matchAll(/Target="([^"]+)"[^>]*?(TargetMode="External")?\/>/g)) {
  if (m[2]) continue;
  const target = m[1].replace(/^\.\.\//, '').replace(/^\//, '');
  const full = target.startsWith('ppt/') ? target : `ppt/${target}`;
  if (!names.includes(full)) fail(`Relationship target missing from package: ${full}`);
}

// Every internal relationship in the package should resolve, not just the ones
// from presentation.xml. Also catch the PowerPoint repair seen in Office: notes
// masters that share the slide/presentation theme part and blank notes text runs.
const relRecords = [];
const attr = (tag, name) => tag.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1] || '';
const relSourceDir = (relName) => {
  if (relName === '_rels/.rels') return '';
  const marker = '/_rels/';
  const idx = relName.indexOf(marker);
  return idx === -1 ? posixPath.dirname(relName) : relName.slice(0, idx);
};
const normalizeTarget = (relName, target) => {
  if (/^[a-z][a-z0-9+.-]*:/i.test(target)) return target;
  return posixPath.normalize(posixPath.join(relSourceDir(relName), target)).replace(/^\.\//, '');
};

for (const relName of names.filter((n) => n.endsWith('.rels'))) {
  const relXml = (await read(relName)) || '';
  const ids = new Set();
  for (const m of relXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const tag = m[0];
    const id = attr(tag, 'Id');
    const type = attr(tag, 'Type');
    const target = attr(tag, 'Target');
    const mode = attr(tag, 'TargetMode');
    if (id) {
      if (ids.has(id)) fail(`${relName} declares duplicate relationship id ${id}`);
      ids.add(id);
    }
    if (!target || mode === 'External') continue;
    const full = normalizeTarget(relName, target);
    relRecords.push({ relName, id, type, target, full });
    if (!names.includes(full)) fail(`${relName} targets missing package part: ${full}`);
  }
}

const contentTypesXml = (await read('[Content_Types].xml')) || '';
const hasContentTypeOverride = (partName, contentType) => new RegExp(`<Override\\s+PartName="/${partName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"\\s+ContentType="${contentType.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"\\s*/>`).test(contentTypesXml);
const THEME_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml';
for (const themePart of names.filter((n) => /^ppt\/theme\/theme\d+\.xml$/.test(n))) {
  if (!hasContentTypeOverride(themePart, THEME_CONTENT_TYPE)) {
    fail(`${themePart} is missing its theme content-type override`);
  }
}

const sharedThemeOwners = new Set(
  relRecords
    .filter((r) => r.type.endsWith('/theme') && (r.relName === 'ppt/_rels/presentation.xml.rels' || r.relName.startsWith('ppt/slideMasters/_rels/')))
    .map((r) => r.full)
);
for (const r of relRecords.filter((rec) => rec.type.endsWith('/theme') && rec.relName.startsWith('ppt/notesMasters/_rels/'))) {
  if (sharedThemeOwners.has(r.full)) {
    fail(`${r.relName} shares theme part ${r.full} with the presentation/slide master; PowerPoint repairs this by adding a dedicated notes theme`);
  }
}

const EMPTY_NOTES_RUN = /<a:r>\s*<a:rPr(?:\s[^>]*)?\/>\s*<a:t(?:\s*\/|><\/a:t>)\s*<\/a:r>/;
for (const notesPart of names.filter((n) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n))) {
  if (EMPTY_NOTES_RUN.test((await read(notesPart)) || '')) {
    fail(`${notesPart} contains an empty notes text run that PowerPoint removes during repair`);
  }
}

/* ── 2. Deck geometry ───────────────────────────────────── */

const EMU_PER_IN = 914400;
const sz = presXml.match(/<p:sldSz\s+cx="(\d+)"\s+cy="(\d+)"/);
if (!sz) fail('presentation.xml has no <p:sldSz> — slide size undefined');
else {
  const cx = Number(sz[1]), cy = Number(sz[2]);
  const aspect = cx / cy;
  info.slideSize = {
    emu: { cx, cy },
    inches: { w: +(cx / EMU_PER_IN).toFixed(4), h: +(cy / EMU_PER_IN).toFixed(4) },
    aspect: +aspect.toFixed(4),
    canonicalWidescreen: cx === 12192000 && cy === 6858000,
  };
  if (Math.abs(aspect - 16 / 9) > 0.01) warn(`Slide aspect is ${aspect.toFixed(3)}, not 16:9 (1.778)`);
  else if (!info.slideSize.canonicalWidescreen) {
    warn(`Slide size ${cx}×${cy} EMU is 16:9 but not PowerPoint's canonical 12192000×6858000 — merging into another deck may prompt to rescale`);
  }
  // CT_Presentation fixes child order; notesMasterIdLst must precede sldIdLst.
  const iNotes = presXml.indexOf('<p:notesMasterIdLst>');
  const iSlides = presXml.indexOf('<p:sldIdLst>');
  if (iNotes !== -1 && iSlides !== -1 && iNotes > iSlides) {
    fail('presentation.xml element order violates CT_Presentation: <p:notesMasterIdLst> must precede <p:sldIdLst>');
  }
}

/* ── 3. Slide inventory ─────────────────────────────────── */

const slideParts = names.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
  .sort((a, b) => Number(a.match(/(\d+)/)[1]) - Number(b.match(/(\d+)/)[1]));
const sldIdCount = [...presXml.matchAll(/<p:sldId\s/g)].length;

info.slideCount = slideParts.length;
if (slideParts.length === 0) fail('Package contains no slide parts');
if (sldIdCount !== slideParts.length) {
  fail(`Slide list mismatch: ${sldIdCount} entries in sldIdLst but ${slideParts.length} slide parts`);
}
if (opts.expectSlides != null && slideParts.length !== opts.expectSlides) {
  fail(`Expected ${opts.expectSlides} slides, found ${slideParts.length}`);
}

/** Minimal PNG/JPEG header parsing — enough to confirm the art is the right size. */
function imageSize(buf) {
  if (buf.length > 24 && buf.readUInt32BE(0) === 0x89504e47) {
    return { type: 'png', w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
  }
  if (buf.length > 4 && buf[0] === 0xff && buf[1] === 0xd8) {
    let o = 2;
    while (o + 9 < buf.length) {
      if (buf[o] !== 0xff) { o++; continue; }
      const marker = buf[o + 1];
      // SOF0-SOF15, excluding the non-frame markers DHT(c4)/JPG(c8)/DAC(cc)
      if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) {
        return { type: 'jpeg', h: buf.readUInt16BE(o + 5), w: buf.readUInt16BE(o + 7) };
      }
      o += 2 + buf.readUInt16BE(o + 2);
    }
  }
  return null;
}

// pptxgenjs emits a notesSlide part for every slide, so count the ones that
// actually carry text rather than the raw part count.
const notesParts = names.filter((n) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n));
let notesWithText = 0;
for (const n of notesParts) {
  const xml = (await read(n)) || '';
  // The slide-number placeholder contributes a run on every notes page; ignore
  // pages whose only text is that numeric field.
  const body = [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]).join('').trim();
  if (body && !/^\d+$/.test(body)) notesWithText++;
}
const slideReport = [];
let totalTextRuns = 0;

for (const [i, part] of slideParts.entries()) {
  const xml = await read(part);
  const rels = (await read(part.replace('slides/', 'slides/_rels/') + '.rels')) || '';

  // Text runs
  const runs = [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]);
  const textChars = runs.join('').replace(/&[a-z]+;/g, ' ').trim().length;
  totalTextRuns += runs.length;

  // Media referenced by this slide
  const media = [...rels.matchAll(/Target="\.\.\/media\/([^"]+)"/g)].map((m) => m[1]);
  const mediaInfo = [];
  for (const m of media) {
    const f = zip.file(`ppt/media/${m}`);
    if (!f) { fail(`slide ${i + 1}: references missing media ppt/media/${m}`); continue; }
    const buf = await f.async('nodebuffer');
    const dim = imageSize(buf);
    mediaInfo.push({ name: m, bytes: buf.length, ...(dim || {}) });
    if (buf.length === 0) fail(`slide ${i + 1}: media ${m} is empty`);
    // 4. Blank-art heuristic: a flat single-colour frame compresses to almost nothing.
    //    Only applied to full-slide artwork — a small logo is legitimately tiny.
    else if (dim && dim.w >= 1000 && buf.length < 8192) {
      warn(`slide ${i + 1}: ${dim.w}×${dim.h} artwork ${m} is only ${buf.length} bytes — likely a blank or near-blank capture`);
    }
  }

  // CT_TextParagraph permits one <a:pPr>, first. Checked directly so the bug is
  // caught even when the schema pass is skipped or offline.
  const PPR = /<a:pPr(?:[^>]*\/>|[^>]*>[\s\S]*?<\/a:pPr>)/g;
  for (const para of xml.match(/<a:p>[\s\S]*?<\/a:p>/g) || []) {
    if ((para.match(PPR) || []).length > 1) {
      fail(`slide ${i + 1}: a paragraph has multiple <a:pPr> elements (CT_TextParagraph allows one, first)`);
      break;
    }
  }

  const hasNotes = /Target="\.\.\/notesSlides\//.test(rels);
  const pics = [...xml.matchAll(/<p:pic>/g)].length;
  const links = [...xml.matchAll(/<a:hlinkClick/g)].length;

  if (pics === 0 && textChars === 0) fail(`slide ${i + 1}: no pictures and no text — slide is empty`);
  if (opts.expectText && textChars === 0) fail(`slide ${i + 1}: no text runs (expected with --expect-text)`);

  slideReport.push({ slide: i + 1, pictures: pics, textRuns: runs.length, textChars, hyperlinks: links, notes: hasNotes, media: mediaInfo });
}

info.slides = slideReport;
info.notesSlides = notesWithText;
info.totalTextRuns = totalTextRuns;

// Slide art should match the deck aspect; a mismatch means letterboxing in PowerPoint.
if (info.slideSize) {
  for (const s of slideReport) {
    const art = s.media.find((m) => m.w && m.h);
    if (!art) continue;
    const artAspect = art.w / art.h;
    if (Math.abs(artAspect - info.slideSize.aspect) > 0.02) {
      warn(`slide ${s.slide}: artwork aspect ${artAspect.toFixed(3)} differs from slide aspect ${info.slideSize.aspect} — expect distortion or bars`);
    }
  }
}

/* ── 5. OOXML schema validation ─────────────────────────── */

if (opts.schema) {
  try {
    const out = execFileSync('npx', ['-y', '@xarsh/ooxml-validator', FILE], {
      encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], timeout: 180000,
    });
    const res = JSON.parse(out);
    info.schema = { ok: res.ok, errors: res.errors || [] };
    for (const e of res.errors || []) fail(`schema: ${e.path} ${e.xPath || ''} — ${e.description}`);
  } catch (e) {
    // Exit code 1 still carries a valid JSON report on stdout.
    const stdout = e.stdout?.toString?.() || '';
    if (stdout.trim().startsWith('{')) {
      try {
        const res = JSON.parse(stdout);
        info.schema = { ok: res.ok, errors: res.errors || [] };
        for (const er of res.errors || []) fail(`schema: ${er.path} ${er.xPath || ''} — ${er.description}`);
      } catch { warn(`schema validator returned unparseable output; skipped`); }
    } else {
      info.schema = { skipped: true, reason: e.message.split('\n')[0] };
      warn('schema validation skipped (needs npx and one-time network access) — run with --no-schema to silence');
    }
  }
}

/* Layout correctness — whether text sits where it should — is checked by
   scripts/inspect-pptx.mjs, which reads the emitted geometry and renders a preview
   in the same headless Chromium the converter already uses. */

/* ── Report ─────────────────────────────────────────────── */

const ok = problems.length === 0;
if (opts.json) {
  console.log(JSON.stringify({ file: FILE, ok, problems, warnings, ...info }, null, 2));
} else {
  console.log('');
  console.log(`  ${FILE}`);
  console.log(`  ${info.slideCount} slides · ${info.notesSlides} notes pages · ${totalTextRuns} text runs`);
  if (info.slideSize) {
    console.log(`  ${info.slideSize.inches.w}×${info.slideSize.inches.h}in (${info.slideSize.emu.cx}×${info.slideSize.emu.cy} EMU)` +
      (info.slideSize.canonicalWidescreen ? ' — canonical widescreen' : ''));
  }
  if (info.schema?.ok) console.log('  schema: valid (ECMA-376)');
  else if (info.schema?.skipped) console.log('  schema: skipped');
  console.log('');
  for (const w of warnings) console.log(`  ⚠ ${w}`);
  for (const p of problems) console.log(`  ✗ ${p}`);
  console.log('');
  console.log(ok ? '  ✓ PASS' : `  ✗ FAIL — ${problems.length} problem(s)`);
  console.log('');
}

process.exit(ok ? 0 : 1);
