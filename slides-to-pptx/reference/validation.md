# Validation

```bash
node scripts/validate-pptx.mjs deck.pptx [--expect-slides N] [--expect-text] [--no-schema] [--json]
```

Exit code `0` = pass, `1` = problems found, `2` = bad usage or unreadable file.
`--json` emits the full report for scripting.

## Why validate at all

A `.pptx` is a zip of XML parts bound by relationship files. It is easy to produce one that
PowerPoint opens (PowerPoint is famously lenient) but that Google Slides or Keynote
reject or render wrong. Two of the three schema bugs this validator catches are produced by
**pptxgenjs itself**, on every file it writes.

## The checks

### 1. Package structure

Zip integrity, the four required parts (`[Content_Types].xml`, `_rels/.rels`,
`ppt/presentation.xml`, `ppt/_rels/presentation.xml.rels`), every `r:id` resolving to a declared
relationship, every internal relationship target existing in the package, and package patterns
that make desktop PowerPoint open with a repair prompt.

**Catches:** truncated downloads, broken hand-edits, missing media, missing theme content types,
shared notes-master themes, and empty notes text runs.

### 2. Deck geometry

Slide size in EMU, 16:9 aspect, and whether it matches PowerPoint's canonical widescreen
`12192000 × 6858000` EMU.

**Why canonical matters:** a 16:9 deck that is not *exactly* 12192000 EMU wide will prompt
PowerPoint to rescale when someone merges its slides into another deck. The converter uses
`13.3333333` inches specifically because it rounds to exactly 12192000 EMU — `13.333` gives
12191695, which is off by enough to trigger the prompt.

Also checks the CT_Presentation child order (see Known bugs).

### 3. Slide inventory

Slide parts vs `sldIdLst` entries, pictures per slide, text runs per slide, hyperlinks, speaker
notes, and the pixel dimensions of every referenced image.

- `--expect-slides N` fails if the count differs from the source deck.
- `--expect-text` fails if any slide has no text runs — use it for `searchable` and `editable`
  mode, where a slide without text means extraction silently failed. The wrapper adds it
  automatically for those modes.
- A slide with **neither** pictures nor text always fails: it is blank.

### 4. Blank-artwork heuristic

A full-slide image (≥1000px wide) that compresses to under 8 KB is almost certainly a flat
blank frame — the usual sign that a slide never became visible before capture.

Reported as a **warning**, because a genuinely minimal slide (solid background, no content) is
legitimate. Check the slide before dismissing it.

### 5. OOXML schema validation

Runs `npx @xarsh/ooxml-validator`, a standalone port of Microsoft's Open XML SDK validation
logic — no .NET required. Validates against the ECMA-376 schemas and reports the exact part,
XPath and rule for each violation.

Needs `npx` and one-time network access to fetch the package. If unavailable, the check is
skipped with a warning rather than failing; `--no-schema` silences it. The duplicate-`<a:pPr>`
bug is *also* checked directly in pure JS, so it is caught even offline.

## What this does *not* check

Validation proves the package is well-formed. It says nothing about whether the content landed
where it should — text can be schema-perfect and still sit off the slide or on top of its
neighbour. That is [inspection](inspection.md)'s job, and the wrapper runs it straight after
this.

## Known pptxgenjs bugs, repaired automatically

The converter repairs these in the written package and reports what it fixed. They are present in
pptxgenjs 4.0.1 and affect most or all files it produces.

### `<p:notesMasterIdLst>` in the wrong position

ECMA-376 `CT_Presentation` fixes the child order:

```
sldMasterIdLst → notesMasterIdLst → handoutMasterIdLst → sldIdLst → sldSz → notesSz → …
```

pptxgenjs writes `notesMasterIdLst` *after* `sldIdLst`. The fix moves it before.

### Duplicate `<a:pPr>` in multi-run paragraphs

`CT_TextParagraph` permits at most one `<a:pPr>`, as the first child. When a paragraph has
several formatting runs — any text containing `<strong>`, `<em>`, a link, or a colour change —
pptxgenjs emits a fresh identical `<a:pPr>` before each run:

```xml
<a:p>
  <a:pPr .../><a:r>…</a:r>
  <a:pPr .../><a:r>…</a:r>   <!-- invalid -->
</a:p>
```

The fix keeps the first and drops the repeats. Since the duplicates are byte-identical, the
paragraph's formatting is unchanged.

### Notes master shares `theme1.xml`

pptxgenjs wires `ppt/notesMasters/_rels/notesMaster1.xml.rels` to `../theme/theme1.xml`, the
same theme part used by the slide master and presentation. Some desktop PowerPoint builds pass
schema validation but still repair the package on open, creating `ppt/theme/theme2.xml` and
retargeting the notes master to that dedicated theme. The converter now mirrors that package
shape up front and adds the required theme content-type override.

### Empty notes text runs

pptxgenjs emits empty `<a:r><a:t/></a:r>` runs in blank notes placeholders. PowerPoint removes
them during repair/save. The converter removes them before writing the final file so a deck with
speaker-note parts opens cleanly.

## Interpreting a failure

| Message | Meaning | Fix |
| ------- | ------- | --- |
| `Not a readable zip/OOXML package` | File is truncated or not a pptx | Re-run the conversion |
| `Slide list mismatch` | `sldIdLst` disagrees with slide parts | Re-run; report if it recurs |
| `Expected N slides, found M` | Slides were missed or duplicated | Check the `--selector` matches only real slides |
| `slide N: no pictures and no text` | Slide is blank | The slide never became visible — check the deck's show mechanism |
| `slide N: no text runs (expected with --expect-text)` | Text extraction found nothing | Text may be inside canvas/SVG, or the slide is genuinely image-only |
| `slide N: references missing media` | Broken relationship | Re-run the conversion |
| `multiple <a:pPr>` | The pptxgenjs bug above | Should be auto-repaired; report if it survives |
| `shares theme part` | Notes master reuses `theme1.xml` | Should be auto-repaired; report if it survives |
| `empty notes text run` | Blank notes placeholder has an empty run | Should be auto-repaired; report if it survives |
| `schema: …` | Any other ECMA-376 violation | Not a known bug — report the exact message |

## Warnings worth mentioning to the user

- **Artwork aspect differs from slide aspect** — the picture will be distorted or letterboxed.
  Usually means the stage was mid-resize during capture.
- **Not canonical widescreen** — fine standalone, awkward when merged into another deck.
- **Blank-artwork heuristic** — check that slide before shipping.
