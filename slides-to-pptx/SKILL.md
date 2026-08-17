---
name: slides-to-pptx
description: Convert an HTML presentation (built with the frontend-slides skill, or any deck of fixed-size slide elements) into a real PowerPoint .pptx file, then validate it. Use when the user asks for a .pptx, a PowerPoint export, or needs to hand a deck to someone who will open it in PowerPoint, Keynote, or Google Slides.
---

# Slides → PowerPoint

Turn a frontend-slides HTML deck into a `.pptx` that opens correctly in PowerPoint, Keynote and
Google Slides — and prove it is right before handing it over.

## Core Principles

1. **Fidelity is the default, editability is a request.** A frontend-slides deck is a designed
   artifact: custom webfonts, gradients, CSS art. Reproducing it pixel-perfect is the honest
   default. Rebuilding it as native shapes is what you do when someone must *edit* it.
2. **Never hand over an unverified file.** Two different questions, two different checks: the
   validator proves the package is well-formed, the inspector proves the content landed where it
   belongs. A deck can pass the first and be unusable. Run both — the wrapper does by default.
3. **Say what was lost.** Every conversion drops something — animations always, fonts usually,
   editability sometimes. Name the losses explicitly instead of letting the user discover them
   in front of an audience.
4. **The deck's own final frame is the source of truth.** Capture what the design intends to
   show after its reveal animations settle — do not stomp opacity/transform to force it.

## When to Use This vs. PDF Export

| User needs | Use |
| ---------- | --- |
| To email/print/archive a static copy | `frontend-slides` PDF export (`scripts/export-pdf.sh`) |
| To present from PowerPoint or Keynote | This skill, `image` mode |
| To hand the deck to a team that edits in PowerPoint | This skill, `editable` mode |
| A file that is accessible / searchable / copy-pasteable | This skill, `searchable` mode |
| A live shareable link | `frontend-slides` Vercel deploy |

If the user just wants "a file I can send", ask whether the recipient needs to **edit** it. If
not, PDF is often the better answer — say so, then do what they ask.

---

## Phase 0: Check the Deck

1. **Locate the HTML.** If the user just generated a deck with `frontend-slides`, use that file.
2. **Confirm the slide selector.** The converter looks for `.slide` elements. Check:
   ```bash
   grep -c 'class="[^"]*slide' deck.html
   ```
   If the deck uses a different class, pass `--selector`.
3. **Check for local assets.** Images must resolve relative to the HTML file. Absolute
   filesystem paths (`src="/Users/name/photo.png"`) will not load — fix them first.
4. **Note the stage size.** Standard frontend-slides decks are 1920×1080. The converter
   auto-detects any size and maps it onto a 16:9 PowerPoint deck.

---

## Phase 1: Choose a Conversion Mode

**Ask the user.** This is the one decision that changes the output materially. Use the
environment's structured-question UI if available, otherwise ask concisely.

**Question — "PowerPoint mode"**: How will this file be used?

| Option | What they get | Pick when |
| ------ | ------------- | --------- |
| **Presenting as-is** (`image`) | Every slide is one full-bleed picture. Pixel-perfect: exact fonts, gradients, CSS art. Nothing is editable. | Default. They will present from PowerPoint or just need a `.pptx` to send. |
| **Searchable & accessible** (`searchable`) | Same pixel-perfect picture, plus an invisible native text layer. Text is searchable, copy-pasteable and readable by screen readers. Still not visually editable. | Corporate/compliance sharing, accessibility requirements, decks that get searched. |
| **Editable in PowerPoint** (`editable`) | Background artwork is captured with the text hidden, then every text block is rebuilt as a real PowerPoint text box. Fully editable. Webfonts get substituted, and text the stand-in font would push out of its box is shrunk to fit. | Someone else must change the words, translate it, or reuse slides. |

Read [reference/conversion-modes.md](reference/conversion-modes.md) before recommending a mode
if you need the detailed trade-offs.

**If the user is unsure, recommend `image`** and mention they can re-run in another mode in
under a minute — conversion is cheap and non-destructive.

---

## Phase 2: Convert

```bash
bash scripts/slides-to-pptx.sh <path-to-html> [output.pptx] [--mode image|searchable|editable]
```

The wrapper installs its Node dependencies once (into `~/.cache/slides-to-pptx`), converts, and
validates. Examples:

```bash
# Default: pixel-perfect, presentation-ready
bash scripts/slides-to-pptx.sh ./deck.html

# Accessible / searchable
bash scripts/slides-to-pptx.sh ./deck.html --mode searchable

# Editable, with <img> tags lifted out as replaceable pictures
bash scripts/slides-to-pptx.sh ./deck.html --mode editable --native-images

# Small file for email
bash scripts/slides-to-pptx.sh ./deck.html --jpeg 85 --scale 1
```

### Options that matter

| Flag | Effect |
| ---- | ------ |
| `--mode` | `image` (default), `searchable`, `editable` |
| `--scale N` | Capture pixel ratio. `2` (default) is retina-crisp; `1` roughly thirds the file size |
| `--jpeg [q]` | JPEG artwork instead of PNG. Dramatic size win; slightly softer text edges |
| `--native-images` | `editable` only: extract `<img>` tags as real, replaceable pictures |
| `--font-map "A=B,C=D"` | Override webfont substitution (see Phase 4) |
| `--selector <css>` | If the deck does not use `.slide` |
| `--keep-chrome` | Keep nav arrows / page counters / edit UI in the artwork (normally hidden) |
| `--no-notes`, `--no-links` | Skip speaker notes / hyperlink hotspots |
| `--preview <dir>` | Render each emitted slide back to a PNG for visual verification |
| `--no-shrink` | `editable` only: keep authored font sizes even where the substituted font overflows |
| `--no-inspect` | Skip the layout inspection pass |

### File size expectations

A 1920×1080 slide at `--scale 2` is about **1.2 MB per slide** as PNG. Budget accordingly:

| Setting | ~Size for an 18-slide deck |
| ------- | -------------------------- |
| default (PNG, scale 2) | ~21 MB |
| `--scale 1` | ~7 MB |
| `--jpeg 85 --scale 1` | ~1 MB |

**If the result exceeds 20 MB, proactively tell the user and offer to re-run with
`--jpeg 85`.** Many mail servers reject attachments over 25 MB.

---

## Phase 3: Read the Validation Output

The wrapper runs the validator automatically. To run it alone:

```bash
node scripts/validate-pptx.mjs deck.pptx --expect-slides 18 --expect-text
```

A pass looks like:

```
  18 slides · 6 notes pages · 214 text runs
  13.3333×7.5in (12192000×6858000 EMU) — canonical widescreen
  schema: valid (ECMA-376)
  ✓ PASS
```

**Do not hand over a file that reports `✗ FAIL`.** See
[reference/validation.md](reference/validation.md) for what each check means and how to fix
failures. Warnings (`⚠`) are worth reading aloud to the user but are not blockers.

---

## Phase 4: Inspect the Layout (Required)

Validation proves the file is *well-formed*. It does not prove it *looks right*. The wrapper runs
`inspect-pptx.mjs` next, which reads the geometry back out of the written package and lints it:

```
  ⚠ slide 16 [overlap] two text boxes overlap by 100% — "For Java, over-approximates" / "CHA"
  ⚠ slide 8 [off-slide] text box runs 0.4in past the right edge: "…"
  ⚠ slide 1 [tiny-font] 7.2pt text (below 8pt): "Denis Bueno · Arlen Cox"
```

**These findings are usually real.** The HTML laid the deck out with nothing overlapping, so an
overlap in the output means something reflowed. See
[reference/inspection.md](reference/inspection.md) for how to read each one.

No office suite is involved and none is needed. Do not add one to check a deck — that file
explains why, and the preview below is the better check.

### Look at it

Add `--preview ./preview` and open 2-3 of the PNGs — the title slide, the densest content slide,
and any slide the lint flagged. Each is the emitted `.pptx` rebuilt from its own coordinates, so
what you see is what was written. Confirm:

- Text sits where the design put it; nothing collides or spills out of a panel
- Gradients, CSS art and any images are present in the background artwork
- Reveal animations landed in their final state, not mid-fade
- No nav chrome or page counters baked in

To compare against the source instead, `--keep-shots ./shots` writes the captured artwork at the
same 1920px width, so the two sets line up 1:1.

In `editable` mode the artwork should show the design with **all words removed**. If words are
still visible there, they will be double-rendered under the native text boxes — report it rather
than shipping it.

### Font substitution (editable mode)

`editable` mode reports every substitution it made, and how many boxes it had to shrink:

```
  334 text boxes
  31 of them shrunk to fit the substituted font (--no-shrink to keep authored sizes)
  ⚠ 2 still overflow at the 80% floor — check these:
      slide 18: "MODELS"
font substitutions (webfonts are not installed on viewers' machines):
  JetBrains Mono → Consolas
  Space Grotesk → Arial
```

Substitution is the biggest fidelity loss in editable mode and **you must tell the user about
it**. Stand-in fonts have different metrics, so text that fitted its box in the browser reflows;
the converter measures each box again with the substitute applied and steps the font size down
until it fits, stopping at 80% and naming whatever still does not fit. Fixes, in order of
quality — see [reference/font-mapping.md](reference/font-mapping.md):

1. The recipient installs the actual webfonts (best fidelity, needs their cooperation)
2. Override the mapping with `--font-map "Clash Display=Oswald"` to a narrower font they have
3. Accept the defaults — always works, always looks different from the HTML

Note the measurement uses the fonts available on *this* machine. Office fonts like Consolas are
often absent, so it measures a same-class stand-in and lands close rather than exact.

---

## Phase 5: Deliver

Tell the user:

- **File location and size**
- **Which mode** and, in one sentence, what that means for them
- **What was lost.** Always: animations and transitions become their final static frame.
  In `image`/`searchable`: text is not visually editable. In `editable`: fonts were substituted,
  and say how many boxes were shrunk to keep the substitute inside its box.
- **Speaker notes carried over**, if any — they live in PowerPoint's notes pane
- **Hyperlinks still work** (they are invisible clickable hotspots over the artwork)

Then offer the natural follow-ups: re-run in a different mode, shrink the file, or export a PDF
instead via the frontend-slides skill.

---

## Troubleshooting

| Symptom | Cause | Fix |
| ------- | ----- | --- |
| `Found 0 slides` | Deck does not use `.slide` | `--selector .your-class` |
| Fonts look wrong in the artwork | Webfont did not load (offline / blocked CDN) | Check network, re-run; the capture needs the font to load |
| Slides captured mid-animation | Animation is infinite, so it is deliberately left running | Expected for looping ambience; add `--keep-shots` and check |
| Blank or near-blank slides | Slide never became visible | Deck uses a nonstandard show mechanism — check it toggles `.active`/`.visible` |
| Images missing | Absolute filesystem paths in `src` | Make paths relative to the HTML file |
| File is enormous | PNG at scale 2 | `--jpeg 85 --scale 1` |
| Text overflows its box in editable mode | Substituted font is wider than the original and the shrink pass bottomed out at 80% | `--font-map` to a narrower font, or use `searchable` mode |
| Editable text looks slightly small | The shrink pass fitted it to the substituted font | Expected; `--no-shrink` restores authored sizes and accepts the overflow |
| `overlap` findings on a flex/grid-heavy deck | A container's border box spans its children | The converter uses the glyphs' own rect here; confirm against `--preview` before reporting |
| Words glued together, or `<pre>` lines merged | An old build — whitespace-only text nodes were being dropped | Fixed; re-run with the current scripts |
| `schema: ... unexpected child element` | A pptxgenjs output bug | The converter auto-repairs the two known ones; report anything else |

---

## How It Works

Worth knowing when something goes wrong:

1. The deck is served over **HTTP** (not `file://`) so webfonts and relative assets load.
2. Headless Chromium opens it at the deck's **authored stage size**, so the stage renders at
   scale 1 and DOM coordinates are already slide coordinates.
3. For each slide, the deck's own `.active`/`.visible` contract is driven, then every finite
   CSS animation and transition is jumped to its end state via `Animation.finish()` — landing
   on the design's intended final frame without overriding any styles.
4. Positioned elements **outside** the slide stage (nav, counters, edit UI) are hidden.
5. Each slide is captured, then emitted at PowerPoint's canonical widescreen size.
   **1 authored px = 0.5 pt = 6350 EMU** at a 1920px stage, so positions and font sizes are exact.
6. Text boxes are cut from the DOM by inline flow, not by element: anything laid out inside its
   parent's line — `inline`, and `inline-block`/`inline-flex` chips and badges too — is absorbed
   into that parent's box, subtree and all. Lifting a mid-sentence badge into a box of its own
   would make the parent's words flow through the space it occupies and land on top of it.
7. A box is positioned by its element's content box, except where that box is shared with text
   that became a box of its own — a flex or grid row — in which case the glyphs' own rect is
   used. Text the browser laid out on one line is emitted `wrap="none"` so no renderer can
   reflow it.
8. In `editable` mode each box is measured a second time with its substituted font applied, in
   an off-screen clone, and the font size stepped down until it fits (floor 80%).
9. Two known pptxgenjs schema bugs are repaired in the written package (see
   [reference/validation.md](reference/validation.md)).

---

## Supporting Files

| File | Purpose | When to Read |
| ---- | ------- | ------------ |
| [reference/conversion-modes.md](reference/conversion-modes.md) | What survives each mode, in detail | Phase 1, when advising on mode |
| [reference/font-mapping.md](reference/font-mapping.md) | Webfont → Office font table and strategy | Phase 4, editable mode |
| [reference/inspection.md](reference/inspection.md) | Geometry report, preview render, and how to read each lint finding | Phase 4, always |
| [reference/validation.md](reference/validation.md) | Every check, what it catches, how to fix | Phase 3, on any failure |
| [scripts/slides-to-pptx.sh](scripts/slides-to-pptx.sh) | Wrapper: deps → convert → validate → inspect | Phase 2 |
| [scripts/html2pptx.mjs](scripts/html2pptx.mjs) | The converter | Direct/scripted use |
| [scripts/validate-pptx.mjs](scripts/validate-pptx.mjs) | The validator: package, geometry, schema | Phase 3 |
| [scripts/inspect-pptx.mjs](scripts/inspect-pptx.mjs) | The inspector: layout lint + preview render | Phase 4 |
