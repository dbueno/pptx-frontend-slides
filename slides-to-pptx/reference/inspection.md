# Inspection — proving the layout is right

```bash
node scripts/inspect-pptx.mjs deck.pptx [--preview <dir>] [--slides 1,10-14] [--quiet] [--json]
```

Exit code `0` = no hard failures, `1` = a slide is empty, `2` = bad usage.

[Validation](validation.md) proves the `.pptx` is a well-formed OOXML package. It cannot tell
you whether the words landed where they belong — text can be schema-perfect and still sit
underneath its neighbour or half off the slide. This is the check for that, and the wrapper runs
it automatically after validating.

Everything here is read back out of **the written package**, not out of the converter's memory,
so it independently confirms what was actually emitted.

## No office suite required

Do not add one, and do not reach for a desktop office suite to check a deck. The offline ones are
bad automation dependencies: on macOS the headless entry point is often a GUI launcher that forks
and **exits 0 immediately**, so a conversion that produced nothing is indistinguishable from a
corrupt file and a valid deck gets reported as broken. Cold starts are slow, the output is
unstructured, and they substitute their own fonts, so what they render is not what PowerPoint
renders anyway.

Neither output here needs one. The geometry comes from the OOXML, and the preview renders in the
same Playwright Chromium the converter already depends on.

## Output 1 — the geometry report

Every picture and text box on every slide, with position and size in inches, the fonts and sizes
in use, whether the box wraps, and the text itself:

```
  ── slide 10 — 1 picture(s), 14 visible text box(es)
     0.45,0.38 5.11×0.20in     Consolas 11pt nowrap        "THE DATA-FLOW GRAPH"
     0.89,3.63 6.53×1.58in     Consolas 124.2pt nowrap     "CTADL_"
     3.85,2.16 4.61×0.91in     Consolas 13.5pt             "a = p;\nb = a.f;   // b <- p.f"
```

`--quiet` suppresses the listing and prints only findings. `--json` emits the whole model,
which is the right form for asserting on it in a script.

## Output 2 — the preview render

`--preview <dir>` rebuilds each slide as HTML from the emitted coordinates, fonts and sizes, and
screenshots it to `<dir>/slide-NNN.png` at 1920px wide — the same size the converter captures
at, so a preview lines up 1:1 with `--keep-shots` artwork for direct comparison.

Fonts get a generic fallback of the same class (`'Consolas', monospace`), matching what the
converter's fit measurement does, so a preview on a machine without the Office fonts still
shows representative line lengths rather than silently falling back to a proportional face.

## The lint

Automatic checks for the failure modes that eyeballing is normally used to catch:

| Finding | What it means |
| ------- | ------------- |
| `off-slide` | A text box runs past a slide edge by more than 0.1in. Usually a substituted font reflowing text past where it was measured. |
| `overlap` | Two visible text boxes overlap by more than 25% of the smaller one's area. The HTML layout had no overlap, so this means something reflowed. `--overlap N` changes the threshold. |
| `tiny-font` | Text below 8pt (`--min-font N`). Typically the shrink-to-fit pass bottoming out. |
| `empty` | A slide with no pictures and no text. The only hard failure. |

`searchable` mode's invisible text layer is excluded from overlap and off-slide checks — it is
fully transparent by design, so it cannot collide with anything visually.

## Reading an overlap finding

Not every overlap is a defect, but on a deck converted from HTML most are. The HTML laid the
text out with nothing on top of anything, so an overlap in the output means a box was measured
in one place and drew in another. The usual causes, in order of likelihood:

1. **A substituted font reflowed a line.** Check the `shrunkTextBoxes` count from the
   conversion and the overflow list it prints; try `--font-map` to a narrower face.
2. **A flex or grid row.** Its border box spans all its children, so a text box drawn from the
   element covers its siblings. The converter uses the glyphs' own rect in this case; a finding
   here means a layout it did not recognise.
3. **Genuinely stacked design.** Deliberately layered text. Confirm against the preview and move
   on.

Always confirm against the preview PNG before reporting an overlap to the user as a defect.
