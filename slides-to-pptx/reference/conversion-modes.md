# Conversion Modes

Three modes trade fidelity against editability. All three produce a valid, PowerPoint-native
`.pptx` at 13.3333×7.5 inches with speaker notes and working hyperlinks.

## The trade-off in one table

| | `image` | `searchable` | `editable` |
| --- | --- | --- | --- |
| Visual fidelity | Pixel-perfect | Pixel-perfect | High (art) + substituted fonts (text) |
| Custom webfonts | Exact | Exact | **Substituted** |
| Gradients, CSS art, effects | Exact | Exact | Exact (captured as artwork) |
| Text is searchable | No | **Yes** | **Yes** |
| Text is copy-pasteable | No | **Yes** | **Yes** |
| Screen-reader accessible | Via alt text only | **Yes, real text** | **Yes, real text** |
| Text is visually editable | No | No | **Yes** |
| Reflows when edited | — | — | **Yes** |
| Speaker notes | Yes | Yes | Yes |
| Hyperlinks | Yes | Yes | Yes |
| Animations | Final frame only | Final frame only | Final frame only |
| File size | Baseline | Baseline | Slightly smaller (text not in art) |
| Risk of looking wrong | Very low | Very low | **Moderate — verify visually** |

## `image` — the default

Each slide becomes one full-bleed picture on an otherwise empty slide.

**What you get:** exactly what the browser rendered. Every font, gradient, shadow, blend mode,
CSS-drawn shape and layout quirk survives, because none of it is being reinterpreted.

**What you lose:** nothing is editable. In PowerPoint the slide is a single picture. Changing a
typo means going back to the HTML and re-converting.

**Accessibility:** each picture gets alt text containing the slide's title and its full text
content, so screen readers announce something meaningful. This is a real mitigation but it is
not the same as native text — `searchable` mode is the better answer when accessibility matters.

**Pick it when:** the deck is being presented or sent as-is. This covers most requests.

## `searchable` — fidelity plus a text layer

Identical artwork to `image`, plus every text block reproduced as a native PowerPoint text box
with its fill set to 100% transparent — the same technique as an OCR layer in a scanned PDF.

**What you get:** the pixel-perfect slide, and text that PowerPoint's search finds, that users
can select and copy, and that screen readers read in reading order.

**What you lose:** the text is invisible, so editing it changes nothing visible. Users who click
into the invisible text box and type will see no effect on the slide — this surprises people, so
mention it.

**Pick it when:** the deck goes into a corporate repository, needs to satisfy an accessibility
requirement, or people will search across decks for a phrase.

## `editable` — real PowerPoint text

The slide is captured **with all text hidden**, so the artwork contains only background,
gradients, borders, cards, icons and images. Every text block is then rebuilt as a native
PowerPoint text box positioned exactly over that artwork.

**What you get:** a deck someone can genuinely edit. Change words, translate it, restyle text,
reuse individual slides. Positions and sizes are exact — a 120px slide padding lands at exactly
120px, and an 88px heading becomes exactly 44pt.

**What you lose:**

- **Fonts are substituted.** Fontshare/Google webfonts are not installed on the recipient's
  machine. See [font-mapping.md](font-mapping.md). This is the dominant fidelity loss.
- **Text effects do not transfer.** Gradient-filled text (`background-clip: text`), text
  shadows, outlined text, and variable-font weights become flat solid-colour text.
- **Layout is frozen.** Text boxes are absolutely positioned. Editing text longer than the
  original will overflow its box rather than reflowing the slide.

**Add `--native-images`** to also lift `<img>` tags out of the artwork as separate, replaceable
pictures. Without it, images stay baked into the background. Note that `object-fit`, CSS
`border-radius` and `box-shadow` on images are approximated.

**Pick it when:** a human must edit the words. Then **always verify visually** — this is the one
mode that can look wrong.

## What every mode drops

- **Animations and transitions.** Slides are captured after finite animations finish, so you get
  the design's intended final frame. Infinite/looping animations (ambient backgrounds, particle
  fields) are deliberately left running and captured wherever they happen to be.
- **Interactivity.** Inline editing, keyboard navigation, hover states, canvas simulations.
- **Video.** A `<video>` element is captured as its current frame.
- **Build steps.** frontend-slides decks reveal all content on slide entry, so there is nothing
  to expand. Decks with click-by-click builds capture the fully-revealed state.

## Switching modes

Conversion is cheap, non-destructive and takes seconds per slide. If the user is unsure, convert
in `image` mode, show them, and re-run in another mode if they want something different. Write
to a different output filename so they can compare.
