# Font Mapping

Only relevant in `editable` mode. In `image` and `searchable` modes the artwork carries the real
rendered glyphs, so fonts are exact regardless of what the recipient has installed.

## The problem

frontend-slides decks are explicitly told to avoid system fonts and use Fontshare or Google
Fonts. That is right for the web and wrong for PowerPoint: a `.pptx` references fonts *by name*
and the viewer's machine resolves them. `Clash Display` is not installed on anyone's laptop, so
PowerPoint silently substitutes something — usually badly, and differently on macOS and Windows.

So `editable` mode substitutes deliberately, up front, and reports what it did. A predictable
substitution you were told about beats an arbitrary one you discover during the talk.

## Built-in substitutions

Mapped to fonts that ship with Microsoft Office on **both** macOS and Windows.

| Webfont | Substitute | Why |
| ------- | ---------- | --- |
| Clash Display | Impact | Heavy, tight, display-weight |
| Oswald, Bebas Neue, Anton | Impact | Condensed display |
| Satoshi, Switzer, General Sans, Space Grotesk | Verdana | Neutral geometric sans, generous width |
| Cabinet Grotesk, Chillax, Poppins, Montserrat | Trebuchet MS | Rounder humanist sans |
| Inter, Roboto, Archivo | Arial | Neutral grotesque |
| Zodiak, Sentient, Gambetta, Bespoke Serif | Georgia | Contrast serif |
| Playfair Display, DM Serif Display, Libre Baskerville, Lora | Georgia | Serif |
| JetBrains Mono, IBM Plex Mono, Space Mono, Fira Code | Consolas | Monospace |
| Courier Prime | Courier New | Typewriter |

Any font not in the table is passed through unchanged — correct behaviour for a deck that
already uses Office-safe families or a font the recipient actually has.

## Choosing better

In order of quality:

### 1. Have the recipient install the real fonts (best)

Fontshare and Google Fonts are free and downloadable. If the deck goes to one team that will
edit it repeatedly, this is worth the five minutes:

```bash
# The deck's fonts are named in its <link> tag and CSS variables
grep -o 'fontshare[^"]*\|fonts.googleapis[^"]*' deck.html
```

Then pass `--font-map` with identity mappings, or simply let the pass-through handle it — the
converter only substitutes names it recognises as webfonts.

### 2. Map to a font they do have

If the recipient's organisation standardises on a font, map to it:

```bash
--font-map "Clash Display=Franklin Gothic Heavy,Satoshi=Segoe UI"
```

Match on **shape and width**, not vibe. The critical property is width: a substitute wider than
the original will overflow the text box, because boxes are sized from the original rendering.
Condensed display faces (Clash Display, Oswald) need condensed substitutes.

### 3. Accept the defaults

Always works, never matches. Fine for decks where the words matter more than the look.

## Safe fonts by category

Present in Office on macOS and Windows:

| Category | Safe choices |
| -------- | ------------ |
| Neutral sans | Arial, Helvetica, Verdana, Tahoma, Segoe UI (Win), Calibri |
| Humanist sans | Trebuchet MS, Gill Sans MT, Candara |
| Condensed / display | Impact, Arial Narrow, Haettenschweiler (Win) |
| Serif | Georgia, Times New Roman, Cambria, Garamond, Book Antiqua |
| Monospace | Consolas, Courier New |

Avoid macOS-only faces (Avenir, Futura, SF Pro, Optima) if the file goes to Windows, and
Windows-only faces (Segoe UI, Corbel, Haettenschweiler) if it goes to macOS.

## Checking the damage

After an `editable` conversion, look at the artwork with `--keep-shots` and compare against the
live HTML deck. Watch for:

- **Text overflowing its box** — substitute is too wide. Pick a narrower font.
- **Headlines looking weedy** — substitute is too light. Impact or a Black weight fixes it.
- **Wrong text case** — CSS `text-transform` is applied to the string during extraction, so
  uppercase styling is baked in correctly. If it looks wrong, the source CSS changed after capture.
- **Lost letter-spacing** — CSS `letter-spacing` maps to PowerPoint `charSpacing` and should
  survive. Wide-tracked eyebrow text is the usual place to check.

## Why fonts are not embedded

PowerPoint supports font embedding, but it is a Windows-PowerPoint-authored feature with poor
cross-platform support, it is blocked by most commercial font licences (Fontshare's included
for redistribution in this form), and it inflates the file substantially. Substituting and
telling the user is the honest trade. If exact fonts genuinely matter, use `image` or
`searchable` mode — the glyphs are already pixel-perfect there.
