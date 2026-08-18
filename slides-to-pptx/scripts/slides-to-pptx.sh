#!/usr/bin/env bash
# slides-to-pptx.sh — Convert a frontend-slides HTML deck to .pptx, then validate it.
#
# Usage:
#   bash scripts/slides-to-pptx.sh <path-to-html> [output.pptx] [options]
#
# Examples:
#   bash scripts/slides-to-pptx.sh ./deck.html
#   bash scripts/slides-to-pptx.sh ./deck.html ./deck.pptx --mode searchable
#   bash scripts/slides-to-pptx.sh ./my-deck/index.html --mode editable --native-images
#
# Options (passed through to html2pptx.mjs):
#   --mode image|searchable|editable   How much of the deck becomes native PowerPoint content
#   --scale N                          Screenshot pixel ratio (default 2; use 1 for smaller files)
#   --jpeg [quality]                   JPEG artwork instead of PNG — much smaller, softer text
#   --native-images                    editable mode: lift <img> tags out as replaceable pictures
#   --no-background-art                editable mode: omit full-slide screenshot artwork
#   --keep-chrome                      Keep nav arrows / page counters / edit UI in the artwork
#   --keep-shots <dir>                 Keep the slide screenshots for visual verification
#   --font-map "Clash Display=Impact"  Override the built-in webfont substitutions
#   --selector <css>                   Slide selector if the deck does not use .slide
#   --no-notes / --no-links            Skip speaker notes / hyperlink hotspots
#   --no-shrink                        editable mode: keep authored font sizes even where
#                                      the substituted font overflows its box
#   --skip-validate                    Convert only, do not run the validator
#   --preview <dir>                    Also render each emitted slide to <dir>/slide-NNN.png
#   --no-inspect                       Skip the layout inspection pass
#
# What this does:
#   1. Installs Node deps (playwright, pptxgenjs, jszip) into a cached local dir
#   2. Serves the deck over HTTP and drives it in headless Chromium
#   3. Captures each slide at its authored stage size and builds the .pptx
#   4. Validates package structure, deck geometry and OOXML schema
#   5. Inspects the emitted layout for off-slide text, collisions and tiny fonts
set -euo pipefail

RED='\033[0;31m'; GREEN='\033[0;32m'; CYAN='\033[0;36m'; YELLOW='\033[1;33m'
BOLD='\033[1m'; NC='\033[0m'
info() { echo -e "${CYAN}ℹ${NC} $*"; }
ok()   { echo -e "${GREEN}✓${NC} $*"; }
warn() { echo -e "${YELLOW}⚠${NC} $*"; }
err()  { echo -e "${RED}✗${NC} $*" >&2; }

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# ─── Parse arguments ──────────────────────────────────────
SKIP_VALIDATE=false
SKIP_INSPECT=false
PREVIEW_DIR=""
MODE=image
POSITIONAL=()
PASSTHROUGH=()

while [[ $# -gt 0 ]]; do
    case "$1" in
        --skip-validate) SKIP_VALIDATE=true; shift ;;
        --no-inspect)    SKIP_INSPECT=true; shift ;;
        --preview)       PREVIEW_DIR="$2"; shift 2 ;;
        --mode)          MODE="$2"; PASSTHROUGH+=("$1" "$2"); shift 2 ;;
        # Options that take a value
        --scale|--selector|--font-map|--keep-shots|--title|--author|--company|--subject)
            PASSTHROUGH+=("$1" "$2"); shift 2 ;;
        --jpeg)
            PASSTHROUGH+=("$1"); shift
            if [[ ${1:-} =~ ^[0-9]+$ ]]; then PASSTHROUGH+=("$1"); shift; fi ;;
        --native-images|--no-background-art|--keep-chrome|--no-notes|--no-links|--no-shrink)
            PASSTHROUGH+=("$1"); shift ;;
        -*) err "Unknown option: $1"; exit 1 ;;
        *)  POSITIONAL+=("$1"); shift ;;
    esac
done

if [[ ${#POSITIONAL[@]} -lt 1 ]]; then
    err "Usage: bash scripts/slides-to-pptx.sh <path-to-html> [output.pptx] [--mode image|searchable|editable]"
    exit 1
fi

INPUT_HTML="${POSITIONAL[0]}"
[[ -f "$INPUT_HTML" ]] || { err "File not found: $INPUT_HTML"; exit 1; }
INPUT_HTML="$(cd "$(dirname "$INPUT_HTML")" && pwd)/$(basename "$INPUT_HTML")"

if [[ ${#POSITIONAL[@]} -ge 2 ]]; then
    OUTPUT_PPTX="${POSITIONAL[1]}"
else
    OUTPUT_PPTX="$(dirname "$INPUT_HTML")/$(basename "${INPUT_HTML%.*}").pptx"
fi
OUT_DIR="$(dirname "$OUTPUT_PPTX")"
mkdir -p "$OUT_DIR"
OUTPUT_PPTX="$(cd "$OUT_DIR" && pwd)/$(basename "$OUTPUT_PPTX")"

echo ""
echo -e "${BOLD}╔══════════════════════════════════════╗${NC}"
echo -e "${BOLD}║      Slides → PowerPoint (.pptx)     ║${NC}"
echo -e "${BOLD}╚══════════════════════════════════════╝${NC}"
echo ""

# ─── Dependencies ─────────────────────────────────────────
# Cached in one place so repeat conversions are instant. Override with SLIDES_PPTX_HOME.

command -v npx >/dev/null 2>&1 || {
    err "Node.js is required but not installed."
    err "  macOS:  brew install node"
    err "  or download from https://nodejs.org"
    exit 1
}

DEPS_DIR="${SLIDES_PPTX_HOME:-$HOME/.cache/slides-to-pptx}"
mkdir -p "$DEPS_DIR"

if [[ ! -d "$DEPS_DIR/node_modules/pptxgenjs" ]] \
   || [[ ! -d "$DEPS_DIR/node_modules/playwright" ]] \
   || [[ ! -d "$DEPS_DIR/node_modules/jszip" ]]; then
    info "Installing converter dependencies (first run only, ~30-60s)..."
    printf '{"name":"slides-to-pptx-deps","private":true,"type":"module"}' > "$DEPS_DIR/package.json"
    (cd "$DEPS_DIR" && npm install --no-audit --no-fund playwright pptxgenjs jszip) >/dev/null 2>&1 || {
        err "Failed to install dependencies. Try manually:"
        err "  cd $DEPS_DIR && npm install playwright pptxgenjs jszip"
        exit 1
    }
    (cd "$DEPS_DIR" && npx playwright install chromium) >/dev/null 2>&1 || {
        err "Failed to download Chromium. Try: npx playwright install chromium"
        exit 1
    }
fi
ok "Dependencies ready"

# ─── Convert ──────────────────────────────────────────────
# SLIDES_PPTX_DEPS tells the scripts where the shared dependency cache lives.
# (NODE_PATH would not work — ESM resolves imports relative to the module file.)

info "Converting $(basename "$INPUT_HTML")..."
echo ""

export SLIDES_PPTX_DEPS="$DEPS_DIR"
if ! node "$SCRIPT_DIR/html2pptx.mjs" "$INPUT_HTML" "$OUTPUT_PPTX" "${PASSTHROUGH[@]+"${PASSTHROUGH[@]}"}"; then
    err "Conversion failed."
    exit 1
fi

# ─── Validate ─────────────────────────────────────────────

if [[ "$SKIP_VALIDATE" == "false" ]]; then
    echo ""
    info "Validating the .pptx..."
    VALIDATE_ARGS=()
    # Editable and searchable decks must carry real text; image decks legitimately do not.
    if [[ "$MODE" == "searchable" || "$MODE" == "editable" ]]; then
        VALIDATE_ARGS+=(--expect-text)
    fi
    if ! node "$SCRIPT_DIR/validate-pptx.mjs" "$OUTPUT_PPTX" "${VALIDATE_ARGS[@]+"${VALIDATE_ARGS[@]}"}"; then
        warn "Validation reported problems — see above before sharing this file."
        exit 1
    fi
fi

# ─── Inspect the layout ───────────────────────────────────
# Validation proves the package is well-formed; this proves the content landed where it
# should. Reads the emitted geometry and, with --preview, renders it back to PNGs.

if [[ "$SKIP_INSPECT" == "false" ]]; then
    echo ""
    info "Inspecting the emitted layout..."
    INSPECT_ARGS=(--quiet)
    [[ -n "$PREVIEW_DIR" ]] && INSPECT_ARGS+=(--preview "$PREVIEW_DIR")
    if ! node "$SCRIPT_DIR/inspect-pptx.mjs" "$OUTPUT_PPTX" "${INSPECT_ARGS[@]}"; then
        warn "Layout inspection found problems — see above before sharing this file."
        exit 1
    fi
fi

# ─── Done ─────────────────────────────────────────────────

FILE_SIZE=$(du -h "$OUTPUT_PPTX" | cut -f1 | xargs)
echo -e "${BOLD}════════════════════════════════════════${NC}"
ok "PowerPoint file ready"
echo ""
echo -e "  ${BOLD}File:${NC} $OUTPUT_PPTX"
echo -e "  ${BOLD}Size:${NC} $FILE_SIZE"
echo ""
echo -e "${BOLD}════════════════════════════════════════${NC}"
echo ""

# Best-effort open in whatever the user has — nothing installed is not a failure.
if command -v open >/dev/null 2>&1; then
    open "$OUTPUT_PPTX" 2>/dev/null || info "No installed app claims .pptx — the file is still valid."
fi
