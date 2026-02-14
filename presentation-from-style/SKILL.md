---
name: presentation-from-style
description: "Create professional presentations (pptx) by combining content (text/markdown/pptx) with visual style from a separate source (pptx template, pdf presentation, or pdf brandbook). Trigger when the user has BOTH a content source AND a style reference and wants to merge them into a finished presentation. Also trigger on: 'create slides from template', 'presentation in our brand style', 'copy the style from this deck', 'make slides matching this brandbook', 'style this content like that template'."
---

# Presentation from Style

Generate presentations that match a given visual style — extracted from a PPTX template, a PDF presentation, or a brand guidelines document.

**Prerequisite:** This skill depends on the `pptx` skill. Read its SKILL.md first.

## Strategy Selection

| Content source | Style source | Strategy | Key files to read |
|----------------|-------------|----------|-------------------|
| MD / text | PPTX template | Unpack template → edit XML | `pptx/editing.md` |
| MD / text | PDF presentation | Extract style → PptxGenJS from scratch | `pptx/pptxgenjs.md` |
| MD / text | PDF brandbook | Extract brand rules → PptxGenJS from scratch | `pptx/pptxgenjs.md` |
| PPTX (content) | PPTX (style) | Unpack style template → replace text in XML | `pptx/editing.md` |

## Required Arguments

Ask the user if not provided:

1. **Content source**: markdown file, text prompt, or .pptx with content
2. **Style source**: template .pptx, presentation .pdf, or brandbook .pdf
3. **Output filename**: `styled_<content filename>.pptx` (default convention)

## Workflow

### Step 0: Setup

```bash
pip install "markitdown[pptx]" Pillow PyMuPDF
npm install -g pptxgenjs
```

Create a `temp/` folder in the working directory for all intermediate files.
Do not delete temp/ after completion — it's useful for debugging.

### Step 1: Analyze inputs

**Content (any source):**
```bash
# From MD: read the file directly
# From PPTX: extract text
python -m markitdown content.pptx
```

Count slides, note content per slide (title, subtitle, body, bullets, tables).

**Style source — PPTX template:**
1. Extract text: `python -m markitdown template.pptx`
2. Export thumbnails: use PowerPoint COM (Windows) or LibreOffice to export all slides as JPG into `temp/template_slides/`
3. Visually inspect all template slides — identify layout types (title, content, 2-column, cards, stats, quote, closing)
4. Count available layouts. Plan which content slides map to which template layouts.

**Style source — PDF:**
1. Convert pages to images: `pdftoppm -jpeg -r 150 style.pdf temp/pdf_preview/slide`
2. Visually analyze: color palette, fonts, layout patterns, background textures
3. Extract textures/backgrounds with PyMuPDF if needed

**Style source — Brandbook PDF:**
1. Convert to images as above
2. Extract: primary/secondary/accent colors, fonts, logo rules, spacing guidelines
3. Download logo images from brandbook pages using PyMuPDF

### Step 2: Plan slide mapping

Create an explicit mapping: which content slide → which template layout.

Rules:
- **Vary layouts** — don't use the same template slide twice in a row
- **Match content complexity** to layout capacity (3 bullets → 3-card layout, not 4-card)
- **Title slide** → template's title layout
- **Closing slide** → template's thank-you/contact layout
- For PPTX editing: note which template slides to keep, delete, and reorder

### Step 3: Generate

**Strategy A — Edit PPTX template** (when style source is .pptx):

Follow the `pptx` skill's editing workflow:
1. `python scripts/office/unpack.py template.pptx temp/unpacked/`
2. Edit `ppt/presentation.xml` — reorder `<p:sldIdLst>` entries, remove unused slides
3. `python scripts/office/clean.py temp/unpacked/` — remove orphaned files
4. Edit each `ppt/slides/slide*.xml` — replace text content
5. `python scripts/office/pack.py temp/unpacked/ output.pptx`

**Critical XML editing rules:**
- Use the Edit tool, not sed/scripts
- Preserve `<a:pPr>` (paragraph properties) — never delete formatting
- Set `b="1"` on bold headers
- Don't use unicode bullets `•` — they display as `â€¢` in some viewers
- Escape `&` → `&amp;`, `<` → `&lt;`, `>` → `&gt;` in all text
- If replacement text is longer than original — check text box width (`cx` in `<a:ext>`)
- If template had a short number (e.g., "9.04K") and you're replacing with longer text — reduce `sz` (font size) or widen the text box

**Strategy B — Build from scratch** (when style source is PDF/brandbook):

Follow the `pptx` skill's PptxGenJS guide:
1. Create `temp/generate.js` with all slide definitions
2. Define color constants matching extracted palette
3. Download and reference fonts (Google Fonts → `fonts/` folder)
4. `node temp/generate.js`

**Critical PptxGenJS rules:**
- Don't use `#` prefix in hex colors → file corruption
- Use `bullet: true`, not unicode `•`
- Use `breakLine: true` between array text items
- Never reuse option objects (PptxGenJS mutates them)
- Images: always use `sizing: { type: "contain", w, h }` to preserve aspect ratio
- Set `margin: 0` on text boxes when aligning with shapes

### Step 4: Fonts and Cyrillic

**If content is in Russian or another Cyrillic language:**
1. Check if template/style fonts support Cyrillic
2. If not — find similar Google Fonts with Cyrillic support:
   - Sans-serif: Inter, PT Sans, Montserrat, Open Sans, Roboto, Noto Sans
   - Serif: PT Serif, Noto Serif, Merriweather
3. Download to `fonts/` folder
4. Russian text is 20-30% longer than English — adjust font sizes or widen text boxes

**If content is in English:**
- Use original template fonts as-is

### Step 5: QA (mandatory)

**Assume there are problems. Your job is to find them.**

**Content QA:**
```bash
python -m markitdown output.pptx
python -m markitdown output.pptx | grep -iE "lorem|ipsum|placeholder|xxxx"
```

**Visual QA:**
Export slides to images (PowerPoint COM on Windows, or LibreOffice → PDF → pdftoppm).
Use a **subagent with fresh eyes** — the editor sees what they expect, not what's there.

Check for:
- Overlapping text (especially where stat numbers were replaced with longer text)
- Leftover placeholder text from template
- Text cut off at edges or overflowing boxes
- Misaligned elements, inconsistent spacing
- Low-contrast text on backgrounds
- Distorted images/logos (wrong aspect ratio)

**Fix-and-verify loop:**
1. Generate → export images → inspect with subagent
2. List all issues found (if zero, look harder)
3. Fix issues
4. Re-export affected slides → re-inspect
5. Repeat until a full pass finds no new issues

**Minimum 2 rounds.** One fix often creates another problem.

## Known Pitfalls

These issues were discovered across 4 test runs. Address them proactively:

| # | Problem | Prevention |
|---|---------|------------|
| 1 | Template fonts lack Cyrillic | Check upfront; substitute with Google Fonts if needed |
| 2 | Russian text overflows text boxes | Reduce font size or widen boxes; budget +25% width |
| 3 | Tables rendered as text | Specify in prompt: "insert as real PowerPoint tables" |
| 4 | Agent invents content | Specify: "use ONLY content from the source file" |
| 5 | Large images (> 2000px) crash | Resize before processing |
| 6 | `&`, `<`, `>` break XML | Escape before editing slide XML |
| 7 | Image aspect ratios distorted | Use `sizing: { type: "contain" }` in PptxGenJS |
| 8 | PDF textures lost | Extract with PyMuPDF, use as slide backgrounds |
| 9 | Complex layouts simplified | Explicitly analyze layout zones in PDF (light/dark areas, proportions) |
| 10 | Text overlaps after replacing short template text with longer content | Check `cx` (width) and `sz` (font size); reduce font or widen box |

## Token Budget

Expect ~1M-1.5M input tokens for a 20-slide presentation:
- Template analysis: ~100-200k
- Content editing (with subagents): ~300-500k
- QA rounds (2-3 cycles): ~200-400k
- Growing conversation context: ~200-400k

**Optimization tips:**
- Use parallel subagents (4 agents × 5 slides) — same tokens, 3x faster
- Choose templates with master slides (lighter XML)
- Limit to 2 QA rounds for drafts, 3 for final versions

## Installation

This skill requires the `pptx` skill to be installed. Verify:

```bash
# Check pptx skill is available
ls .agents/skills/pptx/SKILL.md
ls .agents/skills/pptx/editing.md
ls .agents/skills/pptx/pptxgenjs.md
ls .agents/skills/pptx/scripts/office/unpack.py
ls .agents/skills/pptx/scripts/office/pack.py
ls .agents/skills/pptx/scripts/office/clean.py
```

System dependencies:
- Python 3.10+ with pip
- Node.js 18+ with npm
- LibreOffice (for PDF conversion) — or PowerPoint on Windows
- Poppler utils (`pdftoppm`) — for PDF to image conversion

Python packages:
```bash
pip install "markitdown[pptx]" Pillow PyMuPDF
```

Node packages:
```bash
npm install -g pptxgenjs
```

## Shortened Prompt (with skill)

When this skill is installed, users can use a shorter prompt:

```
В папке <folder>/ находятся:
- <content_file> — контент
- <style_file> — стиль/шаблон

Используя навык presentation-from-style, создай презентацию
styled_<content_name>.pptx. Все временные файлы — в temp/.
```

The skill handles strategy selection, font management, QA workflow, and known pitfalls automatically.
