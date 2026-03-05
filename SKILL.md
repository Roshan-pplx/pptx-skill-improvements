---
name: pptx
description: "Use this skill any time a .pptx file is involved in any way — as input, output, or both. This includes: creating slide decks, pitch decks, or presentations; reading, parsing, or extracting text from any .pptx file (even if the extracted content will be used elsewhere, like in an email or summary); editing, modifying, or updating existing presentations; combining or splitting slide files; working with templates, layouts, speaker notes, or comments. Trigger whenever the user mentions \"deck,\" \"slides,\" \"presentation,\" or references a .pptx filename, regardless of what they plan to do with the content afterward. If a .pptx file needs to be opened, created, or touched, use this skill."
license: Proprietary. LICENSE.txt has complete terms
---

# PPTX Skill

## Quick Reference

| Task | Guide |
|------|-------|
| Read/analyze content | `python -m markitdown presentation.pptx` |
| Edit or create from template | Read [editing.md](editing.md) |
| Create from scratch | Read [pptxgenjs.md](pptxgenjs.md) |

---

## Reading Content

```bash
# Text extraction
python -m markitdown presentation.pptx

# Visual overview
python skills/pptx/scripts/thumbnail.py presentation.pptx

# Raw XML
python skills/pptx/scripts/office/unpack.py presentation.pptx unpacked/
```

---

## Editing Workflow

**Read [editing.md](editing.md) for full details.**

1. Analyze template with `thumbnail.py`
2. Unpack → manipulate slides → edit content → clean → pack

---

## Creating from Scratch

**Read [pptxgenjs.md](pptxgenjs.md) for full details.**

Use when no template or reference presentation is available. **Always define slide masters first** — every from-scratch deck must have at least a title and content master defined before adding slides. Add speaker notes to every slide by default — use `slide.addNotes("...")` with 2-3 talking points per slide.

---

## Creating from Documents

Convert an existing PDF, DOCX, or other document into a slide deck.

### Workflow

**1. Extract content**
```bash
python -m markitdown document.pdf    # or .docx, .txt, .html
```
Review the output — this is your raw material.

**2. Analyze structure**

Identify from the extracted text:
- Major sections / headings → become slide titles or section dividers
- Key statistics and numbers → become stat callouts or chart data
- Lists and processes → become bullet or flow slides
- Conclusions and recommendations → become the takeaway/CTA slide
- Themes and tone → inform color palette and visual style

**3. Map to slide types**

| Document element | Slide type |
|-----------------|------------|
| Document title / executive summary | Title slide |
| Major section headers | Section divider slide |
| Paragraphs with 3-5 key points | Content/bullet slide |
| Tables, statistics, data | Chart or stat callout slide |
| Case studies / examples | Evidence slide (image + text) |
| Conclusions / next steps | Takeaway / CTA slide |

**4. Apply content density rules**

Summarize — never transcribe. Dense prose does not belong on slides.

- Max **5-7 words per bullet point** — distill, don't copy
- Max **4-5 bullets per slide** — split into two slides if needed
- One idea per slide — resist cramming related points together
- Titles should be assertions, not labels: "Revenue grew 23% YoY" not "Revenue"

**5. Decide: visualize or summarize**

| Source content | How to handle it |
|---------------|-----------------|
| Numbers, percentages, trends | → Chart (bar, line, pie) |
| Ranked or comparative lists | → Bar chart or comparison columns |
| Sequential steps / processes | → Timeline or numbered flow |
| Long prose paragraphs | → Distill to 3-4 bullet points |
| Quotes / testimonials | → Callout slide with large pull quote |
| Team / person descriptions | → Multi-column card layout |

**6. Generate using PptxGenJS**

Follow the standard workflow in [pptxgenjs.md](pptxgenjs.md). Use the Content Planning outline (see above) to map each document section to a slide before writing code.

---

## Mathematical Equations

Use Unicode math characters. **Do not use OMML or render equations as images** — LibreOffice (used for visual QA) cannot render either.

---

## Content Planning

**Create a slide-by-slide outline BEFORE writing any code.** Jumping straight into code without a coherent story is the most common failure mode — the outline forces clarity on structure, content density, and narrative flow.

### Outline Format

For each slide, specify:

| Field | Description |
|-------|-------------|
| **Slide #** | Number and purpose (e.g., "Slide 3 — Problem Statement") |
| **Title** | Exact slide title text |
| **Key Points** | 2-4 bullets summarizing the content |
| **Layout Type** | e.g., two-column, full-bleed image, stat callout, chart slide |
| **Visual Elements** | Image, chart type, icons, background treatment needed |
| **Speaker Notes** | 1-2 sentence summary of what to say |

### Narrative Arc

Structure the deck around this flow — every slide should advance the story:

1. **Hook** — compelling opening: bold stat, provocative question, or striking visual
2. **Context** — why this topic matters; frame the problem or opportunity
3. **Core Content** — the main substance: data, findings, framework, or argument
4. **Evidence** — proof points, case studies, charts, testimonials
5. **Takeaway / CTA** — clear conclusion and what you want the audience to do next

**Example outline entry:**
```
Slide 4 — Market Opportunity
Title: "A $47B Market, Mostly Untapped"
Key Points:
  - Total addressable market reached $47B in 2024
  - Top 3 players hold only 28% combined share
  - SMB segment growing 3x faster than enterprise
Layout: Two-column (stat callout left / bar chart right)
Visuals: Stacked bar chart comparing competitor market share
Notes: Pause after the $47B number — let it land before moving to share breakdown.
```

---

## Design Ideas

**Don't create boring slides.** Plain bullets on a white background won't impress anyone. Consider ideas from this list for each slide.

### Before Starting

- **Pick a bold, content-informed color palette**: The palette should feel designed for THIS topic. If swapping your colors into a completely different presentation would still "work," you haven't made specific enough choices.
- **Dominance over equality**: One color should dominate (60-70% visual weight), with 1-2 supporting tones and one sharp accent. Never give all colors equal weight.
- **Dark/light contrast**: Dark backgrounds for title + conclusion slides, light for content ("sandwich" structure). Or commit to dark throughout for a premium feel.
- **Commit to a visual motif**: Pick ONE distinctive element and repeat it — rounded image frames, icons in colored circles, thick single-side borders. Carry it across every slide.

### Color Palettes

Choose colors that match your topic — don't default to generic blue. Use these palettes as inspiration:

| Theme | Primary | Secondary | Accent |
|-------|---------|-----------|--------|
| **Midnight Executive** | `1E2761` (navy) | `CADCFC` (ice blue) | `FFFFFF` (white) |
| **Forest & Moss** | `2C5F2D` (forest) | `97BC62` (moss) | `F5F5F5` (cream) |
| **Coral Energy** | `F96167` (coral) | `F9E795` (gold) | `2F3C7E` (navy) |
| **Warm Terracotta** | `B85042` (terracotta) | `E7E8D1` (sand) | `A7BEAE` (sage) |
| **Ocean Gradient** | `065A82` (deep blue) | `1C7293` (teal) | `21295C` (midnight) |
| **Charcoal Minimal** | `36454F` (charcoal) | `F2F2F2` (off-white) | `212121` (black) |
| **Teal Trust** | `028090` (teal) | `00A896` (seafoam) | `02C39A` (mint) |
| **Berry & Cream** | `6D2E46` (berry) | `A26769` (dusty rose) | `ECE2D0` (cream) |
| **Sage Calm** | `84B59F` (sage) | `69A297` (eucalyptus) | `50808E` (slate) |
| **Cherry Bold** | `990011` (cherry) | `FCF6F5` (off-white) | `2F3C7E` (navy) |

### For Each Slide

**Every slide needs a visual element** — image, chart, icon, or shape. Text-only slides are forgettable.

**Layout options:**
- Two-column (text left, illustration on right)
- Icon + text rows (icon in colored circle, bold header, description below)
- 2x2 or 2x3 grid (image on one side, grid of content blocks on other)
- Half-bleed image (full left or right side) with content overlay

**Data display:**
- Large stat callouts (big numbers 60-72pt with small labels below)
- Comparison columns (before/after, pros/cons, side-by-side options)
- Timeline or process flow (numbered steps, arrows)
- **Data slides**: use two-column layout (narrative left 40% / chart right 60%) or full-width chart with annotation strip — see [pptxgenjs.md](pptxgenjs.md) Chart Layout Integration for sizing rules and callout patterns

**Visual polish:**
- Icons in small colored circles next to section headers
- Italic accent text for key stats or taglines

### Advanced Visual Patterns

Go beyond basic layouts with these high-impact techniques.

#### Background Treatments

**Split background** — half the slide a dark color, half light:
```javascript
// Dark left panel (40% width)
slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 4, h: 5.625, fill: { color: "1E2761" } });
// Content sits on contrasting sides
```

**Color block overlay** — semi-transparent band over an image (40-60% coverage):
```javascript
slide.background = { path: "background.jpg" };
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 1.5, w: 10, h: 2.8, fill: { color: "000000", transparency: 40 }
});
```

**Edge accent band** — narrow strip along one edge:
```javascript
// Left edge band (brand color strip, ~0.15" wide)
slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: "F96167" } });
```

#### Geometric Accents

**Diagonal divider** — rotated rectangle creates a dynamic angle:
```javascript
slide.addShape(pres.shapes.RECTANGLE, {
  x: 4.6, y: -0.5, w: 0.4, h: 7, rotate: 8, fill: { color: "028090" }
});
```

**Corner bracket** — L-shaped accent using two thin rectangles:
```javascript
slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.4, w: 0.5, h: 0.06, fill: { color: "F96167" } });
slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 0.4, w: 0.06, h: 0.5, fill: { color: "F96167" } });
```

**Thick single-side border** — left accent bar on content cards:
```javascript
slide.addShape(pres.shapes.RECTANGLE, { x: 1.0, y: 1.2, w: 0.08, h: 1.4, fill: { color: "0891B2" } });
// Card content starts at x: 1.25 to respect the border
```

#### Advanced Layouts

**Z-pattern / F-pattern eye flow** — place the most important element top-left, second most important top-right, then bottom-left. Natural reading direction guides attention.

**Modular card grid** — equal-sized content blocks with consistent spacing:
```javascript
const cards = [{ x: 0.5 }, { x: 3.5 }, { x: 6.5 }];  // 3-column grid
cards.forEach(({ x }) => {
  slide.addShape(pres.shapes.RECTANGLE, { x, y: 1.5, w: 2.7, h: 2.8, fill: { color: "FFFFFF" },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 135, opacity: 0.12 } });
});
```

**Sidebar navigation strip** — vertical label strip on the left for multi-section slides:
```javascript
slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 1.2, h: 5.625, fill: { color: "1E2761" } });
slide.addText("SECTION 2", { x: 0, y: 2.3, w: 1.2, h: 0.5, rotate: 270,
  color: "FFFFFF", fontSize: 10, align: "center" });
```

#### Image Treatments

**Full-bleed with semi-transparent text overlay** — image covers entire slide, dark band holds readable text:
```javascript
slide.background = { path: "hero.jpg" };
slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 3.5, w: 10, h: 2.125,
  fill: { color: "000000", transparency: 35 } });
slide.addText("Your headline here", { x: 0.5, y: 3.7, w: 9, h: 1.5,
  color: "FFFFFF", fontSize: 40, bold: true });
```

**Circular image frame** — use `rounding: true` for profile photos or focus images:
```javascript
slide.addImage({ path: "headshot.jpg", x: 3.5, y: 1.0, w: 3, h: 3, rounding: true });
```

**Tinted color overlay on image** — place semi-transparent colored shape over image:
```javascript
slide.addImage({ path: "photo.jpg", x: 5, y: 0, w: 5, h: 5.625 });
slide.addShape(pres.shapes.RECTANGLE, { x: 5, y: 0, w: 5, h: 5.625,
  fill: { color: "028090", transparency: 60 } });
```

**Choose an interesting font pairing** — don't default to Arial. Pick a header font with personality and pair it with a clean body font.

| Header Font | Body Font |
|-------------|-----------|
| Georgia | Calibri |
| Arial Black | Arial |
| Calibri | Calibri Light |
| Cambria | Calibri |
| Trebuchet MS | Calibri |
| Impact | Arial |
| Palatino | Garamond |
| Consolas | Calibri |

| Element | Size |
|---------|------|
| Slide title | 36-44pt bold |
| Section header | 20-24pt bold |
| Body text | 14-16pt |
| Captions | 10-12pt muted |

### Spacing

- 0.5" minimum margins
- 0.3-0.5" between content blocks
- Leave breathing room—don't fill every inch

### Avoid (Common Mistakes)

- **Don't repeat the same layout** — vary columns, cards, and callouts across slides
- **Don't center body text** — left-align paragraphs and lists; center only titles
- **Don't skimp on size contrast** — titles need 36pt+ to stand out from 14-16pt body
- **Don't default to blue** — pick colors that reflect the specific topic
- **Don't mix spacing randomly** — choose 0.3" or 0.5" gaps and use consistently
- **Don't style one slide and leave the rest plain** — commit fully or keep it simple throughout
- **Don't create text-only slides** — add images, icons, charts, or visual elements; avoid plain title + bullets
- **Don't forget text box padding** — when aligning lines or shapes with text edges, set `margin: 0` on the text box or offset the shape to account for padding
- **Don't use low-contrast elements** — icons AND text need strong contrast against the background; avoid light text on light backgrounds or dark text on dark backgrounds
- **NEVER use accent lines under titles** — these are a hallmark of AI-generated slides; use whitespace or background color instead

## Source Citations

Every slide that uses information gathered from web sources MUST have a source attribution line at the bottom of the slide using **hyperlinked source names** — each source name is displayed as clickable text linking to the full URL. Always use "Source:" (singular). Use an array of text objects with `hyperlink` options.

```javascript
slide.addText([
  { text: "Source: " },
  { text: "Reuters", options: { hyperlink: { url: "https://reuters.com/article/123" } } },
  { text: ", " },
  { text: "WHO", options: { hyperlink: { url: "https://who.int/publications/m/item/update-42" } } },
  { text: ", " },
  { text: "World Bank", options: { hyperlink: { url: "https://worldbank.org/en/topic/water" } } }
], { x: 0.5, y: 5.2, w: 9, h: 0.3 });
```

- Each source name MUST have a `hyperlink.url` with the full `https://` URL — never omit hyperlinks
- WRONG: `"Sources: WHO, Reuters, UNICEF"` (plain text, no hyperlinks)
- WRONG: `"Source: WHO, https://who.int/report/123"` (raw URL in text instead of hyperlink)
- RIGHT: `[{ text: "WHO", options: { hyperlink: { url: "https://who.int/report/123" } } }]` (clickable name)

---

## QA (Required)

**Assume there are problems. Your job is to find them.**

Your first render is almost never correct. Approach QA as a bug hunt, not a confirmation step. If you found zero issues on first inspection, you weren't looking hard enough.

### Content QA

```bash
python -m markitdown output.pptx
```

Check for missing content, typos, wrong order.

**When using templates, check for leftover placeholder text:**

```bash
python -m markitdown output.pptx | grep -iE "xxxx|lorem|ipsum|this.*(page|slide).*layout"
```

If grep returns results, fix them before declaring success.

### Visual QA

**⚠️ USE SUBAGENTS** — even for 2-3 slides. You've been staring at the code and will see what you expect, not what's there. Subagents have fresh eyes.

Convert slides to images (see [Converting to Images](#converting-to-images)), then use this prompt:

```
Visually inspect these slides. Assume there are issues — find them.

Look for:
- Overlapping elements (text through shapes, lines through words, stacked elements)
- Text overflow or cut off at edges/box boundaries
- Decorative lines positioned for single-line text but title wrapped to two lines
- Source citations or footers colliding with content above
- Elements too close (< 0.3" gaps) or cards/sections nearly touching
- Uneven gaps (large empty area in one place, cramped in another)
- Insufficient margin from slide edges (< 0.5")
- Columns or similar elements not aligned consistently
- Low-contrast text (e.g., light gray text on cream-colored background)
- Low-contrast icons (e.g., dark icons on dark backgrounds without a contrasting circle)
- Text boxes too narrow causing excessive wrapping
- Leftover placeholder content

For each slide, list issues or areas of concern, even if minor.

Read and analyze these images:
1. /path/to/slide-01.jpg (Expected: [brief description])
2. /path/to/slide-02.jpg (Expected: [brief description])

Report ALL issues found, including minor ones.
```

### Verification Loop

1. Generate slides → Convert to images → Inspect
2. **List issues found** (if none found, look again more critically)
3. Fix issues
4. **Re-verify affected slides** — one fix often creates another problem
5. Repeat until a full pass reveals no new issues

**Do not declare success until you've completed at least one fix-and-verify cycle.**

---

## Converting to Images

Convert presentations to individual slide images for visual inspection:

```bash
python skills/pptx/scripts/office/soffice.py --headless --convert-to pdf output.pptx
pdftoppm -jpeg -r 150 output.pdf slide
```

This creates `slide-01.jpg`, `slide-02.jpg`, etc.

To re-render specific slides after fixes:

```bash
pdftoppm -jpeg -r 150 -f N -l N output.pdf slide-fixed
```

---

## Dependencies

All dependencies are pre-installed in the sandbox:

- `markitdown[pptx]` - text extraction
- `Pillow` - thumbnail grids
- `pptxgenjs` - creating from scratch (use `require("pptxgenjs")` directly)
- `react-icons`, `react`, `react-dom`, `sharp` - icon rendering
- LibreOffice (`soffice`) - PDF conversion
- Poppler (`pdftoppm`) - PDF to images
