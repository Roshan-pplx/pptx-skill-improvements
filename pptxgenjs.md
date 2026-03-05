# PptxGenJS Tutorial

## Setup & Basic Structure

```javascript
const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';  // or 'LAYOUT_16x10', 'LAYOUT_4x3', 'LAYOUT_WIDE'
pres.author = 'Your Name';
pres.title = 'Presentation Title';

let slide = pres.addSlide();
slide.addText("Hello World!", { x: 0.5, y: 0.5, fontSize: 36, color: "363636" });

pres.writeFile({ fileName: "Presentation.pptx" });
```

## Layout Dimensions

Slide dimensions (coordinates in inches):
- `LAYOUT_16x9`: 10" × 5.625" (default)
- `LAYOUT_16x10`: 10" × 6.25"
- `LAYOUT_4x3`: 10" × 7.5"
- `LAYOUT_WIDE`: 13.3" × 7.5"

---

## Text & Formatting

```javascript
// Basic text
slide.addText("Simple Text", {
  x: 1, y: 1, w: 8, h: 2, fontSize: 24, fontFace: "Arial",
  color: "363636", bold: true, align: "center", valign: "middle"
});

// Character spacing (use charSpacing, not letterSpacing which is silently ignored)
slide.addText("SPACED TEXT", { x: 1, y: 1, w: 8, h: 1, charSpacing: 6 });

// Rich text arrays
slide.addText([
  { text: "Bold ", options: { bold: true } },
  { text: "Italic ", options: { italic: true } }
], { x: 1, y: 3, w: 8, h: 1 });

// Multi-line text (requires breakLine: true)
slide.addText([
  { text: "Line 1", options: { breakLine: true } },
  { text: "Line 2", options: { breakLine: true } },
  { text: "Line 3" }  // Last item doesn't need breakLine
], { x: 0.5, y: 0.5, w: 8, h: 2 });

// Text box margin (internal padding)
slide.addText("Title", {
  x: 0.5, y: 0.3, w: 9, h: 0.6,
  margin: 0  // Use 0 when aligning text with other elements like shapes or icons
});
```

**Tip:** Text boxes have internal margin by default. Set `margin: 0` when you need text to align precisely with shapes, lines, or icons at the same x-position.

---

## Lists & Bullets

```javascript
// ✅ CORRECT: Multiple bullets
slide.addText([
  { text: "First item", options: { bullet: true, breakLine: true } },
  { text: "Second item", options: { bullet: true, breakLine: true } },
  { text: "Third item", options: { bullet: true } }
], { x: 0.5, y: 0.5, w: 8, h: 3 });

// ❌ WRONG: Never use unicode bullets
slide.addText("• First item", { ... });  // Creates double bullets

// Sub-items and numbered lists
{ text: "Sub-item", options: { bullet: true, indentLevel: 1 } }
{ text: "First", options: { bullet: { type: "number" }, breakLine: true } }
```

---

## Shapes

```javascript
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 0.8, w: 1.5, h: 3.0,
  fill: { color: "FF0000" }, line: { color: "000000", width: 2 }
});

slide.addShape(pres.shapes.OVAL, { x: 4, y: 1, w: 2, h: 2, fill: { color: "0000FF" } });

slide.addShape(pres.shapes.LINE, {
  x: 1, y: 3, w: 5, h: 0, line: { color: "FF0000", width: 3, dashType: "dash" }
});

// With transparency
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "0088CC", transparency: 50 }
});

// Rounded rectangle (rectRadius only works with ROUNDED_RECTANGLE, not RECTANGLE)
// ⚠️ Don't pair with rectangular accent overlays — they won't cover rounded corners. Use RECTANGLE instead.
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "FFFFFF" }, rectRadius: 0.1
});

// With shadow
slide.addShape(pres.shapes.RECTANGLE, {
  x: 1, y: 1, w: 3, h: 2,
  fill: { color: "FFFFFF" },
  shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 }
});
```

Shadow options:

| Property | Type | Range | Notes |
|----------|------|-------|-------|
| `type` | string | `"outer"`, `"inner"` | |
| `color` | string | 6-char hex (e.g. `"000000"`) | No `#` prefix, no 8-char hex — see Common Pitfalls |
| `blur` | number | 0-100 pt | |
| `offset` | number | 0-200 pt | **Must be non-negative** — negative values corrupt the file |
| `angle` | number | 0-359 degrees | Direction the shadow falls (135 = bottom-right, 270 = upward) |
| `opacity` | number | 0.0-1.0 | Use this for transparency, never encode in color string |

To cast a shadow upward (e.g. on a footer bar), use `angle: 270` with a positive offset — do **not** use a negative offset.

**Note**: Gradient fills are not natively supported. Use a gradient image as a background instead.

---

## Images

### Image Sources

```javascript
// From file path
slide.addImage({ path: "images/chart.png", x: 1, y: 1, w: 5, h: 3 });

// From URL
slide.addImage({ path: "https://example.com/image.jpg", x: 1, y: 1, w: 5, h: 3 });

// From base64 (faster, no file I/O)
slide.addImage({ data: "image/png;base64,iVBORw0KGgo...", x: 1, y: 1, w: 5, h: 3 });
```

### Image Options

```javascript
slide.addImage({
  path: "image.png",
  x: 1, y: 1, w: 5, h: 3,
  rotate: 45,              // 0-359 degrees
  rounding: true,          // Circular crop
  transparency: 50,        // 0-100
  flipH: true,             // Horizontal flip
  flipV: false,            // Vertical flip
  altText: "Description",  // Accessibility
  hyperlink: { url: "https://example.com" }
});
```

### Image Sizing Modes

```javascript
// Contain - fit inside, preserve ratio
{ sizing: { type: 'contain', w: 4, h: 3 } }

// Cover - fill area, preserve ratio (may crop)
{ sizing: { type: 'cover', w: 4, h: 3 } }

// Crop - cut specific portion
{ sizing: { type: 'crop', x: 0.5, y: 0.5, w: 2, h: 2 } }
```

### Calculate Dimensions (preserve aspect ratio)

```javascript
const origWidth = 1978, origHeight = 923, maxHeight = 3.0;
const calcWidth = maxHeight * (origWidth / origHeight);
const centerX = (10 - calcWidth) / 2;

slide.addImage({ path: "image.png", x: centerX, y: 1.2, w: calcWidth, h: maxHeight });
```

### Supported Formats

- **Standard**: PNG, JPG, GIF (animated GIFs work in Microsoft 365)
- **SVG**: Works in modern PowerPoint/Microsoft 365

---

## Icons

Use react-icons to generate SVG icons, then rasterize to PNG for universal compatibility.

### Setup

```javascript
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaCheckCircle, FaChartLine } = require("react-icons/fa");

function renderIconSvg(IconComponent, color = "#000000", size = 256) {
  return ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
}

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = renderIconSvg(IconComponent, color, size);
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}
```

### Add Icon to Slide

```javascript
const iconData = await iconToBase64Png(FaCheckCircle, "#4472C4", 256);

slide.addImage({
  data: iconData,
  x: 1, y: 1, w: 0.5, h: 0.5  // Size in inches
});
```

**Note**: Use size 256 or higher for crisp icons. The size parameter controls the rasterization resolution, not the display size on the slide (which is set by `w` and `h` in inches).

### Icon Libraries

Pre-installed: `react-icons`, `react`, `react-dom`, `sharp`

Popular icon sets in react-icons:
- `react-icons/fa` - Font Awesome
- `react-icons/md` - Material Design
- `react-icons/hi` - Heroicons
- `react-icons/bi` - Bootstrap Icons

---

## Slide Backgrounds

```javascript
// Solid color
slide.background = { color: "F1F1F1" };

// Color with transparency
slide.background = { color: "FF3399", transparency: 50 };

// Image from URL
slide.background = { path: "https://example.com/bg.jpg" };

// Image from base64
slide.background = { data: "image/png;base64,iVBORw0KGgo..." };
```

---

## Tables

```javascript
slide.addTable([
  ["Header 1", "Header 2"],
  ["Cell 1", "Cell 2"]
], {
  x: 1, y: 1, w: 8, h: 2,
  border: { pt: 1, color: "999999" }, fill: { color: "F1F1F1" }
});

// Advanced with merged cells
let tableData = [
  [{ text: "Header", options: { fill: { color: "6699CC" }, color: "FFFFFF", bold: true } }, "Cell"],
  [{ text: "Merged", options: { colspan: 2 } }]
];
slide.addTable(tableData, { x: 1, y: 3.5, w: 8, colW: [4, 4] });
```

---

## Charts

```javascript
// Bar chart
slide.addChart(pres.charts.BAR, [{
  name: "Sales", labels: ["Q1", "Q2", "Q3", "Q4"], values: [4500, 5500, 6200, 7100]
}], {
  x: 0.5, y: 0.6, w: 6, h: 3, barDir: 'col',
  showTitle: true, title: 'Quarterly Sales'
});

// Line chart
slide.addChart(pres.charts.LINE, [{
  name: "Temp", labels: ["Jan", "Feb", "Mar"], values: [32, 35, 42]
}], { x: 0.5, y: 4, w: 6, h: 3, lineSize: 3, lineSmooth: true });

// Pie chart
slide.addChart(pres.charts.PIE, [{
  name: "Share", labels: ["A", "B", "Other"], values: [35, 45, 20]
}], { x: 7, y: 1, w: 5, h: 4, showPercent: true });
```

### Better-Looking Charts

Default charts look dated. Apply these options for a modern, clean appearance:

```javascript
slide.addChart(pres.charts.BAR, chartData, {
  x: 0.5, y: 1, w: 9, h: 4, barDir: "col",

  // Custom colors (match your presentation palette)
  chartColors: ["0D9488", "14B8A6", "5EEAD4"],

  // Clean background
  chartArea: { fill: { color: "FFFFFF" }, roundedCorners: true },

  // Muted axis labels
  catAxisLabelColor: "64748B",
  valAxisLabelColor: "64748B",

  // Subtle grid (value axis only)
  valGridLine: { color: "E2E8F0", size: 0.5 },
  catGridLine: { style: "none" },

  // Data labels on bars
  showValue: true,
  dataLabelPosition: "outEnd",
  dataLabelColor: "1E293B",

  // Hide legend for single series
  showLegend: false,
});
```

**Key styling options:**
- `chartColors: [...]` - hex colors for series/segments
- `chartArea: { fill, border, roundedCorners }` - chart background
- `catGridLine/valGridLine: { color, style, size }` - grid lines (`style: "none"` to hide)
- `lineSmooth: true` - curved lines (line charts)
- `legendPos: "r"` - legend position: "b", "t", "l", "r", "tr"

### Chart Layout Integration

Where you place a chart matters as much as how it looks.

#### Chart Type Selection

| Goal | Chart type |
|------|-----------|
| Compare values across categories | Bar (`BAR`, `barDir: "col"`) |
| Show trends over time | Line (`LINE`) |
| Show composition / parts of a whole | Pie or Donut (`PIE`, `DOUGHNUT`) |
| Show correlation between two variables | Scatter (`SCATTER`) |
| Compare multiple categories side-by-side | Clustered bar (`barGrouping: "clustered"`) |

#### Layout Patterns

**Two-column (narrative left / chart right)** — preferred for most data slides:
```javascript
// Left: headline + 2-3 context bullets (40% width)
slide.addText("Revenue accelerating", { x: 0.5, y: 0.7, w: 3.5, h: 0.8, fontSize: 24, bold: true });
slide.addText([...bullets], { x: 0.5, y: 1.7, w: 3.5, h: 3.2 });
// Right: chart (60% width, at least 4" wide)
slide.addChart(pres.charts.BAR, data, { x: 4.2, y: 0.6, w: 5.5, h: 4.5, barDir: "col" });
```

**Full-width chart with annotation strip below:**
```javascript
slide.addChart(pres.charts.LINE, data, { x: 0.5, y: 0.7, w: 9, h: 3.8 });
// Annotation strip
slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 4.6, w: 10, h: 1.0, fill: { color: "F1F5F9" } });
slide.addText("Key insight: growth inflected in Q3 following product launch", {
  x: 0.5, y: 4.7, w: 9, h: 0.7, fontSize: 13, italic: true, color: "475569" });
```

**Dashboard grid (2×2 for multiple metrics):**
```javascript
const charts = [
  { data: revenueData, x: 0.3, y: 0.7, w: 4.5, h: 2.2 },
  { data: growthData,  x: 5.2, y: 0.7, w: 4.5, h: 2.2 },
  { data: retentionData, x: 0.3, y: 3.1, w: 4.5, h: 2.2 },
  { data: npsData,    x: 5.2, y: 3.1, w: 4.5, h: 2.2 },
];
charts.forEach(({ data, x, y, w, h }) =>
  slide.addChart(pres.charts.BAR, data, { x, y, w, h, barDir: "col", showLegend: false })
);
```

#### Sizing Rules

- Charts should occupy **at least 60% of the content area** — small charts are hard to read
- Minimum **4" wide** for any chart; 5.5"+ preferred
- Never stack a chart directly below a long text block — use two-column layout instead

#### Callout Annotations

Highlight the key number with a bold callout next to or over the chart:

```javascript
// Bold annotation callout
slide.addText("↑ 23%", { x: 7.8, y: 1.2, w: 1.8, h: 0.6,
  fontSize: 22, bold: true, color: "0D9488", align: "center" });
slide.addText("YoY growth", { x: 7.8, y: 1.75, w: 1.8, h: 0.4,
  fontSize: 11, color: "64748B", align: "center" });
```

---

## Speaker Notes

Add presenter notes to any slide with `addNotes()`:

```javascript
slide.addNotes("Key talking points:\n- Revenue grew 23% YoY driven by enterprise segment\n- Pause here — ask audience: what's driving their growth?\n- Transition: next slide shows the regional breakdown");
```

### Best Practices

- **2-3 talking points per slide**, not a full script — notes are a memory aid, not a teleprompter
- **Include audience engagement cues**: questions to pose, moments to pause, transitions to the next slide
- **Use plain language**: write how you'd actually speak, not how you'd write a report
- **Flag complex visuals**: brief note explaining what the chart shows before interpreting it

```javascript
// Title slide
slide.addNotes("Welcome audience, briefly introduce yourself.\nToday's goal: leave with one actionable insight on X.");

// Data slide
slide.addNotes("This chart shows 18-month trend — note the inflection in Q3.\nAsk: does this match what you're seeing in your market?");

// Conclusion slide
slide.addNotes("Recap the 3 main points. End with the call to action.\nLeave 5 minutes for Q&A.");
```

---

## Slide Masters (Required)

**Every presentation created from scratch MUST define at least one slide master before adding any slides.** This ensures consistent branding, font defaults, and background styling across all slides.

### Why Slide Masters Matter

- **Consistency** — backgrounds, fonts, and colors apply automatically to every slide that references the master, so you don't repeat styling on each slide
- **Professional appearance** — a unified look signals intentional design; mismatched slides signal laziness
- **Easier theme changes** — update the master definition and every referencing slide updates with it
- **Default fonts and colors propagate** — text added to placeholders inherits the master's font settings, reducing per-element styling

### Defining Masters

Call `pres.defineSlideMaster(...)` for each master **before** any `pres.addSlide(...)` calls. At minimum, define a title slide master and a content slide master.

```javascript
const pptxgen = require("pptxgenjs");
let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";

// ─── Define all slide masters FIRST ───────────────────────────

// 1. Title slide — dark background, centered title and subtitle
pres.defineSlideMaster({
  title: "TITLE_SLIDE",
  background: { color: "1E2761" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 1, y: 1.5, w: 8, h: 1.5, fontFace: "Georgia", fontSize: 40, color: "FFFFFF", bold: true, align: "center" } } },
    { placeholder: { options: { name: "subtitle", type: "body", x: 1, y: 3.2, w: 8, h: 1, fontFace: "Calibri", fontSize: 20, color: "CADCFC", align: "center" } } }
  ]
});

// 2. Content slide — light background, top title area, body area with standard margins
pres.defineSlideMaster({
  title: "CONTENT_SLIDE",
  background: { color: "F2F2F2" },
  objects: [
    // Title bar background
    { rect: { x: 0, y: 0, w: 10, h: 0.9, fill: { color: "1E2761" } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.1, w: 9, h: 0.7, fontFace: "Georgia", fontSize: 24, color: "FFFFFF", bold: true } } },
    { placeholder: { options: { name: "body", type: "body", x: 0.5, y: 1.2, w: 9, h: 4.0, fontFace: "Calibri", fontSize: 16, color: "363636" } } }
  ]
});

// 3. Section divider (optional) — bold color break between sections
pres.defineSlideMaster({
  title: "SECTION_DIVIDER",
  background: { color: "283A5E" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 1, y: 2, w: 8, h: 1.5, fontFace: "Georgia", fontSize: 36, color: "FFFFFF", bold: true, align: "center" } } }
  ]
});

// ─── Now add slides using the masters ─────────────────────────

let slide1 = pres.addSlide({ masterName: "TITLE_SLIDE" });
slide1.addText("Quarterly Business Review", { placeholder: "title" });
slide1.addText("Q4 2025 Results & Outlook", { placeholder: "subtitle" });

let slide2 = pres.addSlide({ masterName: "CONTENT_SLIDE" });
slide2.addText("Revenue Summary", { placeholder: "title" });
slide2.addText("Body content here — or skip the placeholder and add elements manually.", { placeholder: "body" });

let slide3 = pres.addSlide({ masterName: "SECTION_DIVIDER" });
slide3.addText("Regional Breakdown", { placeholder: "title" });
```

### Using Masters with Slides

Every `addSlide()` call should reference a master:

```javascript
// ✅ CORRECT: slide inherits master's background, fonts, and layout
let slide = pres.addSlide({ masterName: "CONTENT_SLIDE" });

// ❌ AVOID: bare slide with no master — falls back to blank default
let slide = pres.addSlide();
```

Slides without a `masterName` fall back to a blank white default, which defeats the purpose of defining masters. You can still add any elements (text, shapes, charts, images) on top of a mastered slide — the master provides the base layer, not a constraint.

**Note:** You do not have to use placeholders. A master can define just a background and default objects (like a logo or footer bar), and you can add all content with regular `addText`, `addShape`, etc. calls. Placeholders are a convenience for common title/body positions, not a requirement.

---

## Common Pitfalls

⚠️ These issues cause file corruption, visual bugs, or broken output. Avoid them.

1. **NEVER use "#" with hex colors** - causes file corruption
   ```javascript
   color: "FF0000"      // ✅ CORRECT
   color: "#FF0000"     // ❌ WRONG
   ```

2. **NEVER encode opacity in hex color strings** - 8-char colors (e.g., `"00000020"`) corrupt the file. Use the `opacity` property instead.
   ```javascript
   shadow: { type: "outer", blur: 6, offset: 2, color: "00000020" }          // ❌ CORRUPTS FILE
   shadow: { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.12 }  // ✅ CORRECT
   ```

3. **Use `bullet: true`** - NEVER unicode symbols like "•" (creates double bullets)

4. **Use `breakLine: true`** between array items or text runs together

5. **Avoid `lineSpacing` with bullets** - causes excessive gaps; use `paraSpaceAfter` instead

6. **Each presentation needs fresh instance** - don't reuse `pptxgen()` objects

7. **NEVER reuse option objects across calls** - PptxGenJS mutates objects in-place (e.g. converting shadow values to EMU). Sharing one object between multiple calls corrupts the second shape.
   ```javascript
   const shadow = { type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.15 };
   slide.addShape(pres.shapes.RECTANGLE, { shadow, ... });  // ❌ second call gets already-converted values
   slide.addShape(pres.shapes.RECTANGLE, { shadow, ... });

   const makeShadow = () => ({ type: "outer", blur: 6, offset: 2, color: "000000", opacity: 0.15 });
   slide.addShape(pres.shapes.RECTANGLE, { shadow: makeShadow(), ... });  // ✅ fresh object each time
   slide.addShape(pres.shapes.RECTANGLE, { shadow: makeShadow(), ... });
   ```

8. **Don't use `ROUNDED_RECTANGLE` with accent borders** - rectangular overlay bars won't cover rounded corners. Use `RECTANGLE` instead.
   ```javascript
   // ❌ WRONG: Accent bar doesn't cover rounded corners
   slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 1, y: 1, w: 3, h: 1.5, fill: { color: "FFFFFF" } });
   slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 0.08, h: 1.5, fill: { color: "0891B2" } });

   // ✅ CORRECT: Use RECTANGLE for clean alignment
   slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 3, h: 1.5, fill: { color: "FFFFFF" } });
   slide.addShape(pres.shapes.RECTANGLE, { x: 1, y: 1, w: 0.08, h: 1.5, fill: { color: "0891B2" } });
   ```

---

## Quick Reference

- **Shapes**: RECTANGLE, OVAL, LINE, ROUNDED_RECTANGLE
- **Charts**: BAR, LINE, PIE, DOUGHNUT, SCATTER, BUBBLE, RADAR
- **Layouts**: LAYOUT_16x9 (10"×5.625"), LAYOUT_16x10, LAYOUT_4x3, LAYOUT_WIDE
- **Alignment**: "left", "center", "right"
- **Chart data labels**: "outEnd", "inEnd", "center"
