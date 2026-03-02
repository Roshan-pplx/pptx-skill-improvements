# 5 Proposed Improvements to the PPTX Skill

Based on analysis of competing skills: partme-ai/full-stack-skills pptx-agent (Playbooks), ai-forever/slides_generator, Presenton, SlideDeck AI (slide-deck-ai), and darrenhinde/slidev-ai-presentations.

---

## 1. Add Speaker Notes Generation Guidance

**Gap:** Our skill has zero guidance on generating speaker notes. Competing skills like Presenton, SlideSpeak, and the Playbooks pptx-agent all emphasize speaker notes as a first-class feature. Presenton auto-generates detailed presenter notes with talking points per slide. The Playbooks skill explicitly reads/writes notes via `ppt/notesSlides/notesSlide{N}.xml`.

**Proposed Change:** Add a "Speaker Notes" section to `pptxgenjs.md` showing how to add notes via PptxGenJS (`slide.addNotes("...")`), and add guidance to `SKILL.md` encouraging notes generation as a default behavior when creating from scratch. Include best practices (2-3 sentences per slide, talking points not script, audience cues).

**Source Skills:** Presenton (auto-generated notes), Playbooks pptx-agent (XML notes access), SlideSpeak (notes as feature).

---

## 2. Add Slide Outline / Content Planning Phase

**Gap:** Our skill jumps straight from "Design Ideas" to code generation. Competing skills like SlideDeck AI, Presenton, and the Slidev skill all enforce a mandatory outline/planning step before generation. SlideDeck AI generates structured JSON outlines first. The Slidev skill creates a `plan.txt` with sections before generating slides. Presenton shows outlines for user review/editing before rendering.

**Proposed Change:** Add a "Content Planning" section to `SKILL.md` before the Design Ideas section. This should instruct the agent to: (a) create a slide-by-slide outline with titles, key points, and layout type before writing any code, (b) specify which visual elements each slide needs, (c) plan the narrative arc (hook → problem → solution → evidence → CTA). This prevents the common failure of jumping into code without a coherent story.

**Source Skills:** SlideDeck AI (JSON outline), Slidev AI (plan.txt), Presenton (outline review step).

---

## 3. Add Advanced Visual Design Patterns Section (Geometric Shapes, Backgrounds, Visual Treatments)

**Gap:** The Playbooks pptx-agent skill has a significantly richer design vocabulary than ours: diagonal dividers, asymmetric columns, rotated text, circular image frames, triangular accents, split backgrounds, color block overlays (40-60%), L-shaped borders, edge bands, floating boxes, Z/F-pattern layouts, and modular grids. Our skill covers basic layout options but lacks these advanced visual patterns that make slides look truly designed rather than generic.

**Proposed Change:** Expand the "Design Ideas > For Each Slide" section in `SKILL.md` with an "Advanced Visual Patterns" subsection covering: (a) background treatments (split backgrounds, color blocks, gradient images, edge bands), (b) geometric accents (diagonal dividers, corner brackets, thick single-side borders), (c) advanced layouts (Z/F-pattern eye flow, modular grids, sidebar nav, floating content boxes), (d) image treatments (full-bleed with overlay, circular frames, tinted overlays). Include PptxGenJS code snippets for each pattern.

**Source Skills:** Playbooks pptx-agent (18 palettes + extensive visual details section), Presenton (HTML/Tailwind CSS template system with rich visual patterns).

---

## 4. Add Document-to-Slides Workflow (PDF/DOCX → PPTX)

**Gap:** Multiple competing skills support converting existing documents into presentations. SlideDeck AI can create presentations from PDF files. Presenton supports PDF, DOCX, PPTX uploads as source material. The Playbooks skill supports creating from any existing content. Our skill has no guidance for this common use case — taking a report, paper, or document and converting it into a slide deck.

**Proposed Change:** Add a "Creating from Documents" section to `SKILL.md` with a workflow: (a) extract text from PDF/DOCX using `markitdown`, (b) use an LLM to identify key sections, themes, and data points, (c) map sections to slide types (title → title slide, data → chart slide, conclusion → CTA slide), (d) generate the presentation using the existing PptxGenJS workflow. Include guidance on content density (5-7 words per bullet, 4-5 bullets per slide max) and what to summarize vs. visualize.

**Source Skills:** SlideDeck AI (PDF input), Presenton (PDF/DOCX/PPTX upload), Aurora (manuscript-to-slides).

---

## 5. Add Chart/Table Best Practices and Layout Rules

**Gap:** The Playbooks pptx-agent has explicit rules about chart/table layout that prevent common failures: "Two-column layout PREFERRED for charts (header full-width, 40%/60% split)", "NEVER vertically stack charts below text", and "full-slide layout for maximum impact." Our skill shows how to create charts but gives no guidance on where to place them or how to integrate them with surrounding content. The current chart section in pptxgenjs.md only covers API syntax.

**Proposed Change:** Expand the Charts section in `pptxgenjs.md` with layout integration rules: (a) preferred layout patterns for data slides (two-column with narrative left / chart right, full-width chart with annotation bar, dashboard grid for multiple metrics), (b) sizing guidelines (charts should occupy at least 60% of content area), (c) annotation best practices (callout boxes for key numbers, trend arrows), (d) when to use which chart type (comparisons → bar, trends → line, composition → pie/donut, relationship → scatter). Add a "Data Slides" subsection to the Design Ideas in SKILL.md.

**Source Skills:** Playbooks pptx-agent (explicit chart layout rules), SlideSpeak/Copilot (consulting-grade chart generation), Presenton (smart data visualization).
