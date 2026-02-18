---
name: docx-msft
description: "Use this skill whenever creating Word documents (.docx) that should use the Microsoft brand template. This skill extends the base docx skill with the MS Simplicity brand design system — Segoe UI typography, MS brand color palette, and pre-defined paragraph styles. Triggers when user asks for branded documents, Microsoft-styled reports, memos, or any .docx that should follow corporate formatting. This skill provides the style definitions so you apply them correctly via docx-js or XML editing."
---

# Microsoft Brand DOCX Skill

This skill extends the base `docx` skill with your Microsoft brand template (`Brand_template_20200929.dotx`). **Read the base docx skill first** for general docx creation/editing workflows — this document only covers brand-specific additions.

## Quick Reference

| Task | Approach |
|------|-------|
| Create branded document (programmatic) | Use `docx-js` with the style definitions below |
| Create branded document (from template) | `python scripts/new_from_template.py output.docx` then unpack/edit/repack |
| **Extract docx → Markdown (for AI)** | `python scripts/docx_to_markdown.py input.docx output.md` |
| **Reassemble Markdown → docx** | `python scripts/markdown_to_docx.py input.md output.docx` |
| Inspect template styles | `python scripts/extract_styles.py` |
| All other docx tasks | Use the base `docx` skill |

---

## Creating from Template

The fastest way to create a branded document is to start from the template:

```bash
# Create a new .docx from the brand template
python scripts/new_from_template.py output.docx

# Then use the base docx skill's unpack/edit/repack workflow
python ../docx/scripts/office/unpack.py output.docx unpacked/
# ... edit XML in unpacked/word/ ...
python ../docx/scripts/office/pack.py unpacked/ output.docx --original output.docx
```

When editing an unpacked template-based document, all the styles below are already available in the XML. Apply them with `<w:pStyle w:val="StyleId"/>` in paragraph properties.

---

## MS Brand Design System

### Color Palette (Theme: "MS Simplicity")

| Role | Hex | Name | Usage |
|------|-----|------|-------|
| **Dark 1 (dk1)** | `#000000` | Black | Primary text |
| **Light 1 (lt1)** | `#FFFFFF` | White | Backgrounds |
| **Dark 2 (dk2)** | `#636466` | Gray | Secondary text |
| **Light 2 (lt2)** | `#D2D2D2` | Light Gray | Subtle backgrounds, table alternating rows |
| **Accent 1** | `#00BCF2` | Azure Blue | Primary accent — links, table headers, borders |
| **Accent 2** | `#E81123` | Red | Alerts, emphasis |
| **Accent 3** | `#7FBA00` | Green | Success, table first-column highlight |
| **Accent 4** | `#5C005C` | Purple | Secondary accent |
| **Accent 5** | `#FFB900` | Yellow/Amber | Warnings, callouts |
| **Accent 6** | `#737373` | Medium Gray | Neutral elements |
| **Hyperlink** | `#505050` | Dark Gray | In-text hyperlinks |
| **Followed Hyperlink** | `#00BCF2` | Azure Blue | Visited links |

### Typography

| Role | Font | Fallback |
|------|------|----------|
| **Major (headings)** | Segoe UI Light | Segoe UI |
| **Minor (body)** | Segoe UI | Arial |

**Default document font:** Arial 11pt (from `<w:docDefaults>`)

---

## Paragraph Styles

These are the styles defined in the template. **Use the Style ID** (not the display name) when referencing in XML or docx-js.

### Primary Content Styles (use these most often)

| Style ID | Display Name | Font | Size | Weight | Spacing | Usage |
|----------|-------------|------|------|--------|---------|-------|
| `Covertitle` | Cover title | Segoe UI Semibold | 42pt | Semibold | before: 2400 twips, line: 228% | Document cover page title |
| `TitlePageSubhead` | Title Page Subhead | Segoe UI Semibold | 16pt | Semibold | before: 240 twips | Subtitle on cover page, has white top border for visual separation |
| `Heading1` | Heading 1 | Segoe UI Semibold | 26pt | Semibold | line: 228% | Major section headings (outline level 0) |
| `Heading2` | Heading 2 | Segoe UI Semibold | 16pt | Semibold | line: 228% | Sub-section headings (outline level 1), keepNext/keepLines |
| `Heading3` | Heading 3 | Segoe UI Semibold | 10pt | Semibold | before: 120, after: 120 | Minor headings (outline level 2) |
| `BodyCopy` | Body Copy | Segoe UI | 10pt | Regular | after: 120 twips | **Primary body text** — use this instead of Normal |
| `IntroCopy` | Intro Copy | Segoe UI | 16pt | Regular | after: 120 twips | Lead-in paragraph, larger body text for introductions |
| `BodyCopyBulleted` | Body Copy_Bulleted | Segoe UI | 10pt | Regular | after: 120, indent: 216 twips | Bulleted list items |
| `BodyCopyNumbered` | Body Copy_Numbered | Segoe UI | 10pt | Regular | after: 120, indent: 288 twips | Numbered list items |
| `Subheading1` | Subheading 1 | Segoe UI (theme minor) | 10pt | Bold | after: 120 twips | Bold inline subheading within body |
| `Subheading2Nospaceafter` | Subheading 2_No space after | Segoe UI (theme minor) | 10pt | Bold | after: 0 | Subheading immediately followed by content (no gap) |
| `Attribution` | Attribution | Segoe UI | 10pt | Regular | before: 120, after: 120 | Source citations, attributions |
| `Legalese` | Legalese | Segoe UI | 7pt | Regular | after: 120, line: 180 twips min | Fine print, disclaimers, legal text |

### Heading 5 / Quote Attribution

| Style ID | Display Name | Font | Size | Usage |
|----------|-------------|------|------|-------|
| `Heading5` | Quote_Attribution | Segoe UI (theme minor) | 10pt | Used as quote attribution (aliased), outline level 4 |

### Table Styles

| Style ID | Display Name | Description |
|----------|-------------|-------------|
| `TableGrid` | Table Grid | Clean table with no cell margins (left/right = 0), rows can't split |
| `Placeholder` | Placeholder | Branded table — Azure Blue (#00BCF2) borders, blue header row with white bold text, green first column, alternating gray rows |

### System/Internal Styles (rarely used directly)

| Style ID | Usage |
|----------|-------|
| `Normal` | Default paragraph — Segoe UI 10pt, after: 120 twips. **Prefer `BodyCopy` instead** |
| `Header` | Page header paragraphs |
| `Footer` | Page footer paragraphs — no spacing after, negative right indent |
| `TOC1` | Table of contents level 1 — dot leader at 7891 twips |
| `TOC2` | Table of contents level 2 — indented 220 twips |

---

## Using Styles in docx-js

When creating documents programmatically with docx-js, define the styles to match the template:

```javascript
const { Document, Packer, Paragraph, TextRun, HeadingLevel,
        AlignmentType, LevelFormat, TableOfContents } = require('docx');
const fs = require('fs');

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "Segoe UI", size: 20, color: "000000" }  // 10pt
      }
    },
    paragraphStyles: [
      // Cover title — 42pt Segoe UI Semibold
      {
        id: "Covertitle", name: "Cover title",
        run: { font: "Segoe UI Semibold", size: 84, color: "000000" },
        paragraph: { spacing: { before: 2400, line: 228, lineRule: "auto" } }
      },
      // Title Page Subhead — 16pt Segoe UI Semibold
      {
        id: "TitlePageSubhead", name: "Title Page Subhead",
        run: { font: "Segoe UI Semibold", size: 32, color: "000000" },
        paragraph: { spacing: { before: 240 } }
      },
      // Heading 1 — 26pt Segoe UI Semibold
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "BodyCopy",
        quickFormat: true,
        run: { font: "Segoe UI Semibold", size: 52, color: "000000" },
        paragraph: { spacing: { line: 228, lineRule: "auto" }, outlineLevel: 0 }
      },
      // Heading 2 — 16pt Segoe UI Semibold
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "BodyCopy",
        quickFormat: true,
        run: { font: "Segoe UI Semibold", size: 32, color: "000000" },
        paragraph: {
          keepNext: true, keepLines: true,
          spacing: { line: 228, lineRule: "auto" }, outlineLevel: 1
        }
      },
      // Heading 3 — 10pt Segoe UI Semibold
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "BodyCopy",
        run: { font: "Segoe UI Semibold", size: 20, color: "000000" },
        paragraph: { spacing: { before: 120, after: 120 }, outlineLevel: 2 }
      },
      // Body Copy — 10pt Segoe UI (PRIMARY body style)
      {
        id: "BodyCopy", name: "Body Copy",
        quickFormat: true,
        run: { font: "Segoe UI", size: 20, color: "000000" },
        paragraph: { spacing: { after: 120 } }
      },
      // Intro Copy — 16pt Segoe UI
      {
        id: "IntroCopy", name: "Intro Copy", basedOn: "Normal",
        quickFormat: true,
        run: { font: "Segoe UI", size: 32, color: "000000" },
        paragraph: { spacing: { after: 120 } }
      },
      // Attribution — 10pt Segoe UI
      {
        id: "Attribution", name: "Attribution", basedOn: "BodyCopy",
        quickFormat: true,
        run: { font: "Segoe UI", size: 20 },
        paragraph: { spacing: { before: 120, after: 120 } }
      },
      // Legalese — 7pt Segoe UI
      {
        id: "Legalese", name: "Legalese",
        quickFormat: true,
        run: { font: "Segoe UI", size: 14, color: "000000" },
        paragraph: { spacing: { after: 120 } }
      },
    ]
  },
  numbering: {
    config: [
      // Bulleted list (matches Body Copy_Bulleted)
      {
        reference: "ms-bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: { indent: { left: 216, hanging: 216 } },
            run: { font: "Segoe UI", size: 20 }
          }
        }]
      },
      // Numbered list (matches Body Copy_Numbered)
      {
        reference: "ms-numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.",
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: { indent: { left: 288, hanging: 288 } },
            run: { font: "Segoe UI", size: 20 }
          }
        }]
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },  // US Letter
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      // Cover title
      new Paragraph({ style: "Covertitle",
        children: [new TextRun("Document Title")] }),
      // Subhead
      new Paragraph({ style: "TitlePageSubhead",
        children: [new TextRun("Subtitle or date")] }),
      // Body content
      new Paragraph({ style: "BodyCopy",
        children: [new TextRun("Body text goes here.")] }),
      // Bullet
      new Paragraph({
        numbering: { reference: "ms-bullets", level: 0 },
        children: [new TextRun("First bullet point")]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => fs.writeFileSync("branded-doc.docx", buf));
```

### Style Selection Guide

| Content Type | Style to Use |
|-------------|-------------|
| Document title on cover page | `Covertitle` |
| Subtitle / date on cover page | `TitlePageSubhead` |
| Major section heading | `Heading1` (or `heading: HeadingLevel.HEADING_1`) |
| Sub-section heading | `Heading2` (or `heading: HeadingLevel.HEADING_2`) |
| Minor heading | `Heading3` |
| Regular body paragraphs | `BodyCopy` |
| Opening/intro paragraph (larger) | `IntroCopy` |
| Bold inline label | `Subheading1` |
| Bold label with no gap after | `Subheading2Nospaceafter` |
| Bullet list items | numbering ref `ms-bullets` |
| Numbered list items | numbering ref `ms-numbers` |
| Source/attribution line | `Attribution` |
| Fine print / disclaimers | `Legalese` |
| Quote attribution | `Heading5` |

---

## Using Styles in XML (Template-Based Editing)

When editing an unpacked template-based document, apply styles in the paragraph properties:

```xml
<!-- Body Copy paragraph -->
<w:p>
  <w:pPr>
    <w:pStyle w:val="BodyCopy"/>
  </w:pPr>
  <w:r>
    <w:t>This text will use Body Copy formatting.</w:t>
  </w:r>
</w:p>

<!-- Heading 1 -->
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
  </w:pPr>
  <w:r>
    <w:t>Section Title</w:t>
  </w:r>
</w:p>

<!-- Bulleted item (uses numId 39 from template) -->
<w:p>
  <w:pPr>
    <w:pStyle w:val="BodyCopyBulleted"/>
  </w:pPr>
  <w:r>
    <w:t>A bullet point</w:t>
  </w:r>
</w:p>

<!-- Numbered item (uses numId 38 from template) -->
<w:p>
  <w:pPr>
    <w:pStyle w:val="BodyCopyNumbered"/>
  </w:pPr>
  <w:r>
    <w:t>Step one</w:t>
  </w:r>
</w:p>
```

---

## Brand Table (Placeholder Style)

The template includes a branded table style called "Placeholder" with:
- **Azure Blue (#00BCF2) borders** on all sides
- **Header row**: Azure Blue background with white bold text
- **First column**: Green (#7FBA00) background with bold text
- **Alternating rows**: Light gray (#B2B2B2) shading on even rows
- **Cell margins**: 115 twips left/right
- **Rows can't split** across pages

To apply in XML:
```xml
<w:tbl>
  <w:tblPr>
    <w:tblStyle w:val="Placeholder"/>
    <w:tblW w:w="9360" w:type="dxa"/>
    <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0"
               w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
  </w:tblPr>
  <!-- rows... -->
</w:tbl>
```

---

## Round-Trip Markdown Workflow (AI Processing)

This skill supports a round-trip workflow for AI-assisted document editing:

```
.docx  ──→  .md (with style tags)  ──→  AI processes  ──→  .docx (styled)
```

### Extract: docx → Markdown

```bash
python scripts/docx_to_markdown.py document.docx output.md
```

Produces a markdown file with:
- **YAML front matter** — title, author, company, creation date, template name, and all styles used
- **Style annotations** — `<!-- style: StyleName -->` HTML comments before each styled section
- **Standard markdown** — `#` headings, `**bold**`, `*italic*`, `- bullets`, `1. numbered`, `[links](url)`, `| tables |`
- **Extracted images** — saved to `output_media/` folder, referenced as `![alt](output_media/image.png)`

**Example extracted output:**

```markdown
---
title: Quarterly Report
author: David B
company: Microsoft
source: report.docx
extracted: 2026-02-17T21:05:58
template: Brand_template_20200929.dotx
styles_used:
  - BodyCopy
  - BodyCopyBulleted
  - Heading1
  - Heading2
  - IntroCopy
---

<!-- style: Covertitle -->
Quarterly Business Report
<!-- style: TitlePageSubhead -->
Q4 2025 Results

<!-- pagebreak -->

<!-- style: Heading1 -->
# Executive Summary

<!-- style: IntroCopy -->
This quarter delivered strong results across all business units.

<!-- style: BodyCopy -->
Revenue grew **15%** year-over-year driven by cloud services.

<!-- style: Heading2 -->
## Key Metrics

<!-- style: BodyCopyBulleted -->
- Cloud revenue: $28.5B
- Operating income: $12.1B
- Free cash flow: $9.8B
```

### Reassemble: Markdown → docx

```bash
python scripts/markdown_to_docx.py output.md final.docx
```

Reads the annotated markdown and creates a fully styled `.docx` from the brand template:
- Applies paragraph styles from `<!-- style: X -->` annotations
- Maps `#`/`##`/`###` to Heading1/2/3 (or uses explicit style annotation)
- Converts `- ` to BodyCopyBulleted, `1. ` to BodyCopyNumbered (with template numbering)
- Converts `**bold**`, `*italic*`, `~~strike~~` to formatted runs
- Embeds hyperlinks with proper OOXML relationships
- Embeds images from the media folder with computed dimensions
- Preserves template headers, footers, theme, and page layout

### Style Annotation Rules

| Annotation | Effect |
|-----------|--------|
| `<!-- style: BodyCopy -->` | Sets style for subsequent paragraphs until next annotation |
| `<!-- style: Heading2 -->` followed by `## text` | Heading2 style applied (annotation + markdown heading) |
| `<!-- style: BodyCopyBulleted -->` followed by `- items` | Bulleted list with correct numbering |
| `<!-- pagebreak -->` | Inserts a page break |
| No annotation | Defaults to `BodyCopy` |

### Typical AI Workflow

1. **Extract** the source document:
   ```bash
   python scripts/docx_to_markdown.py original.docx working.md
   ```
2. **Process** `working.md` with AI (edit content, restructure, translate, summarize, etc.)
3. **Reassemble** into a new branded document:
   ```bash
   python scripts/markdown_to_docx.py working.md final.docx
   ```
4. **Review** in Word — all styles, headers, footers, and theme are applied

### Limitations

- **Complex tables** — Merged cells and multi-line cells are simplified to single-line markdown tables
- **Drawing objects** — Text boxes, shapes, and SmartArt are not extracted (only inline images)
- **Section breaks** — Multiple sections with different layouts are collapsed to the template's default
- **Comments/tracked changes** — Not preserved in the markdown round-trip
- **Image sizing** — Reassembled images use computed dimensions from pixel data; may differ from original layout

---

## Avoid (Common Mistakes)

- **Don't use `Normal` for body text** — use `BodyCopy` instead (same formatting but proper style hierarchy)
- **Don't use Arial** — use Segoe UI (body) and Segoe UI Semibold (headings)
- **Don't hardcode font sizes** — apply the named styles; they carry the correct sizes
- **Don't use generic blue** — use the brand Azure Blue `#00BCF2` (Accent 1) for any accent color
- **Don't skip the cover page styles** — use `Covertitle` + `TitlePageSubhead` for the first page
- **Don't use `TableGrid` for branded tables** — use `Placeholder` for the full branded look
- **Don't mix Segoe UI Light into body text** — it's the theme major font (headings only)

---

## Dependencies

Same as the base `docx` skill, plus:
- **Python 3** — for all scripts (no external packages required — stdlib only)
- The template file `Brand_template_20200929.dotx` (included in this skill directory)

### Scripts

| Script | Purpose |
|--------|--------|
| `scripts/new_from_template.py` | Create a new clean `.docx` from the brand template (strips sample content) |
| `scripts/extract_styles.py` | Display all styles, theme colors, and fonts defined in the template |
| `scripts/docx_to_markdown.py` | Extract a `.docx` to Markdown with YAML front matter and style annotations |
| `scripts/markdown_to_docx.py` | Reassemble annotated Markdown into a styled `.docx` using the brand template |
