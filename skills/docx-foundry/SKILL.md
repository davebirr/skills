---
name: docx-foundry
description: "Use this skill when the user wants to decompose, disassemble, or modularize Word documents (.docx) into reusable AI-editable components, or reassemble modular components back into a polished .docx. Triggers include: breaking a document into sections, extracting styles and formatting from a source document, creating a component library from multiple source documents, splitting a docx into markdown parts for AI editing, analyzing document structure for remixing or rebuilding, improving individual sections of a professional document, or combining sections from multiple sources into a new document. Also triggers when the user wants to maintain consistent look and feel across documents built from modular parts. This skill extends the base docx skill — use it for structural decomposition and reassembly workflows, not basic document creation or editing."
---

# DOCX Foundry — Decompose, Improve, Reassemble

Extends the base `docx` skill. **Read the base docx skill first** for general docx creation/editing — this skill covers the decompose → edit → reassemble workflow.

## Quick Reference

| Task | Approach |
|------|----------|
| Decompose a document | `python scripts/decompose.py input.docx [output_dir]` |
| Review structure | Read `manifest.yaml` for section index and stats |
| Review/edit styles | Read and modify `styles.yaml` |
| Edit a section | Modify markdown files in `sections/` |
| Create new section | Create new `.md` with YAML front matter matching schema |
| Mix sources | Decompose multiple docs, combine sections in a new manifest |
| Reassemble to .docx | Use base `docx` skill (docx-js or unpack/edit/repack) with styles from `styles.yaml` |
| Validate output | Create sample .docx, ask user to review spacing/layout |

## Step 1: Decompose

### Run the Decomposer

```bash
python scripts/decompose.py input.docx output_dir
python scripts/decompose.py input.docx output_dir --split-on 2   # split on Heading2
```

Auto-detects the highest heading level used for splitting. Override with `--split-on N`.

### Output Structure

```
output_dir/
  manifest.yaml          # Section index, metadata, stats
  styles.yaml            # Full style definitions (editable)
  sections/
    00-front-matter.md   # Content before first heading
    01-section-name.md   # One file per heading section
    ...
  media/
    image1.png           # All extracted images
    ...
```

See [references/schemas.md](references/schemas.md) for complete field definitions.

### Review After Decomposition

1. **Check manifest.yaml** — verify sections were split correctly. If sections are too granular or too broad, re-run with a different `--split-on` level.
2. **Check styles.yaml** — look for the `anomalies` section. If present, **alert the user**: the source document has inconsistent styling (e.g., paragraphs that should be "BodyCopy" but have direct font-size overrides). Ask if they want to normalize these.
3. **Spot-check 2-3 section files** — verify content and style annotations look correct.

## Step 2: Work with Components

### Editing Sections

Each section file is standalone markdown with YAML front matter and `<!-- style: StyleName -->` annotations:

```markdown
---
section_index: 1
title: "Introduction"
heading_level: 1
source_styles:
  - BodyCopy
  - BodyCopyBulleted
  - Heading2
word_count: 342
images:
  - image1.png
---

# Introduction

<!-- style: IntroCopy -->
This guide provides an overview of the security features.

<!-- style: BodyCopy -->
Microsoft 365 Business Premium includes comprehensive tools.

<!-- style: BodyCopyBulleted -->
- Advanced threat protection
- Device management
```

**Rules for editing:**
- Keep `<!-- style: X -->` annotations — they drive formatting during reassembly
- Update `word_count` in front matter after significant edits
- If adding images, add the filename to the `images` list and place the file in `media/`
- Use standard markdown: `#` headings, `**bold**`, `*italic*`, `- bullets`, `1. numbered`

### Creating New Sections

Copy the front matter pattern from an existing section. Set `section_index` to the desired position and update `manifest.yaml` to include the new file.

### Mixing Components from Multiple Sources

Decompose each source document into its own output directory:

```bash
python scripts/decompose.py source1.docx source1_parts
python scripts/decompose.py source2.docx source2_parts
```

Then build a new manifest:
1. Pick sections from each decomposed source
2. Copy selected section files and their images into a combined output directory
3. Create a new `manifest.yaml` listing the sections in desired order
4. Merge `styles.yaml` files — when styles conflict, choose which source's definitions to keep. **Alert the user** if the same style name (e.g., "Heading1") has different definitions across sources.

### Improving Content with AI

The decomposed format is designed for AI editing. For each section:
1. Read the section markdown
2. Understand the style annotations (what formatting will be applied)
3. Improve the content while preserving structure and style annotations
4. For substantial rewrites, preserve the same heading hierarchy

**When improving sections, ask the user what they want changed** — don't assume. The content may be authoritative professional material where accuracy matters more than prose style.

## Step 3: Reassemble

Use the base `docx` skill to build a new .docx from the modular components. Two approaches:

### Option A: docx-js (Programmatic)

Best for building from scratch without a template. Read `styles.yaml` to generate the style definitions.

1. Read `manifest.yaml` for section order
2. Read `styles.yaml` to build the `styles` config object
3. For each section, parse the markdown and style annotations
4. Generate docx-js code that applies the correct styles

**Mapping styles.yaml to docx-js:**

```javascript
// From styles.yaml paragraph_styles.Heading1:
//   font: { name: "Segoe UI Semibold", size_pt: 26, bold: true }
//   paragraph: { spacing: { before: 240 }, outline_level: 0 }

paragraphStyles: [{
  id: "Heading1", name: "Heading 1",
  basedOn: "Normal", next: "BodyCopy",
  quickFormat: true,
  run: { font: "Segoe UI Semibold", size: 52, bold: true },  // size = pt * 2
  paragraph: {
    spacing: { before: 240 },
    outlineLevel: 0,
  }
}]
```

Key conversions from `styles.yaml` to docx-js:
- `size_pt` → multiply by 2 for docx-js `size` (half-points)
- `spacing.before/after` → use directly (already in DXA/twips)
- `outline_level` → `outlineLevel`
- `keep_next` → `keepNext`

### Option B: Template-Based (Unpack/Edit/Repack)

Best when you have a source .docx or .dotx with the right styles already defined.

1. Start from the source document or a template:
   ```bash
   python ../docx/scripts/office/unpack.py template.docx unpacked/
   ```
2. Replace the body content in `unpacked/word/document.xml` with content from the section files
3. Apply styles by setting `<w:pStyle w:val="StyleName"/>` per the `<!-- style: X -->` annotations
4. Add images to `unpacked/word/media/` and update relationships
5. Repack:
   ```bash
   python ../docx/scripts/office/pack.py unpacked/ output.docx --original template.docx
   ```

### Choosing an Approach

| Factor | docx-js | Template-based |
|--------|---------|----------------|
| No existing template | Best choice | Need to create one first |
| Complex headers/footers | Must code manually | Inherited from template |
| Branded tables | Must code styling | Style already defined |
| Theme colors/fonts | Must define in code | Inherited from template |
| Full control over output | High | High |
| Speed for simple docs | Faster | More setup |

**Recommendation:** If the source document has complex formatting (branded tables, headers with logos, custom numbering), use template-based. If building a clean document from content, use docx-js.

## Step 4: Validate

### Create Output Samples

After reassembly, **always create a sample and ask the user to review it.** This catches issues that are hard to detect programmatically:

- Paragraph spacing (too tight? too loose?)
- Table column widths
- Image sizing and placement
- Header/footer layout
- Font rendering
- Page break positions

### What to Check

1. **Open the .docx in Word** — verify it renders correctly
2. **Compare side-by-side** with source document (if improving, not creating new)
3. **Check page count** — significant differences suggest layout/spacing issues
4. **Print preview** — catches margin and header/footer problems

### Handling Style Inconsistencies

When `styles.yaml` contains an `anomalies` section:

1. **Show anomalies to the user**: "The source document has inconsistent styling — some paragraphs with style 'BodyCopy' override the font size directly. Want to normalize these?"
2. **If normalizing**: Remove the direct overrides and rely on the style definition
3. **If keeping**: Note in the manifest that direct formatting exists
4. **Ask for screenshots** if uncertain: "Can you screenshot pages 3-4? I want to verify whether the font size variation is intentional or accidental."

### Style Discoverability

When styles.yaml doesn't fully capture the source look and feel:

1. Ask the user for details: "What font should headings use? What's the accent color?"
2. Ask for screenshots of the source document for reference
3. Create a 1-page test document with sample headings, body text, bullets, and a table
4. Ask the user to compare: "Does this match the source? What needs adjusting?"
5. Iterate on `styles.yaml` until the output looks right

## Edge Cases

- **No headings in document**: Everything goes into `00-front-matter.md`. Suggest manual splitting.
- **Cover pages**: Typically captured as front matter (section 00). May use unique styles like "CoverTitle" — preserved in style annotations.
- **Table of Contents**: Decomposed as content with ToC styles. When reassembling, use `TableOfContents` from base docx skill instead of the extracted text.
- **Embedded charts/SmartArt**: Not extracted as editable content. Images are extracted as PNG fallbacks. Alert the user that these won't round-trip as editable objects.
- **Track changes / comments**: Not preserved during decomposition. If needed, work with the base docx skill's unpack/edit workflow directly.
- **Headers/footers with images**: Image references captured in `styles.yaml` headers_footers section. For reassembly, embed these separately.

## Dependencies

- **Python 3.6+**: For decompose.py (uses only stdlib: zipfile, xml.etree, argparse)
- **Base docx skill**: For reassembly (docx-js, unpack/pack scripts, pandoc)
