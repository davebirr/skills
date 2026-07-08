# DOCX Foundry Schemas

Field definitions for the files produced by `decompose.py`.

## manifest.yaml

| Field | Type | Description |
|-------|------|-------------|
| `source` | string | Original .docx filename |
| `decomposed` | string | ISO timestamp of decomposition |
| `metadata` | object | Document properties from docProps |
| `metadata.title` | string | Document title |
| `metadata.author` | string | Author name |
| `metadata.company` | string | Company name |
| `metadata.template` | string | Template name used by source |
| `metadata.created` | string | Creation timestamp |
| `metadata.modified` | string | Last modified timestamp |
| `metadata.pages` | int | Page count |
| `page_layout` | object | Page dimensions and margins (DXA units, 1440 = 1 inch) |
| `page_layout.width` | int | Page width (12240 = US Letter) |
| `page_layout.height` | int | Page height (15840 = US Letter) |
| `page_layout.orientation` | string | "landscape" if set, absent for portrait |
| `page_layout.margins` | object | top, right, bottom, left, header, footer, gutter |
| `page_layout.columns` | int | Column count (only if > 1) |
| `sections` | list | Ordered list of section entries |
| `sections[].file` | string | Relative path to section markdown file |
| `sections[].title` | string | Section heading text |
| `sections[].heading_level` | int | Heading level (0 = front matter) |
| `sections[].word_count` | int | Approximate word count |
| `sections[].images` | list | Filenames of images used in this section |
| `styles_file` | string | Path to styles.yaml |
| `media_dir` | string | Path to media directory |
| `stats` | object | Summary statistics |
| `anomalies_count` | int | Number of style inconsistencies (if any) |

## styles.yaml

| Field | Type | Description |
|-------|------|-------------|
| `theme` | object | Theme colors and fonts |
| `theme.name` | string | Theme name |
| `theme.colors` | object | Color scheme: dk1, dk2, lt1, lt2, accent1-6, hlink, folHlink |
| `theme.fonts` | object | major (headings), minor (body) font families |
| `defaults` | object | Document default font properties |
| `defaults.name` | string | Default font name |
| `defaults.size_pt` | number | Default font size in points |
| `paragraph_styles` | object | Keyed by style ID |
| `paragraph_styles.{ID}.name` | string | Display name |
| `paragraph_styles.{ID}.based_on` | string | Parent style ID |
| `paragraph_styles.{ID}.next` | string | Next paragraph style ID |
| `paragraph_styles.{ID}.quick_format` | bool | Available in Quick Styles gallery |
| `paragraph_styles.{ID}.font` | object | Font properties |
| `paragraph_styles.{ID}.font.name` | string | Font family name |
| `paragraph_styles.{ID}.font.theme` | string | Theme font reference (e.g., "majorHAnsi") |
| `paragraph_styles.{ID}.font.size_pt` | number | Font size in points |
| `paragraph_styles.{ID}.font.color` | string | Hex color (e.g., "2E74B5") |
| `paragraph_styles.{ID}.font.theme_color` | string | Theme color reference |
| `paragraph_styles.{ID}.font.bold` | bool | Bold flag |
| `paragraph_styles.{ID}.font.italic` | bool | Italic flag |
| `paragraph_styles.{ID}.paragraph` | object | Paragraph properties |
| `paragraph_styles.{ID}.paragraph.spacing` | object | before, after (DXA), line, line_rule |
| `paragraph_styles.{ID}.paragraph.indent` | object | left, right, hanging, firstLine (DXA) |
| `paragraph_styles.{ID}.paragraph.alignment` | string | left, center, right, both |
| `paragraph_styles.{ID}.paragraph.outline_level` | int | 0-8 (for TOC/navigation) |
| `paragraph_styles.{ID}.paragraph.keep_next` | bool | Keep with next paragraph |
| `character_styles` | object | Same font sub-structure as paragraph styles |
| `table_styles` | object | Table style definitions |
| `numbering` | object | List/numbering definitions |
| `numbering.abstract` | object | Abstract numbering templates keyed by ID |
| `numbering.instances` | object | Numbering instances mapping numId → abstractNumId |
| `headers_footers` | object | Extracted text content |
| `anomalies` | list | Style inconsistencies found in source |
| `anomalies[].style` | string | Style ID with the issue |
| `anomalies[].note` | string | Human-readable description |

## Section Markdown (.md) Front Matter

| Field | Type | Description |
|-------|------|-------------|
| `section_index` | int | Position in document (0-based) |
| `title` | string | Section heading text |
| `heading_level` | int | Heading level (0 = front matter / pre-heading content) |
| `source_styles` | list | Style IDs used in this section |
| `word_count` | int | Approximate word count |
| `images` | list | Image filenames referenced (optional, only if images present) |

### Style Annotations in Section Body

| Annotation | Meaning |
|------------|---------|
| `<!-- style: StyleName -->` | Following paragraphs use this paragraph style |
| `<!-- table-style: StyleName -->` | Following table uses this table style |
| `<!-- pagebreak -->` | Page break in source document |

### Content Conventions

| Markdown | Maps to |
|----------|---------|
| `# Heading` | Heading1 (or style annotation above) |
| `## Heading` | Heading2 |
| `- item` | Bulleted list (style annotation specifies which list style) |
| `1. item` | Numbered list |
| `**bold**` | Bold run |
| `*italic*` | Italic run |
| `~~strike~~` | Strikethrough |
| `[text](url)` | External hyperlink |
| `![alt](media/file.png)` | Inline image |
| `\| table \|` | Markdown table (table-style annotation specifies formatting) |

## Units Reference

| Unit | Conversion | Used for |
|------|-----------|----------|
| DXA (twips) | 1440 = 1 inch, 20 = 1 point | Spacing, margins, page size, indents |
| Half-points | 2 = 1 point | Internal XML font size (converted to pt in styles.yaml) |
| EMU | 914400 = 1 inch | Image dimensions in XML |
| Points (pt) | 72 = 1 inch | Font sizes in styles.yaml |
