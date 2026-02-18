#!/usr/bin/env python3
"""Extract a .docx to Markdown with YAML front matter and style annotations.

Creates a markdown file that preserves document structure and style information,
suitable for AI processing. The output can be reassembled into a styled .docx
using markdown_to_docx.py.

Output format:
  - YAML front matter: document metadata (title, author, template, styles_used)
  - Markdown body with <!-- style: StyleName --> annotations before paragraphs
  - Standard markdown for headings (#), lists (- / 1.), bold (**), italic (*)
  - Images extracted to {output_stem}_media/ directory

Usage:
    python docx_to_markdown.py input.docx [output.md]

If output.md is omitted, uses the same stem as the input (input.md).
"""

import sys
import os
import re
import struct
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

# ──────────────────────────────────────────────────────────────────────
# XML Namespaces
# ──────────────────────────────────────────────────────────────────────

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
DC = 'http://purl.org/dc/elements/1.1/'
DCTERMS = 'http://purl.org/dc/terms/'

# Shorthand for Clark notation
def w(tag): return f'{{{W}}}{tag}'
def r(tag): return f'{{{R}}}{tag}'

# ──────────────────────────────────────────────────────────────────────
# Style → Markdown Mappings
# ──────────────────────────────────────────────────────────────────────

# Styles that naturally map to markdown headings
HEADING_MAP = {
    'Heading1': 1,
    'Heading2': 2,
    'Heading3': 3,
    'Heading4': 4,
    'Heading5': 5,
}

# Styles that map to list markers
LIST_STYLES = {
    'BodyCopyBulleted': 'bullet',
    'BodyCopyNumbered': 'numbered',
    'ListBullet': 'bullet',
    'ListNumber': 'numbered',
    'ListParagraph': 'bullet',
}

# Default body style (annotation omitted when this is current)
DEFAULT_STYLE = 'BodyCopy'


# ──────────────────────────────────────────────────────────────────────
# Metadata Extraction
# ──────────────────────────────────────────────────────────────────────

def extract_metadata(zf):
    """Extract document properties from docProps/core.xml and app.xml."""
    meta = {}

    # Core properties
    try:
        core = ET.fromstring(zf.read('docProps/core.xml'))
        for child in core:
            local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            ns = child.tag.split('}')[0][1:] if '}' in child.tag else ''
            if ns == DC:
                if local in ('title', 'creator', 'subject', 'description'):
                    key = 'author' if local == 'creator' else local
                    if child.text and child.text.strip():
                        meta[key] = child.text.strip()
            elif ns == DCTERMS:
                if local in ('created', 'modified') and child.text:
                    meta[local] = child.text.strip()
    except (KeyError, ET.ParseError):
        pass

    # App properties
    try:
        app = ET.fromstring(zf.read('docProps/app.xml'))
        for child in app:
            local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if local == 'Company' and child.text:
                meta['company'] = child.text.strip()
    except (KeyError, ET.ParseError):
        pass

    return meta


# ──────────────────────────────────────────────────────────────────────
# Relationship Parsing
# ──────────────────────────────────────────────────────────────────────

def parse_relationships(zf):
    """Parse word/_rels/document.xml.rels for hyperlink and image targets."""
    rels = {}
    try:
        root = ET.fromstring(zf.read('word/_rels/document.xml.rels'))
        for el in root:
            rid = el.get('Id', '')
            rels[rid] = {
                'target': el.get('Target', ''),
                'type': el.get('Type', ''),
                'external': el.get('TargetMode', '') == 'External',
            }
    except (KeyError, ET.ParseError):
        pass
    return rels


# ──────────────────────────────────────────────────────────────────────
# Paragraph Helpers
# ──────────────────────────────────────────────────────────────────────

def get_style(el):
    """Get the style ID from a paragraph element."""
    pPr = el.find(w('pPr'))
    if pPr is not None:
        pStyle = pPr.find(w('pStyle'))
        if pStyle is not None:
            return pStyle.get(w('val'), '')
    return ''


def get_table_style(tbl):
    """Get the style ID from a table element."""
    tblPr = tbl.find(w('tblPr'))
    if tblPr is not None:
        tblStyle = tblPr.find(w('tblStyle'))
        if tblStyle is not None:
            return tblStyle.get(w('val'), '')
    return ''


def get_num_info(el):
    """Get numbering info (numId, ilvl) from a paragraph."""
    pPr = el.find(w('pPr'))
    if pPr is not None:
        numPr = pPr.find(w('numPr'))
        if numPr is not None:
            ilvl_el = numPr.find(w('ilvl'))
            numId_el = numPr.find(w('numId'))
            ilvl = int(ilvl_el.get(w('val'), '0')) if ilvl_el is not None else 0
            numId = numId_el.get(w('val'), '') if numId_el is not None else ''
            return {'ilvl': ilvl, 'numId': numId}
    return None


def has_page_break(el):
    """Check if paragraph contains a page break."""
    # In runs
    for run in el.findall(w('r')):
        br = run.find(w('br'))
        if br is not None and br.get(w('type')) == 'page':
            return True
    # In paragraph properties
    pPr = el.find(w('pPr'))
    if pPr is not None:
        if pPr.find(w('pageBreakBefore')) is not None:
            return True
    return False


# ──────────────────────────────────────────────────────────────────────
# Run Extraction
# ──────────────────────────────────────────────────────────────────────

def extract_run_format(run_el):
    """Get bold/italic/strike/underline flags from a run."""
    fmt = {}
    rPr = run_el.find(w('rPr'))
    if rPr is not None:
        for tag, key in [('b', 'bold'), ('i', 'italic'),
                         ('strike', 'strike'), ('u', 'underline')]:
            el = rPr.find(w(tag))
            if el is not None:
                val = el.get(w('val'), 'true')
                if val not in ('false', '0'):
                    fmt[key] = True
        rStyle = rPr.find(w('rStyle'))
        if rStyle is not None:
            fmt['char_style'] = rStyle.get(w('val'), '')
    return fmt


def extract_run_text(run_el):
    """Get text content from a run, handling tabs and breaks."""
    parts = []
    for child in run_el:
        local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if local == 't':
            parts.append(child.text or '')
        elif local == 'tab':
            parts.append('\t')
        elif local == 'br':
            if child.get(w('type')) == 'page':
                parts.append('\n<!-- pagebreak -->\n')
            else:
                parts.append('  \n')
    return ''.join(parts)


def apply_md_formatting(text, fmt):
    """Wrap text with markdown formatting."""
    if not text or not text.strip():
        return text
    stripped = text.strip()
    leading = text[:len(text) - len(text.lstrip())]
    trailing = text[len(text.rstrip()):]

    if fmt.get('bold') and fmt.get('italic'):
        stripped = f'***{stripped}***'
    elif fmt.get('bold'):
        stripped = f'**{stripped}**'
    elif fmt.get('italic'):
        stripped = f'*{stripped}*'
    if fmt.get('strike'):
        stripped = f'~~{stripped}~~'

    return f'{leading}{stripped}{trailing}'


# ──────────────────────────────────────────────────────────────────────
# Image Extraction
# ──────────────────────────────────────────────────────────────────────

def extract_image_from_run(run_el, rels, zf, media_dir):
    """If a run contains an image, extract it and return markdown ref."""
    # Look for drawing > inline or anchor > blip
    blip = run_el.find(f'.//{{{A}}}blip')
    if blip is None:
        return None

    embed_id = blip.get(f'{{{R}}}embed', '')
    if not embed_id or embed_id not in rels:
        return None

    target = rels[embed_id]['target']        # e.g. "media/image1.png"
    zip_path = f'word/{target}'

    if zf is None or media_dir is None:
        return f'![image]({target})'

    try:
        image_data = zf.read(zip_path)
    except KeyError:
        return f'![image]({target})'

    os.makedirs(media_dir, exist_ok=True)
    filename = os.path.basename(target)
    out_path = os.path.join(media_dir, filename)
    with open(out_path, 'wb') as f:
        f.write(image_data)

    media_dirname = os.path.basename(media_dir)
    return f'![{filename}]({media_dirname}/{filename})'


# ──────────────────────────────────────────────────────────────────────
# Paragraph → Markdown
# ──────────────────────────────────────────────────────────────────────

def paragraph_to_text(p_el, rels, zf=None, media_dir=None):
    """Convert a w:p element to markdown text (content only, no heading #)."""
    parts = []

    for child in p_el:
        local = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if local == 'r':
            # Check for image
            img_ref = extract_image_from_run(child, rels, zf, media_dir)
            if img_ref:
                parts.append(img_ref)
                continue
            text = extract_run_text(child)
            fmt = extract_run_format(child)
            parts.append(apply_md_formatting(text, fmt))

        elif local == 'hyperlink':
            r_id = child.get(f'{{{R}}}id', '')
            link_parts = []
            for run in child.findall(w('r')):
                link_parts.append(extract_run_text(run))
            link_text = ''.join(link_parts).strip()

            if r_id and r_id in rels and rels[r_id].get('external'):
                url = rels[r_id]['target']
                parts.append(f'[{link_text}]({url})')
            else:
                parts.append(link_text)

    return ''.join(parts)


# ──────────────────────────────────────────────────────────────────────
# Table → Markdown
# ──────────────────────────────────────────────────────────────────────

def table_to_markdown(tbl_el, rels, zf=None, media_dir=None):
    """Convert a w:tbl to markdown table syntax."""
    rows_data = []
    max_cols = 0

    for tr in tbl_el.findall(w('tr')):
        row = []
        for tc in tr.findall(w('tc')):
            cell_parts = []
            for p in tc.findall(w('p')):
                text = paragraph_to_text(p, rels, zf, media_dir).strip()
                if text:
                    cell_parts.append(text)
            row.append(' '.join(cell_parts) if cell_parts else '')
        rows_data.append(row)
        max_cols = max(max_cols, len(row))

    if not rows_data or max_cols == 0:
        return ''

    # Normalize row lengths
    for row in rows_data:
        while len(row) < max_cols:
            row.append('')

    # Column widths
    widths = [max(3, max(len(rows_data[r][c]) for r in range(len(rows_data))))
              for c in range(max_cols)]

    lines = []
    # Header
    lines.append('| ' + ' | '.join(
        rows_data[0][c].ljust(widths[c]) for c in range(max_cols)) + ' |')
    # Separator
    lines.append('|' + '|'.join('-' * (w + 2) for w in widths) + '|')
    # Body rows
    for row in rows_data[1:]:
        lines.append('| ' + ' | '.join(
            row[c].ljust(widths[c]) for c in range(max_cols)) + ' |')

    return '\n'.join(lines)


# ──────────────────────────────────────────────────────────────────────
# Body Processing
# ──────────────────────────────────────────────────────────────────────

def process_body(body, rels, zf=None, media_dir=None):
    """Walk the document body and produce (markdown_lines, styles_used)."""
    lines = []
    styles_used = set()
    current_style = None
    in_list = False
    list_style = None
    list_counter = 0

    for child in body:
        local = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        # ── Paragraphs ──
        if local == 'p':
            style = get_style(child) or 'Normal'
            text = paragraph_to_text(child, rels, zf, media_dir)
            num_info = get_num_info(child)

            if style:
                styles_used.add(style)

            # Page break
            if has_page_break(child):
                lines.append('')
                lines.append('<!-- pagebreak -->')
                lines.append('')

            # Empty paragraph → blank line
            if not text.strip():
                if in_list:
                    in_list = False
                    list_style = None
                    list_counter = 0
                lines.append('')
                continue

            # ── List items ──
            is_list = style in LIST_STYLES or num_info is not None
            if is_list:
                list_type = LIST_STYLES.get(style, 'bullet')
                indent = '  ' * (num_info['ilvl'] if num_info else 0)

                if not in_list or list_style != style:
                    if current_style != style:
                        lines.append(f'<!-- style: {style} -->')
                        current_style = style
                    in_list = True
                    list_style = style
                    list_counter = 0

                list_counter += 1
                if list_type == 'numbered':
                    lines.append(f'{indent}{list_counter}. {text}')
                else:
                    lines.append(f'{indent}- {text}')
                continue

            # End list if active
            if in_list:
                in_list = False
                list_style = None
                list_counter = 0
                lines.append('')

            # ── Style annotation ──
            if style != current_style:
                lines.append(f'<!-- style: {style} -->')
                current_style = style

            # ── Headings ──
            heading_level = HEADING_MAP.get(style)
            if heading_level:
                lines.append(f'{"#" * heading_level} {text}')
            else:
                lines.append(text)

        # ── Tables ──
        elif local == 'tbl':
            if in_list:
                in_list = False
                list_style = None
                list_counter = 0

            tbl_style = get_table_style(child)
            if tbl_style:
                styles_used.add(tbl_style)
                lines.append(f'<!-- style: {tbl_style} -->')

            table_md = table_to_markdown(child, rels, zf, media_dir)
            if table_md:
                lines.append(table_md)
            lines.append('')
            current_style = None

        # ── Structured Document Tags (content controls) ──
        elif local == 'sdt':
            # Process the content inside
            sdt_content = child.find(w('sdtContent'))
            if sdt_content is not None:
                inner_lines, inner_styles = process_body(
                    sdt_content, rels, zf, media_dir)
                lines.extend(inner_lines)
                styles_used.update(inner_styles)

        # ── sectPr and other elements: skip ──

    return lines, styles_used


# ──────────────────────────────────────────────────────────────────────
# YAML Front Matter
# ──────────────────────────────────────────────────────────────────────

def build_front_matter(meta, source_path, styles_used, template_name=None):
    """Serialize metadata as YAML front matter."""
    lines = ['---']

    def add(key, val):
        if isinstance(val, str) and any(c in val for c in ':"\'{}\n'):
            lines.append(f'{key}: "{val}"')
        else:
            lines.append(f'{key}: {val}')

    if meta.get('title'):
        add('title', meta['title'])
    if meta.get('author'):
        add('author', meta['author'])
    if meta.get('company'):
        add('company', meta['company'])
    if meta.get('created'):
        add('created', meta['created'])

    add('source', os.path.basename(source_path))
    add('extracted', datetime.now().isoformat(timespec='seconds'))

    if template_name:
        add('template', template_name)

    if styles_used:
        lines.append('styles_used:')
        for s in sorted(styles_used):
            lines.append(f'  - {s}')

    lines.append('---')
    return '\n'.join(lines)


# ──────────────────────────────────────────────────────────────────────
# Main Extraction
# ──────────────────────────────────────────────────────────────────────

def extract(docx_path, output_md=None, template_name=None):
    """Extract a .docx to markdown with YAML front matter and style tags."""
    if output_md is None:
        output_md = os.path.splitext(docx_path)[0] + '.md'

    media_dir = os.path.splitext(output_md)[0] + '_media'

    with zipfile.ZipFile(docx_path, 'r') as zf:
        meta = extract_metadata(zf)
        rels = parse_relationships(zf)

        doc = ET.fromstring(zf.read('word/document.xml'))
        body = doc.find(w('body'))

        if body is None:
            print("Error: No <w:body> found in document.xml", file=sys.stderr)
            sys.exit(1)

        md_lines, styles_used = process_body(body, rels, zf, media_dir)

        front_matter = build_front_matter(
            meta, docx_path, styles_used, template_name)

        # Assemble
        output = front_matter + '\n\n' + '\n'.join(md_lines) + '\n'
        # Collapse excessive blank lines
        output = re.sub(r'\n{4,}', '\n\n\n', output)

        with open(output_md, 'w', encoding='utf-8') as f:
            f.write(output)

    print(f"Extracted: {output_md}")
    print(f"  Source: {docx_path}")
    print(f"  Styles found: {', '.join(sorted(styles_used))}")
    if os.path.isdir(media_dir):
        n = len(os.listdir(media_dir))
        print(f"  Media files: {n} in {media_dir}/")

    return output_md


def main():
    if len(sys.argv) < 2:
        print("Usage: python docx_to_markdown.py input.docx [output.md]",
              file=sys.stderr)
        sys.exit(1)

    docx_path = sys.argv[1]
    output_md = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(docx_path):
        print(f"Error: File not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    extract(docx_path, output_md)


if __name__ == '__main__':
    main()
