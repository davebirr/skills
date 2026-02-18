#!/usr/bin/env python3
"""Reassemble a styled Markdown file into a branded .docx.

Reads a markdown file produced by docx_to_markdown.py (with YAML front matter
and <!-- style: X --> annotations) and creates a new .docx from the brand
template with all styles correctly applied.

Usage:
    python markdown_to_docx.py input.md output.docx [template_path]

Features:
  - Parses YAML front matter for metadata
  - Applies paragraph styles from <!-- style: X --> annotations
  - Maps markdown headings (#) to Heading1-5 styles
  - Maps markdown lists (-, 1.) to BodyCopyBulleted/BodyCopyNumbered
  - Preserves bold (**), italic (*), strikethrough (~~)
  - Embeds hyperlinks with proper relationship entries
  - Embeds images from media folder with relationship entries
  - Preserves template headers, footers, theme, and page layout
"""

import sys
import os
import re
import struct
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as _xml_escape

# ──────────────────────────────────────────────────────────────────────
# XML Namespaces
# ──────────────────────────────────────────────────────────────────────

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

REL_TYPE_HYPERLINK = ('http://schemas.openxmlformats.org/officeDocument/'
                      '2006/relationships/hyperlink')
REL_TYPE_IMAGE = ('http://schemas.openxmlformats.org/officeDocument/'
                  '2006/relationships/image')

# Numbering IDs from the MS brand template (Brand_template_20200929.dotx)
BULLET_NUM_ID = 39      # abstractNumId 28  — bullet
NUMBERED_NUM_ID = 38    # abstractNumId 32  — decimal

# Heading level → default style
HEADING_STYLES = {1: 'Heading1', 2: 'Heading2', 3: 'Heading3',
                  4: 'Heading4', 5: 'Heading5'}

DEFAULT_STYLE = 'BodyCopy'

# ──────────────────────────────────────────────────────────────────────
# XML Helpers
# ──────────────────────────────────────────────────────────────────────


def esc(text):
    """Escape text for XML content."""
    return _xml_escape(text or '', {'"': '&quot;', "'": '&apos;'})


def esc_attr(text):
    """Escape text for XML attribute values."""
    return _xml_escape(text or '', {'"': '&quot;'})


# ──────────────────────────────────────────────────────────────────────
# Markdown Parsing
# ──────────────────────────────────────────────────────────────────────


def parse_markdown(md_path):
    """Parse markdown into (front_matter_dict, list_of_blocks).

    Block types:
      paragraph  — {type, style, runs: [(text, fmt_dict), ...]}
      heading    — {type, style, level, runs}
      list       — {type, style, list_type: 'bullet'|'numbered',
                     items: [{indent, runs}, ...]}
      table      — {type, style, data: [[cell_text, ...], ...]}
      pagebreak  — {type: 'pagebreak'}
      image      — {type, style, alt, path}
    """
    with open(md_path, 'r', encoding='utf-8') as f:
        text = f.read()

    # ── Front matter ──
    front_matter = {}
    body = text
    fm_match = re.match(r'^---\s*\n(.*?)\n---\s*\n', text, re.DOTALL)
    if fm_match:
        front_matter = _parse_yaml(fm_match.group(1))
        body = text[fm_match.end():]

    # ── Body ──
    blocks = _parse_body(body)
    return front_matter, blocks


def _parse_yaml(yaml_text):
    """Minimal YAML parser for flat key-value pairs and simple lists."""
    result = {}
    current_key = None
    current_list = None

    for line in yaml_text.split('\n'):
        line = line.rstrip()
        if not line:
            continue
        if re.match(r'^\s+-\s+', line):
            if current_key is not None and current_list is not None:
                current_list.append(line.strip().lstrip('- '))
            continue
        m = re.match(r'^(\w[\w_]*)\s*:\s*(.*)', line)
        if m:
            if current_key and current_list is not None:
                result[current_key] = current_list
            current_key = m.group(1)
            val = m.group(2).strip()
            if val == '':
                current_list = []
            else:
                if (val.startswith('"') and val.endswith('"')) or \
                   (val.startswith("'") and val.endswith("'")):
                    val = val[1:-1]
                result[current_key] = val
                current_list = None

    if current_key and current_list is not None:
        result[current_key] = current_list
    return result


def _parse_body(body_text):
    """Parse markdown body into structured blocks."""
    lines = body_text.split('\n')
    blocks = []
    current_style = DEFAULT_STYLE
    i = 0

    while i < len(lines):
        line = lines[i]

        # ── Style annotation ──
        sm = re.match(r'<!--\s*style:\s*(\S+)\s*-->', line)
        if sm:
            current_style = sm.group(1)
            i += 1
            continue

        # ── Page break ──
        if re.match(r'<!--\s*pagebreak\s*-->', line):
            blocks.append({'type': 'pagebreak'})
            i += 1
            continue

        # ── Table / other HTML comments — skip ──
        if re.match(r'<!--.*-->', line):
            i += 1
            continue

        # ── Blank line ──
        if not line.strip():
            i += 1
            continue

        # ── Markdown table ──
        if line.strip().startswith('|'):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i])
                i += 1
            blocks.append({
                'type': 'table',
                'style': current_style,
                'data': _parse_md_table(table_lines),
            })
            continue

        # ── Heading ──
        hm = re.match(r'^(#{1,6})\s+(.*)', line)
        if hm:
            level = len(hm.group(1))
            text = hm.group(2)
            style = (current_style
                     if current_style not in (DEFAULT_STYLE, 'Normal')
                     else HEADING_STYLES.get(level, f'Heading{level}'))
            blocks.append({
                'type': 'heading',
                'style': style,
                'level': level,
                'runs': _parse_inline(text),
            })
            # Reset to default after a heading (next annotation applies)
            if style == current_style:
                current_style = DEFAULT_STYLE
            i += 1
            continue

        # ── Bulleted list ──
        if re.match(r'^(\s*)[-*]\s+', line):
            items = []
            while i < len(lines) and re.match(r'^(\s*)[-*]\s+', lines[i]):
                m = re.match(r'^(\s*)[-*]\s+(.*)', lines[i])
                items.append({
                    'indent': len(m.group(1)) // 2,
                    'runs': _parse_inline(m.group(2)),
                })
                i += 1
            style = (current_style
                     if current_style in ('BodyCopyBulleted', 'ListBullet')
                     else 'BodyCopyBulleted')
            blocks.append({
                'type': 'list', 'list_type': 'bullet',
                'style': style, 'items': items,
            })
            current_style = DEFAULT_STYLE
            continue

        # ── Numbered list ──
        if re.match(r'^(\s*)\d+[.)]\s+', line):
            items = []
            while i < len(lines) and re.match(r'^(\s*)\d+[.)]\s+', lines[i]):
                m = re.match(r'^(\s*)\d+[.)]\s+(.*)', lines[i])
                items.append({
                    'indent': len(m.group(1)) // 2,
                    'runs': _parse_inline(m.group(2)),
                })
                i += 1
            style = (current_style
                     if current_style in ('BodyCopyNumbered', 'ListNumber')
                     else 'BodyCopyNumbered')
            blocks.append({
                'type': 'list', 'list_type': 'numbered',
                'style': style, 'items': items,
            })
            current_style = DEFAULT_STYLE
            continue

        # ── Image ──
        img_m = re.match(r'!\[([^\]]*)\]\(([^)]+)\)', line)
        if img_m:
            blocks.append({
                'type': 'image',
                'style': current_style,
                'alt': img_m.group(1),
                'path': img_m.group(2),
            })
            i += 1
            continue

        # ── Regular paragraph ──
        blocks.append({
            'type': 'paragraph',
            'style': current_style,
            'runs': _parse_inline(line),
        })
        i += 1

    return blocks


def _parse_inline(text):
    """Parse markdown inline formatting into [(text, fmt_dict), ...].

    Handles: ***bold+italic***, **bold**, *italic*, ~~strike~~, [text](url)
    """
    runs = []
    pattern = re.compile(
        r'(\*\*\*(.+?)\*\*\*)'
        r'|(\*\*(.+?)\*\*)'
        r'|(\*(.+?)\*)'
        r'|(~~(.+?)~~)'
        r'|(\[([^\]]+)\]\(([^)]+)\))'
    )
    pos = 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            runs.append((text[pos:m.start()], {}))
        if m.group(2):
            runs.append((m.group(2), {'bold': True, 'italic': True}))
        elif m.group(4):
            runs.append((m.group(4), {'bold': True}))
        elif m.group(6):
            runs.append((m.group(6), {'italic': True}))
        elif m.group(8):
            runs.append((m.group(8), {'strike': True}))
        elif m.group(10):
            runs.append((m.group(10), {'link': m.group(11)}))
        pos = m.end()
    if pos < len(text):
        runs.append((text[pos:], {}))
    if not runs and text:
        runs.append((text, {}))
    return runs


def _parse_md_table(table_lines):
    """Parse markdown table lines into list of rows (list of cell strings)."""
    rows = []
    for line in table_lines:
        line = line.strip()
        if line.startswith('|'):
            line = line[1:]
        if line.endswith('|'):
            line = line[:-1]
        if re.match(r'^[\s|:\-]+$', line):
            continue
        rows.append([c.strip() for c in line.split('|')])
    return rows


# ──────────────────────────────────────────────────────────────────────
# XML Generation (string-based for reliable namespace handling)
# ──────────────────────────────────────────────────────────────────────


def _run_xml(text, fmt, hyperlink_rels=None, md_dir=None):
    """Build XML string for a single run.

    If fmt contains 'link', returns a w:hyperlink element instead.
    hyperlink_rels is a dict we mutate: {url: rId, ...}
    """
    if fmt.get('link'):
        url = fmt['link']
        if hyperlink_rels is not None and url not in hyperlink_rels:
            rid = f'rId{200 + len(hyperlink_rels)}'
            hyperlink_rels[url] = rid
        rid = hyperlink_rels.get(url, 'rId999') if hyperlink_rels else 'rId999'

        rpr = '<w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>'
        return (f'<w:hyperlink r:id="{rid}">'
                f'<w:r>{rpr}'
                f'<w:t xml:space="preserve">{esc(text)}</w:t>'
                f'</w:r></w:hyperlink>')

    # Normal run
    rpr_parts = []
    if fmt.get('bold'):
        rpr_parts.append('<w:b/>')
    if fmt.get('italic'):
        rpr_parts.append('<w:i/>')
    if fmt.get('strike'):
        rpr_parts.append('<w:strike/>')
    rpr = f'<w:rPr>{"".join(rpr_parts)}</w:rPr>' if rpr_parts else ''

    return (f'<w:r>{rpr}'
            f'<w:t xml:space="preserve">{esc(text)}</w:t>'
            f'</w:r>')


def _paragraph_xml(style, runs, hyperlink_rels=None, md_dir=None):
    """Build XML string for a styled paragraph."""
    ppr = f'<w:pPr><w:pStyle w:val="{esc_attr(style)}"/></w:pPr>'
    runs_xml = ''.join(_run_xml(t, f, hyperlink_rels, md_dir)
                       for t, f in runs)
    return f'<w:p>{ppr}{runs_xml}</w:p>'


def _list_paragraph_xml(style, runs, num_id, ilvl,
                        hyperlink_rels=None, md_dir=None):
    """Build XML for a list item paragraph with numbering reference."""
    ppr = (f'<w:pPr>'
           f'<w:pStyle w:val="{esc_attr(style)}"/>'
           f'<w:numPr>'
           f'<w:ilvl w:val="{ilvl}"/>'
           f'<w:numId w:val="{num_id}"/>'
           f'</w:numPr>'
           f'</w:pPr>')
    runs_xml = ''.join(_run_xml(t, f, hyperlink_rels, md_dir)
                       for t, f in runs)
    return f'<w:p>{ppr}{runs_xml}</w:p>'


def _page_break_xml():
    """Build XML for a page-break paragraph."""
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def _table_xml(data, style):
    """Build XML for a table."""
    tbl_style = (f'<w:tblStyle w:val="{esc_attr(style)}"/>'
                 if style else '')
    rows_xml = []
    for row in data:
        cells_xml = []
        for cell_text in row:
            t_xml = (f'<w:r><w:t xml:space="preserve">{esc(cell_text)}</w:t>'
                     f'</w:r>' if cell_text else '')
            cells_xml.append(f'<w:tc><w:p>{t_xml}</w:p></w:tc>')
        rows_xml.append(f'<w:tr>{"".join(cells_xml)}</w:tr>')
    return (f'<w:tbl><w:tblPr>{tbl_style}'
            f'<w:tblW w:w="0" w:type="auto"/>'
            f'</w:tblPr>{"".join(rows_xml)}</w:tbl>')


def _get_image_dimensions(data):
    """Read width, height from PNG or JPEG binary data (no PIL needed)."""
    # PNG
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        w, h = struct.unpack('>II', data[16:24])
        return w, h
    # JPEG
    if data[:2] == b'\xff\xd8':
        i = 2
        while i < len(data) - 9:
            if data[i] != 0xFF:
                break
            marker = data[i + 1]
            if marker in (0xC0, 0xC2):
                h, w = struct.unpack('>HH', data[i + 5:i + 9])
                return w, h
            if marker == 0xD9:
                break
            length = struct.unpack('>H', data[i + 2:i + 4])[0]
            i += 2 + length
    return None, None


def _image_paragraph_xml(r_id, alt, cx, cy, img_id):
    """Build XML for an inline image paragraph.

    cx, cy are in EMUs (English Metric Units, 914400 EMU = 1 inch).
    """
    return (
        f'<w:p><w:r><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0"'
        f' xmlns:wp="{WP}">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:docPr id="{img_id}" name="Picture {img_id}"'
        f' descr="{esc_attr(alt)}"/>'
        f'<a:graphic xmlns:a="{A}">'
        f'<a:graphicData uri="{PIC}">'
        f'<pic:pic xmlns:pic="{PIC}">'
        f'<pic:nvPicPr>'
        f'<pic:cNvPr id="{img_id}" name="Picture {img_id}"/>'
        f'<pic:cNvPicPr/>'
        f'</pic:nvPicPr>'
        f'<pic:blipFill>'
        f'<a:blip r:embed="{r_id}"/>'
        f'<a:stretch><a:fillRect/></a:stretch>'
        f'</pic:blipFill>'
        f'<pic:spPr>'
        f'<a:xfrm>'
        f'<a:off x="0" y="0"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/>'
        f'</a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</pic:spPr>'
        f'</pic:pic>'
        f'</a:graphicData>'
        f'</a:graphic>'
        f'</wp:inline>'
        f'</w:drawing></w:r></w:p>'
    )


# ──────────────────────────────────────────────────────────────────────
# Document Assembly
# ──────────────────────────────────────────────────────────────────────


def _extract_sect_pr(template_doc_xml):
    """Extract the last <w:sectPr ...>...</w:sectPr> from template XML."""
    # Use regex to get the full sectPr element with all namespace attrs
    m = re.search(r'(<w:sectPr[^>]*>.*?</w:sectPr>)',
                  template_doc_xml, re.DOTALL)
    if m:
        return m.group(1)
    # Try self-closing
    m = re.search(r'(<w:sectPr[^/]*/\s*>)', template_doc_xml)
    return m.group(1) if m else '<w:sectPr/>'


def _extract_doc_open_tag(template_doc_xml):
    """Extract the opening <w:document ...> tag with all namespace decls."""
    m = re.search(r'(<w:document[^>]*>)', template_doc_xml, re.DOTALL)
    if not m:
        raise ValueError("No <w:document> element found in template")
    return m.group(1)


def _extract_xml_decl(template_doc_xml):
    """Extract the XML declaration if present."""
    if template_doc_xml.startswith('<?xml'):
        end = template_doc_xml.index('?>') + 2
        return template_doc_xml[:end]
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'


def build_document_xml(blocks, template_doc_bytes, md_dir=None,
                       image_rels=None, extra_media=None):
    """Build complete document.xml from blocks, using template for structure.

    Args:
        blocks: parsed markdown blocks
        template_doc_bytes: raw bytes of template's document.xml
        md_dir: directory of the source markdown (for finding images)
        image_rels: dict we mutate — {media_path_in_zip: rId}
        extra_media: list we mutate — [(zip_path, file_bytes), ...]

    Returns:
        document.xml content as a string
    """
    template_xml = template_doc_bytes.decode('utf-8')

    xml_decl = _extract_xml_decl(template_xml)
    doc_open = _extract_doc_open_tag(template_xml)
    sect_pr = _extract_sect_pr(template_xml)

    hyperlink_rels = {}  # url → rId
    if image_rels is None:
        image_rels = {}
    if extra_media is None:
        extra_media = []
    img_counter = 0

    body_parts = []

    for block in blocks:
        btype = block['type']

        if btype == 'pagebreak':
            body_parts.append(_page_break_xml())

        elif btype in ('paragraph', 'heading'):
            body_parts.append(
                _paragraph_xml(block['style'], block.get('runs', []),
                               hyperlink_rels, md_dir))

        elif btype == 'list':
            lt = block.get('list_type', 'bullet')
            num_id = NUMBERED_NUM_ID if lt == 'numbered' else BULLET_NUM_ID
            for item in block['items']:
                body_parts.append(
                    _list_paragraph_xml(
                        block['style'], item['runs'],
                        num_id, item.get('indent', 0),
                        hyperlink_rels, md_dir))

        elif btype == 'table':
            style = block.get('style', '')
            if style in (DEFAULT_STYLE, 'Normal', ''):
                style = 'Placeholder'
            body_parts.append(_table_xml(block['data'], style))

        elif btype == 'image':
            img_path = block.get('path', '')
            alt = block.get('alt', '')

            # Resolve image path relative to markdown file
            if md_dir:
                abs_path = os.path.join(md_dir, img_path)
            else:
                abs_path = img_path

            if os.path.isfile(abs_path):
                img_counter += 1
                ext = os.path.splitext(img_path)[1].lower() or '.png'
                media_name = f'image_inserted_{img_counter}{ext}'
                zip_path = f'media/{media_name}'
                rid = f'rId{300 + img_counter}'

                with open(abs_path, 'rb') as f:
                    img_data = f.read()

                extra_media.append((f'word/{zip_path}', img_data))
                image_rels[zip_path] = rid

                # Compute dimensions (default to 6 inches wide)
                px_w, px_h = _get_image_dimensions(img_data)
                if px_w and px_h:
                    max_cx = 5486400  # ~6 inches
                    cx = min(px_w * 9525, max_cx)  # px → EMU (approx)
                    cy = int(cx * px_h / px_w)
                else:
                    cx, cy = 5486400, 3657600  # 6×4 inches

                body_parts.append(
                    _image_paragraph_xml(rid, alt, cx, cy, img_counter))
            else:
                # Image file not found — insert placeholder text
                body_parts.append(
                    _paragraph_xml(block.get('style', DEFAULT_STYLE),
                                   [(f'[Image: {alt or img_path}]', {})],
                                   hyperlink_rels, md_dir))

    # Assemble document
    body_xml = ''.join(body_parts)
    doc_xml = f'{xml_decl}\n{doc_open}<w:body>{body_xml}{sect_pr}</w:body></w:document>'

    return doc_xml, hyperlink_rels, image_rels


# ──────────────────────────────────────────────────────────────────────
# Relationship Management
# ──────────────────────────────────────────────────────────────────────


def _update_relationships(original_rels_bytes, hyperlink_rels, image_rels):
    """Add new hyperlink and image relationships to the rels XML.

    Returns updated rels XML as bytes.
    """
    rels_str = original_rels_bytes.decode('utf-8')

    new_rels = []
    for url, rid in hyperlink_rels.items():
        new_rels.append(
            f'<Relationship Id="{rid}" '
            f'Type="{REL_TYPE_HYPERLINK}" '
            f'Target="{esc_attr(url)}" TargetMode="External"/>')

    for media_path, rid in image_rels.items():
        new_rels.append(
            f'<Relationship Id="{rid}" '
            f'Type="{REL_TYPE_IMAGE}" '
            f'Target="{media_path}"/>')

    if new_rels:
        insert = '\n'.join(new_rels)
        rels_str = rels_str.replace('</Relationships>',
                                    f'{insert}\n</Relationships>')

    return rels_str.encode('utf-8')


def _update_content_types(ct_bytes, extra_media):
    """Ensure [Content_Types].xml has entries for any new image extensions."""
    ct_str = ct_bytes.decode('utf-8')

    extensions_needed = set()
    for zip_path, _ in extra_media:
        ext = os.path.splitext(zip_path)[1].lower().lstrip('.')
        if ext:
            extensions_needed.add(ext)

    ext_to_mime = {
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'gif': 'image/gif',
        'bmp': 'image/bmp',
        'tiff': 'image/tiff',
        'tif': 'image/tiff',
        'svg': 'image/svg+xml',
        'emf': 'image/x-emf',
        'wmf': 'image/x-wmf',
    }

    for ext in extensions_needed:
        if f'Extension="{ext}"' not in ct_str:
            mime = ext_to_mime.get(ext, 'application/octet-stream')
            entry = f'<Default Extension="{ext}" ContentType="{mime}"/>'
            ct_str = ct_str.replace('<Types',
                                    f'<Types>{entry}', 1) if '<Types>' not in ct_str else \
                     ct_str.replace('<Types>', f'<Types>{entry}', 1)

    return ct_str.encode('utf-8')


# ──────────────────────────────────────────────────────────────────────
# Main Assembly
# ──────────────────────────────────────────────────────────────────────


def assemble(md_path, output_path, template_path=None):
    """Assemble a styled markdown file into a branded .docx."""
    if template_path is None:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(
            script_dir, '..', 'Brand_template_20200929.dotx')

    if not os.path.exists(template_path):
        print(f"Error: Template not found: {template_path}", file=sys.stderr)
        sys.exit(1)

    # Parse markdown
    front_matter, blocks = parse_markdown(md_path)
    md_dir = os.path.dirname(os.path.abspath(md_path))

    # Content type replacements
    tpl_ct = ('application/vnd.openxmlformats-officedocument.'
              'wordprocessingml.template.main+xml')
    doc_ct = ('application/vnd.openxmlformats-officedocument.'
              'wordprocessingml.document.main+xml')

    image_rels = {}
    extra_media = []

    with zipfile.ZipFile(template_path, 'r') as zin:
        template_doc_xml = zin.read('word/document.xml')

        doc_xml, hyperlink_rels, image_rels = build_document_xml(
            blocks, template_doc_xml, md_dir, image_rels, extra_media)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename == '[Content_Types].xml':
                    data = data.replace(
                        tpl_ct.encode(), doc_ct.encode())
                    if extra_media:
                        data = _update_content_types(data, extra_media)

                elif item.filename == 'word/document.xml':
                    data = doc_xml.encode('utf-8')

                elif item.filename == 'word/_rels/document.xml.rels':
                    if hyperlink_rels or image_rels:
                        data = _update_relationships(
                            data, hyperlink_rels, image_rels)

                zout.writestr(item, data)

            # Add extra media files (inserted images)
            for zip_path, file_data in extra_media:
                zout.writestr(zip_path, file_data)

    print(f"Created: {output_path}")
    print(f"  From markdown: {md_path}")
    print(f"  Template: {os.path.basename(template_path)}")
    print(f"  Content blocks: {len(blocks)}")
    if hyperlink_rels:
        print(f"  Hyperlinks: {len(hyperlink_rels)}")
    if extra_media:
        print(f"  Embedded images: {len(extra_media)}")


def main():
    if len(sys.argv) < 3:
        print("Usage: python markdown_to_docx.py input.md output.docx "
              "[template_path]", file=sys.stderr)
        sys.exit(1)

    md_path = sys.argv[1]
    output_path = sys.argv[2]
    template_path = sys.argv[3] if len(sys.argv) > 3 else None

    if not os.path.exists(md_path):
        print(f"Error: File not found: {md_path}", file=sys.stderr)
        sys.exit(1)

    if os.path.exists(output_path):
        print(f"Error: Output exists: {output_path}", file=sys.stderr)
        sys.exit(1)

    assemble(md_path, output_path, template_path)


if __name__ == '__main__':
    main()
