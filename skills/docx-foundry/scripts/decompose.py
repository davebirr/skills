#!/usr/bin/env python3
"""Decompose a .docx into modular, AI-editable components.

Splits a Word document into:
  - manifest.yaml: document metadata, section index, stats
  - styles.yaml: style definitions, theme colors/fonts, numbering, page layout
  - sections/*.md: one markdown file per section with YAML front matter
  - media/*: extracted images (png, jpg, etc.)

Usage:
    python decompose.py input.docx [output_dir] [--split-on LEVEL]

Options:
    output_dir       Output directory (default: input filename without extension)
    --split-on N     Heading level to split on (default: auto-detect highest used)
"""

import argparse
import os
import re
import struct
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime

# ═══════════════════════════════════════════════════════════════════════
# XML Namespaces
# ═══════════════════════════════════════════════════════════════════════

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
DC = 'http://purl.org/dc/elements/1.1/'
DCTERMS = 'http://purl.org/dc/terms/'

def w(tag): return f'{{{W}}}{tag}'
def a(tag): return f'{{{A}}}{tag}'

HEADING_MAP = {
    'Heading1': 1, 'Heading2': 2, 'Heading3': 3,
    'Heading4': 4, 'Heading5': 5, 'Heading6': 6,
    'heading1': 1, 'heading2': 2, 'heading3': 3,
}

LIST_STYLES = {
    'ListBullet': 'bullet', 'ListNumber': 'numbered',
    'ListParagraph': 'bullet',
    'BodyCopyBulleted': 'bullet', 'BodyCopyNumbered': 'numbered',
}

# ═══════════════════════════════════════════════════════════════════════
# YAML Utilities (no pyyaml dependency)
# ═══════════════════════════════════════════════════════════════════════

def _needs_quote(s):
    if not isinstance(s, str):
        return False
    if not s:
        return True
    if s.lower() in ('true', 'false', 'null', 'yes', 'no', 'on', 'off'):
        return True
    if s[0] in ' \t-?:,[]{}#&*!|>\'"@`' or s[-1] in ' \t':
        return True
    if ':' in s or '#' in s or '\n' in s or '\t' in s:
        return True
    try:
        float(s)
        return True
    except ValueError:
        pass
    return False


def _quote(s):
    return '"' + s.replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\n') + '"'


def yaml_val(v):
    if v is None:
        return 'null'
    if isinstance(v, bool):
        return 'true' if v else 'false'
    if isinstance(v, (int, float)):
        return str(v)
    if isinstance(v, str):
        return _quote(v) if _needs_quote(v) else v
    return str(v)


def yaml_dump(data, indent=0):
    """Serialize Python dict/list to YAML string."""
    pad = '  ' * indent
    if data is None:
        return 'null'
    if isinstance(data, bool):
        return 'true' if data else 'false'
    if isinstance(data, (int, float, str)):
        return yaml_val(data)

    if isinstance(data, dict):
        if not data:
            return '{}'
        lines = []
        for k, v in data.items():
            if isinstance(v, (dict, list)) and v:
                lines.append(f'{pad}{k}:')
                lines.append(yaml_dump(v, indent + 1))
            else:
                val = '{}' if isinstance(v, dict) else ('[]' if isinstance(v, list) else yaml_val(v))
                lines.append(f'{pad}{k}: {val}')
        return '\n'.join(lines)

    if isinstance(data, (list, tuple)):
        if not data:
            return '[]'
        lines = []
        for item in data:
            if isinstance(item, dict) and item:
                first = True
                for k, v in item.items():
                    pfx = f'{pad}- ' if first else f'{pad}  '
                    first = False
                    if isinstance(v, (dict, list)) and v:
                        lines.append(f'{pfx}{k}:')
                        lines.append(yaml_dump(v, indent + 2))
                    else:
                        val = '{}' if isinstance(v, dict) else ('[]' if isinstance(v, list) else yaml_val(v))
                        lines.append(f'{pfx}{k}: {val}')
            else:
                lines.append(f'{pad}- {yaml_val(item)}')
        return '\n'.join(lines)

    return yaml_val(data)


def yaml_front_matter(data):
    """Generate YAML front matter block for markdown files."""
    return '---\n' + yaml_dump(data) + '\n---'

# ═══════════════════════════════════════════════════════════════════════
# Metadata Extraction (docProps/core.xml, docProps/app.xml)
# ═══════════════════════════════════════════════════════════════════════

def extract_metadata(zf):
    meta = {}
    try:
        core = ET.fromstring(zf.read('docProps/core.xml'))
        for child in core:
            local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            ns = child.tag.split('}')[0][1:] if '}' in child.tag else ''
            if ns == DC and local in ('title', 'creator', 'subject', 'description'):
                key = 'author' if local == 'creator' else local
                if child.text and child.text.strip():
                    meta[key] = child.text.strip()
            elif ns == DCTERMS and local in ('created', 'modified') and child.text:
                meta[local] = child.text.strip()
    except (KeyError, ET.ParseError):
        pass
    try:
        app = ET.fromstring(zf.read('docProps/app.xml'))
        for child in app:
            local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if local == 'Company' and child.text:
                meta['company'] = child.text.strip()
            elif local == 'Template' and child.text:
                meta['template'] = child.text.strip()
            elif local == 'Pages' and child.text:
                meta['pages'] = int(child.text.strip())
    except (KeyError, ET.ParseError, ValueError):
        pass
    return meta

# ═══════════════════════════════════════════════════════════════════════
# Relationship Parsing
# ═══════════════════════════════════════════════════════════════════════

def parse_relationships(zf):
    rels = {}
    try:
        root = ET.fromstring(zf.read('word/_rels/document.xml.rels'))
        for el in root:
            rid = el.get('Id', '')
            rels[rid] = {
                'target': el.get('Target', ''),
                'type': el.get('Type', '').split('/')[-1],
                'external': el.get('TargetMode', '') == 'External',
            }
    except (KeyError, ET.ParseError):
        pass
    return rels

# ═══════════════════════════════════════════════════════════════════════
# Style Extraction (word/styles.xml)
# ═══════════════════════════════════════════════════════════════════════

def _parse_run_props(rpr):
    """Extract font properties from a w:rPr element."""
    if rpr is None:
        return {}
    props = {}
    fonts = rpr.find(w('rFonts'))
    if fonts is not None:
        for attr in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
            val = fonts.get(w(attr))
            if val:
                props.setdefault('font', {})['name'] = val
                break
        for attr in ('asciiTheme', 'hAnsiTheme'):
            val = fonts.get(w(attr))
            if val:
                props.setdefault('font', {})['theme'] = val
                break
    sz = rpr.find(w('sz'))
    if sz is not None:
        half_pt = int(sz.get(w('val'), '0'))
        if half_pt:
            props.setdefault('font', {})['size_pt'] = half_pt / 2
    color = rpr.find(w('color'))
    if color is not None:
        val = color.get(w('val'))
        if val and val != 'auto':
            props.setdefault('font', {})['color'] = val
        tc = color.get(w('themeColor'))
        if tc:
            props.setdefault('font', {})['theme_color'] = tc
    for tag, key in [('b', 'bold'), ('i', 'italic'), ('u', 'underline'), ('strike', 'strikethrough')]:
        el = rpr.find(w(tag))
        if el is not None:
            val = el.get(w('val'), 'true')
            if val not in ('false', '0'):
                props.setdefault('font', {})[key] = True
    return props


def _parse_para_props(ppr):
    """Extract paragraph properties from a w:pPr element."""
    if ppr is None:
        return {}
    props = {}
    spacing = ppr.find(w('spacing'))
    if spacing is not None:
        sp = {}
        for attr in ('before', 'after', 'line'):
            val = spacing.get(w(attr))
            if val:
                sp[attr] = int(val)
        lr = spacing.get(w('lineRule'))
        if lr:
            sp['line_rule'] = lr
        if sp:
            props['spacing'] = sp
    ind = ppr.find(w('ind'))
    if ind is not None:
        indent = {}
        for attr in ('left', 'right', 'hanging', 'firstLine'):
            val = ind.get(w(attr))
            if val:
                indent[attr] = int(val)
        if indent:
            props['indent'] = indent
    jc = ppr.find(w('jc'))
    if jc is not None:
        props['alignment'] = jc.get(w('val'))
    outline = ppr.find(w('outlineLvl'))
    if outline is not None:
        props['outline_level'] = int(outline.get(w('val'), '0'))
    for tag, key in [('keepNext', 'keep_next'), ('keepLines', 'keep_lines'),
                     ('pageBreakBefore', 'page_break_before')]:
        if ppr.find(w(tag)) is not None:
            props[key] = True
    return props


def extract_styles(zf):
    """Extract style definitions from word/styles.xml."""
    result = {'defaults': {}, 'paragraph': {}, 'character': {}, 'table': {}}
    try:
        root = ET.fromstring(zf.read('word/styles.xml'))
    except (KeyError, ET.ParseError):
        return result

    # Document defaults
    doc_defaults = root.find(w('docDefaults'))
    if doc_defaults is not None:
        rpr_default = doc_defaults.find(f'{w("rPrDefault")}/{w("rPr")}')
        if rpr_default is not None:
            rp = _parse_run_props(rpr_default)
            if 'font' in rp:
                result['defaults'] = rp['font']

    # Styles
    for style_el in root.findall(w('style')):
        stype = style_el.get(w('type'), 'paragraph')
        sid = style_el.get(w('styleId'), '')
        if not sid:
            continue

        sdef = {}
        name_el = style_el.find(w('name'))
        if name_el is not None:
            sdef['name'] = name_el.get(w('val'), sid)
        based = style_el.find(w('basedOn'))
        if based is not None:
            sdef['based_on'] = based.get(w('val'))
        nxt = style_el.find(w('next'))
        if nxt is not None:
            sdef['next'] = nxt.get(w('val'))
        if style_el.find(w('qFormat')) is not None:
            sdef['quick_format'] = True

        rpr = style_el.find(w('rPr'))
        rp = _parse_run_props(rpr)
        if 'font' in rp:
            sdef['font'] = rp['font']

        ppr = style_el.find(w('pPr'))
        pp = _parse_para_props(ppr)
        if pp:
            sdef['paragraph'] = pp

        # Categorize
        bucket = 'paragraph' if stype == 'paragraph' else ('character' if stype == 'character' else 'table')
        if bucket in result:
            result[bucket][sid] = sdef

    return result

# ═══════════════════════════════════════════════════════════════════════
# Theme Extraction (word/theme/theme1.xml)
# ═══════════════════════════════════════════════════════════════════════

def extract_theme(zf):
    """Extract color scheme and font scheme from the theme."""
    theme = {'colors': {}, 'fonts': {}}
    try:
        root = ET.fromstring(zf.read('word/theme/theme1.xml'))
    except (KeyError, ET.ParseError):
        return theme

    # Color scheme
    clr_scheme = root.find(f'.//{a("clrScheme")}')
    if clr_scheme is not None:
        theme['name'] = clr_scheme.get('name', '')
        color_names = ['dk1', 'dk2', 'lt1', 'lt2',
                       'accent1', 'accent2', 'accent3', 'accent4',
                       'accent5', 'accent6', 'hlink', 'folHlink']
        for cname in color_names:
            el = clr_scheme.find(a(cname))
            if el is not None:
                srgb = el.find(a('srgbClr'))
                sys_clr = el.find(a('sysClr'))
                if srgb is not None:
                    theme['colors'][cname] = srgb.get('val', '')
                elif sys_clr is not None:
                    theme['colors'][cname] = sys_clr.get('lastClr', sys_clr.get('val', ''))

    # Font scheme
    font_scheme = root.find(f'.//{a("fontScheme")}')
    if font_scheme is not None:
        major = font_scheme.find(f'{a("majorFont")}/{a("latin")}')
        minor = font_scheme.find(f'{a("minorFont")}/{a("latin")}')
        if major is not None:
            theme['fonts']['major'] = major.get('typeface', '')
        if minor is not None:
            theme['fonts']['minor'] = minor.get('typeface', '')

    return theme

# ═══════════════════════════════════════════════════════════════════════
# Numbering Extraction (word/numbering.xml)
# ═══════════════════════════════════════════════════════════════════════

def extract_numbering(zf):
    """Extract numbering/list definitions."""
    numbering = {'abstract': {}, 'instances': {}}
    try:
        root = ET.fromstring(zf.read('word/numbering.xml'))
    except (KeyError, ET.ParseError):
        return numbering

    for abs_num in root.findall(w('abstractNum')):
        abs_id = abs_num.get(w('abstractNumId'), '')
        levels = []
        for lvl in abs_num.findall(w('lvl')):
            ilvl = lvl.get(w('ilvl'), '0')
            fmt_el = lvl.find(w('numFmt'))
            txt_el = lvl.find(w('lvlText'))
            level_def = {
                'level': int(ilvl),
                'format': fmt_el.get(w('val'), '') if fmt_el is not None else '',
                'text': txt_el.get(w('val'), '') if txt_el is not None else '',
            }
            ppr = lvl.find(w('pPr'))
            if ppr is not None:
                ind = ppr.find(w('ind'))
                if ind is not None:
                    left = ind.get(w('left'))
                    hanging = ind.get(w('hanging'))
                    if left:
                        level_def['indent_left'] = int(left)
                    if hanging:
                        level_def['indent_hanging'] = int(hanging)
            levels.append(level_def)
        numbering['abstract'][abs_id] = levels

    for num in root.findall(w('num')):
        num_id = num.get(w('numId'), '')
        abs_ref = num.find(w('abstractNumId'))
        if abs_ref is not None:
            numbering['instances'][num_id] = abs_ref.get(w('val'), '')

    return numbering

# ═══════════════════════════════════════════════════════════════════════
# Page Layout Extraction
# ═══════════════════════════════════════════════════════════════════════

def extract_page_layout(body):
    """Extract page layout from the last sectPr in body."""
    layout = {}
    sect_pr = body.find(w('sectPr'))
    if sect_pr is None:
        # Try last child
        for child in reversed(list(body)):
            if child.tag == w('sectPr'):
                sect_pr = child
                break
    if sect_pr is None:
        return layout

    pg_sz = sect_pr.find(w('pgSz'))
    if pg_sz is not None:
        layout['width'] = int(pg_sz.get(w('w'), '12240'))
        layout['height'] = int(pg_sz.get(w('h'), '15840'))
        orient = pg_sz.get(w('orient'))
        if orient:
            layout['orientation'] = orient

    pg_mar = sect_pr.find(w('pgMar'))
    if pg_mar is not None:
        layout['margins'] = {}
        for attr in ('top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'):
            val = pg_mar.get(w(attr))
            if val:
                layout['margins'][attr] = int(val)

    cols = sect_pr.find(w('cols'))
    if cols is not None:
        num = cols.get(w('num'))
        if num and int(num) > 1:
            layout['columns'] = int(num)
            space = cols.get(w('space'))
            if space:
                layout['column_space'] = int(space)

    return layout

# ═══════════════════════════════════════════════════════════════════════
# Header/Footer Extraction
# ═══════════════════════════════════════════════════════════════════════

def extract_headers_footers(zf, rels):
    """Extract header/footer text content."""
    hf = {'headers': {}, 'footers': {}}
    for rid, info in rels.items():
        rtype = info['type']
        target = info['target']
        if rtype not in ('header', 'footer'):
            continue
        try:
            root = ET.fromstring(zf.read(f'word/{target}'))
            texts = []
            for p in root.iter(w('t')):
                if p.text:
                    texts.append(p.text)
            content = ' '.join(texts).strip()
            if content:
                bucket = 'headers' if rtype == 'header' else 'footers'
                hf[bucket][target] = content
        except (KeyError, ET.ParseError):
            pass
    return hf

# ═══════════════════════════════════════════════════════════════════════
# Run & Paragraph → Markdown Conversion
# ═══════════════════════════════════════════════════════════════════════

def _get_style(p_el):
    pPr = p_el.find(w('pPr'))
    if pPr is not None:
        ps = pPr.find(w('pStyle'))
        if ps is not None:
            return ps.get(w('val'), '')
    return ''


def _get_num_info(p_el):
    pPr = p_el.find(w('pPr'))
    if pPr is not None:
        numPr = pPr.find(w('numPr'))
        if numPr is not None:
            ilvl_el = numPr.find(w('ilvl'))
            numId_el = numPr.find(w('numId'))
            return {
                'ilvl': int(ilvl_el.get(w('val'), '0')) if ilvl_el is not None else 0,
                'numId': numId_el.get(w('val'), '') if numId_el is not None else '',
            }
    return None


def _has_page_break(p_el):
    for run in p_el.findall(w('r')):
        br = run.find(w('br'))
        if br is not None and br.get(w('type')) == 'page':
            return True
    pPr = p_el.find(w('pPr'))
    if pPr is not None and pPr.find(w('pageBreakBefore')) is not None:
        return True
    return False


def _extract_run_format(run_el):
    fmt = {}
    rPr = run_el.find(w('rPr'))
    if rPr is not None:
        for tag, key in [('b', 'bold'), ('i', 'italic'), ('strike', 'strike')]:
            el = rPr.find(w(tag))
            if el is not None and el.get(w('val'), 'true') not in ('false', '0'):
                fmt[key] = True
    return fmt


def _extract_run_text(run_el):
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


def _apply_md_format(text, fmt):
    if not text or not text.strip():
        return text
    stripped = text.strip()
    lead = text[:len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()):]
    if fmt.get('bold') and fmt.get('italic'):
        stripped = f'***{stripped}***'
    elif fmt.get('bold'):
        stripped = f'**{stripped}**'
    elif fmt.get('italic'):
        stripped = f'*{stripped}*'
    if fmt.get('strike'):
        stripped = f'~~{stripped}~~'
    return f'{lead}{stripped}{trail}'


def _find_images_in_element(el, rels):
    """Find all image references in an element tree."""
    images = []
    for blip in el.iter(f'{{{A}}}blip'):
        embed_id = blip.get(f'{{{R}}}embed', '')
        if embed_id and embed_id in rels:
            target = rels[embed_id]['target']
            if target and rels[embed_id]['type'] == 'image':
                images.append(target)
    return images


def _paragraph_to_text(p_el, rels):
    """Convert a w:p element to markdown text."""
    parts = []
    for child in p_el:
        local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if local == 'r':
            # Check for image
            imgs = _find_images_in_element(child, rels)
            if imgs:
                for img in imgs:
                    fname = os.path.basename(img)
                    parts.append(f'![{fname}](media/{fname})')
                continue
            text = _extract_run_text(child)
            fmt = _extract_run_format(child)
            parts.append(_apply_md_format(text, fmt))
        elif local == 'hyperlink':
            r_id = child.get(f'{{{R}}}id', '')
            link_parts = []
            for run in child.findall(w('r')):
                link_parts.append(_extract_run_text(run))
            link_text = ''.join(link_parts).strip()
            if r_id and r_id in rels and rels[r_id].get('external'):
                url = rels[r_id]['target']
                parts.append(f'[{link_text}]({url})')
            else:
                parts.append(link_text)
    return ''.join(parts)


def _table_to_md(tbl_el, rels):
    """Convert a w:tbl to markdown table."""
    rows_data = []
    max_cols = 0
    # Capture table style
    tbl_style = ''
    tblPr = tbl_el.find(w('tblPr'))
    if tblPr is not None:
        ts = tblPr.find(w('tblStyle'))
        if ts is not None:
            tbl_style = ts.get(w('val'), '')

    for tr in tbl_el.findall(w('tr')):
        row = []
        for tc in tr.findall(w('tc')):
            cell_parts = []
            for p in tc.findall(w('p')):
                text = _paragraph_to_text(p, rels).strip()
                if text:
                    cell_parts.append(text)
            row.append(' '.join(cell_parts) if cell_parts else '')
        rows_data.append(row)
        max_cols = max(max_cols, len(row))

    if not rows_data or max_cols == 0:
        return '', tbl_style

    for row in rows_data:
        while len(row) < max_cols:
            row.append('')

    widths = [max(3, max(len(rows_data[r][c]) for r in range(len(rows_data))))
              for c in range(max_cols)]
    lines = []
    lines.append('| ' + ' | '.join(rows_data[0][c].ljust(widths[c]) for c in range(max_cols)) + ' |')
    lines.append('|' + '|'.join('-' * (wd + 2) for wd in widths) + '|')
    for row in rows_data[1:]:
        lines.append('| ' + ' | '.join(row[c].ljust(widths[c]) for c in range(max_cols)) + ' |')
    return '\n'.join(lines), tbl_style

# ═══════════════════════════════════════════════════════════════════════
# Section Splitting & Body Processing
# ═══════════════════════════════════════════════════════════════════════

def detect_split_level(body):
    """Auto-detect the highest heading level used in the document."""
    for level in range(1, 7):
        target_styles = {k for k, v in HEADING_MAP.items() if v == level}
        for p in body.iter(w('p')):
            style = _get_style(p)
            if style in target_styles:
                return level
    return 1  # default


def process_body(body, rels, split_level):
    """Walk document body, split into sections, return list of section dicts."""
    split_styles = {k for k, v in HEADING_MAP.items() if v == split_level}
    sections = []
    current = {
        'title': 'Front Matter',
        'heading_level': 0,
        'lines': [],
        'styles_used': set(),
        'images': [],
        'word_count': 0,
    }
    current_style = None

    def flush_section():
        if current['lines'] or current['images']:
            sections.append(current.copy())
            current['styles_used'] = set(current['styles_used'])

    for child in body:
        local = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if local == 'p':
            style = _get_style(child) or 'Normal'
            text = _paragraph_to_text(child, rels)
            num_info = _get_num_info(child)

            # Check for section split
            if style in split_styles and text.strip():
                flush_section()
                current = {
                    'title': text.strip().replace('**', '').replace('*', ''),
                    'heading_level': split_level,
                    'lines': [],
                    'styles_used': set(),
                    'images': [],
                    'word_count': 0,
                }
                current_style = None
                current['styles_used'].add(style)
                continue  # heading is captured as section title, not body content

            current['styles_used'].add(style)

            # Track images
            imgs = _find_images_in_element(child, rels)
            for img in imgs:
                fname = os.path.basename(img)
                if fname not in current['images']:
                    current['images'].append(fname)

            # Page break
            if _has_page_break(child):
                current['lines'].append('')
                current['lines'].append('<!-- pagebreak -->')
                current['lines'].append('')

            # Empty paragraph
            if not text.strip():
                current['lines'].append('')
                continue

            # Word count
            current['word_count'] += len(text.split())

            # List items
            is_list = style in LIST_STYLES or num_info is not None
            if is_list:
                list_type = LIST_STYLES.get(style, 'bullet')
                indent = '  ' * (num_info['ilvl'] if num_info else 0)
                if style != current_style:
                    current['lines'].append(f'<!-- style: {style} -->')
                    current_style = style
                if list_type == 'numbered':
                    current['lines'].append(f'{indent}1. {text}')
                else:
                    current['lines'].append(f'{indent}- {text}')
                continue

            # Style annotation
            if style != current_style:
                current['lines'].append(f'<!-- style: {style} -->')
                current_style = style

            # Headings
            heading_level = HEADING_MAP.get(style)
            if heading_level:
                current['lines'].append(f'{"#" * heading_level} {text}')
            else:
                current['lines'].append(text)

        elif local == 'tbl':
            table_md, tbl_style = _table_to_md(child, rels)
            if tbl_style:
                current['styles_used'].add(tbl_style)
                current['lines'].append(f'<!-- table-style: {tbl_style} -->')
            if table_md:
                current['lines'].append(table_md)
            current['lines'].append('')
            current_style = None

            # Track table images
            imgs = _find_images_in_element(child, rels)
            for img in imgs:
                fname = os.path.basename(img)
                if fname not in current['images']:
                    current['images'].append(fname)

        elif local == 'sdt':
            sdt_content = child.find(w('sdtContent'))
            if sdt_content is not None:
                inner_sections = process_body(sdt_content, rels, split_level)
                if inner_sections:
                    # Merge first inner section into current
                    first = inner_sections[0]
                    current['lines'].extend(first['lines'])
                    current['styles_used'].update(first['styles_used'])
                    current['images'].extend(first['images'])
                    current['word_count'] += first['word_count']
                    # Additional sections become new sections
                    for s in inner_sections[1:]:
                        flush_section()
                        current = s

    # Flush last section
    flush_section()
    return sections

# ═══════════════════════════════════════════════════════════════════════
# Image Extraction
# ═══════════════════════════════════════════════════════════════════════

def extract_images(zf, rels, media_dir):
    """Extract all images from the docx to media_dir."""
    extracted = []
    os.makedirs(media_dir, exist_ok=True)
    for rid, info in rels.items():
        if info['type'] != 'image':
            continue
        target = info['target']
        zip_path = f'word/{target}'
        try:
            data = zf.read(zip_path)
            fname = os.path.basename(target)
            out_path = os.path.join(media_dir, fname)
            with open(out_path, 'wb') as f:
                f.write(data)
            extracted.append(fname)
        except KeyError:
            pass
    return extracted

# ═══════════════════════════════════════════════════════════════════════
# Anomaly Detection
# ═══════════════════════════════════════════════════════════════════════

def detect_anomalies(body, styles_def):
    """Detect style inconsistencies: direct formatting that overrides styles."""
    anomalies = []
    seen_overrides = {}  # style -> set of override descriptions

    for p in body.iter(w('p')):
        style = _get_style(p) or 'Normal'
        pPr = p.find(w('pPr'))
        if pPr is None:
            continue

        # Check for direct paragraph formatting that overrides the style
        rPr_in_pPr = pPr.find(w('rPr'))
        if rPr_in_pPr is not None:
            overrides = []
            for el in rPr_in_pPr:
                local = el.tag.split('}')[-1] if '}' in el.tag else el.tag
                if local in ('b', 'i', 'sz', 'color', 'rFonts'):
                    overrides.append(local)
            if overrides:
                key = f'{style}:{",".join(sorted(overrides))}'
                if key not in seen_overrides:
                    seen_overrides[key] = True
                    anomalies.append({
                        'type': 'direct_override',
                        'style': style,
                        'overrides': overrides,
                        'note': f'Paragraphs with style "{style}" have direct formatting overrides ({", ".join(overrides)}) — may indicate inconsistent styling in source',
                    })

    return anomalies

# ═══════════════════════════════════════════════════════════════════════
# Output Generation
# ═══════════════════════════════════════════════════════════════════════

def write_output(output_dir, source_name, metadata, sections, styles, theme,
                 numbering, layout, headers_footers, anomalies, images):
    """Write all output files."""
    os.makedirs(output_dir, exist_ok=True)
    sections_dir = os.path.join(output_dir, 'sections')
    os.makedirs(sections_dir, exist_ok=True)

    # ── manifest.yaml ──
    section_entries = []
    for i, sec in enumerate(sections):
        fname = f'{i:02d}-{_slugify(sec["title"])}.md'
        sec['_filename'] = fname
        section_entries.append({
            'file': f'sections/{fname}',
            'title': sec['title'],
            'heading_level': sec['heading_level'],
            'word_count': sec['word_count'],
            'images': sec['images'],
        })

    manifest = {
        'source': source_name,
        'decomposed': datetime.now().isoformat(timespec='seconds'),
        'metadata': metadata,
        'page_layout': layout,
        'sections': section_entries,
        'styles_file': 'styles.yaml',
        'media_dir': 'media',
        'stats': {
            'total_sections': len(sections),
            'total_words': sum(s['word_count'] for s in sections),
            'total_images': len(images),
            'styles_used': len(set().union(*(s['styles_used'] for s in sections))),
        },
    }
    if anomalies:
        manifest['anomalies_count'] = len(anomalies)

    with open(os.path.join(output_dir, 'manifest.yaml'), 'w', encoding='utf-8') as f:
        f.write(yaml_dump(manifest) + '\n')

    # ── styles.yaml ──
    styles_data = {
        'theme': theme,
        'defaults': styles.get('defaults', {}),
        'paragraph_styles': styles.get('paragraph', {}),
        'character_styles': styles.get('character', {}),
        'table_styles': styles.get('table', {}),
        'numbering': numbering,
        'headers_footers': headers_footers,
    }
    if anomalies:
        styles_data['anomalies'] = [
            {'style': a['style'], 'note': a['note']} for a in anomalies
        ]

    with open(os.path.join(output_dir, 'styles.yaml'), 'w', encoding='utf-8') as f:
        f.write('# Style definitions extracted from source document\n')
        f.write('# Edit these to customize formatting for reassembly\n\n')
        f.write(yaml_dump(styles_data) + '\n')

    # ── Section files ──
    for i, sec in enumerate(sections):
        fm = {
            'section_index': i,
            'title': sec['title'],
            'heading_level': sec['heading_level'],
            'source_styles': sorted(sec['styles_used']),
            'word_count': sec['word_count'],
        }
        if sec['images']:
            fm['images'] = sec['images']

        content = yaml_front_matter(fm) + '\n\n'
        # Add heading for non-front-matter sections
        if sec['heading_level'] > 0:
            content += '#' * sec['heading_level'] + ' ' + sec['title'] + '\n\n'

        # Strip duplicate heading from beginning of lines (can happen when
        # heading is inside an SDT content control or processed twice)
        lines = list(sec['lines'])
        while lines and not lines[0].strip():
            lines.pop(0)
        # Remove leading style annotation + heading that matches section title
        if lines and lines[0].startswith('<!-- style:'):
            heading_prefix = '#' * sec['heading_level'] + ' '
            if len(lines) > 1 and lines[1].startswith(heading_prefix):
                candidate = lines[1][len(heading_prefix):].strip()
                if candidate == sec['title']:
                    lines.pop(0)  # remove style annotation
                    lines.pop(0)  # remove duplicate heading
        elif lines and lines[0].startswith('#'):
            heading_prefix = '#' * sec['heading_level'] + ' '
            if lines[0].startswith(heading_prefix):
                candidate = lines[0][len(heading_prefix):].strip()
                if candidate == sec['title']:
                    lines.pop(0)

        content += '\n'.join(lines) + '\n'
        # Clean up excessive blank lines
        content = re.sub(r'\n{4,}', '\n\n\n', content)

        fpath = os.path.join(sections_dir, sec['_filename'])
        with open(fpath, 'w', encoding='utf-8') as f:
            f.write(content)

    return manifest


def _slugify(text):
    """Convert text to a filename-safe slug."""
    text = text.lower().strip()
    text = re.sub(r'[^\w\s-]', '', text)
    text = re.sub(r'[\s_]+', '-', text)
    text = re.sub(r'-+', '-', text)
    return text[:50].rstrip('-') or 'untitled'

# ═══════════════════════════════════════════════════════════════════════
# Main
# ═══════════════════════════════════════════════════════════════════════

def decompose(input_path, output_dir=None, split_level=None):
    """Main decomposition entry point."""
    if output_dir is None:
        output_dir = os.path.splitext(input_path)[0]

    source_name = os.path.basename(input_path)
    print(f'Decomposing: {source_name}')
    print(f'Output: {output_dir}/')

    with zipfile.ZipFile(input_path, 'r') as zf:
        metadata = extract_metadata(zf)
        rels = parse_relationships(zf)
        styles = extract_styles(zf)
        theme = extract_theme(zf)
        numbering = extract_numbering(zf)

        doc = ET.fromstring(zf.read('word/document.xml'))
        body = doc.find(w('body'))
        if body is None:
            print('Error: No <w:body> found in document.xml', file=sys.stderr)
            sys.exit(1)

        layout = extract_page_layout(body)
        headers_footers = extract_headers_footers(zf, rels)

        # Auto-detect or use specified split level
        if split_level is None:
            split_level = detect_split_level(body)
        print(f'Splitting on: Heading{split_level}')

        # Process body into sections
        sections = process_body(body, rels, split_level)

        # Detect anomalies
        anomalies = detect_anomalies(body, styles)

        # Extract images
        media_dir = os.path.join(output_dir, 'media')
        images = extract_images(zf, rels, media_dir)

    # Write output
    manifest = write_output(
        output_dir, source_name, metadata, sections, styles, theme,
        numbering, layout, headers_footers, anomalies, images,
    )

    # Summary
    print(f'\nDone:')
    print(f'  Sections: {manifest["stats"]["total_sections"]}')
    print(f'  Words: {manifest["stats"]["total_words"]}')
    print(f'  Images: {manifest["stats"]["total_images"]}')
    print(f'  Styles: {manifest["stats"]["styles_used"]}')
    if anomalies:
        print(f'  ⚠ Anomalies: {len(anomalies)} style inconsistencies detected (see styles.yaml)')
    print(f'\nFiles:')
    print(f'  manifest.yaml')
    print(f'  styles.yaml')
    for sec in sections:
        print(f'  sections/{sec["_filename"]}')
    if images:
        print(f'  media/ ({len(images)} images)')

    return output_dir


def main():
    parser = argparse.ArgumentParser(
        description='Decompose a .docx into modular, AI-editable components.')
    parser.add_argument('input', help='Path to .docx file')
    parser.add_argument('output', nargs='?', default=None, help='Output directory')
    parser.add_argument('--split-on', type=int, default=None, metavar='N',
                        help='Heading level to split on (default: auto-detect)')
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f'Error: File not found: {args.input}', file=sys.stderr)
        sys.exit(1)

    decompose(args.input, args.output, args.split_on)


if __name__ == '__main__':
    main()
