#!/usr/bin/env python3
"""Extract and display style definitions from a .dotx or .docx template.

Usage:
    python extract_styles.py [template_path]

If no path is given, uses Brand_template_20200929.dotx in the same directory.
"""

import sys
import os
import zipfile
import xml.etree.ElementTree as ET

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}


def get_attr(el, attr):
    """Get a w: namespaced attribute value."""
    return el.get(f'{{{NS["w"]}}}{attr}') if el is not None else None


def extract_run_props(rpr):
    """Extract run (character) formatting properties."""
    if rpr is None:
        return {}
    props = {}

    fonts = rpr.find('w:rFonts', NS)
    if fonts is not None:
        ascii_font = fonts.get(f'{{{NS["w"]}}}ascii')
        theme_font = fonts.get(f'{{{NS["w"]}}}asciiTheme')
        if ascii_font:
            props['font'] = ascii_font
        elif theme_font:
            props['font'] = f'(theme: {theme_font})'

    sz = rpr.find('w:sz', NS)
    if sz is not None:
        half_pts = int(get_attr(sz, 'val'))
        props['size'] = f'{half_pts / 2}pt'

    bold = rpr.find('w:b', NS)
    if bold is not None:
        val = get_attr(bold, 'val')
        props['bold'] = val != '0' if val else True

    italic = rpr.find('w:i', NS)
    if italic is not None:
        val = get_attr(italic, 'val')
        props['italic'] = val != '0' if val else True

    color = rpr.find('w:color', NS)
    if color is not None:
        props['color'] = get_attr(color, 'val')
        theme_color = get_attr(color, 'themeColor')
        if theme_color:
            props['themeColor'] = theme_color

    return props


def extract_para_props(ppr):
    """Extract paragraph formatting properties."""
    if ppr is None:
        return {}
    props = {}

    spacing = ppr.find('w:spacing', NS)
    if spacing is not None:
        before = get_attr(spacing, 'before')
        after = get_attr(spacing, 'after')
        line = get_attr(spacing, 'line')
        line_rule = get_attr(spacing, 'lineRule')
        if before:
            props['spaceBefore'] = f'{before} twips'
        if after:
            props['spaceAfter'] = f'{after} twips'
        if line:
            rule = f' ({line_rule})' if line_rule else ''
            props['lineSpacing'] = f'{line}{rule}'

    outline = ppr.find('w:outlineLvl', NS)
    if outline is not None:
        props['outlineLevel'] = get_attr(outline, 'val')

    ind = ppr.find('w:ind', NS)
    if ind is not None:
        left = get_attr(ind, 'left')
        hanging = get_attr(ind, 'hanging')
        if left:
            props['indentLeft'] = f'{left} twips'
        if hanging:
            props['hanging'] = f'{hanging} twips'

    jc = ppr.find('w:jc', NS)
    if jc is not None:
        props['alignment'] = get_attr(jc, 'val')

    for flag in ['keepNext', 'keepLines']:
        el = ppr.find(f'w:{flag}', NS)
        if el is not None:
            props[flag] = True

    return props


def extract_styles(zip_path):
    """Extract all styles from a .docx/.dotx file."""
    with zipfile.ZipFile(zip_path, 'r') as z:
        with z.open('word/styles.xml') as f:
            tree = ET.parse(f)

    root = tree.getroot()
    styles = []

    for style_el in root.findall('w:style', NS):
        style_type = get_attr(style_el, 'type')
        style_id = get_attr(style_el, 'styleId')
        is_default = get_attr(style_el, 'default') == '1'
        is_custom = get_attr(style_el, 'customStyle') == '1'

        name_el = style_el.find('w:name', NS)
        name = get_attr(name_el, 'val') if name_el is not None else style_id

        aliases_el = style_el.find('w:aliases', NS)
        aliases = get_attr(aliases_el, 'val') if aliases_el is not None else None

        based_on_el = style_el.find('w:basedOn', NS)
        based_on = get_attr(based_on_el, 'val') if based_on_el is not None else None

        next_el = style_el.find('w:next', NS)
        next_style = get_attr(next_el, 'val') if next_el is not None else None

        semi_hidden = style_el.find('w:semiHidden', NS) is not None
        qformat = style_el.find('w:qFormat', NS) is not None

        ui_priority_el = style_el.find('w:uiPriority', NS)
        ui_priority = get_attr(ui_priority_el, 'val') if ui_priority_el is not None else None

        rpr = style_el.find('w:rPr', NS)
        ppr = style_el.find('w:pPr', NS)

        run_props = extract_run_props(rpr)
        para_props = extract_para_props(ppr)

        styles.append({
            'type': style_type,
            'id': style_id,
            'name': name,
            'aliases': aliases,
            'basedOn': based_on,
            'next': next_style,
            'default': is_default,
            'custom': is_custom,
            'semiHidden': semi_hidden,
            'qFormat': qformat,
            'uiPriority': ui_priority,
            'run': run_props,
            'paragraph': para_props,
        })

    return styles


def extract_theme(zip_path):
    """Extract theme colors and fonts."""
    with zipfile.ZipFile(zip_path, 'r') as z:
        with z.open('word/theme/theme1.xml') as f:
            tree = ET.parse(f)

    root = tree.getroot()
    info = {'colors': {}, 'fonts': {}}

    scheme = root.find('.//a:clrScheme', NS)
    if scheme is not None:
        info['colorSchemeName'] = scheme.get('name', '')
        for child in scheme:
            tag = child.tag.split('}')[-1]
            clr = child.find('a:srgbClr', NS)
            if clr is None:
                clr = child.find('a:sysClr', NS)
            if clr is not None:
                val = clr.get('val', '') or clr.get('lastClr', '')
                info['colors'][tag] = val

    font_scheme = root.find('.//a:fontScheme', NS)
    if font_scheme is not None:
        info['fontSchemeName'] = font_scheme.get('name', '')
        major = font_scheme.find('.//a:majorFont/a:latin', NS)
        minor = font_scheme.find('.//a:minorFont/a:latin', NS)
        if major is not None:
            info['fonts']['major'] = major.get('typeface', '')
        if minor is not None:
            info['fonts']['minor'] = minor.get('typeface', '')

    return info


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_template = os.path.join(script_dir, '..', 'Brand_template_20200929.dotx')

    template_path = sys.argv[1] if len(sys.argv) > 1 else default_template

    if not os.path.exists(template_path):
        print(f"Error: Template not found: {template_path}", file=sys.stderr)
        sys.exit(1)

    print(f"Analyzing: {template_path}\n")

    # Theme
    theme = extract_theme(template_path)
    print("=" * 60)
    print(f"THEME: {theme.get('colorSchemeName', 'Unknown')}")
    print("=" * 60)

    print(f"\nFont Scheme: {theme.get('fontSchemeName', 'Unknown')}")
    for role, font in theme['fonts'].items():
        print(f"  {role}: {font}")

    print(f"\nColor Palette:")
    for name, val in theme['colors'].items():
        print(f"  {name:12s} #{val}")

    # Styles
    styles = extract_styles(template_path)

    # Group by type
    by_type = {}
    for s in styles:
        by_type.setdefault(s['type'], []).append(s)

    for stype in ['paragraph', 'character', 'table', 'numbering']:
        group = by_type.get(stype, [])
        if not group:
            continue

        print(f"\n{'=' * 60}")
        print(f"{stype.upper()} STYLES ({len(group)})")
        print("=" * 60)

        # Show visible/important styles first
        visible = [s for s in group if not s['semiHidden']]
        hidden = [s for s in group if s['semiHidden']]

        for s in visible:
            flags = []
            if s['default']:
                flags.append('DEFAULT')
            if s['custom']:
                flags.append('CUSTOM')
            if s['qFormat']:
                flags.append('QUICK')
            flag_str = f" [{', '.join(flags)}]" if flags else ''

            print(f"\n  {s['id']}{flag_str}")
            print(f"    Name: {s['name']}")
            if s['aliases']:
                print(f"    Aliases: {s['aliases']}")
            if s['basedOn']:
                print(f"    Based on: {s['basedOn']}")
            if s['next']:
                print(f"    Next: {s['next']}")

            if s['run']:
                parts = []
                for k, v in s['run'].items():
                    parts.append(f"{k}={v}")
                print(f"    Run: {', '.join(parts)}")

            if s['paragraph']:
                parts = []
                for k, v in s['paragraph'].items():
                    parts.append(f"{k}={v}")
                print(f"    Para: {', '.join(parts)}")

        if hidden:
            print(f"\n  --- Hidden styles ({len(hidden)}): ", end='')
            print(', '.join(s['id'] for s in hidden))


if __name__ == '__main__':
    main()
