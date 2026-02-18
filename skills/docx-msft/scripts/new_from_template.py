#!/usr/bin/env python3
"""Create a new .docx from the brand .dotx template.

A .dotx is identical to .docx but with a different content type.
This script copies the template, adjusts the content type from template to
document format, strips all sample content from document.xml (keeping styles,
theme, headers, footers, and media), and writes a ready-to-edit .docx file.

Usage:
    python new_from_template.py output.docx [template_path]
"""

import sys
import os
import zipfile
import xml.etree.ElementTree as ET

# XML namespaces used in docx
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
}


def _strip_document_body(xml_bytes):
    """Replace document.xml body with an empty body, preserving the root element
    and all namespace declarations.

    Returns the modified XML as bytes.
    """
    # Register all namespaces so ET doesn't mangle prefixes
    for prefix, uri in NAMESPACES.items():
        ET.register_namespace(prefix, uri)
    # Also handle common namespaces that appear in docx
    ET.register_namespace('', NAMESPACES['w'])

    # Parse — but we need to preserve namespace prefixes on the root element.
    # ET.parse drops some, so we do a hybrid approach: rebuild a minimal document.
    tree = ET.parse_from_bytes(xml_bytes) if hasattr(ET, 'parse_from_bytes') else None

    # Safer approach: manually rebuild minimal document.xml
    # Read root tag and its namespace declarations from original bytes
    text = xml_bytes.decode('utf-8')

    # Find the opening <w:document ...> tag (may span multiple lines)
    import re
    doc_open_match = re.search(r'<w:document[^>]*>', text, re.DOTALL)
    if not doc_open_match:
        # Fallback: try without namespace prefix
        doc_open_match = re.search(
            r'<\w+:document[^>]*>|<document[^>]*>', text, re.DOTALL
        )

    if not doc_open_match:
        raise ValueError("Could not find <w:document> element in document.xml")

    doc_open_tag = doc_open_match.group(0)

    # Find the closing tag
    close_match = re.search(r'</w:document>|</\w+:document>|</document>', text)
    close_tag = close_match.group(0) if close_match else '</w:document>'

    # Extract XML declaration if present
    xml_decl = ''
    if text.startswith('<?xml'):
        decl_end = text.index('?>') + 2
        xml_decl = text[:decl_end]

    # Build minimal document with empty body
    minimal = f'{xml_decl}\n{doc_open_tag}<w:body><w:sectPr/></w:body>{close_tag}'

    # Try to preserve the original sectPr (page size, margins, columns, headers/footers refs)
    tree = ET.ElementTree(ET.fromstring(xml_bytes))
    root = tree.getroot()
    ns_w = NAMESPACES['w']
    body = root.find(f'{{{ns_w}}}body')
    if body is not None:
        sect_pr = body.find(f'{{{ns_w}}}sectPr')
        if sect_pr is not None:
            # Re-serialize sectPr and inject it
            sect_xml = ET.tostring(sect_pr, encoding='unicode')
            minimal = f'{xml_decl}\n{doc_open_tag}<w:body>{sect_xml}</w:body>{close_tag}'

    return minimal.encode('utf-8')


def dotx_to_docx(template_path, output_path):
    """Convert a .dotx template to a clean .docx document.

    - Changes content type from template to document
    - Strips all sample content from document.xml body (keeps sectPr for
      page layout, headers/footers references)
    - Preserves: styles.xml, theme, numbering, headers, footers, media
    """
    template_content_type = (
        'application/vnd.openxmlformats-officedocument.'
        'wordprocessingml.template.main+xml'
    )
    document_content_type = (
        'application/vnd.openxmlformats-officedocument.'
        'wordprocessingml.document.main+xml'
    )

    with zipfile.ZipFile(template_path, 'r') as zin:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                # Fix content type declaration
                if item.filename == '[Content_Types].xml':
                    content = data.decode('utf-8')
                    content = content.replace(
                        template_content_type,
                        document_content_type
                    )
                    data = content.encode('utf-8')

                # Strip sample content from document body
                elif item.filename == 'word/document.xml':
                    data = _strip_document_body(data)

                zout.writestr(item, data)

    print(f"Created: {output_path}")
    print(f"  From template: {template_path}")
    print(f"  Sample content stripped — clean empty document.")
    print(f"  All brand styles, theme, headers, footers, and media preserved.")


def main():
    if len(sys.argv) < 2:
        print("Usage: python new_from_template.py output.docx [template_path]",
              file=sys.stderr)
        sys.exit(1)

    output_path = sys.argv[1]

    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_template = os.path.join(
        script_dir, '..', 'Brand_template_20200929.dotx'
    )
    template_path = sys.argv[2] if len(sys.argv) > 2 else default_template

    if not os.path.exists(template_path):
        print(f"Error: Template not found: {template_path}", file=sys.stderr)
        sys.exit(1)

    if os.path.exists(output_path):
        print(f"Error: Output file already exists: {output_path}",
              file=sys.stderr)
        print("Delete it first or choose a different name.", file=sys.stderr)
        sys.exit(1)

    dotx_to_docx(template_path, output_path)


if __name__ == '__main__':
    main()
