"""Change paragraph styles in docx with tracked changes."""

import copy
import os
import re
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

from .constants import (
    W_NS, BODY_TAG, P_TAG, PPR_TAG, PSTYLE_TAG, VAL_ATTR,
    ID_ATTR, AUTHOR_ATTR, DATE_ATTR, PPR_CHANGE_TAG,
    _register_namespaces,
)


def _find_max_revision_id(root):
    """Find maximum w:id value used in the document."""
    max_id = 0
    for elem in root.iter():
        val = elem.get(ID_ATTR)
        if val is not None:
            try:
                max_id = max(max_id, int(val))
            except ValueError:
                pass
    return max_id


def _find_paragraph_in_raw(raw_xml, p_elem):
    """Find the byte range of a paragraph in the raw XML by matching structure.

    Returns (start, end) indices or None.
    """
    # Collect text from direct runs for fingerprinting
    from .constants import R_TAG, T_TAG
    run_texts = []
    for child in p_elem:
        if child.tag == R_TAG:
            for sub in child:
                if sub.tag == T_TAG and sub.text:
                    run_texts.append(sub.text)

    # Also get pPr/pStyle value for matching empty paragraphs
    ppr = p_elem.find(PPR_TAG)
    pstyle_val = None
    if ppr is not None:
        pstyle = ppr.find(PSTYLE_TAG)
        if pstyle is not None:
            pstyle_val = pstyle.get(VAL_ATTR)

    p_starts = [m.start() for m in re.finditer(r"<w:p[ >]", raw_xml)]

    for p_start in p_starts:
        depth = 0
        i = p_start
        p_end = None
        while i < len(raw_xml):
            if raw_xml[i:i+4] == "<w:p" and (i + 4 < len(raw_xml) and raw_xml[i+4] in " >/"):
                depth += 1
            elif raw_xml[i:i+6] == "</w:p>":
                depth -= 1
                if depth == 0:
                    p_end = i + 6
                    break
            i += 1

        if p_end is None:
            continue

        block = raw_xml[p_start:p_end]

        if run_texts:
            pos = 0
            found_all = True
            for txt in run_texts:
                escaped = txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                idx = block.find(escaped, pos)
                if idx == -1:
                    found_all = False
                    break
                pos = idx + len(escaped)
            if found_all:
                return p_start, p_end
        elif pstyle_val:
            # Empty paragraph — match by style
            if f'w:val="{pstyle_val}"' in block:
                return p_start, p_end

    return None


def _serialize_paragraph(p_elem):
    """Serialize a paragraph element preserving w: namespace prefix."""
    wrapper = ET.Element("_wrapper")
    wrapper.set("xmlns:w", W_NS)
    wrapper.append(p_elem)
    raw = ET.tostring(wrapper, encoding="unicode")
    start = raw.index("<w:p")
    end = raw.rindex("</w:p>") + len("</w:p>")
    return raw[start:end]


def change_style_in_docx(input_path, output_path, paragraph, new_style, author):
    """Change the paragraph style as a tracked change.

    Returns (success, message).
    """
    date_str = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    with zipfile.ZipFile(input_path, "r") as zf:
        raw_xml = zf.read("word/document.xml")
    _register_namespaces(raw_xml)
    root = ET.fromstring(raw_xml)

    body = root.find(BODY_TAG)
    if body is None:
        return False, "Could not find document body"

    paragraphs = list(body.iter(P_TAG))
    if paragraph < 1 or paragraph > len(paragraphs):
        return False, f"Paragraph {paragraph} out of range (1-{len(paragraphs)})"

    p_elem = paragraphs[paragraph - 1]

    # Get current style
    ppr = p_elem.find(PPR_TAG)
    old_style = None
    if ppr is not None:
        pstyle = ppr.find(PSTYLE_TAG)
        if pstyle is not None:
            old_style = pstyle.get(VAL_ATTR)

    if old_style == new_style:
        return False, f"Paragraph {paragraph} already has style '{new_style}'"

    # Find paragraph in raw XML before modifying
    raw_str = raw_xml.decode("utf-8")
    span = _find_paragraph_in_raw(raw_str, p_elem)
    if span is None:
        return False, "Could not locate paragraph in raw XML for splicing"

    rev_id = _find_max_revision_id(root) + 1

    # Build the pPrChange element (stores old properties)
    if ppr is None:
        ppr = ET.SubElement(p_elem, PPR_TAG)
        # Insert pPr as first child
        p_elem.remove(ppr)
        p_elem.insert(0, ppr)

    # Remove any existing pPrChange (shouldn't normally be there)
    existing_change = ppr.find(PPR_CHANGE_TAG)
    if existing_change is not None:
        ppr.remove(existing_change)

    # Create pPrChange with the OLD properties
    ppr_change = ET.SubElement(ppr, PPR_CHANGE_TAG)
    ppr_change.set(ID_ATTR, str(rev_id))
    ppr_change.set(AUTHOR_ATTR, author)
    ppr_change.set(DATE_ATTR, date_str)

    # Store old pPr inside pPrChange
    old_ppr = ET.SubElement(ppr_change, PPR_TAG)
    if old_style is not None:
        old_pstyle = ET.SubElement(old_ppr, PSTYLE_TAG)
        old_pstyle.set(VAL_ATTR, old_style)

    # Now set the new style on the actual pPr
    pstyle = ppr.find(PSTYLE_TAG)
    if pstyle is not None:
        pstyle.set(VAL_ATTR, new_style)
    else:
        pstyle = ET.Element(PSTYLE_TAG)
        pstyle.set(VAL_ATTR, new_style)
        # Insert pStyle as first child of pPr (before pPrChange)
        ppr.insert(0, pstyle)

    # Splice modified paragraph into raw XML
    start, end = span
    new_p_xml = _serialize_paragraph(p_elem)
    output_str = raw_str[:start] + new_p_xml + raw_str[end:]
    output_bytes = output_str.encode("utf-8")

    # Write new docx
    use_temp = os.path.abspath(input_path) == os.path.abspath(output_path)
    dest = output_path
    if use_temp:
        fd, dest = tempfile.mkstemp(suffix=".docx")
        os.close(fd)

    with zipfile.ZipFile(input_path, "r") as zin:
        with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, output_bytes)
                else:
                    zout.writestr(item, zin.read(item.filename))

    if use_temp:
        shutil.move(dest, output_path)

    old_label = old_style or "(default)"
    return True, f"Changed paragraph {paragraph} style: {old_label} -> {new_style}"
