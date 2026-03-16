"""Replace text in docx as tracked changes."""

import copy
import os
import re
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

from .constants import (
    W_NS, BODY_TAG, DEL_TAG, DELTEXT_TAG, INS_TAG, P_TAG, R_TAG, T_TAG,
    RPR_TAG, AUTHOR_ATTR, DATE_ATTR, ID_ATTR, XML_SPACE_ATTR,
    _register_namespaces,
)


def _make_t_elem(text):
    """Create a w:t element with proper xml:space handling."""
    t = ET.Element(T_TAG)
    t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set(XML_SPACE_ATTR, "preserve")
    return t


def _make_deltext_elem(text):
    """Create a w:delText element."""
    t = ET.Element(DELTEXT_TAG)
    t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set(XML_SPACE_ATTR, "preserve")
    return t


def _get_run_text(run):
    """Get concatenated text from all w:t elements in a run."""
    parts = []
    for sub in run:
        if sub.tag == T_TAG and sub.text:
            parts.append(sub.text)
    return "".join(parts)


def _clone_run_with_text(run, text, use_deltext=False):
    """Create a new run preserving rPr but with new text content."""
    new_run = ET.Element(R_TAG)
    rpr = run.find(RPR_TAG)
    if rpr is not None:
        new_run.append(copy.deepcopy(rpr))
    if use_deltext:
        new_run.append(_make_deltext_elem(text))
    else:
        new_run.append(_make_t_elem(text))
    return new_run


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


def _get_paragraph_plain_text(p_elem):
    """Get plain text from direct runs in a paragraph (skipping ins/del)."""
    parts = []
    for child in p_elem:
        if child.tag == R_TAG:
            parts.append(_get_run_text(child))
    return "".join(parts)


def _replace_in_paragraph(p_elem, old_text, new_text, author, date_str, rev_id,
                           context=None):
    """Replace old_text with new_text as tracked change in paragraph.

    Returns (success, next_rev_id).
    """
    # Collect direct run children with their text
    run_info = []  # (run_elem, text)
    for child in p_elem:
        if child.tag == R_TAG:
            text = _get_run_text(child)
            run_info.append((child, text))

    if not run_info:
        return False, rev_id

    full_text = "".join(ri[1] for ri in run_info)

    # Find match position, optionally scoped by context
    if context is not None:
        ctx_pos = full_text.find(context)
        if ctx_pos == -1:
            return False, rev_id
        pos = full_text.find(old_text, ctx_pos, ctx_pos + len(context))
    else:
        pos = full_text.find(old_text)

    if pos == -1:
        return False, rev_id

    match_end = pos + len(old_text)

    # Map match boundaries to run indices
    char_offset = 0
    first_ri = last_ri = None
    first_offset = last_end_offset = 0

    for ri_idx, (_, text) in enumerate(run_info):
        run_start = char_offset
        run_end = char_offset + len(text)

        if first_ri is None and run_end > pos:
            first_ri = ri_idx
            first_offset = pos - run_start

        if run_end >= match_end:
            last_ri = ri_idx
            last_end_offset = match_end - run_start
            break

        char_offset = run_end

    if first_ri is None or last_ri is None:
        return False, rev_id

    # Build replacement elements
    new_elements = []

    first_run, first_text = run_info[first_ri]
    before_text = first_text[:first_offset]
    if before_text:
        new_elements.append(_clone_run_with_text(first_run, before_text))

    # <w:del> element
    del_elem = ET.Element(DEL_TAG)
    del_elem.set(ID_ATTR, str(rev_id))
    del_elem.set(AUTHOR_ATTR, author)
    del_elem.set(DATE_ATTR, date_str)
    rev_id += 1

    if first_ri == last_ri:
        del_text = first_text[first_offset:last_end_offset]
        del_elem.append(_clone_run_with_text(first_run, del_text, use_deltext=True))
    else:
        del_elem.append(_clone_run_with_text(
            first_run, first_text[first_offset:], use_deltext=True))
        for ri_idx in range(first_ri + 1, last_ri):
            mid_run, mid_text = run_info[ri_idx]
            del_elem.append(_clone_run_with_text(mid_run, mid_text, use_deltext=True))
        last_run, last_text = run_info[last_ri]
        del_elem.append(_clone_run_with_text(
            last_run, last_text[:last_end_offset], use_deltext=True))

    new_elements.append(del_elem)

    # <w:ins> element (only if new_text is non-empty)
    if new_text:
        ins_elem = ET.Element(INS_TAG)
        ins_elem.set(ID_ATTR, str(rev_id))
        ins_elem.set(AUTHOR_ATTR, author)
        ins_elem.set(DATE_ATTR, date_str)
        rev_id += 1
        ins_elem.append(_clone_run_with_text(first_run, new_text))
        new_elements.append(ins_elem)

    # "after" part of the last run
    if first_ri == last_ri:
        after_text = first_text[last_end_offset:]
    else:
        _, last_text = run_info[last_ri]
        after_text = last_text[last_end_offset:]

    if after_text:
        after_run = run_info[last_ri][0]
        new_elements.append(_clone_run_with_text(after_run, after_text))

    # Find insertion point, remove old runs, insert new elements
    children = list(p_elem)
    insert_pos = children.index(run_info[first_ri][0])

    for ri_idx in range(first_ri, last_ri + 1):
        p_elem.remove(run_info[ri_idx][0])

    for i, elem in enumerate(new_elements):
        p_elem.insert(insert_pos + i, elem)

    return True, rev_id


def _serialize_paragraph(p_elem):
    """Serialize a paragraph element to an XML string using w: namespace prefix.

    Uses a temporary wrapper with the w: namespace declaration so that
    ET.tostring produces w: prefixed tags, then extracts just the <w:p> part.
    """
    wrapper = ET.Element("_wrapper")
    wrapper.set("xmlns:w", W_NS)
    wrapper.append(p_elem)
    raw = ET.tostring(wrapper, encoding="unicode")
    # Extract the <w:p ...>...</w:p> from inside the wrapper
    start = raw.index("<w:p")
    end = raw.rindex("</w:p>") + len("</w:p>")
    return raw[start:end]


def _find_paragraph_in_raw(raw_xml, p_elem):
    """Find the byte range of a paragraph in the raw XML.

    Uses unique text content to locate the paragraph reliably.
    Returns (start, end) indices or None.
    """
    # Get all direct run texts to build a unique fingerprint
    run_texts = []
    for child in p_elem:
        if child.tag == R_TAG:
            for sub in child:
                if sub.tag == T_TAG and sub.text:
                    run_texts.append(sub.text)

    # Strategy: find the <w:p> that contains these exact run texts
    # Search for all <w:p ...>...</w:p> blocks
    p_starts = [m.start() for m in re.finditer(r"<w:p[ >]", raw_xml)]

    for p_start in p_starts:
        # Find the matching </w:p>
        depth = 0
        i = p_start
        p_end = None
        while i < len(raw_xml):
            if raw_xml[i:i+4] == "<w:p" and (raw_xml[i+4] in " >/"):
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
        # Check if this block contains all our run texts in order
        if not run_texts:
            continue
        pos = 0
        found_all = True
        for txt in run_texts:
            # The text appears as content of <w:t> elements, possibly with xml:space
            escaped = txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            idx = block.find(escaped, pos)
            if idx == -1:
                found_all = False
                break
            pos = idx + len(escaped)

        if found_all:
            return p_start, p_end

    return None


def replace_in_docx(input_path, output_path, old_text, new_text, author,
                    paragraph=None, context=None):
    """Replace old_text with new_text as a tracked change.

    Returns (success, message).
    """
    date_str = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    with zipfile.ZipFile(input_path, "r") as zf:
        raw_doc = zf.read("word/document.xml")
    _register_namespaces(raw_doc)
    root = ET.fromstring(raw_doc)

    body = root.find(BODY_TAG)
    if body is None:
        return False, "Could not find document body"

    # Collect all paragraphs (including inside tables)
    all_paragraphs = list(body.iter(P_TAG))

    if paragraph is not None:
        if paragraph < 1 or paragraph > len(all_paragraphs):
            return False, f"Paragraph {paragraph} out of range (1-{len(all_paragraphs)})"
        candidates = [all_paragraphs[paragraph - 1]]
    else:
        candidates = all_paragraphs

    rev_id = _find_max_revision_id(root) + 1
    replaced = False
    target_p = None

    for p_elem in candidates:
        # Quick check: does this paragraph contain the text?
        p_text = _get_paragraph_plain_text(p_elem)
        if old_text not in p_text:
            continue
        if context is not None and context not in p_text:
            continue

        # Save original run texts for locating in raw XML
        orig_run_texts = []
        for child in p_elem:
            if child.tag == R_TAG:
                for sub in child:
                    if sub.tag == T_TAG and sub.text:
                        orig_run_texts.append(sub.text)

        success, rev_id = _replace_in_paragraph(
            p_elem, old_text, new_text, author, date_str, rev_id, context)
        if success:
            replaced = True
            target_p = p_elem
            target_run_texts = orig_run_texts
            break

    if not replaced:
        return False, "Text not found"

    # Locate the original paragraph in raw XML and splice in the modified version
    raw_str = raw_doc.decode("utf-8")

    # Build a dummy p_elem with the original run texts for matching
    dummy_p = ET.Element(P_TAG)
    for txt in target_run_texts:
        r = ET.SubElement(dummy_p, R_TAG)
        t = ET.SubElement(r, T_TAG)
        t.text = txt

    span = _find_paragraph_in_raw(raw_str, dummy_p)
    if span is None:
        return False, "Could not locate paragraph in raw XML for splicing"

    start, end = span
    new_p_xml = _serialize_paragraph(target_p)
    output_str = raw_str[:start] + new_p_xml + raw_str[end:]
    output_bytes = output_str.encode("utf-8")

    # Write new docx (use temp file for safe in-place overwrite)
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
        os.replace(dest, output_path)

    return True, f"Replaced in {output_path}"
