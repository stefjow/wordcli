"""Apply/remove run formatting in docx with tracked changes."""

import copy
import os
import re
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

from .constants import (
    W_NS, BODY_TAG, P_TAG, R_TAG, T_TAG, RPR_TAG,
    ID_ATTR, AUTHOR_ATTR, DATE_ATTR, XML_SPACE_ATTR,
    _register_namespaces,
)
from .matching import find_matching_paragraphs, select_match

RPR_CHANGE_TAG = f"{{{W_NS}}}rPrChange"
B_TAG = f"{{{W_NS}}}b"
B_CS_TAG = f"{{{W_NS}}}bCs"
I_TAG = f"{{{W_NS}}}i"
I_CS_TAG = f"{{{W_NS}}}iCs"
U_TAG = f"{{{W_NS}}}u"
STRIKE_TAG = f"{{{W_NS}}}strike"
VAL_ATTR = f"{{{W_NS}}}val"

# Formatting property tags and their "complex script" counterparts
FORMAT_PROPS = {
    "bold": (B_TAG, B_CS_TAG),
    "italic": (I_TAG, I_CS_TAG),
    "underline": (U_TAG, None),
    "strike": (STRIKE_TAG, None),
}


def _get_run_text(run):
    parts = []
    for sub in run:
        if sub.tag == T_TAG and sub.text:
            parts.append(sub.text)
    return "".join(parts)


def _make_t_elem(text):
    t = ET.Element(T_TAG)
    t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set(XML_SPACE_ATTR, "preserve")
    return t


def _clone_run_with_text(run, text):
    """Create a new run preserving rPr but with new text content."""
    new_run = ET.Element(R_TAG)
    rpr = run.find(RPR_TAG)
    if rpr is not None:
        new_run.append(copy.deepcopy(rpr))
    new_run.append(_make_t_elem(text))
    return new_run


def _has_prop(rpr, tag):
    """Check if rPr has a formatting property enabled."""
    if rpr is None:
        return False
    elem = rpr.find(tag)
    if elem is None:
        return False
    # <w:b/> means true; <w:b w:val="false"/> or <w:b w:val="0"/> means false
    val = elem.get(VAL_ATTR)
    if val is not None and val.lower() in ("false", "0"):
        return False
    return True


def _set_prop(rpr, tag, enable):
    """Set or remove a formatting property on rPr."""
    existing = rpr.find(tag)
    if enable:
        if existing is None:
            ET.SubElement(rpr, tag)
        else:
            # Remove val="false" if present
            if VAL_ATTR in existing.attrib:
                del existing.attrib[VAL_ATTR]
    else:
        if existing is not None:
            rpr.remove(existing)


def _apply_format_to_run(run, changes, author, date_str, rev_id):
    """Apply formatting changes to a run, adding rPrChange for tracking.

    changes is a dict like {"bold": True, "italic": False}.
    Returns next_rev_id.
    """
    rpr = run.find(RPR_TAG)

    # Check if any change is actually needed
    needs_change = False
    for prop_name, enable in changes.items():
        tag, _ = FORMAT_PROPS[prop_name]
        currently_set = _has_prop(rpr, tag)
        if currently_set != enable:
            needs_change = True
            break

    if not needs_change:
        return rev_id

    # Save old rPr for tracking
    if rpr is not None:
        old_rpr = copy.deepcopy(rpr)
    else:
        old_rpr = ET.Element(RPR_TAG)

    # Create rPr if needed
    if rpr is None:
        rpr = ET.Element(RPR_TAG)
        run.insert(0, rpr)

    # Remove any existing rPrChange
    existing_change = rpr.find(RPR_CHANGE_TAG)
    if existing_change is not None:
        rpr.remove(existing_change)

    # Apply new formatting
    for prop_name, enable in changes.items():
        tag, cs_tag = FORMAT_PROPS[prop_name]
        _set_prop(rpr, tag, enable)
        if cs_tag:
            _set_prop(rpr, cs_tag, enable)

    # Add rPrChange with old properties
    rpr_change = ET.SubElement(rpr, RPR_CHANGE_TAG)
    rpr_change.set(ID_ATTR, str(rev_id))
    rpr_change.set(AUTHOR_ATTR, author)
    rpr_change.set(DATE_ATTR, date_str)
    # Remove any rPrChange from the old copy before nesting
    old_change = old_rpr.find(RPR_CHANGE_TAG)
    if old_change is not None:
        old_rpr.remove(old_change)
    rpr_change.append(old_rpr)

    return rev_id + 1


def _format_in_paragraph(p_elem, text, changes, author, date_str, rev_id,
                         context=None):
    """Apply formatting to matched text in a paragraph.

    Returns (success, next_rev_id).
    """
    # Collect direct run children with their text
    run_info = []
    for child in p_elem:
        if child.tag == R_TAG:
            run_text = _get_run_text(child)
            run_info.append((child, run_text))

    if not run_info:
        return False, rev_id

    full_text = "".join(ri[1] for ri in run_info)

    # Find match position
    if context is not None:
        ctx_pos = full_text.find(context)
        if ctx_pos == -1:
            return False, rev_id
        pos = full_text.find(text, ctx_pos, ctx_pos + len(context))
    else:
        pos = full_text.find(text)

    if pos == -1:
        return False, rev_id

    match_end = pos + len(text)

    # Map match boundaries to run indices
    char_offset = 0
    first_ri = last_ri = None
    first_offset = last_end_offset = 0

    for ri_idx, (_, run_text) in enumerate(run_info):
        run_start = char_offset
        run_end = char_offset + len(run_text)

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

    # Split runs at match boundaries and apply formatting
    new_elements = []

    first_run, first_text = run_info[first_ri]
    before_text = first_text[:first_offset]
    if before_text:
        new_elements.append(_clone_run_with_text(first_run, before_text))

    # Build matched runs (split at boundaries, apply formatting)
    if first_ri == last_ri:
        matched_text = first_text[first_offset:last_end_offset]
        matched_run = _clone_run_with_text(first_run, matched_text)
        rev_id = _apply_format_to_run(matched_run, changes, author, date_str, rev_id)
        new_elements.append(matched_run)
    else:
        # First partial run
        partial = _clone_run_with_text(first_run, first_text[first_offset:])
        rev_id = _apply_format_to_run(partial, changes, author, date_str, rev_id)
        new_elements.append(partial)
        # Middle full runs
        for ri_idx in range(first_ri + 1, last_ri):
            mid_run, mid_text = run_info[ri_idx]
            cloned = _clone_run_with_text(mid_run, mid_text)
            rev_id = _apply_format_to_run(cloned, changes, author, date_str, rev_id)
            new_elements.append(cloned)
        # Last partial run
        last_run, last_text = run_info[last_ri]
        partial = _clone_run_with_text(last_run, last_text[:last_end_offset])
        rev_id = _apply_format_to_run(partial, changes, author, date_str, rev_id)
        new_elements.append(partial)

    # After part
    if first_ri == last_ri:
        after_text = first_text[last_end_offset:]
    else:
        _, last_text = run_info[last_ri]
        after_text = last_text[last_end_offset:]

    if after_text:
        after_run = run_info[last_ri][0]
        new_elements.append(_clone_run_with_text(after_run, after_text))

    # Replace old runs with new elements
    children = list(p_elem)
    insert_pos = children.index(run_info[first_ri][0])

    for ri_idx in range(first_ri, last_ri + 1):
        p_elem.remove(run_info[ri_idx][0])

    for i, elem in enumerate(new_elements):
        p_elem.insert(insert_pos + i, elem)

    return True, rev_id


def _find_max_revision_id(root):
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
    """Find the byte range of a paragraph in the raw XML."""
    run_texts = []
    for child in p_elem:
        if child.tag == R_TAG:
            for sub in child:
                if sub.tag == T_TAG and sub.text:
                    run_texts.append(sub.text)

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
        if not run_texts:
            continue
        found_pos = 0
        found_all = True
        for txt in run_texts:
            escaped = txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            idx = block.find(escaped, found_pos)
            if idx == -1:
                found_all = False
                break
            found_pos = idx + len(escaped)
        if found_all:
            return p_start, p_end
    return None


def _serialize_paragraph(p_elem):
    wrapper = ET.Element("_wrapper")
    wrapper.set("xmlns:w", W_NS)
    wrapper.append(p_elem)
    raw = ET.tostring(wrapper, encoding="unicode")
    start = raw.index("<w:p")
    end = raw.rindex("</w:p>") + len("</w:p>")
    return raw[start:end]


def format_in_docx(input_path, output_path, text, changes, author,
                   paragraph=None, context=None, occurrence=None):
    """Apply formatting changes to matched text as tracked changes.

    changes is a dict like {"bold": True, "italic": False}.
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

    matches, err = find_matching_paragraphs(body, text, paragraph, context)
    if err:
        return False, err

    p_elem, _, err = select_match(matches, text, occurrence)
    if err:
        return False, err

    # Save original run texts for locating in raw XML
    orig_run_texts = []
    for child in p_elem:
        if child.tag == R_TAG:
            for sub in child:
                if sub.tag == T_TAG and sub.text:
                    orig_run_texts.append(sub.text)

    rev_id = _find_max_revision_id(root) + 1

    success, rev_id = _format_in_paragraph(
        p_elem, text, changes, author, date_str, rev_id, context)
    if not success:
        return False, "Text not found in paragraph"

    # Splice into raw XML
    raw_str = raw_xml.decode("utf-8")

    # Build dummy paragraph with original texts for locating
    dummy_p = ET.Element(P_TAG)
    for txt in orig_run_texts:
        r = ET.SubElement(dummy_p, R_TAG)
        t = ET.SubElement(r, T_TAG)
        t.text = txt

    span = _find_paragraph_in_raw(raw_str, dummy_p)
    if span is None:
        return False, "Could not locate paragraph in raw XML for splicing"

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

    props = []
    for prop_name, enable in changes.items():
        props.append(f"+{prop_name}" if enable else f"-{prop_name}")
    return True, f"Formatted \"{text}\": {', '.join(props)}"
