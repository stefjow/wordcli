"""Add bookmarks and cross-references to docx documents."""

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
    DEL_TAG, DELTEXT_TAG, INS_TAG,
    BOOKMARK_START_TAG, BOOKMARK_END_TAG,
    FLDCHAR_TAG, FLDCHARTYPE_ATTR, INSTRTEXT_TAG,
    NAME_ATTR, ID_ATTR, XML_SPACE_ATTR,
    AUTHOR_ATTR, DATE_ATTR,
    _register_namespaces,
)
from .matching import find_matching_paragraphs, select_match, get_run_text


def _clone_run_with_text(run, text):
    """Create a new run preserving rPr but with new text content."""
    new_run = ET.Element(R_TAG)
    rpr = run.find(RPR_TAG)
    if rpr is not None:
        new_run.append(copy.deepcopy(rpr))
    t = ET.Element(T_TAG)
    t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set(XML_SPACE_ATTR, "preserve")
    new_run.append(t)
    return new_run


def _find_max_id(root):
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


def _serialize_paragraph(p_elem):
    """Serialize a paragraph element to an XML string using w: namespace prefix."""
    wrapper = ET.Element("_wrapper")
    wrapper.set("xmlns:w", W_NS)
    wrapper.append(p_elem)
    raw = ET.tostring(wrapper, encoding="unicode")
    start = raw.index("<w:p")
    end = raw.rindex("</w:p>") + len("</w:p>")
    return raw[start:end]


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
        if not run_texts:
            continue
        search_pos = 0
        found_all = True
        for txt in run_texts:
            escaped = txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            idx = block.find(escaped, search_pos)
            if idx == -1:
                found_all = False
                break
            search_pos = idx + len(escaped)

        if found_all:
            return p_start, p_end

    return None


def _add_bookmark_to_paragraph(p_elem, anchor_text, bookmark_id, bookmark_name,
                                context=None):
    """Insert bookmarkStart/End markers around anchor_text in paragraph.

    Non-destructive: preserves all existing elements (including field codes).
    Only splits runs when the match boundary falls mid-run.
    Returns True if successful.
    """
    # Collect ALL direct children, tracking which are runs with text
    children = list(p_elem)
    run_info = []  # (child_index, run_elem, text)
    for ci, child in enumerate(children):
        if child.tag == R_TAG:
            text = get_run_text(child)
            run_info.append((ci, child, text))

    if not run_info:
        return False

    full_text = "".join(ri[2] for ri in run_info)

    if context is not None:
        ctx_pos = full_text.find(context)
        if ctx_pos == -1:
            return False
        pos = full_text.find(anchor_text, ctx_pos, ctx_pos + len(context))
    else:
        pos = full_text.find(anchor_text)

    if pos == -1:
        return False

    match_end = pos + len(anchor_text)

    # Map match boundaries to run indices
    char_offset = 0
    first_ri = last_ri = None
    first_offset = last_end_offset = 0

    for ri_idx, (_, _, text) in enumerate(run_info):
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
        return False

    first_child_idx, first_run, first_text = run_info[first_ri]
    last_child_idx, last_run, last_text = run_info[last_ri]

    # bookmarkEnd — insert AFTER the last run (or after split)
    # Handle split at end of match if needed
    if last_end_offset < len(last_text):
        # Split last run: keep matched part, create new run for after
        after_text = last_text[last_end_offset:]
        after_run = _clone_run_with_text(last_run, after_text)
        # Truncate the original run's text
        for t_elem in last_run.iter(T_TAG):
            if t_elem.text:
                # Set to just the matched portion
                orig = t_elem.text
                if last_end_offset <= len(orig):
                    t_elem.text = orig[:last_end_offset]
                    if t_elem.text and (t_elem.text[0] == " " or t_elem.text[-1] == " "):
                        t_elem.set(XML_SPACE_ATTR, "preserve")
                break
        # Insert after_run after last_run
        p_elem.insert(last_child_idx + 1, after_run)
        # Refresh children list
        children = list(p_elem)
        last_child_idx = children.index(last_run)

    bm_end = ET.Element(BOOKMARK_END_TAG)
    bm_end.set(ID_ATTR, str(bookmark_id))
    # Insert bookmarkEnd right after the last matched run
    end_pos = list(p_elem).index(last_run) + 1
    p_elem.insert(end_pos, bm_end)

    # bookmarkStart — insert BEFORE the first run (or after split)
    # Handle split at start of match if needed
    if first_offset > 0:
        before_text = first_text[:first_offset]
        before_run = _clone_run_with_text(first_run, before_text)
        # Truncate the original run's text to start at match
        for t_elem in first_run.iter(T_TAG):
            if t_elem.text:
                orig = t_elem.text
                t_elem.text = orig[first_offset:]
                if t_elem.text and (t_elem.text[0] == " " or t_elem.text[-1] == " "):
                    t_elem.set(XML_SPACE_ATTR, "preserve")
                break
        # Insert before_run before first_run
        first_pos = list(p_elem).index(first_run)
        p_elem.insert(first_pos, before_run)

    bm_start = ET.Element(BOOKMARK_START_TAG)
    bm_start.set(ID_ATTR, str(bookmark_id))
    bm_start.set(NAME_ATTR, f"_Ref_{bookmark_name}")
    # Insert bookmarkStart right before the first matched run
    start_pos = list(p_elem).index(first_run)
    p_elem.insert(start_pos, bm_start)

    return True


def _clone_run_with_deltext(run, text):
    """Create a new run preserving rPr but with w:delText content."""
    new_run = ET.Element(R_TAG)
    rpr = run.find(RPR_TAG)
    if rpr is not None:
        new_run.append(copy.deepcopy(rpr))
    dt = ET.Element(DELTEXT_TAG)
    dt.text = text
    if text and (text[0] == " " or text[-1] == " "):
        dt.set(XML_SPACE_ATTR, "preserve")
    new_run.append(dt)
    return new_run


def _replace_text_with_ref_field(p_elem, ref_text, bookmark_name, author,
                                  date_str, rev_id, context=None,
                                  display_text=None):
    """Replace ref_text with a clickable REF field as tracked change.

    Returns (success, next_rev_id).
    """
    run_info = []
    for child in p_elem:
        if child.tag == R_TAG:
            text = get_run_text(child)
            run_info.append((child, text))

    if not run_info:
        return False, rev_id

    full_text = "".join(ri[1] for ri in run_info)

    if context is not None:
        ctx_pos = full_text.find(context)
        if ctx_pos == -1:
            return False, rev_id
        pos = full_text.find(ref_text, ctx_pos, ctx_pos + len(context))
    else:
        pos = full_text.find(ref_text)

    if pos == -1:
        return False, rev_id

    match_end = pos + len(ref_text)

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

    first_run = run_info[first_ri][0]
    first_text = run_info[first_ri][1]
    rpr = first_run.find(RPR_TAG)

    new_elements = []

    # Before text
    before_text = first_text[:first_offset]
    if before_text:
        new_elements.append(_clone_run_with_text(first_run, before_text))

    # w:del — wrap old text as deletion
    del_elem = ET.Element(DEL_TAG)
    del_elem.set(ID_ATTR, str(rev_id))
    del_elem.set(AUTHOR_ATTR, author)
    del_elem.set(DATE_ATTR, date_str)
    rev_id += 1

    if first_ri == last_ri:
        del_text = first_text[first_offset:last_end_offset]
        del_elem.append(_clone_run_with_deltext(first_run, del_text))
    else:
        del_elem.append(_clone_run_with_deltext(first_run, first_text[first_offset:]))
        for ri_idx in range(first_ri + 1, last_ri):
            mid_run, mid_text = run_info[ri_idx]
            del_elem.append(_clone_run_with_deltext(mid_run, mid_text))
        last_run, last_text_val = run_info[last_ri]
        del_elem.append(_clone_run_with_deltext(last_run, last_text_val[:last_end_offset]))

    new_elements.append(del_elem)

    # w:ins — wrap REF field as insertion
    ins_elem = ET.Element(INS_TAG)
    ins_elem.set(ID_ATTR, str(rev_id))
    ins_elem.set(AUTHOR_ATTR, author)
    ins_elem.set(DATE_ATTR, date_str)
    rev_id += 1

    # fldChar begin
    r_begin = ET.Element(R_TAG)
    if rpr is not None:
        r_begin.append(copy.deepcopy(rpr))
    fc_begin = ET.SubElement(r_begin, FLDCHAR_TAG)
    fc_begin.set(FLDCHARTYPE_ATTR, "begin")
    ins_elem.append(r_begin)

    # instrText
    r_instr = ET.Element(R_TAG)
    if rpr is not None:
        r_instr.append(copy.deepcopy(rpr))
    instr = ET.SubElement(r_instr, INSTRTEXT_TAG)
    instr.set(XML_SPACE_ATTR, "preserve")
    instr.text = f" REF _Ref_{bookmark_name} \\h "
    ins_elem.append(r_instr)

    # fldChar separate
    r_sep = ET.Element(R_TAG)
    if rpr is not None:
        r_sep.append(copy.deepcopy(rpr))
    fc_sep = ET.SubElement(r_sep, FLDCHAR_TAG)
    fc_sep.set(FLDCHARTYPE_ATTR, "separate")
    ins_elem.append(r_sep)

    # Display text run
    r_display = _clone_run_with_text(first_run, display_text or ref_text)
    ins_elem.append(r_display)

    # fldChar end
    r_end = ET.Element(R_TAG)
    if rpr is not None:
        r_end.append(copy.deepcopy(rpr))
    fc_end = ET.SubElement(r_end, FLDCHAR_TAG)
    fc_end.set(FLDCHARTYPE_ATTR, "end")
    ins_elem.append(r_end)

    new_elements.append(ins_elem)

    # After text
    if first_ri == last_ri:
        after_text = first_text[last_end_offset:]
    else:
        _, last_text_val = run_info[last_ri]
        after_text = last_text_val[last_end_offset:]

    if after_text:
        after_run = run_info[last_ri][0]
        new_elements.append(_clone_run_with_text(after_run, after_text))

    # Replace runs
    children = list(p_elem)
    insert_pos = children.index(run_info[first_ri][0])

    for ri_idx in range(first_ri, last_ri + 1):
        p_elem.remove(run_info[ri_idx][0])

    for i, elem in enumerate(new_elements):
        p_elem.insert(insert_pos + i, elem)

    return True, rev_id


def add_bookmark_to_docx(input_path, output_path, bookmark_name, anchor_text,
                          paragraph=None, context=None, occurrence=None):
    """Add a bookmark around anchor_text in the document.

    Returns (success, message).
    """
    if not re.match(r'^[A-Za-z0-9_]+$', bookmark_name):
        return False, "Bookmark name must contain only letters, digits, and underscores"

    with zipfile.ZipFile(input_path, "r") as zf:
        raw_doc = zf.read("word/document.xml")

    _register_namespaces(raw_doc)
    root = ET.fromstring(raw_doc)

    body = root.find(BODY_TAG)
    if body is None:
        return False, "Could not find document body"

    matches, err = find_matching_paragraphs(body, anchor_text, paragraph, context)
    if err:
        return False, err

    p_elem, _, err = select_match(matches, anchor_text, occurrence)
    if err:
        return False, err

    # Save original run texts for raw XML matching
    target_run_texts = []
    for child in p_elem:
        if child.tag == R_TAG:
            for sub in child:
                if sub.tag == T_TAG and sub.text:
                    target_run_texts.append(sub.text)

    bookmark_id = _find_max_id(root) + 1
    success = _add_bookmark_to_paragraph(
        p_elem, anchor_text, bookmark_id, bookmark_name, context)
    if not success:
        return False, "Anchor text not found"

    # Splice into raw XML
    raw_str = raw_doc.decode("utf-8")

    dummy_p = ET.Element(P_TAG)
    for txt in target_run_texts:
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

    return True, f"Bookmark '{bookmark_name}' added in {output_path}"


def add_crossref_to_docx(input_path, output_path, bookmark_name, ref_text,
                          paragraph=None, context=None, occurrence=None,
                          display_text=None, author="wordcli"):
    """Replace ref_text with a clickable REF field pointing to a bookmark.

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

    matches, err = find_matching_paragraphs(body, ref_text, paragraph, context)
    if err:
        return False, err

    p_elem, _, err = select_match(matches, ref_text, occurrence)
    if err:
        return False, err

    # Save original run texts for raw XML matching
    target_run_texts = []
    for child in p_elem:
        if child.tag == R_TAG:
            for sub in child:
                if sub.tag == T_TAG and sub.text:
                    target_run_texts.append(sub.text)

    rev_id = _find_max_id(root) + 1
    success, rev_id = _replace_text_with_ref_field(
        p_elem, ref_text, bookmark_name, author, date_str, rev_id,
        context, display_text)
    if not success:
        return False, "Reference text not found"

    # Splice into raw XML
    raw_str = raw_doc.decode("utf-8")

    dummy_p = ET.Element(P_TAG)
    for txt in target_run_texts:
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

    return True, f"Cross-reference to '{bookmark_name}' added in {output_path}"
