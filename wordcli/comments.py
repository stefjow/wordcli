"""Add comments to docx documents."""

import copy
import os
import re
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

from .constants import (
    W_NS, BODY_TAG, P_TAG, R_TAG, T_TAG, RPR_TAG, RSTYLE_TAG,
    COMMENT_TAG, COMMENT_RANGE_START_TAG, COMMENT_RANGE_END_TAG,
    COMMENT_REFERENCE_TAG, ANNOTATION_REF_TAG, FOOTNOTE_REF_TAG,
    AUTHOR_ATTR, DATE_ATTR, ID_ATTR, INITIALS_ATTR, VAL_ATTR,
    XML_SPACE_ATTR, PPR_TAG, PSTYLE_TAG,
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


def _find_max_comment_id(comments_root):
    """Find maximum comment id in comments.xml."""
    max_id = -1
    if comments_root is not None:
        for c in comments_root.iter(COMMENT_TAG):
            val = c.get(ID_ATTR)
            if val is not None:
                try:
                    max_id = max(max_id, int(val))
                except ValueError:
                    pass
    return max_id


def _make_initials(author):
    """Generate initials from author name."""
    parts = author.split()
    if not parts:
        return "X"
    return "".join(p[0].upper() for p in parts)


def _add_comment_to_paragraph(p_elem, anchor_text, comment_id, context=None):
    """Insert commentRangeStart/End markers around anchor_text in paragraph.

    Returns True if successful, False otherwise.
    """
    # Collect direct run children with their text
    run_info = []  # (run_elem, text)
    for child in p_elem:
        if child.tag == R_TAG:
            text = get_run_text(child)
            run_info.append((child, text))

    if not run_info:
        return False

    full_text = "".join(ri[1] for ri in run_info)

    # Find match position, optionally scoped by context
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
        return False

    # Build new elements to replace the affected runs
    new_elements = []

    first_run, first_text = run_info[first_ri]

    # Text before the match in the first run
    before_text = first_text[:first_offset]
    if before_text:
        new_elements.append(_clone_run_with_text(first_run, before_text))

    # commentRangeStart
    range_start = ET.Element(COMMENT_RANGE_START_TAG)
    range_start.set(ID_ATTR, str(comment_id))
    new_elements.append(range_start)

    # The matched runs (preserved as-is, possibly split)
    if first_ri == last_ri:
        matched_text = first_text[first_offset:last_end_offset]
        new_elements.append(_clone_run_with_text(first_run, matched_text))
    else:
        # First partial run
        new_elements.append(_clone_run_with_text(first_run, first_text[first_offset:]))
        # Middle runs (full)
        for ri_idx in range(first_ri + 1, last_ri):
            mid_run, mid_text = run_info[ri_idx]
            new_elements.append(_clone_run_with_text(mid_run, mid_text))
        # Last partial run
        last_run, last_text = run_info[last_ri]
        new_elements.append(_clone_run_with_text(last_run, last_text[:last_end_offset]))

    # commentRangeEnd
    range_end = ET.Element(COMMENT_RANGE_END_TAG)
    range_end.set(ID_ATTR, str(comment_id))
    new_elements.append(range_end)

    # commentReference run
    ref_run = ET.Element(R_TAG)
    ref_rpr = ET.SubElement(ref_run, RPR_TAG)
    ref_style = ET.SubElement(ref_rpr, RSTYLE_TAG)
    ref_style.set(VAL_ATTR, "CommentReference")
    ref_ref = ET.SubElement(ref_run, COMMENT_REFERENCE_TAG)
    ref_ref.set(ID_ATTR, str(comment_id))
    new_elements.append(ref_run)

    # Text after the match in the last run
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

    return True


def _build_comment_element(comment_id, author, date_str, text):
    """Build a w:comment XML element."""
    comment = ET.Element(COMMENT_TAG)
    comment.set(ID_ATTR, str(comment_id))
    comment.set(AUTHOR_ATTR, author)
    comment.set(DATE_ATTR, date_str)
    comment.set(INITIALS_ATTR, _make_initials(author))

    p = ET.SubElement(comment, P_TAG)
    ppr = ET.SubElement(p, PPR_TAG)
    pstyle = ET.SubElement(ppr, PSTYLE_TAG)
    pstyle.set(VAL_ATTR, "CommentText")

    # annotationRef run
    ref_run = ET.SubElement(p, R_TAG)
    ref_rpr = ET.SubElement(ref_run, RPR_TAG)
    ref_style = ET.SubElement(ref_rpr, RSTYLE_TAG)
    ref_style.set(VAL_ATTR, "CommentReference")
    ET.SubElement(ref_run, ANNOTATION_REF_TAG)

    # Text run
    text_run = ET.SubElement(p, R_TAG)
    t = ET.SubElement(text_run, T_TAG)
    t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set(XML_SPACE_ATTR, "preserve")

    return comment


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


def _serialize_comment(comment_elem):
    """Serialize a comment element to XML string with w: prefix."""
    wrapper = ET.Element("_wrapper")
    wrapper.set("xmlns:w", W_NS)
    wrapper.append(comment_elem)
    raw = ET.tostring(wrapper, encoding="unicode")
    start = raw.index("<w:comment")
    end = raw.rindex("</w:comment>") + len("</w:comment>")
    return raw[start:end]


def _ensure_comments_xml(zf):
    """Check if word/comments.xml exists in the zip. Return raw bytes or None."""
    try:
        return zf.read("word/comments.xml")
    except KeyError:
        return None


def _ensure_comments_relationship(rels_xml):
    """Ensure comments.xml relationship exists. Returns (updated_xml, changed)."""
    if b"comments.xml" in rels_xml:
        return rels_xml, False

    # Find max rId
    rels_str = rels_xml.decode("utf-8")
    ids = [int(m.group(1)) for m in re.finditer(r'Id="rId(\d+)"', rels_str)]
    next_id = max(ids) + 1 if ids else 1

    rel_tag = (
        f'<Relationship Id="rId{next_id}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" '
        f'Target="comments.xml"/>'
    )
    rels_str = rels_str.replace("</Relationships>", rel_tag + "</Relationships>")
    return rels_str.encode("utf-8"), True


def _ensure_comments_content_type(content_types_xml):
    """Ensure comments content type exists. Returns (updated_xml, changed)."""
    if b"comments.xml" in content_types_xml:
        return content_types_xml, False

    ct_str = content_types_xml.decode("utf-8")
    override = (
        '<Override PartName="/word/comments.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
    )
    ct_str = ct_str.replace("</Types>", override + "</Types>")
    return ct_str.encode("utf-8"), True


def _create_empty_comments_xml():
    """Create a minimal word/comments.xml."""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '</w:comments>'
    ).encode("utf-8")


def _add_comment_to_footnote_ref(body, footnote_id, comment_id):
    """Insert commentRangeStart/End markers around the footnote reference run.

    Finds the run containing <w:footnoteReference w:id="footnote_id"/> and wraps it.
    Returns (success, p_elem_modified, orig_run_texts).
    """
    for p_elem in body.iter(P_TAG):
        children = list(p_elem)
        for idx, child in enumerate(children):
            if child.tag != R_TAG:
                continue
            # Check if this run contains the footnote reference
            fn_ref = None
            for sub in child:
                if sub.tag == FOOTNOTE_REF_TAG:
                    ref_id = sub.get(ID_ATTR)
                    if ref_id is not None and int(ref_id) == footnote_id:
                        fn_ref = sub
                        break
            if fn_ref is None:
                continue

            # Save original run texts for raw XML matching
            orig_run_texts = []
            for c in p_elem:
                if c.tag == R_TAG:
                    for sub in c:
                        if sub.tag == T_TAG and sub.text:
                            orig_run_texts.append(sub.text)

            # Insert comment markers around this run
            range_start = ET.Element(COMMENT_RANGE_START_TAG)
            range_start.set(ID_ATTR, str(comment_id))

            range_end = ET.Element(COMMENT_RANGE_END_TAG)
            range_end.set(ID_ATTR, str(comment_id))

            ref_run = ET.Element(R_TAG)
            ref_rpr = ET.SubElement(ref_run, RPR_TAG)
            ref_style = ET.SubElement(ref_rpr, RSTYLE_TAG)
            ref_style.set(VAL_ATTR, "CommentReference")
            ref_ref = ET.SubElement(ref_run, COMMENT_REFERENCE_TAG)
            ref_ref.set(ID_ATTR, str(comment_id))

            # Insert: rangeStart before the run, rangeEnd + refRun after
            p_elem.insert(idx, range_start)
            # After inserting rangeStart, the run moved to idx+1
            run_pos = idx + 1
            p_elem.insert(run_pos + 1, range_end)
            p_elem.insert(run_pos + 2, ref_run)

            return True, p_elem, orig_run_texts

    return False, None, None


def add_comment_to_docx(input_path, output_path, anchor_text, comment_text,
                         author, paragraph=None, context=None, occurrence=None,
                         footnote=None):
    """Add a comment anchored to anchor_text in the document.

    Returns (success, message).
    """
    date_str = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    with zipfile.ZipFile(input_path, "r") as zf:
        raw_doc = zf.read("word/document.xml")
        raw_comments = _ensure_comments_xml(zf)

    _register_namespaces(raw_doc)
    root = ET.fromstring(raw_doc)

    body = root.find(BODY_TAG)
    if body is None:
        return False, "Could not find document body"

    # Determine comment ID
    doc_max_id = _find_max_id(root)
    if raw_comments is not None:
        _register_namespaces(raw_comments)
        comments_root = ET.fromstring(raw_comments)
        comment_max_id = _find_max_comment_id(comments_root)
    else:
        comments_root = None
        comment_max_id = -1

    comment_id = max(doc_max_id, comment_max_id) + 1

    if footnote:
        # Anchor comment on the footnote reference in the main text
        success, p_elem, target_run_texts = _add_comment_to_footnote_ref(
            body, footnote, comment_id)
        if not success:
            return False, f"Footnote {footnote} reference not found in document"
    else:
        # Normal text-anchored comment
        matches, err = find_matching_paragraphs(body, anchor_text, paragraph, context)
        if err:
            return False, err

        p_elem, _, err = select_match(matches, anchor_text, occurrence)
        if err:
            return False, err

        # Save original run texts for locating in raw XML
        target_run_texts = []
        for child in p_elem:
            if child.tag == R_TAG:
                for sub in child:
                    if sub.tag == T_TAG and sub.text:
                        target_run_texts.append(sub.text)

        success = _add_comment_to_paragraph(p_elem, anchor_text, comment_id, context)
        if not success:
            return False, "Anchor text not found"

    # Splice modified paragraph into raw document XML
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
    output_doc_bytes = output_str.encode("utf-8")

    # Build the comment element and update comments.xml
    comment_elem = _build_comment_element(comment_id, author, date_str, comment_text)
    new_comment_xml = _serialize_comment(comment_elem)

    if raw_comments is not None:
        comments_str = raw_comments.decode("utf-8")
        comments_str = comments_str.replace(
            "</w:comments>", new_comment_xml + "</w:comments>"
        )
        output_comments_bytes = comments_str.encode("utf-8")
    else:
        output_comments_bytes = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            + new_comment_xml +
            '</w:comments>'
        ).encode("utf-8")

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
                    zout.writestr(item, output_doc_bytes)
                elif item.filename == "word/comments.xml":
                    zout.writestr(item, output_comments_bytes)
                elif item.filename == "word/_rels/document.xml.rels":
                    rels_data = zin.read(item.filename)
                    rels_data, _ = _ensure_comments_relationship(rels_data)
                    zout.writestr(item, rels_data)
                elif item.filename == "[Content_Types].xml":
                    ct_data = zin.read(item.filename)
                    ct_data, _ = _ensure_comments_content_type(ct_data)
                    zout.writestr(item, ct_data)
                else:
                    zout.writestr(item, zin.read(item.filename))

            # If comments.xml didn't exist before, add it now
            if raw_comments is None:
                zout.writestr("word/comments.xml", output_comments_bytes)

    if use_temp:
        shutil.move(dest, output_path)

    return True, f"Comment added in {output_path}"
