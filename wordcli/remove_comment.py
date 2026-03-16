"""Remove comments from docx documents."""

import os
import re
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET

from .constants import (
    COMMENT_TAG, ID_ATTR, _register_namespaces,
)


def _find_and_remove_comment_block(raw_xml, comment_id):
    """Remove <w:comment w:id="N"...>...</w:comment> from raw XML string.

    Returns (updated_xml, success).
    """
    id_str = str(comment_id)
    pattern = re.compile(r'<w:comment\b')
    pos = 0
    while True:
        m = pattern.search(raw_xml, pos)
        if not m:
            return raw_xml, False

        start = m.start()
        # Check if this comment has the right ID
        # Find the end of the opening tag to extract attributes
        tag_end = raw_xml.index('>', start)
        tag_str = raw_xml[start:tag_end + 1]
        id_match = re.search(r'w:id="(\d+)"', tag_str)
        if not id_match or id_match.group(1) != id_str:
            pos = tag_end
            continue

        # Find the matching </w:comment>
        depth = 1
        i = tag_end + 1
        while i < len(raw_xml):
            if raw_xml[i:i + 10] == '<w:comment' and i + 10 < len(raw_xml) and raw_xml[i + 10] in ' >':
                depth += 1
            elif raw_xml[i:i + 12] == '</w:comment>':
                depth -= 1
                if depth == 0:
                    end = i + 12
                    return raw_xml[:start] + raw_xml[end:], True
            i += 1

        return raw_xml, False


def _remove_range_markers(raw_xml, comment_id):
    """Remove commentRangeStart, commentRangeEnd, and commentReference run."""
    id_str = str(comment_id)

    # Remove commentRangeStart (self-closing)
    raw_xml = re.sub(
        r'<w:commentRangeStart\b[^/]*?w:id="' + id_str + r'"[^/]*?/>',
        '', raw_xml)

    # Remove commentRangeEnd (self-closing)
    raw_xml = re.sub(
        r'<w:commentRangeEnd\b[^/]*?w:id="' + id_str + r'"[^/]*?/>',
        '', raw_xml)

    # Remove the w:r containing w:commentReference with this ID
    # Pattern: <w:r>...<w:commentReference w:id="N"/>...</w:r>
    # The run typically contains only rPr + commentReference
    ref_pattern = r'<w:r\b[^>]*>(?:(?!</w:r>).)*?<w:commentReference\b[^/]*?w:id="' + id_str + r'"[^/]*?/>(?:(?!</w:r>).)*?</w:r>'
    raw_xml = re.sub(ref_pattern, '', raw_xml, flags=re.DOTALL)

    return raw_xml


def remove_comment_from_docx(input_path, output_path, comment_id):
    """Remove a comment by numeric ID.

    Returns (success, message).
    """
    with zipfile.ZipFile(input_path, "r") as zf:
        raw_doc = zf.read("word/document.xml")
        try:
            raw_comments = zf.read("word/comments.xml")
        except KeyError:
            return False, "No comments found in document"

    _register_namespaces(raw_doc)
    _register_namespaces(raw_comments)

    # Verify comment exists
    comments_root = ET.fromstring(raw_comments)
    found = False
    for c in comments_root.iter(COMMENT_TAG):
        if c.get(ID_ATTR) is not None and int(c.get(ID_ATTR)) == comment_id:
            found = True
            break
    if not found:
        return False, f"Comment {comment_id} not found"

    # Remove from comments.xml
    comments_str = raw_comments.decode("utf-8")
    comments_str, removed = _find_and_remove_comment_block(comments_str, comment_id)
    if not removed:
        return False, f"Could not locate comment {comment_id} in raw XML"

    # Remove markers from document.xml
    doc_str = raw_doc.decode("utf-8")
    doc_str = _remove_range_markers(doc_str, comment_id)

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
                    zout.writestr(item, doc_str.encode("utf-8"))
                elif item.filename == "word/comments.xml":
                    zout.writestr(item, comments_str.encode("utf-8"))
                else:
                    zout.writestr(item, zin.read(item.filename))

    if use_temp:
        shutil.move(dest, output_path)

    return True, f"Comment {comment_id} removed from {output_path}"
