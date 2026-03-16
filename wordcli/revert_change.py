"""Revert tracked changes in docx documents."""

import os
import re
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET

from .constants import (
    W_NS, INS_TAG, DEL_TAG, DELTEXT_TAG, T_TAG, R_TAG,
    AUTHOR_ATTR, ID_ATTR,
    _register_namespaces,
)


def _collect_changes(root):
    """Collect all tracked changes with their parent info.

    Returns list of dicts with keys: type, author, text, id, elem, parent.
    """
    parent_map = {child: parent for parent in root.iter() for child in parent}
    changes = []

    for elem in root.iter():
        if elem.tag == INS_TAG:
            parts = []
            for run in elem:
                if run.tag == R_TAG:
                    for sub in run:
                        if sub.tag == T_TAG and sub.text:
                            parts.append(sub.text)
            if parts:
                changes.append({
                    "type": "INS",
                    "author": elem.get(AUTHOR_ATTR, ""),
                    "text": "".join(parts),
                    "id": elem.get(ID_ATTR),
                    "elem": elem,
                    "parent": parent_map.get(elem),
                })
        elif elem.tag == DEL_TAG:
            parts = []
            for run in elem:
                if run.tag == R_TAG:
                    for sub in run:
                        if sub.tag == DELTEXT_TAG and sub.text:
                            parts.append(sub.text)
            if parts:
                changes.append({
                    "type": "DEL",
                    "author": elem.get(AUTHOR_ATTR, ""),
                    "text": "".join(parts),
                    "id": elem.get(ID_ATTR),
                    "elem": elem,
                    "parent": parent_map.get(elem),
                })

    return changes


def _select_change(changes, author=None, text=None, change_type=None, occurrence=None):
    """Filter and select a single change.

    Returns (change_dict, error_message).
    """
    filtered = changes

    if author is not None:
        filtered = [c for c in filtered if author.lower() in c["author"].lower()]
    if text is not None:
        filtered = [c for c in filtered if text in c["text"]]
    if change_type is not None:
        filtered = [c for c in filtered if c["type"] == change_type.upper()]

    if not filtered:
        return None, "No matching tracked change found"

    if occurrence is not None:
        if occurrence < 1 or occurrence > len(filtered):
            return None, (
                f"Occurrence {occurrence} out of range "
                f"(found {len(filtered)} match{'es' if len(filtered) > 1 else ''})"
            )
        return filtered[occurrence - 1], None

    if len(filtered) == 1:
        return filtered[0], None

    # Multiple matches — build error message
    lines = [f"Found {len(filtered)} matching changes:"]
    for i, c in enumerate(filtered, 1):
        snippet = c["text"][:60]
        if len(c["text"]) > 60:
            snippet += "..."
        lines.append(f"  [{i}] [{c['type']}] {c['author']}: \"{snippet}\"")
    lines.append("Use --occurrence to select one.")
    return None, "\n".join(lines)


def _find_change_block(raw_xml, tag_name, change_id):
    """Find a w:ins or w:del block by w:id in raw XML.

    tag_name should be 'ins' or 'del'.
    Returns (start, end) indices or None.
    """
    id_str = str(change_id)
    pattern = re.compile(r'<w:' + tag_name + r'\b')
    pos = 0
    while True:
        m = pattern.search(raw_xml, pos)
        if not m:
            return None

        start = m.start()
        tag_end = raw_xml.index('>', start)
        tag_str = raw_xml[start:tag_end + 1]
        id_match = re.search(r'w:id="([^"]*)"', tag_str)
        if not id_match or id_match.group(1) != id_str:
            pos = tag_end
            continue

        # Find matching close tag
        close_tag = f'</w:{tag_name}>'
        open_tag = f'<w:{tag_name}'
        depth = 1
        i = tag_end + 1
        while i < len(raw_xml):
            if raw_xml[i:].startswith(open_tag) and i + len(open_tag) < len(raw_xml) and raw_xml[i + len(open_tag)] in ' >':
                depth += 1
            elif raw_xml[i:].startswith(close_tag):
                depth -= 1
                if depth == 0:
                    end = i + len(close_tag)
                    return start, end
            i += 1

        return None


def revert_change_in_docx(input_path, output_path, author=None, text=None,
                           occurrence=None, change_type=None, footnote=None):
    """Revert a tracked change.

    Returns (success, message).
    """
    target_file = "word/footnotes.xml" if footnote else "word/document.xml"

    with zipfile.ZipFile(input_path, "r") as zf:
        raw_xml = zf.read(target_file)

    _register_namespaces(raw_xml)
    root = ET.fromstring(raw_xml)
    changes = _collect_changes(root)

    change, err = _select_change(changes, author, text, change_type, occurrence)
    if err:
        return False, err

    # Revert using raw XML splicing
    raw_str = raw_xml.decode("utf-8")
    tag_name = "ins" if change["type"] == "INS" else "del"
    span = _find_change_block(raw_str, tag_name, change["id"])
    if span is None:
        return False, "Could not locate tracked change in raw XML"

    start, end = span
    block = raw_str[start:end]

    if change["type"] == "INS":
        # Remove the entire insertion
        replacement = ""
    else:
        # Unwrap deletion: extract inner runs, convert delText -> t
        # Remove the outer <w:del ...> and </w:del> tags
        inner_start = block.index('>') + 1
        inner_end = block.rindex(f'</w:{tag_name}>')
        inner = block[inner_start:inner_end]
        # Convert w:delText to w:t
        inner = inner.replace('<w:delText', '<w:t')
        inner = inner.replace('</w:delText>', '</w:t>')
        replacement = inner

    output_str = raw_str[:start] + replacement + raw_str[end:]

    # Write new docx
    use_temp = os.path.abspath(input_path) == os.path.abspath(output_path)
    dest = output_path
    if use_temp:
        fd, dest = tempfile.mkstemp(suffix=".docx")
        os.close(fd)

    with zipfile.ZipFile(input_path, "r") as zin:
        with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == target_file:
                    zout.writestr(item, output_str.encode("utf-8"))
                else:
                    zout.writestr(item, zin.read(item.filename))

    if use_temp:
        shutil.move(dest, output_path)

    action = "reverted (insertion removed)" if change["type"] == "INS" else "reverted (deletion restored)"
    return True, f"Change {action} in {output_path}"
