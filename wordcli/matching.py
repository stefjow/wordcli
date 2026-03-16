"""Shared logic for finding text matches in paragraphs."""

from .constants import P_TAG, R_TAG, T_TAG, BODY_TAG, FOOTNOTE_TAG, ID_ATTR


def get_run_text(run):
    """Get concatenated text from all w:t elements in a run."""
    parts = []
    for sub in run:
        if sub.tag == T_TAG and sub.text:
            parts.append(sub.text)
    return "".join(parts)


def get_paragraph_plain_text(p_elem):
    """Get plain text from direct runs in a paragraph (skipping ins/del)."""
    parts = []
    for child in p_elem:
        if child.tag == R_TAG:
            parts.append(get_run_text(child))
    return "".join(parts)


def find_matching_paragraphs(body, search_text, paragraph=None, context=None):
    """Find all paragraphs containing search_text.

    Returns list of (paragraph_number, p_elem, snippet) for each match.
    """
    all_paragraphs = list(body.iter(P_TAG))

    if paragraph is not None:
        if paragraph < 1 or paragraph > len(all_paragraphs):
            return None, f"Paragraph {paragraph} out of range (1-{len(all_paragraphs)})"
        indexed = [(paragraph, all_paragraphs[paragraph - 1])]
    else:
        indexed = [(i + 1, p) for i, p in enumerate(all_paragraphs)]

    matches = []
    for para_nr, p_elem in indexed:
        p_text = get_paragraph_plain_text(p_elem)
        if search_text not in p_text:
            continue
        if context is not None and context not in p_text:
            continue
        # Build a snippet around the match
        idx = p_text.find(search_text)
        start = max(0, idx - 30)
        end = min(len(p_text), idx + len(search_text) + 30)
        snippet = p_text[start:end]
        if start > 0:
            snippet = "..." + snippet
        if end < len(p_text):
            snippet = snippet + "..."
        matches.append((para_nr, p_elem, snippet))

    return matches, None


def select_match(matches, search_text, occurrence=None):
    """Select a single match from the list, enforcing uniqueness.

    Returns (p_elem, para_nr, error_message).
    If error_message is set, the other values are None.
    """
    if not matches:
        return None, None, "Text not found"

    if occurrence is not None:
        if occurrence < 1 or occurrence > len(matches):
            return None, None, (
                f"Occurrence {occurrence} out of range "
                f"(found {len(matches)} match{'es' if len(matches) > 1 else ''})"
            )
        para_nr, p_elem, _ = matches[occurrence - 1]
        return p_elem, para_nr, None

    if len(matches) == 1:
        para_nr, p_elem, _ = matches[0]
        return p_elem, para_nr, None

    # Multiple matches — build error message
    lines = [f'"{search_text}" found {len(matches)} times:']
    for para_nr, _, snippet in matches:
        lines.append(f"  [{para_nr}] {snippet}")
    lines.append("Use --paragraph, --context, or --occurrence to disambiguate.")
    return None, None, "\n".join(lines)


def find_matching_paragraphs_in_footnote(fn_root, footnote_id, search_text, context=None):
    """Find paragraphs containing search_text within a specific footnote.

    Returns (matches, error_message) where matches is list of (label, p_elem, snippet).
    """
    for fn in fn_root.findall(f".//{FOOTNOTE_TAG}"):
        fn_id = fn.get(ID_ATTR)
        if fn_id is not None and int(fn_id) == footnote_id:
            paragraphs = list(fn.findall(f".//{P_TAG}"))
            matches = []
            for i, p_elem in enumerate(paragraphs):
                p_text = get_paragraph_plain_text(p_elem)
                if search_text not in p_text:
                    continue
                if context is not None and context not in p_text:
                    continue
                idx = p_text.find(search_text)
                start = max(0, idx - 30)
                end = min(len(p_text), idx + len(search_text) + 30)
                snippet = p_text[start:end]
                if start > 0:
                    snippet = "..." + snippet
                if end < len(p_text):
                    snippet = snippet + "..."
                matches.append((f"fn{footnote_id}", p_elem, snippet))
            return matches, None
    return None, f"Footnote {footnote_id} not found"
