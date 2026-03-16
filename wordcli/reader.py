"""DocxReader — reading and parsing docx documents."""

import copy
import zipfile
import xml.etree.ElementTree as ET

from .constants import (
    BODY_TAG, COMMENT_TAG, DEL_TAG, DELTEXT_TAG, INS_TAG, P_TAG, R_TAG,
    T_TAG, TBL_TAG, TR_TAG, TC_TAG, TCPR_TAG, GRIDSPAN_TAG, VMERGE_TAG,
    SECTPR_TAG, FOOTNOTE_TAG, FOOTNOTE_REF_TAG, PPR_TAG, PSTYLE_TAG,
    AUTHOR_ATTR, DATE_ATTR, ID_ATTR, VAL_ATTR, HEADING_RE,
)
from .formatting import table_to_markdown


class DocxReader:
    def __init__(self, path):
        self.path = path
        self.zf = zipfile.ZipFile(path, "r")
        self._cache = {}

    def _parse_xml(self, entry):
        if entry not in self._cache:
            try:
                with self.zf.open(entry) as f:
                    self._cache[entry] = ET.parse(f).getroot()
            except KeyError:
                return None
        return self._cache[entry]

    def _text_from_element(self, elem, accept_changes=False, footnote_markers=False):
        """Extract text from an element, handling tracked changes."""
        parts = []
        for child in elem:
            if child.tag == R_TAG:
                for sub in child:
                    if sub.tag == T_TAG and sub.text:
                        parts.append(sub.text)
                    elif sub.tag == DELTEXT_TAG and sub.text and not accept_changes:
                        parts.append(sub.text)
                    elif sub.tag == FOOTNOTE_REF_TAG and footnote_markers:
                        fn_id = sub.get(ID_ATTR)
                        if fn_id and int(fn_id) >= 2:
                            parts.append(f"[^{fn_id}]")
            elif child.tag == INS_TAG:
                if accept_changes:
                    for run in child:
                        if run.tag == R_TAG:
                            for sub in run:
                                if sub.tag == T_TAG and sub.text:
                                    parts.append(sub.text)
                                elif sub.tag == FOOTNOTE_REF_TAG and footnote_markers:
                                    fn_id = sub.get(ID_ATTR)
                                    if fn_id and int(fn_id) >= 2:
                                        parts.append(f"[^{fn_id}]")
                else:
                    ins_parts = []
                    for run in child:
                        if run.tag == R_TAG:
                            for sub in run:
                                if sub.tag == T_TAG and sub.text:
                                    ins_parts.append(sub.text)
                    if ins_parts:
                        parts.append(f"[+{''.join(ins_parts)}+]")
            elif child.tag == DEL_TAG:
                if not accept_changes:
                    del_parts = []
                    for run in child:
                        if run.tag == R_TAG:
                            for sub in run:
                                if sub.tag == DELTEXT_TAG and sub.text:
                                    del_parts.append(sub.text)
                    if del_parts:
                        parts.append(f"[-{''.join(del_parts)}-]")
            else:
                parts.append(self._text_from_element(child, accept_changes, footnote_markers))
        return "".join(parts)

    def _get_heading_level(self, p_elem):
        """Return heading level (1-9) or 0 if not a heading."""
        ppr = p_elem.find(PPR_TAG)
        if ppr is None:
            return 0
        pstyle = ppr.find(PSTYLE_TAG)
        if pstyle is None:
            return 0
        val = pstyle.get(VAL_ATTR, "")
        m = HEADING_RE.match(val)
        return int(m.group(1)) if m else 0

    def _extract_table(self, tbl_elem, accept_changes=False):
        """Extract a table as list of rows, each row a list of cell texts."""
        rows = []
        for tr in tbl_elem:
            if tr.tag != TR_TAG:
                continue
            row = []
            for tc in tr:
                if tc.tag != TC_TAG:
                    continue
                # Check for vertical merge continuation
                tcpr = tc.find(TCPR_TAG)
                if tcpr is not None:
                    vmerge = tcpr.find(VMERGE_TAG)
                    if vmerge is not None and vmerge.get(VAL_ATTR) is None:
                        row.append("")
                        continue
                # Check gridSpan for horizontal merge
                span = 1
                if tcpr is not None:
                    gs = tcpr.find(GRIDSPAN_TAG)
                    if gs is not None:
                        span = int(gs.get(VAL_ATTR, "1"))
                # Extract cell text (join paragraphs with space)
                cell_parts = []
                for p in tc.findall(f".//{P_TAG}"):
                    cell_parts.append(self._text_from_element(p, accept_changes))
                cell_text = " ".join(cell_parts).strip()
                row.append(cell_text)
                for _ in range(span - 1):
                    row.append("")
            rows.append(row)
        return rows

    def _table_to_markdown(self, rows):
        """Convert table rows to markdown table string."""
        return table_to_markdown(rows)

    def extract_paragraphs(self, accept_changes=False):
        """Return list of (paragraph_number, text)."""
        root = self._parse_xml("word/document.xml")
        if root is None:
            return []
        body = root.find(BODY_TAG)
        if body is None:
            return []
        result = []
        nr = 0
        for p in body.iter(P_TAG):
            nr += 1
            text = self._text_from_element(p, accept_changes)
            result.append((nr, text))
        return result

    def extract_document_structure(self, accept_changes=True):
        """Walk body children in order, yielding ('p', heading_level, text),
        ('table', rows) or ('break',) for each element."""
        root = self._parse_xml("word/document.xml")
        if root is None:
            return
        body = root.find(BODY_TAG)
        if body is None:
            return
        for child in body:
            if child.tag == P_TAG:
                level = self._get_heading_level(child)
                text = self._text_from_element(child, accept_changes, footnote_markers=True)
                yield ("p", level, text)
            elif child.tag == TBL_TAG:
                rows = self._extract_table(child, accept_changes)
                yield ("table", rows)
            elif child.tag == SECTPR_TAG:
                pass  # skip section properties

    def extract_footnotes(self):
        """Return dict of {id: text}."""
        root = self._parse_xml("word/footnotes.xml")
        if root is None:
            return {}
        result = {}
        for fn in root.findall(f".//{FOOTNOTE_TAG}"):
            fn_id = fn.get(ID_ATTR)
            if fn_id is not None and int(fn_id) >= 2:
                parts = []
                for p in fn.findall(f".//{P_TAG}"):
                    parts.append(self._text_from_element(p, accept_changes=True))
                result[int(fn_id)] = " ".join(parts).strip()
        return result

    def extract_footnotes_raw(self, accept_changes=False):
        """Return dict of {id: text} without accepting changes (for verify)."""
        root = self._parse_xml("word/footnotes.xml")
        if root is None:
            return {}
        result = {}
        for fn in root.findall(f".//{FOOTNOTE_TAG}"):
            fn_id = fn.get(ID_ATTR)
            if fn_id is not None and int(fn_id) >= 2:
                parts = []
                for p in fn.findall(f".//{P_TAG}"):
                    text_parts = []
                    for t_elem in p.findall(f".//{T_TAG}"):
                        if t_elem.text:
                            text_parts.append(t_elem.text)
                    parts.append("".join(text_parts))
                text = " ".join(parts).strip()
                if text:
                    result[int(fn_id)] = text
        return result

    def extract_comments(self):
        """Return list of {id, author, date, text}."""
        root = self._parse_xml("word/comments.xml")
        if root is None:
            return []
        result = []
        for c in root.findall(f".//{COMMENT_TAG}"):
            c_id = c.get(ID_ATTR)
            author = c.get(AUTHOR_ATTR, "")
            date = c.get(DATE_ATTR, "")
            parts = []
            for p in c.findall(f".//{P_TAG}"):
                parts.append(self._text_from_element(p, accept_changes=True))
            result.append({
                "id": c_id,
                "author": author,
                "date": date,
                "text": " ".join(parts).strip(),
            })
        return result

    def extract_changes(self):
        """Return list of {type, author, date, text}."""
        root = self._parse_xml("word/document.xml")
        if root is None:
            return []
        result = []
        for elem in root.iter():
            if elem.tag == INS_TAG:
                author = elem.get(AUTHOR_ATTR, "")
                date = elem.get(DATE_ATTR, "")
                parts = []
                for run in elem:
                    if run.tag == R_TAG:
                        for sub in run:
                            if sub.tag == T_TAG and sub.text:
                                parts.append(sub.text)
                if parts:
                    result.append({
                        "type": "INS",
                        "author": author,
                        "date": date,
                        "text": "".join(parts),
                    })
            elif elem.tag == DEL_TAG:
                author = elem.get(AUTHOR_ATTR, "")
                date = elem.get(DATE_ATTR, "")
                parts = []
                for run in elem:
                    if run.tag == R_TAG:
                        for sub in run:
                            if sub.tag == DELTEXT_TAG and sub.text:
                                parts.append(sub.text)
                if parts:
                    result.append({
                        "type": "DEL",
                        "author": author,
                        "date": date,
                        "text": "".join(parts),
                    })
        return result

    def extract_accepted_text(self):
        """Extract full accepted text (deletions removed, insertions kept)."""
        paras = self.extract_paragraphs(accept_changes=True)
        return "\n".join(text for _, text in paras if text)

    def stats(self):
        """Return dict with document statistics."""
        paras = self.extract_paragraphs(accept_changes=True)
        fns = self.extract_footnotes()
        comments = self.extract_comments()
        changes = self.extract_changes()

        # Count changes by author and type
        changes_by_author = {}
        for c in changes:
            key = c["author"] or "(unknown)"
            if key not in changes_by_author:
                changes_by_author[key] = {"INS": 0, "DEL": 0}
            changes_by_author[key][c["type"]] += 1

        # Count comments by author
        comments_by_author = {}
        for c in comments:
            key = c["author"] or "(unknown)"
            comments_by_author[key] = comments_by_author.get(key, 0) + 1

        # Count tables
        root = self._parse_xml("word/document.xml")
        table_count = 0
        if root is not None:
            table_count = len(root.findall(f".//{TBL_TAG}"))

        return {
            "paragraphs": len(paras),
            "paragraphs_nonempty": len([p for _, p in paras if p.strip()]),
            "footnotes": len(fns),
            "tables": table_count,
            "comments_total": len(comments),
            "comments_by_author": comments_by_author,
            "changes_total": len(changes),
            "changes_by_author": changes_by_author,
        }

    def verify_against_original(self, original_path, author=None):
        """Verify no text was lost by removing tracked changes.

        If author is given, only that author's changes are removed.
        If author is None, ALL tracked changes are removed (accept all).
        Checks both main text and footnotes.
        Returns (ok, main_missing, main_extra, fn_missing, fn_extra).
        """
        # --- Main text ---
        root = self._parse_xml("word/document.xml")
        if root is None:
            return False, ["Could not parse edited document"], [], [], []
        edited_root = copy.deepcopy(root)
        self._remove_tracked_changes(edited_root, author)
        edited_text = self._extract_plain_text(edited_root)

        orig = DocxReader(original_path)
        orig_root = orig._parse_xml("word/document.xml")
        if orig_root is None:
            return False, ["Could not parse original document"], [], [], []
        original_text = self._extract_plain_text(orig_root)

        orig_lines = [l for l in original_text.split("\n") if l.strip()]
        edit_lines = [l for l in edited_text.split("\n") if l.strip()]
        orig_set = set(orig_lines)
        edit_set = set(edit_lines)
        main_missing = [l for l in orig_lines if l not in edit_set]
        main_extra = [l for l in edit_lines if l not in orig_set]

        # --- Footnotes ---
        fn_root = self._parse_xml("word/footnotes.xml")
        fn_missing = []
        fn_extra = []
        if fn_root is not None:
            edited_fn_root = copy.deepcopy(fn_root)
            self._remove_tracked_changes(edited_fn_root, author)
            edited_fn_text = self._extract_footnote_texts(edited_fn_root)

            orig_fn_root = orig._parse_xml("word/footnotes.xml")
            if orig_fn_root is not None:
                orig_fn_text = self._extract_footnote_texts(orig_fn_root)
                orig_fn_set = set(orig_fn_text.values())
                edit_fn_set = set(edited_fn_text.values())
                fn_missing = [t for t in orig_fn_text.values() if t not in edit_fn_set]
                fn_extra = [t for t in edited_fn_text.values() if t not in orig_fn_set]

        ok = len(main_missing) == 0 and len(fn_missing) == 0
        return ok, main_missing, main_extra, fn_missing, fn_extra

    def _remove_tracked_changes(self, root, author=None):
        """Remove tracked changes. If author is given, only that author's.
        If author is None, remove ALL tracked changes."""
        # Remove insertions (drop the inserted content)
        for parent in root.iter():
            to_remove = []
            for child in parent:
                if child.tag == INS_TAG:
                    if author is None or child.get(AUTHOR_ATTR) == author:
                        to_remove.append(child)
            for elem in to_remove:
                parent.remove(elem)

        # Unwrap deletions (restore the deleted content)
        for parent in root.iter():
            to_process = []
            for child in parent:
                if child.tag == DEL_TAG:
                    if author is None or child.get(AUTHOR_ATTR) == author:
                        to_process.append((child, list(parent).index(child)))
            for del_elem, del_index in reversed(to_process):
                for elem in del_elem.iter():
                    if elem.tag == DELTEXT_TAG:
                        elem.tag = T_TAG
                for child in reversed(list(del_elem)):
                    parent.insert(del_index, child)
                parent.remove(del_elem)

    def _extract_plain_text(self, root):
        """Extract plain text from XML root."""
        paragraphs = []
        for p_elem in root.findall(f".//{P_TAG}"):
            text_parts = []
            for t_elem in p_elem.findall(f".//{T_TAG}"):
                if t_elem.text:
                    text_parts.append(t_elem.text)
            paragraph_text = "".join(text_parts)
            if paragraph_text:
                paragraphs.append(paragraph_text)
        return "\n".join(paragraphs)

    def _extract_footnote_texts(self, fn_root):
        """Extract footnote texts from a footnotes.xml root, returns {id: text}."""
        result = {}
        for fn in fn_root.findall(f".//{FOOTNOTE_TAG}"):
            fn_id = fn.get(ID_ATTR)
            if fn_id is not None and int(fn_id) >= 2:
                parts = []
                for p in fn.findall(f".//{P_TAG}"):
                    text_parts = []
                    for t_elem in p.findall(f".//{T_TAG}"):
                        if t_elem.text:
                            text_parts.append(t_elem.text)
                    parts.append("".join(text_parts))
                text = " ".join(parts).strip()
                if text:
                    result[int(fn_id)] = text
        return result
