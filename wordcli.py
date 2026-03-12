#!/usr/bin/env python3
"""wordcli — CLI tool for inspecting Word (.docx) documents."""

import argparse
import copy
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

# Fully qualified tag names
P_TAG = f"{{{W_NS}}}p"
R_TAG = f"{{{W_NS}}}r"
T_TAG = f"{{{W_NS}}}t"
DEL_TAG = f"{{{W_NS}}}del"
INS_TAG = f"{{{W_NS}}}ins"
DELTEXT_TAG = f"{{{W_NS}}}delText"
AUTHOR_ATTR = f"{{{W_NS}}}author"
DATE_ATTR = f"{{{W_NS}}}date"
ID_ATTR = f"{{{W_NS}}}id"
FOOTNOTE_TAG = f"{{{W_NS}}}footnote"
FOOTNOTE_REF_TAG = f"{{{W_NS}}}footnoteReference"
COMMENT_TAG = f"{{{W_NS}}}comment"
RPR_TAG = f"{{{W_NS}}}rPr"
PPR_TAG = f"{{{W_NS}}}pPr"
PSTYLE_TAG = f"{{{W_NS}}}pStyle"
VAL_ATTR = f"{{{W_NS}}}val"
TBL_TAG = f"{{{W_NS}}}tbl"
TR_TAG = f"{{{W_NS}}}tr"
TC_TAG = f"{{{W_NS}}}tc"
TCPR_TAG = f"{{{W_NS}}}tcPr"
GRIDSPAN_TAG = f"{{{W_NS}}}gridSpan"
VMERGE_TAG = f"{{{W_NS}}}vMerge"
BODY_TAG = f"{{{W_NS}}}body"
SECTPR_TAG = f"{{{W_NS}}}sectPr"

# Heading style patterns (English and German)
HEADING_RE = re.compile(
    r"^(?:Heading|berschrift|Überschrift)(\d+)$", re.IGNORECASE
)


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
        if not rows:
            return ""
        # Determine column count
        max_cols = max(len(r) for r in rows)
        # Normalize row lengths
        for r in rows:
            while len(r) < max_cols:
                r.append("")
        # Column widths
        widths = []
        for col in range(max_cols):
            w = max((len(rows[r][col]) for r in range(len(rows))), default=3)
            widths.append(max(w, 3))
        lines = []
        for i, row in enumerate(rows):
            cells = [row[c].ljust(widths[c]) for c in range(max_cols)]
            lines.append("| " + " | ".join(cells) + " |")
            if i == 0:
                seps = ["-" * widths[c] for c in range(max_cols)]
                lines.append("| " + " | ".join(seps) + " |")
        return "\n".join(lines)

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


# --- CLI Commands ---

def cmd_text(args):
    doc = DocxReader(args.file)
    paras = doc.extract_paragraphs(accept_changes=args.accept)
    if args.paragraph is not None:
        paras = [(n, t) for n, t in paras if n == args.paragraph]
    elif args.paragraphs is not None:
        start, end = map(int, args.paragraphs.split("-"))
        paras = [(n, t) for n, t in paras if start <= n <= end]
    for nr, text in paras:
        print(f"[{nr}] {text}")


def cmd_search(args):
    doc = DocxReader(args.file)
    query = args.query.lower()
    paras = doc.extract_paragraphs(accept_changes=True)
    for nr, text in paras:
        idx = text.lower().find(query)
        if idx != -1:
            start = max(0, idx - 50)
            end = min(len(text), idx + len(args.query) + 50)
            snippet = text[start:end]
            if start > 0:
                snippet = "..." + snippet
            if end < len(text):
                snippet = snippet + "..."
            print(f"[{nr}] {snippet}")
    if args.footnotes:
        fns = doc.extract_footnotes()
        for fn_id, text in sorted(fns.items()):
            idx = text.lower().find(query)
            if idx != -1:
                start = max(0, idx - 50)
                end = min(len(text), idx + len(args.query) + 50)
                snippet = text[start:end]
                if start > 0:
                    snippet = "..." + snippet
                if end < len(text):
                    snippet = snippet + "..."
                print(f"[Footnote {fn_id}] {snippet}")


def cmd_footnotes(args):
    doc = DocxReader(args.file)
    fns = doc.extract_footnotes()
    if args.id is not None:
        if args.id in fns:
            print(f"[{args.id}] {fns[args.id]}")
        else:
            print(f"Footnote {args.id} not found.", file=sys.stderr)
            sys.exit(1)
    else:
        for fn_id, text in sorted(fns.items()):
            print(f"[{fn_id}] {text}")


def cmd_comments(args):
    doc = DocxReader(args.file)
    comments = doc.extract_comments()
    if args.author:
        comments = [c for c in comments if args.author.lower() in c["author"].lower()]
    if args.json:
        print(json.dumps(comments, ensure_ascii=False, indent=2))
    else:
        for c in comments:
            print(f"[{c['id']}] {c['author']} ({c['date']}): {c['text']}")


def cmd_changes(args):
    doc = DocxReader(args.file)
    changes = doc.extract_changes()
    if args.author:
        changes = [c for c in changes if args.author.lower() in c["author"].lower()]
    for c in changes:
        print(f"[{c['type']}] {c['author']}: \"{c['text']}\"")


def cmd_diff(args):
    doc1 = DocxReader(args.file1)
    doc2 = DocxReader(args.file2)
    text1 = doc1.extract_accepted_text().splitlines()
    text2 = doc2.extract_accepted_text().splitlines()
    max_len = max(len(text1), len(text2))
    for i in range(max_len):
        l1 = text1[i] if i < len(text1) else ""
        l2 = text2[i] if i < len(text2) else ""
        if l1 != l2:
            print(f"--- Paragraph {i + 1} ---")
            if l1:
                print(f"  < {l1}")
            if l2:
                print(f"  > {l2}")


def cmd_verify(args):
    doc = DocxReader(args.file)
    ok, main_missing, main_extra, fn_missing, fn_extra = doc.verify_against_original(args.original, author=args.author)
    if ok:
        print("OK: No text loss detected.")
        sys.exit(0)

    if main_missing:
        print(f"TEXT LOSS IN MAIN BODY: {len(main_missing)} paragraph(s) missing.")
        for line in main_missing:
            print(f"  MISSING: {line[:120]}{'...' if len(line) > 120 else ''}")
    if main_extra:
        print(f"\n{len(main_extra)} unexpected paragraph(s) in main body:")
        for line in main_extra:
            print(f"  EXTRA: {line[:120]}{'...' if len(line) > 120 else ''}")
    if fn_missing:
        print(f"\nTEXT LOSS IN FOOTNOTES: {len(fn_missing)} footnote(s) missing.")
        for line in fn_missing:
            print(f"  MISSING FN: {line[:120]}{'...' if len(line) > 120 else ''}")
    if fn_extra:
        print(f"\n{len(fn_extra)} unexpected footnote(s):")
        for line in fn_extra:
            print(f"  EXTRA FN: {line[:120]}{'...' if len(line) > 120 else ''}")

    sys.exit(1)


def cmd_extract(args):
    doc = DocxReader(args.file)
    lines = []
    for item in doc.extract_document_structure(accept_changes=True):
        if item[0] == "p":
            _, level, text = item
            if level > 0:
                lines.append("")
                lines.append("#" * level + " " + text)
                lines.append("")
            elif text.strip():
                lines.append(text)
            else:
                lines.append("")
        elif item[0] == "table":
            _, rows = item
            if rows:
                lines.append("")
                lines.append(doc._table_to_markdown(rows))
                lines.append("")

    # Append footnotes
    fns = doc.extract_footnotes()
    if fns:
        lines.append("")
        lines.append("---")
        lines.append("")
        for fn_id, text in sorted(fns.items()):
            lines.append(f"[^{fn_id}]: {text}")

    output = "\n".join(lines).strip() + "\n"
    output = re.sub(r"\n{3,}", "\n\n", output)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(output)
        print(f"Written to {args.output}", file=sys.stderr)
    else:
        print(output)


def cmd_tables(args):
    doc = DocxReader(args.file)
    table_nr = 0
    for item in doc.extract_document_structure(accept_changes=True):
        if item[0] == "table":
            _, rows = item
            if rows:
                table_nr += 1
                if args.number is not None and table_nr != args.number:
                    continue
                if table_nr > 1 and args.number is None:
                    print()
                print(f"### Table {table_nr}")
                print()
                print(doc._table_to_markdown(rows))
    if table_nr == 0:
        print("No tables found.", file=sys.stderr)
    elif args.number is not None and args.number > table_nr:
        print(f"Table {args.number} not found (document has {table_nr} table(s)).", file=sys.stderr)
        sys.exit(1)


def cmd_stats(args):
    doc = DocxReader(args.file)
    s = doc.stats()
    if args.json:
        print(json.dumps(s, ensure_ascii=False, indent=2))
        return

    print(f"Paragraphs:       {s['paragraphs']} ({s['paragraphs_nonempty']} non-empty)")
    print(f"Tables:           {s['tables']}")
    print(f"Footnotes:        {s['footnotes']}")
    print(f"Comments:         {s['comments_total']}")
    if s['comments_by_author']:
        for author, count in sorted(s['comments_by_author'].items()):
            print(f"  {author}: {count}")
    print(f"Tracked Changes:  {s['changes_total']}")
    if s['changes_by_author']:
        for author, counts in sorted(s['changes_by_author'].items()):
            print(f"  {author}: {counts['INS']} insertions, {counts['DEL']} deletions")


def main():
    parser = argparse.ArgumentParser(
        prog="wordcli",
        description="CLI tool for inspecting Word (.docx) documents.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # text
    p_text = sub.add_parser("text", help="Extract full text")
    p_text.add_argument("file")
    p_text.add_argument("--paragraph", type=int, default=None)
    p_text.add_argument("--paragraphs", default=None, help="Range, e.g. 3-7")
    p_text.add_argument("--accept", action="store_true", help="Show accepted text only")
    p_text.set_defaults(func=cmd_text)

    # search
    p_search = sub.add_parser("search", help="Search text with context")
    p_search.add_argument("file")
    p_search.add_argument("query")
    p_search.add_argument("--footnotes", action="store_true", help="Also search footnotes")
    p_search.set_defaults(func=cmd_search)

    # footnotes
    p_fn = sub.add_parser("footnotes", help="List footnotes")
    p_fn.add_argument("file")
    p_fn.add_argument("id", nargs="?", type=int, default=None)
    p_fn.set_defaults(func=cmd_footnotes)

    # comments
    p_com = sub.add_parser("comments", help="List comments")
    p_com.add_argument("file")
    p_com.add_argument("--author", default=None)
    p_com.add_argument("--json", action="store_true")
    p_com.set_defaults(func=cmd_comments)

    # changes
    p_chg = sub.add_parser("changes", help="Show tracked changes")
    p_chg.add_argument("file")
    p_chg.add_argument("--author", default=None)
    p_chg.set_defaults(func=cmd_changes)

    # diff
    p_diff = sub.add_parser("diff", help="Compare two documents")
    p_diff.add_argument("file1")
    p_diff.add_argument("file2")
    p_diff.set_defaults(func=cmd_diff)

    # verify
    p_ver = sub.add_parser("verify", help="Check for text loss (main text + footnotes)")
    p_ver.add_argument("file")
    p_ver.add_argument("--original", required=True)
    p_ver.add_argument("--author", default=None, help="Only remove this author's changes (default: all)")
    p_ver.set_defaults(func=cmd_verify)

    # extract
    p_ext = sub.add_parser("extract", help="Extract structured markdown with tables")
    p_ext.add_argument("file")
    p_ext.add_argument("-o", "--output", default=None, help="Output file (default: stdout)")
    p_ext.set_defaults(func=cmd_extract)

    # tables
    p_tbl = sub.add_parser("tables", help="Extract tables as markdown")
    p_tbl.add_argument("file")
    p_tbl.add_argument("number", nargs="?", type=int, default=None, help="Table number (1-based)")
    p_tbl.set_defaults(func=cmd_tables)

    # stats
    p_stats = sub.add_parser("stats", help="Document statistics")
    p_stats.add_argument("file")
    p_stats.add_argument("--json", action="store_true")
    p_stats.set_defaults(func=cmd_stats)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
