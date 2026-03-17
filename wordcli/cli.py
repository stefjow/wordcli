"""CLI commands and argument parser."""

import argparse
import json
import re
import sys

from .reader import DocxReader
from .replace import replace_in_docx
from .comments import add_comment_to_docx
from .remove_comment import remove_comment_from_docx
from .revert_change import revert_change_in_docx
from .crossref import add_bookmark_to_docx, add_crossref_to_docx
from .style import change_style_in_docx
from .formatting import show_nbsp, parse_nbsp, table_to_markdown


def cmd_text(args):
    doc = DocxReader(args.file)
    paras = doc.extract_paragraphs(accept_changes=args.accept,
                                   include_styles=args.styles)
    if args.styles:
        if args.paragraph is not None:
            paras = [(n, t, s) for n, t, s in paras if n == args.paragraph]
        elif args.paragraphs is not None:
            start, end = map(int, args.paragraphs.split("-"))
            paras = [(n, t, s) for n, t, s in paras if start <= n <= end]
        for nr, text, style in paras:
            label = f"{nr}:{style}" if style else str(nr)
            print(f"[{label}] {show_nbsp(text)}")
    else:
        if args.paragraph is not None:
            paras = [(n, t) for n, t in paras if n == args.paragraph]
        elif args.paragraphs is not None:
            start, end = map(int, args.paragraphs.split("-"))
            paras = [(n, t) for n, t in paras if start <= n <= end]
        for nr, text in paras:
            print(f"[{nr}] {show_nbsp(text)}")


def cmd_search(args):
    doc = DocxReader(args.file)
    query = args.query.lower()
    ctx = args.context_size
    paras = doc.extract_paragraphs(accept_changes=True)
    for nr, text in paras:
        idx = text.lower().find(query)
        if idx != -1:
            start = max(0, idx - ctx)
            end = min(len(text), idx + len(args.query) + ctx)
            snippet = text[start:end]
            if start > 0:
                snippet = "..." + snippet
            if end < len(text):
                snippet = snippet + "..."
            print(f"[{nr}] {show_nbsp(snippet)}")
    if args.footnotes:
        fns = doc.extract_footnotes()
        for fn_id, text in sorted(fns.items()):
            idx = text.lower().find(query)
            if idx != -1:
                start = max(0, idx - ctx)
                end = min(len(text), idx + len(args.query) + ctx)
                snippet = text[start:end]
                if start > 0:
                    snippet = "..." + snippet
                if end < len(text):
                    snippet = snippet + "..."
                print(f"[Footnote {fn_id}] {show_nbsp(snippet)}")


def cmd_footnotes(args):
    doc = DocxReader(args.file)
    fns = doc.extract_footnotes()
    if args.id is not None:
        if args.id in fns:
            print(f"[{args.id}] {show_nbsp(fns[args.id])}")
        else:
            print(f"Footnote {args.id} not found.", file=sys.stderr)
            sys.exit(1)
    else:
        for fn_id, text in sorted(fns.items()):
            print(f"[{fn_id}] {show_nbsp(text)}")


def cmd_comments(args):
    doc = DocxReader(args.file)
    comments = doc.extract_comments()
    if args.author:
        comments = [c for c in comments if args.author.lower() in c["author"].lower()]
    if args.json:
        print(json.dumps(comments, ensure_ascii=False, indent=2))
    else:
        for c in comments:
            print(f"[{c['id']}] {c['author']} ({c['date']}): {show_nbsp(c['text'])}")


def cmd_changes(args):
    doc = DocxReader(args.file)
    changes = doc.extract_changes()
    if args.author:
        changes = [c for c in changes if args.author.lower() in c["author"].lower()]
    for c in changes:
        print(f"[{c['type']}] {c['author']}: \"{show_nbsp(c['text'])}\"")


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
    trunc = args.truncate
    ok, main_missing, main_extra, fn_missing, fn_extra = doc.verify_against_original(args.original, author=args.author)
    if ok:
        print("OK: No text loss detected.")
        sys.exit(0)

    if main_missing:
        print(f"TEXT LOSS IN MAIN BODY: {len(main_missing)} paragraph(s) missing.")
        for line in main_missing:
            print(f"  MISSING: {line[:trunc]}{'...' if len(line) > trunc else ''}")
    if main_extra:
        print(f"\n{len(main_extra)} unexpected paragraph(s) in main body:")
        for line in main_extra:
            print(f"  EXTRA: {line[:trunc]}{'...' if len(line) > trunc else ''}")
    if fn_missing:
        print(f"\nTEXT LOSS IN FOOTNOTES: {len(fn_missing)} footnote(s) missing.")
        for line in fn_missing:
            print(f"  MISSING FN: {line[:trunc]}{'...' if len(line) > trunc else ''}")
    if fn_extra:
        print(f"\n{len(fn_extra)} unexpected footnote(s):")
        for line in fn_extra:
            print(f"  EXTRA FN: {line[:trunc]}{'...' if len(line) > trunc else ''}")

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
                lines.append(table_to_markdown(rows))
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
    output = show_nbsp(output)

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
                print(table_to_markdown(rows))
    if table_nr == 0:
        print("No tables found.", file=sys.stderr)
    elif args.number is not None and args.number > table_nr:
        print(f"Table {args.number} not found (document has {table_nr} table(s)).", file=sys.stderr)
        sys.exit(1)


def cmd_replace(args):
    output = args.output or args.file
    old = parse_nbsp(args.old)
    new = parse_nbsp(args.new or "")
    context = parse_nbsp(args.context) if args.context else None
    ok, msg = replace_in_docx(
        args.file, output, old, new,
        author=args.author,
        paragraph=args.paragraph,
        context=context,
        occurrence=args.occurrence,
        footnote=args.footnote,
    )
    if ok:
        print(msg)
    else:
        print(f"Error: {msg}", file=sys.stderr)
        sys.exit(1)


def cmd_comment(args):
    output = args.output or args.file
    if not args.anchor and not args.footnote:
        print("Error: --anchor or --footnote is required", file=sys.stderr)
        sys.exit(1)
    anchor = parse_nbsp(args.anchor) if args.anchor else None
    context = parse_nbsp(args.context) if args.context else None
    ok, msg = add_comment_to_docx(
        args.file, output, anchor, args.text,
        author=args.author,
        paragraph=args.paragraph,
        context=context,
        occurrence=args.occurrence,
        footnote=args.footnote,
    )
    if ok:
        print(msg)
    else:
        print(f"Error: {msg}", file=sys.stderr)
        sys.exit(1)


def cmd_remove_comment(args):
    output = args.output or args.file
    ok, msg = remove_comment_from_docx(args.file, output, args.id)
    if ok:
        print(msg)
    else:
        print(f"Error: {msg}", file=sys.stderr)
        sys.exit(1)


def cmd_revert_change(args):
    output = args.output or args.file
    ok, msg = revert_change_in_docx(
        args.file, output,
        author=args.author,
        text=parse_nbsp(args.text) if args.text else None,
        occurrence=args.occurrence,
        change_type=args.type,
        footnote=args.footnote,
    )
    if ok:
        print(msg)
    else:
        print(f"Error: {msg}", file=sys.stderr)
        sys.exit(1)


def cmd_bookmark(args):
    output = args.output or args.file
    anchor = parse_nbsp(args.anchor)
    context = parse_nbsp(args.context) if args.context else None
    ok, msg = add_bookmark_to_docx(
        args.file, output, args.name, anchor,
        paragraph=args.paragraph,
        context=context,
        occurrence=args.occurrence,
    )
    if ok:
        print(msg)
    else:
        print(f"Error: {msg}", file=sys.stderr)
        sys.exit(1)


def cmd_crossref(args):
    output = args.output or args.file
    text = parse_nbsp(args.text)
    context = parse_nbsp(args.context) if args.context else None
    display = parse_nbsp(args.display) if args.display else None
    ok, msg = add_crossref_to_docx(
        args.file, output, args.bookmark, text,
        paragraph=args.paragraph,
        context=context,
        occurrence=args.occurrence,
        display_text=display,
        author=args.author,
    )
    if ok:
        print(msg)
    else:
        print(f"Error: {msg}", file=sys.stderr)
        sys.exit(1)


def cmd_fields(args):
    doc = DocxReader(args.file)
    fields = doc.extract_fields()
    if args.seq:
        fields = [f for f in fields if f["field_code"].startswith("SEQ")]
    if not fields:
        print("No fields found.", file=sys.stderr)
        return
    for f in fields:
        ctx = show_nbsp(f["context"])
        if len(ctx) > 80:
            ctx = ctx[:80] + "..."
        print(f"[{f['paragraph']}] {f['field_code']} = \"{f['display']}\"  ->  {ctx}")


def cmd_style(args):
    doc = DocxReader(args.file)
    if args.list:
        styles = doc.extract_styles()
        if args.type:
            styles = [s for s in styles if s["type"] == args.type]
        else:
            styles = [s for s in styles if s["type"] == "paragraph"]
        for s in styles:
            print(f"{s['id']:30s}  {s['name']}")
        return
    if args.paragraph is None:
        print("Error: --paragraph is required (or use --list)", file=sys.stderr)
        sys.exit(1)
    if args.set is None:
        # Query mode: show current style for the paragraph
        paras = doc.extract_paragraphs(include_styles=True)
        for nr, text, style in paras:
            if nr == args.paragraph:
                label = style or "(default)"
                print(f"[{nr}] style={label}")
                return
        print(f"Paragraph {args.paragraph} not found.", file=sys.stderr)
        sys.exit(1)
    # Set mode
    output = args.output or args.file
    ok, msg = change_style_in_docx(
        args.file, output, args.paragraph, args.set, author=args.author)
    if ok:
        print(msg)
    else:
        print(f"Error: {msg}", file=sys.stderr)
        sys.exit(1)


def cmd_xml(args):
    import zipfile
    import xml.etree.ElementTree as ET
    from xml.dom.minidom import parseString
    from .constants import W_NS, P_TAG, BODY_TAG, FOOTNOTE_TAG, _register_namespaces

    PART_MAP = {
        "document": "word/document.xml",
        "footnotes": "word/footnotes.xml",
        "comments": "word/comments.xml",
        "styles": "word/styles.xml",
        "numbering": "word/numbering.xml",
        "settings": "word/settings.xml",
        "rels": "word/_rels/document.xml.rels",
    }

    part = args.part
    if part in PART_MAP:
        zip_path = PART_MAP[part]
    else:
        zip_path = part  # allow raw zip paths like word/header1.xml

    try:
        with zipfile.ZipFile(args.file, "r") as zf:
            if args.list:
                for name in sorted(zf.namelist()):
                    print(name)
                return
            raw = zf.read(zip_path)
    except KeyError:
        print(f"Part not found: {zip_path}", file=sys.stderr)
        sys.exit(1)

    # Filter to paragraph range if requested
    if args.paragraph is not None or args.paragraphs is not None:
        _register_namespaces(raw)
        root = ET.fromstring(raw)
        # Find paragraphs in body or footnotes
        if part == "footnotes":
            containers = list(root.iter(FOOTNOTE_TAG))
            paragraphs = []
            for fn in containers:
                paragraphs.extend(fn.findall(P_TAG))
        else:
            body = root.find(BODY_TAG)
            paragraphs = list(body.findall(P_TAG)) if body is not None else []

        if args.paragraph is not None:
            start = end = args.paragraph
        else:
            start, end = map(int, args.paragraphs.split("-"))

        selected = paragraphs[start - 1:end]  # 1-based to 0-based
        if not selected:
            print(f"No paragraphs in range.", file=sys.stderr)
            sys.exit(1)

        for p in selected:
            p_xml = ET.tostring(p, encoding="unicode")
            pretty = parseString(p_xml).toprettyxml(indent="  ")
            # Remove xml declaration line
            lines = pretty.split("\n")
            print("\n".join(lines[1:]).strip())
            print()
        return

    # Pretty-print the whole part
    try:
        pretty = parseString(raw).toprettyxml(indent="  ")
        lines = pretty.split("\n")
        print("\n".join(lines[1:]).strip())
    except Exception:
        # If XML parsing fails, dump raw
        print(raw.decode("utf-8", errors="replace"))


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
    from . import __version__

    parser = argparse.ArgumentParser(
        prog="wordcli",
        description="CLI tool for inspecting Word (.docx) documents.",
    )
    parser.add_argument("--version", action="version", version=f"wordcli {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    # text
    p_text = sub.add_parser("text", help="Extract full text")
    p_text.add_argument("file")
    p_text.add_argument("--paragraph", type=int, default=None)
    p_text.add_argument("--paragraphs", default=None, help="Range, e.g. 3-7")
    p_text.add_argument("--accept", action="store_true", help="Show accepted text only")
    p_text.add_argument("--styles", action="store_true", help="Show paragraph style IDs")
    p_text.set_defaults(func=cmd_text)

    # search
    p_search = sub.add_parser("search", help="Search text with context")
    p_search.add_argument("file")
    p_search.add_argument("query")
    p_search.add_argument("--footnotes", action="store_true", help="Also search footnotes")
    p_search.add_argument("--context-size", type=int, default=50, help="Characters of context around match (default: 50)")
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
    p_ver.add_argument("--truncate", type=int, default=120, help="Truncate preview lines to N characters (default: 120)")
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

    # replace
    p_rep = sub.add_parser("replace", help="Replace text as tracked change")
    p_rep.add_argument("file")
    p_rep.add_argument("--old", required=True, help="Text to replace")
    p_rep.add_argument("--new", default=None, help="Replacement text (empty = delete)")
    p_rep.add_argument("--author", default="wordcli", help="Author name for the change")
    p_rep.add_argument("--paragraph", type=int, default=None, help="Limit to paragraph number")
    p_rep.add_argument("--context", default=None, help="Unique surrounding text to locate the match")
    p_rep.add_argument("--occurrence", type=int, default=None, help="Match the Nth occurrence (1-based)")
    p_rep.add_argument("--footnote", type=int, default=None, help="Replace within footnote N")
    p_rep.add_argument("-o", "--output", default=None, help="Output file (default: overwrite input)")
    p_rep.set_defaults(func=cmd_replace)

    # comment
    p_comment = sub.add_parser("comment", help="Add a comment anchored to text")
    p_comment.add_argument("file")
    p_comment.add_argument("--anchor", default=None, help="Text to anchor the comment to")
    p_comment.add_argument("--text", required=True, help="Comment text")
    p_comment.add_argument("--author", default="wordcli", help="Author name for the comment")
    p_comment.add_argument("--paragraph", type=int, default=None, help="Limit to paragraph number")
    p_comment.add_argument("--context", default=None, help="Unique surrounding text to locate the anchor")
    p_comment.add_argument("--occurrence", type=int, default=None, help="Match the Nth occurrence (1-based)")
    p_comment.add_argument("--footnote", type=int, default=None, help="Comment on footnote N (anchors to footnote reference in main text)")
    p_comment.add_argument("-o", "--output", default=None, help="Output file (default: overwrite input)")
    p_comment.set_defaults(func=cmd_comment)

    # remove-comment
    p_rmcom = sub.add_parser("remove-comment", help="Remove a comment by ID")
    p_rmcom.add_argument("file")
    p_rmcom.add_argument("id", type=int, help="Comment ID to remove")
    p_rmcom.add_argument("-o", "--output", default=None, help="Output file (default: overwrite input)")
    p_rmcom.set_defaults(func=cmd_remove_comment)

    # revert-change
    p_revert = sub.add_parser("revert-change", help="Revert a tracked change")
    p_revert.add_argument("file")
    p_revert.add_argument("--author", default=None, help="Filter by author")
    p_revert.add_argument("--text", default=None, help="Filter by content text")
    p_revert.add_argument("--type", choices=["ins", "del"], default=None, help="Filter by change type (ins or del)")
    p_revert.add_argument("--occurrence", type=int, default=None, help="Pick the Nth matching change (1-based)")
    p_revert.add_argument("--footnote", type=int, default=None, help="Revert change in footnote N")
    p_revert.add_argument("-o", "--output", default=None, help="Output file (default: overwrite input)")
    p_revert.set_defaults(func=cmd_revert_change)

    # bookmark
    p_bm = sub.add_parser("bookmark", help="Add a bookmark around text")
    p_bm.add_argument("file")
    p_bm.add_argument("--name", required=True, help="Bookmark name (letters, digits, underscores)")
    p_bm.add_argument("--anchor", required=True, help="Text to wrap with the bookmark")
    p_bm.add_argument("--paragraph", type=int, default=None, help="Limit to paragraph number")
    p_bm.add_argument("--context", default=None, help="Unique surrounding text")
    p_bm.add_argument("--occurrence", type=int, default=None, help="Match the Nth occurrence (1-based)")
    p_bm.add_argument("-o", "--output", default=None, help="Output file (default: overwrite input)")
    p_bm.set_defaults(func=cmd_bookmark)

    # crossref
    p_xref = sub.add_parser("crossref", help="Replace text with a clickable cross-reference")
    p_xref.add_argument("file")
    p_xref.add_argument("--bookmark", required=True, help="Bookmark name to reference")
    p_xref.add_argument("--text", required=True, help="Text to find and replace with the REF field")
    p_xref.add_argument("--display", default=None, help="Display text for the field (default: same as --text)")
    p_xref.add_argument("--author", default="wordcli", help="Author name for the tracked change")
    p_xref.add_argument("--paragraph", type=int, default=None, help="Limit to paragraph number")
    p_xref.add_argument("--context", default=None, help="Unique surrounding text")
    p_xref.add_argument("--occurrence", type=int, default=None, help="Match the Nth occurrence (1-based)")
    p_xref.add_argument("-o", "--output", default=None, help="Output file (default: overwrite input)")
    p_xref.set_defaults(func=cmd_crossref)

    # fields
    p_fields = sub.add_parser("fields", help="List document fields (SEQ, REF, etc.)")
    p_fields.add_argument("file")
    p_fields.add_argument("--seq", action="store_true", help="Only show SEQ fields")
    p_fields.set_defaults(func=cmd_fields)

    # style
    p_style = sub.add_parser("style", help="Show or change paragraph style")
    p_style.add_argument("file")
    p_style.add_argument("--list", action="store_true", help="List available styles")
    p_style.add_argument("--type", default=None, help="Filter --list by type (paragraph, character, table)")
    p_style.add_argument("--paragraph", type=int, default=None, help="Paragraph number")
    p_style.add_argument("--set", default=None, help="Style ID to apply")
    p_style.add_argument("--author", default="wordcli", help="Author name for the change")
    p_style.add_argument("-o", "--output", default=None, help="Output file (default: overwrite input)")
    p_style.set_defaults(func=cmd_style)

    # xml
    p_xml = sub.add_parser("xml", help="Show raw XML of a document part")
    p_xml.add_argument("file")
    p_xml.add_argument("part", nargs="?", default="document",
                       help="Part name (document, footnotes, comments, styles, numbering, settings, rels) or zip path (default: document)")
    p_xml.add_argument("--paragraph", type=int, default=None, help="Show XML for a single paragraph (1-based)")
    p_xml.add_argument("--paragraphs", default=None, help="Range, e.g. 3-7")
    p_xml.add_argument("--list", action="store_true", help="List all parts in the docx archive")
    p_xml.set_defaults(func=cmd_xml)

    # stats
    p_stats = sub.add_parser("stats", help="Document statistics")
    p_stats.add_argument("file")
    p_stats.add_argument("--json", action="store_true")
    p_stats.set_defaults(func=cmd_stats)

    args = parser.parse_args()
    args.func(args)
