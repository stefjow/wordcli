"""CLI commands and argument parser."""

import argparse
import json
import re
import sys

from .reader import DocxReader
from .replace import replace_in_docx
from .comments import add_comment_to_docx
from .formatting import show_nbsp, parse_nbsp, table_to_markdown


def cmd_text(args):
    doc = DocxReader(args.file)
    paras = doc.extract_paragraphs(accept_changes=args.accept)
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
    )
    if ok:
        print(msg)
    else:
        print(f"Error: {msg}", file=sys.stderr)
        sys.exit(1)


def cmd_comment(args):
    output = args.output or args.file
    anchor = parse_nbsp(args.anchor)
    context = parse_nbsp(args.context) if args.context else None
    ok, msg = add_comment_to_docx(
        args.file, output, anchor, args.text,
        author=args.author,
        paragraph=args.paragraph,
        context=context,
        occurrence=args.occurrence,
    )
    if ok:
        print(msg)
    else:
        print(f"Error: {msg}", file=sys.stderr)
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
    p_rep.add_argument("-o", "--output", default=None, help="Output file (default: overwrite input)")
    p_rep.set_defaults(func=cmd_replace)

    # comment
    p_comment = sub.add_parser("comment", help="Add a comment anchored to text")
    p_comment.add_argument("file")
    p_comment.add_argument("--anchor", required=True, help="Text to anchor the comment to")
    p_comment.add_argument("--text", required=True, help="Comment text")
    p_comment.add_argument("--author", default="wordcli", help="Author name for the comment")
    p_comment.add_argument("--paragraph", type=int, default=None, help="Limit to paragraph number")
    p_comment.add_argument("--context", default=None, help="Unique surrounding text to locate the anchor")
    p_comment.add_argument("--occurrence", type=int, default=None, help="Match the Nth occurrence (1-based)")
    p_comment.add_argument("-o", "--output", default=None, help="Output file (default: overwrite input)")
    p_comment.set_defaults(func=cmd_comment)

    # stats
    p_stats = sub.add_parser("stats", help="Document statistics")
    p_stats.add_argument("file")
    p_stats.add_argument("--json", action="store_true")
    p_stats.set_defaults(func=cmd_stats)

    args = parser.parse_args()
    args.func(args)
