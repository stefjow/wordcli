---
name: wordcli
description: >
  CLI tool for inspecting and editing Word (.docx) documents via the wordcli command.
  Use when the user asks to edit a docx, review a Word document, inspect a .docx file,
  add comments to a document, replace text with tracked changes, remove comments,
  revert tracked changes, search document text, extract text or tables, compare
  documents, verify text integrity, or explicitly mentions wordcli. Supports tracked
  changes, comments, footnotes, tables, and search.
---

# wordcli

## Setup

Ensure wordcli is available before use:

```bash
python -m wordcli --version 2>/dev/null || (git clone https://gitea.wsr.ac.at/sweing/wordcli /tmp/wordcli && pip install /tmp/wordcli)
```

Run via `python -m wordcli <command>`. Use `python -m wordcli --help` or `python -m wordcli <command> --help` for full flag details.

## Commands

| Command | Purpose |
|---------|---------|
| `text <file>` | Extract text with paragraph numbers. `--paragraph N`, `--paragraphs N-M`, `--accept` |
| `search <file> "query"` | Search with context snippets. `--footnotes`, `--context-size N` |
| `extract <file>` | Structured markdown export. `-o file.md` |
| `stats <file>` | Document statistics. `--json` |
| `footnotes <file> [id]` | List or show footnotes |
| `comments <file>` | List comments. `--author`, `--json` |
| `changes <file>` | Show tracked changes. `--author` |
| `tables <file> [N]` | Extract tables as markdown |
| `diff <file1> <file2>` | Compare accepted text paragraph by paragraph |
| `verify <file> --original <orig>` | Check for text loss. `--author`, `--truncate N` |
| `replace <file> --old "X" --new "Y"` | Replace as tracked change. `--author`, `--paragraph`, `--context`, `--occurrence`, `--footnote N`, `-o` |
| `comment <file> --anchor "X" --text "Y"` | Add comment anchored to text. `--author`, `--paragraph`, `--context`, `--occurrence`, `-o` |
| `comment <file> --footnote N --text "Y"` | Add comment on footnote reference in main text. `--author`, `-o` |
| `remove-comment <file> <id>` | Remove a comment by ID. `-o` |
| `revert-change <file>` | Revert a tracked change. `--author`, `--text`, `--type ins\|del`, `--occurrence`, `--footnote N`, `-o` |

## Key workflow: search before replace/comment

Both `replace` and `comment` **refuse if the target text matches multiple locations**. Always search first to find the paragraph number, then scope:

```bash
# 1. Find where the text appears
python -m wordcli search document.docx "target phrase"
# Output: [14] ...context around target phrase...
#         [87] ...another target phrase occurrence...

# 2. Use --paragraph to scope
python -m wordcli replace document.docx --old "target phrase" --new "fixed" --author Claude --paragraph 14

# Or use --occurrence for the Nth match
python -m wordcli comment document.docx --anchor "target phrase" --text "Review this" --author Claude --occurrence 1

# Or use --context with a longer unique string
python -m wordcli replace document.docx --old "phrase" --new "term" --context "unique surrounding target phrase text" --author Claude
```

## NBSP handling

Non-breaking spaces (U+00A0) appear as `[NBSP]` in all text output. When passing text containing non-breaking spaces to `--old`, `--new`, `--anchor`, or `--context`, use the `[NBSP]` marker — wordcli converts it back automatically.

```bash
# Search output shows: [5] 100[NBSP]000 inhabitants
# To replace, use the marker:
python -m wordcli replace document.docx --old "100[NBSP]000" --new "100[NBSP]001" --author Claude --paragraph 5
```

## Footnotes

Use `footnotes <file>` to list footnote text. To replace within a footnote, use `--footnote N`:

```bash
python -m wordcli replace document.docx --old "typo" --new "fixed" --author Claude --footnote 3
```

To comment on a footnote, use `--footnote N` without `--anchor` — the comment anchors to the footnote reference in the main text:

```bash
python -m wordcli comment document.docx --footnote 3 --text "Check this" --author Claude
```

## Undoing mistakes

Use `remove-comment` and `revert-change` to fix errors without rebuilding the document:

```bash
# List comments to find the ID, then remove it
python -m wordcli comments document.docx
python -m wordcli remove-comment document.docx 5

# List changes, then revert a specific one
python -m wordcli changes document.docx
python -m wordcli revert-change document.docx --text "typo" --type del
python -m wordcli revert-change document.docx --text "typo" --type ins

# Revert by occurrence when multiple match
python -m wordcli revert-change document.docx --author Claude --occurrence 3
```

A replace creates paired DEL+INS changes. To fully undo a replace, revert both the DEL (restores original) and the INS (removes replacement).

## Notes

- If a write command fails with `PermissionError`, the file is likely open in Word. Ask the user to close it or offer to write to a new file with `-o`.
- `replace`, `comment`, `remove-comment`, and `revert-change` overwrite the input file by default. Use `-o output.docx` to write to a separate file.
- `--author` defaults to `wordcli`. Set it explicitly (e.g. `--author Claude`) for attribution.
- `verify` exit code: 0 = OK, 1 = text loss detected. Use `--author` to only remove one author's changes.
- Without `--accept`, tracked changes show as `[+inserted+]` and `[-deleted-]` in text output.
