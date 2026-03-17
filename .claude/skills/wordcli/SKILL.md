---
name: wordcli
description: >
  CLI tool for inspecting and editing Word (.docx) documents via the wordcli command.
  Use when the user asks to edit a docx, review a Word document, inspect a .docx file,
  add comments to a document, replace text with tracked changes, remove comments,
  revert tracked changes, add bookmarks, insert cross-references, detect field codes,
  search document text, extract text or tables, compare documents, verify text
  integrity, or explicitly mentions wordcli. Supports tracked changes, comments,
  footnotes, tables, bookmarks, cross-references, field codes, and search.
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
| `text <file>` | Extract text with paragraph numbers. `--paragraph N`, `--paragraphs N-M`, `--accept`, `--styles` |
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
| `bookmark <file> --anchor "X" --name "id"` | Add bookmark around text. `--paragraph`, `--context`, `--occurrence`, `-o` |
| `crossref <file> --bookmark "id" --text "X"` | Replace text with clickable REF field (tracked change). `--display`, `--author`, `--paragraph`, `--context`, `--occurrence`, `-o` |
| `fields <file>` | Show all field codes (SEQ, REF, etc.). `--seq` for SEQ fields only |
| `format <file> --text "X"` | Apply/remove bold, italic, underline, strike as tracked change. `--bold`, `--no-bold`, `--italic`, `--no-italic`, `--underline`, `--no-underline`, `--strike`, `--no-strike`, `--paragraph`, `--context`, `--occurrence`, `--author`, `-o` |
| `style <file>` | Show or change paragraph style. `--list`, `--paragraph N`, `--set StyleId`, `--author`, `-o`. Validates style ID against styles.xml |
| `xml <file> [part]` | Show raw XML of a document part. `--paragraph N`, `--paragraphs N-M`, `--list` |

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

Inline elements in `text` output use markdown-style markers:
- Footnote references: `[^N]` (e.g. `...ist.[^2]) Durch...`)
- Field codes: `[display](FIELD instruction)` (e.g. `[1](SEQ Übersicht \* ARABIC)`, `[Abbildung 1](REF _Ref_fig1 \h)`)
- Hyperlinks: `[text](url)`

Field markers help identify text that is a field (don't replace with `replace` — use `crossref` instead).

**Field overlap warning:** Both `replace` and `comment` emit a warning when the matched text overlaps a field code (REF, SEQ, HYPERLINK, etc.). For `comment`, fields are preserved — the warning is informational. For `replace`, the field structure will be destroyed — prefer `crossref` to modify field display text, or acknowledge that the field is intentionally being removed.

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

## Cross-references workflow

To replace placeholders (e.g. "Abbildung X") with clickable cross-references:

1. Check existing field codes: `fields <file> --seq` to find SEQ-numbered captions
2. Add bookmarks to the targets: `bookmark <file> --anchor "Übersicht 1" --name "uebersicht1" --paragraph 11`
3. Replace placeholders with REF fields: `crossref <file> --bookmark uebersicht1 --text "Übersicht[NBSP]X" --paragraph 9 --display "Übersicht 1" --author Claude`

Important: do crossrefs BEFORE text corrections, since `crossref` cannot find text inside tracked changes.

## Run formatting (bold, italic, etc.)

Use `format` to apply or remove run-level formatting as a tracked change:

```bash
python -m wordcli format document.docx --text "important" --bold --paragraph 5 --author Claude
python -m wordcli format document.docx --text "not italic" --no-italic --paragraph 5 --author Claude
python -m wordcli format document.docx --text "emphasis" --bold --italic --underline --author Claude
```

Supports `--bold/--no-bold`, `--italic/--no-italic`, `--underline/--no-underline`, `--strike/--no-strike`. Multiple can be combined. Uses the same disambiguation as `replace` (`--paragraph`, `--context`, `--occurrence`). Changes show as tracked formatting changes in Word's review pane and appear as `[FORMAT]` in `changes` output.

## Paragraph styles

Use `text --styles` to see paragraph style IDs alongside text, and `style` to query or change them:

```bash
# Scan document structure with styles
python -m wordcli text document.docx --styles
# [1:berschrift1] Der nominell-effektive...
# [2] Der nominell-effektive Wechselkurs ist...    <- no style = default/Normal
# [3:Beschriftung] Abbildung 1: Entwicklung...

# List available paragraph styles (ID + display name)
python -m wordcli style document.docx --list

# Query a single paragraph's style
python -m wordcli style document.docx --paragraph 3

# Change a paragraph's style (tracked change)
python -m wordcli style document.docx --paragraph 2 --set berschrift2 --author Claude
```

Style IDs are internal names (e.g. `berschrift1`, not "Heading 1") and vary by document language. Always use `--list` to discover valid IDs. The style change appears as a tracked formatting change in Word's review pane.

## Inspecting raw XML

The `xml` command shows the raw OOXML of any document part. Use it to inspect formatting (bold, italic, styles) that is not visible in `text` output, or to debug unexpected results from editing commands.

```bash
python -m wordcli xml document.docx --list                  # List all parts in the archive
python -m wordcli xml document.docx --paragraph 5           # XML for paragraph 5 (document.xml)
python -m wordcli xml document.docx --paragraphs 3-7        # Range of paragraphs
python -m wordcli xml document.docx styles                  # Full styles.xml
python -m wordcli xml document.docx footnotes               # Full footnotes.xml
python -m wordcli xml document.docx comments                # Full comments.xml
python -m wordcli xml document.docx word/footer1.xml        # Any zip path
```

Named parts: `document` (default), `footnotes`, `comments`, `styles`, `numbering`, `settings`, `rels`. You can also pass any zip path directly (use `--list` to discover them).

**When to use:** Do NOT look at raw XML before every operation — `replace` already preserves run formatting. Use `xml` only when:
- You need to check what formatting (bold, italic, font) applies to specific text
- An editing command produced unexpected results and you want to understand why
- You need to inspect structure not exposed by other commands (headers, footers, numbering, relationships)

## Notes

- If a write command fails with `PermissionError`, the file is likely open in Word. Ask the user to close it or offer to write to a new file with `-o`.
- `replace`, `comment`, `remove-comment`, and `revert-change` overwrite the input file by default. Use `-o output.docx` to write to a separate file.
- `--author` defaults to `wordcli`. Set it explicitly (e.g. `--author Claude`) for attribution.
- `verify` exit code: 0 = OK, 1 = text loss detected. Use `--author` to only remove one author's changes.
- Without `--accept`, tracked changes show as `[+inserted+]` and `[-deleted-]` in text output.
