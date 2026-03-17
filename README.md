# wordcli

CLI tool for inspecting and editing Word (.docx) documents. No dependencies beyond Python 3 stdlib.

## Usage

```
python -m wordcli <command> [options]
```

No installation required. Optionally run `pip install -e .` to get a `wordcli` command on PATH.

```
python -m wordcli --help              # List all commands
python -m wordcli <command> --help    # Show flags for a specific command
python -m wordcli --version           # Show version
```

## Commands

### extract — Structured markdown export

```
python -m wordcli extract document.docx                # Markdown to stdout
python -m wordcli extract document.docx -o output.md   # Write to file
```

Produces markdown with headings (from Word styles), footnote markers (`[^2]`) with definitions at the end, and tables as markdown tables (including merged cells).

### stats — Document statistics

```
python -m wordcli stats document.docx          # Human-readable
python -m wordcli stats document.docx --json   # JSON
```

Counts paragraphs, tables, footnotes, comments (by author), and tracked changes (by author, INS/DEL).

### text — Extract full text

```
python -m wordcli text document.docx                    # All paragraphs with numbers
python -m wordcli text document.docx --paragraph 5      # Single paragraph
python -m wordcli text document.docx --paragraphs 3-7   # Range
python -m wordcli text document.docx --accept           # Accepted text only (no change markers)
python -m wordcli text document.docx --styles           # Show paragraph style IDs
```

With `--styles`, the output shows `[nr:StyleId]` instead of `[nr]`:
```
[1:berschrift1] Der nominell-effektive Wechselkurs...
[2] Der nominell-effektive Wechselkurs ist...         <- no style = default/Normal
[3:Beschriftung] Abbildung 1: Entwicklung...
```

Tracked changes are shown as `[+inserted+]` and `[-deleted-]` by default. Inline markers use markdown-style syntax:
- Footnote references: `[^N]` (e.g. `...text.[^3]) More text...`)
- Field codes: `[display](FIELD instruction)` (e.g. `[1](SEQ Übersicht \* ARABIC)`, `[Abbildung 1](REF _Ref_fig1 \h)`)
- Hyperlinks: `[text](url)`

### search — Search with context

```
python -m wordcli search document.docx "query"
python -m wordcli search document.docx "query" --footnotes        # Also search footnotes
python -m wordcli search document.docx "query" --context-size 80  # More context (default: 50)
```

Shows paragraph number and surrounding context.

### footnotes — List footnotes

```
python -m wordcli footnotes document.docx        # All footnotes
python -m wordcli footnotes document.docx 3      # Single footnote by ID
```

### comments — List comments

```
python -m wordcli comments document.docx
python -m wordcli comments document.docx --author Claude
python -m wordcli comments document.docx --json
```

### changes — Show tracked changes

```
python -m wordcli changes document.docx
python -m wordcli changes document.docx --author Claude
```

### tables — Extract tables as markdown

```
python -m wordcli tables document.docx      # All tables
python -m wordcli tables document.docx 1    # Single table by number
```

### diff — Compare two documents

```
python -m wordcli diff original.docx edited.docx
```

Compares accepted text paragraph by paragraph.

### verify — Check for text loss (main text + footnotes)

```
python -m wordcli verify edited.docx --original original.docx                    # Remove ALL tracked changes
python -m wordcli verify edited.docx --original original.docx --author Claude    # Only remove Claude's changes
python -m wordcli verify edited.docx --original original.docx --truncate 200     # Longer preview lines (default: 120)
```

Removes tracked changes from the edited document and compares both main text and footnotes against the original. Without `--author`, all tracked changes are removed. With `--author`, only that author's changes are removed (useful when the document already had tracked changes from other reviewers). Exit code 0 = OK, 1 = text loss detected.

### replace — Replace text as tracked change

```
python -m wordcli replace document.docx --old "typo" --new "fixed" --author Claude
python -m wordcli replace document.docx --old "word" --new "term" --paragraph 5
python -m wordcli replace document.docx --old "X" --new "2" --context "Figure X)" --author Claude
python -m wordcli replace document.docx --old "word" --new "term" -o output.docx
```

Replaces text as a tracked change (insertion + deletion) visible in Word's review mode. Handles text that spans multiple runs.

If the text matches multiple locations, the command refuses with an error listing all matches. Scoping options to disambiguate:
- `--paragraph N`: Only search in paragraph N (use `text` or `search` to find the number)
- `--context "..."`: Provide a longer unique string that contains `--old`; only the `--old` portion is replaced
- `--occurrence N`: Select the Nth match (1-based)

Without `-o`, the input file is overwritten. With `-o`, a new file is created.

Use `--footnote N` to replace within a footnote instead of the main text:

```
python -m wordcli replace document.docx --old "typo" --new "fixed" --author Claude --footnote 3
```

### comment — Add a comment anchored to text

```
python -m wordcli comment document.docx --anchor "some phrase" --text "Please review" --author Claude
python -m wordcli comment document.docx --anchor "word" --text "Clarify this" --paragraph 5
python -m wordcli comment document.docx --anchor "X" --text "Update number" --context "Figure X)" --author Claude
python -m wordcli comment document.docx --anchor "phrase" --text "Note" --occurrence 2
python -m wordcli comment document.docx --footnote 3 --text "Check this footnote" --author Claude
```

Adds a comment anchored to the matched text, visible in Word's review pane. Uses the same disambiguation as `replace` (`--paragraph`, `--context`, `--occurrence`). If the anchor text matches multiple locations without scoping, the command refuses and lists all matches.

Use `--footnote N` (without `--anchor`) to place the comment on the footnote reference number in the main text.

### remove-comment — Remove a comment by ID

```
python -m wordcli remove-comment document.docx 5              # Remove comment with ID 5
python -m wordcli remove-comment document.docx 5 -o out.docx  # Write to separate file
```

Removes the comment and its range markers from the document. Use `comments` to find the ID.

### revert-change — Revert a tracked change

```
python -m wordcli revert-change document.docx --text "typo" --type del    # Revert a specific deletion
python -m wordcli revert-change document.docx --author Claude --occurrence 1  # Revert first change by Claude
python -m wordcli revert-change document.docx --text "fixed" --type ins   # Remove an insertion
python -m wordcli revert-change document.docx --footnote 3 --occurrence 1 # Revert change in footnote
```

Reverts a tracked change: for insertions, the inserted text is removed; for deletions, the original text is restored. If multiple changes match the filters, the command lists all matches and requires `--occurrence` to disambiguate.

### bookmark — Add a bookmark around text

```
python -m wordcli bookmark document.docx --anchor "Übersicht 1" --name "uebersicht1" --paragraph 11
python -m wordcli bookmark document.docx --anchor "Figure 1" --name "fig1" --context "Figure 1: Title" -o out.docx
```

Wraps the matched text with `bookmarkStart`/`bookmarkEnd` markers. Non-destructive — preserves existing elements (including SEQ field codes). Uses the same disambiguation as `replace` (`--paragraph`, `--context`, `--occurrence`).

### crossref — Insert a clickable cross-reference

```
python -m wordcli crossref document.docx --bookmark fig1 --text "Figure X" --paragraph 5 --display "Figure 1" --author Claude
python -m wordcli crossref document.docx --bookmark uebersicht1 --text "Übersicht[NBSP]X" --paragraph 9 --display "Übersicht 1"
```

Replaces the matched text with a clickable REF field pointing to the named bookmark. The replacement is shown as a tracked change (deletion of old text + insertion of REF field). `--display` sets the cached display text shown in the field.

### fields — Show field codes in the document

```
python -m wordcli fields document.docx        # All fields
python -m wordcli fields document.docx --seq   # Only SEQ fields (captions)
```

Lists all field codes (SEQ, REF, TOC, etc.) with their paragraph number, field instruction, and display text. Useful for checking existing caption numbering before adding bookmarks or cross-references.

### style — Show or change paragraph style

```
python -m wordcli style document.docx --list                                        # List available paragraph styles
python -m wordcli style document.docx --list --type character                       # List character styles
python -m wordcli style document.docx --paragraph 5                                 # Show current style of paragraph 5
python -m wordcli style document.docx --paragraph 5 --set Heading2 --author Claude  # Change style (tracked change)
python -m wordcli style document.docx --paragraph 5 --set Standard -o out.docx      # Write to separate file
```

Changes the paragraph style as a tracked formatting change visible in Word's review pane. Use `--list` to discover valid style IDs (these are internal names like `berschrift1`, not display names like "Heading 1"). Without `--set`, queries the current style.

### xml — Show raw XML of a document part

```
python -m wordcli xml document.docx                        # Full document.xml (pretty-printed)
python -m wordcli xml document.docx --paragraph 5          # Single paragraph
python -m wordcli xml document.docx --paragraphs 3-7       # Range of paragraphs
python -m wordcli xml document.docx styles                  # Full styles.xml
python -m wordcli xml document.docx comments                # Full comments.xml
python -m wordcli xml document.docx footnotes               # Full footnotes.xml
python -m wordcli xml document.docx word/footer1.xml        # Any zip path
python -m wordcli xml document.docx --list                  # List all parts in the archive
```

Shows the raw OOXML with proper namespace prefixes (`w:`, `w14:`, etc.). Useful for inspecting formatting (bold, italic, fonts, styles) that is not visible in `text` output, or debugging unexpected results from editing commands.

Named parts: `document` (default), `footnotes`, `comments`, `styles`, `numbering`, `settings`, `rels`. Any zip path works too (use `--list` to discover available parts).

## Non-breaking spaces

Non-breaking spaces (U+00A0) are displayed as `[NBSP]` in all text output. When using `--old`, `--new`, `--anchor`, or `--context`, write `[NBSP]` and wordcli converts it back to a real non-breaking space automatically.

```
python -m wordcli text document.docx --paragraph 5
# Output: [5] 100[NBSP]000 inhabitants

python -m wordcli replace document.docx --old "100[NBSP]000" --new "100[NBSP]001" --author Claude --paragraph 5
```

## Claude Code integration

This project includes a [Claude Code](https://claude.com/claude-code) skill (`.claude/skills/wordcli/`) that enables LLMs to use wordcli autonomously — searching, replacing, and commenting with proper disambiguation.
