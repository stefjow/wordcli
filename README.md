# wordcli

CLI tool for inspecting and editing Word (.docx) documents. No dependencies beyond Python 3 stdlib.

## Installation

```
pixi add --path .
```

Or without pixi:

```
pip install -e .
```

## Usage

```
wordcli <command> [options]
```

Also works as `python -m wordcli` or `python wordcli.py` (backwards compatible).

## Commands

### extract — Structured markdown export

```
wordcli extract document.docx                # Markdown to stdout
wordcli extract document.docx -o output.md   # Write to file
```

Produces markdown with headings (from Word styles), footnote markers (`[^2]`) with definitions at the end, and tables as markdown tables (including merged cells).

### stats — Document statistics

```
wordcli stats document.docx          # Human-readable
wordcli stats document.docx --json   # JSON
```

Counts paragraphs, tables, footnotes, comments (by author), and tracked changes (by author, INS/DEL).

### text — Extract full text

```
wordcli text document.docx                    # All paragraphs with numbers
wordcli text document.docx --paragraph 5      # Single paragraph
wordcli text document.docx --paragraphs 3-7   # Range
wordcli text document.docx --accept           # Accepted text only (no change markers)
```

Tracked changes are shown as `[+inserted+]` and `[-deleted-]` by default.

### search — Search with context

```
wordcli search document.docx "query"
wordcli search document.docx "query" --footnotes   # Also search footnotes
```

Shows paragraph number and 50 characters of surrounding context.

### footnotes — List footnotes

```
wordcli footnotes document.docx        # All footnotes
wordcli footnotes document.docx 3      # Single footnote by ID
```

### comments — List comments

```
wordcli comments document.docx
wordcli comments document.docx --author Claude
wordcli comments document.docx --json
```

### changes — Show tracked changes

```
wordcli changes document.docx
wordcli changes document.docx --author Claude
```

### tables — Extract tables as markdown

```
wordcli tables document.docx      # All tables
wordcli tables document.docx 1    # Single table by number
```

### diff — Compare two documents

```
wordcli diff original.docx edited.docx
```

Compares accepted text paragraph by paragraph.

### verify — Check for text loss (main text + footnotes)

```
wordcli verify edited.docx --original original.docx              # Remove ALL tracked changes
wordcli verify edited.docx --original original.docx --author Claude  # Only remove Claude's changes
```

Removes tracked changes from the edited document and compares both main text and footnotes against the original. Without `--author`, all tracked changes are removed. With `--author`, only that author's changes are removed (useful when the document already had tracked changes from other reviewers). Exit code 0 = OK, 1 = text loss detected.

### replace — Replace text as tracked change

```
wordcli replace document.docx --old "typo" --new "fixed" --author Claude
wordcli replace document.docx --old "word" --new "term" --paragraph 5
wordcli replace document.docx --old "X" --new "2" --context "Figure X)" --author Claude
wordcli replace document.docx --old "word" --new "term" -o output.docx
```

Replaces text as a tracked change (insertion + deletion) visible in Word's review mode. Handles text that spans multiple runs.

Scoping options to avoid ambiguous matches:
- `--paragraph N`: Only search in paragraph N (use `text` or `search` to find the number)
- `--context "..."`: Provide a longer unique string that contains `--old`; only the `--old` portion is replaced

Without `-o`, the input file is overwritten. With `-o`, a new file is created.

### comment — Add a comment anchored to text

```
wordcli comment document.docx --anchor "some phrase" --text "Please review" --author Claude
wordcli comment document.docx --anchor "word" --text "Clarify this" --paragraph 5
wordcli comment document.docx --anchor "X" --text "Update number" --context "Figure X)" --author Claude
wordcli comment document.docx --anchor "phrase" --text "Note" -o output.docx
```

Adds a comment anchored to the matched text, visible in Word's review pane. Uses the same scoping options as `replace` (`--paragraph`, `--context`) to target specific occurrences.
