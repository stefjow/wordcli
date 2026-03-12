# wordcli

CLI tool for inspecting Word (.docx) documents. No dependencies beyond Python 3 stdlib.

## Usage

```
python wordcli.py <command> [options]
```

## Commands

### extract — Structured markdown export

```
wordcli extract file.docx                # Markdown to stdout
wordcli extract file.docx -o output.md   # Write to file
```

Produces markdown with headings (from Word styles), footnote markers (`[^2]`) with definitions at the end, and tables as markdown tables (including merged cells).

### stats — Document statistics

```
wordcli stats file.docx          # Human-readable
wordcli stats file.docx --json   # JSON
```

Counts paragraphs, tables, footnotes, comments (by author), and tracked changes (by author, INS/DEL).

### text — Extract full text

```
wordcli text file.docx                    # All paragraphs with numbers
wordcli text file.docx --paragraph 5      # Single paragraph
wordcli text file.docx --paragraphs 3-7   # Range
wordcli text file.docx --accept           # Accepted text only (no change markers)
```

Tracked changes are shown as `[+inserted+]` and `[-deleted-]` by default.

### search — Search with context

```
wordcli search file.docx "search term"
wordcli search file.docx "search term" --footnotes   # Also search footnotes
```

Shows paragraph number and 50 characters of surrounding context.

### footnotes — List footnotes

```
wordcli footnotes file.docx        # All footnotes
wordcli footnotes file.docx 3      # Single footnote by ID
```

### comments — List comments

```
wordcli comments file.docx
wordcli comments file.docx --author Claude
wordcli comments file.docx --json
```

### changes — Show tracked changes

```
wordcli changes file.docx
wordcli changes file.docx --author Claude
```

### tables — Extract tables as markdown

```
wordcli tables file.docx      # All tables
wordcli tables file.docx 1    # Single table by number
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
wordcli replace file.docx --old "Tpyo" --new "Typo" --author Claude
wordcli replace file.docx --old "word" --new "term" --paragraph 5
wordcli replace file.docx --old "X" --new "2" --context "Figure X)" --author Claude
wordcli replace file.docx --old "word" --new "term" -o output.docx
```

Replaces text as a tracked change (insertion + deletion) visible in Word's review mode. Handles text that spans multiple runs.

Scoping options to avoid ambiguous matches:
- `--paragraph N`: Only search in paragraph N (use `text` or `search` to find the number)
- `--context "..."`: Provide a longer unique string that contains `--old`; only the `--old` portion is replaced

Without `-o`, the input file is overwritten. With `-o`, a new file is created.
