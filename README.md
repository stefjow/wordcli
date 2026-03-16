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
python -m wordcli <command> [options]
```

After `pip install -e .`, also available as `wordcli <command>`. Legacy `python wordcli.py` still works.

```
python -m wordcli --help              # List all commands
python -m wordcli <command> --help    # Show flags for a specific command
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
```

Tracked changes are shown as `[+inserted+]` and `[-deleted-]` by default.

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
python -m python -m wordcli comments document.docx
python -m python -m wordcli comments document.docx --author Claude
python -m python -m wordcli comments document.docx --json
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

### comment — Add a comment anchored to text

```
python -m wordcli comment document.docx --anchor "some phrase" --text "Please review" --author Claude
python -m wordcli comment document.docx --anchor "word" --text "Clarify this" --paragraph 5
python -m wordcli comment document.docx --anchor "X" --text "Update number" --context "Figure X)" --author Claude
python -m wordcli comment document.docx --anchor "phrase" --text "Note" --occurrence 2
```

Adds a comment anchored to the matched text, visible in Word's review pane. Uses the same disambiguation as `replace` (`--paragraph`, `--context`, `--occurrence`). If the anchor text matches multiple locations without scoping, the command refuses and lists all matches.

## Non-breaking spaces

Non-breaking spaces (U+00A0) are displayed as `[NBSP]` in all text output. When using `--old`, `--new`, `--anchor`, or `--context`, write `[NBSP]` and wordcli converts it back to a real non-breaking space automatically.

```
python -m wordcli text document.docx --paragraph 5
# Output: [5] 100[NBSP]000 inhabitants

python -m wordcli replace document.docx --old "100[NBSP]000" --new "100[NBSP]001" --author Claude --paragraph 5
```

## Claude Code integration

This project includes a [Claude Code](https://claude.com/claude-code) skill (`.claude/skills/wordcli/`) that enables LLMs to use wordcli autonomously — searching, replacing, and commenting with proper disambiguation.
