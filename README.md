# wordcli

CLI tool for inspecting Word (.docx) documents. No dependencies beyond Python 3 stdlib.

## Usage

```
python wordcli.py <command> [options]
```

## Commands

### extract — Structured markdown export

```
wordcli extract datei.docx                # Markdown to stdout
wordcli extract datei.docx -o output.md   # Write to file
```

Produces markdown with headings (from Word styles), footnote markers (`[^2]`) with definitions at the end, and tables as markdown tables (including merged cells).

### stats — Document statistics

```
wordcli stats datei.docx          # Human-readable
wordcli stats datei.docx --json   # JSON
```

Counts paragraphs, tables, footnotes, comments (by author), and tracked changes (by author, INS/DEL).

### text — Extract full text

```
wordcli text datei.docx                    # All paragraphs with numbers
wordcli text datei.docx --paragraph 5      # Single paragraph
wordcli text datei.docx --paragraphs 3-7   # Range
wordcli text datei.docx --accept           # Accepted text only (no change markers)
```

Tracked changes are shown as `[+inserted+]` and `[-deleted-]` by default.

### search — Search with context

```
wordcli search datei.docx "Suchbegriff"
wordcli search datei.docx "Suchbegriff" --footnotes   # Also search footnotes
```

Shows paragraph number and 50 characters of surrounding context.

### footnotes — List footnotes

```
wordcli footnotes datei.docx        # All footnotes
wordcli footnotes datei.docx 3      # Single footnote by ID
```

### comments — List comments

```
wordcli comments datei.docx
wordcli comments datei.docx --author Claude
wordcli comments datei.docx --json
```

### changes — Show tracked changes

```
wordcli changes datei.docx
wordcli changes datei.docx --author Claude
```

### tables — Extract tables as markdown

```
wordcli tables datei.docx      # All tables
wordcli tables datei.docx 1    # Single table by number
```

### diff — Compare two documents

```
wordcli diff original.docx edited.docx
```

Compares accepted text paragraph by paragraph.

### verify — Check for text loss (main text + footnotes)

```
wordcli verify edited.docx --original original.docx
```

Removes Claude's tracked changes from the edited document and compares both main text and footnotes against the original. Exit code 0 = OK, 1 = text loss detected.
