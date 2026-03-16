"""Output helpers and table formatting."""

NBSP = "\u00a0"
NBSP_MARKER = "[NBSP]"


def show_nbsp(text):
    """Replace non-breaking spaces with a visible marker."""
    return text.replace(NBSP, NBSP_MARKER)


def parse_nbsp(text):
    """Replace [NBSP] markers with actual non-breaking spaces."""
    return text.replace(NBSP_MARKER, NBSP)


def table_to_markdown(rows):
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
