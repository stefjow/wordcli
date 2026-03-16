"""wordcli — CLI tool for inspecting Word (.docx) documents."""

from .reader import DocxReader
from .replace import replace_in_docx
from .comments import add_comment_to_docx

__version__ = "0.1.0"
