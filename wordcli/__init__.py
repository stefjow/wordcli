"""wordcli — CLI tool for inspecting Word (.docx) documents."""

from .reader import DocxReader
from .replace import replace_in_docx
from .comments import add_comment_to_docx
from .remove_comment import remove_comment_from_docx
from .revert_change import revert_change_in_docx

__version__ = "0.1.0"
