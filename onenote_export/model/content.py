"""Content elements that make up a OneNote page."""

from dataclasses import dataclass, field


@dataclass(frozen=True)
class TextRun:
    """A run of text with uniform formatting."""
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strikethrough: bool = False
    superscript: bool = False
    subscript: bool = False
    font: str = ""
    font_size: int = 0
    hyperlink_url: str = ""


@dataclass
class ContentElement:
    """Base class for content elements."""
    pass


@dataclass
class RichText(ContentElement):
    """A rich text element containing formatted text runs."""
    runs: list[TextRun] = field(default_factory=list)
    indent_level: int = 0
    is_title: bool = False
    alignment: str = ""  # "left", "center", "right"
    heading_level: int = 0  # 1-6 for headings, 0 for normal text
    list_type: str = ""  # "ordered", "unordered", "" for non-list


@dataclass
class ImageElement(ContentElement):
    """An image embedded in the page."""
    data: bytes = b""
    filename: str = ""
    alt_text: str = ""
    width: int = 0
    height: int = 0
    format: str = ""  # "png", "jpeg", "gif", "bmp"


@dataclass
class TableElement(ContentElement):
    """A table with rows and cells."""
    rows: list[list[list[ContentElement]]] = field(default_factory=list)
    column_widths: list[float] = field(default_factory=list)
    borders_visible: bool = True


@dataclass
class EmbeddedFile(ContentElement):
    """An embedded file attachment."""
    data: bytes = b""
    filename: str = ""
    source_path: str = ""
