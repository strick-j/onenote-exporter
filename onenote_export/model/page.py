"""Page model representing a single OneNote page."""

from dataclasses import dataclass, field

from onenote_export.model.content import ContentElement


@dataclass
class Page:
    """A single page in a OneNote section."""

    title: str = ""
    level: int = 0
    creation_time: int = 0
    last_modified_time: int = 0
    author: str = ""
    elements: list[ContentElement] = field(default_factory=list)
