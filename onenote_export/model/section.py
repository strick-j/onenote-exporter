"""Section model representing a single .one file."""

from dataclasses import dataclass, field

from onenote_export.model.page import Page


@dataclass
class Section:
    """A section (corresponds to a single .one file) containing pages."""

    name: str = ""
    file_path: str = ""
    pages: list[Page] = field(default_factory=list)
