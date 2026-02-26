"""Notebook model representing a directory of sections."""

from dataclasses import dataclass, field

from onenote_export.model.section import Section


@dataclass
class Notebook:
    """A notebook (corresponds to a directory) containing sections."""

    name: str = ""
    dir_path: str = ""
    sections: list[Section] = field(default_factory=list)
