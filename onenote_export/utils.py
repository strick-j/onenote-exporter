"""Utility functions for the OneNote export tool."""

import re
from pathlib import Path


def discover_one_files(input_dir: Path) -> list[Path]:
    """Recursively find all .one files in a directory.

    Excludes .onetoc2 table-of-contents files.
    """
    files = sorted(
        p
        for p in input_dir.rglob("*.one")
        if p.is_file() and not p.name.endswith(".onetoc2")
    )
    return files


def section_name_from_filename(filename: str) -> str:
    """Extract a clean section name from a .one filename.

    Examples:
        'ADI (On 2-25-26).one' -> 'ADI'
        'ADP.one (On 8-24-25).one' -> 'ADP'
        'BMS (On 2-25-26).one' -> 'BMS'
    """
    name = Path(filename).stem

    # Strip ' (On M-D-YY)' or ' (On M-D-YY - N)' suffix
    name = re.sub(r"\s*\(On\s+\d+-\d+-\d+(?:\s*-\s*\d+)?\)", "", name)

    # Strip trailing '.one' that appears in 'Name.one (On date)' pattern
    name = re.sub(r"\.one$", "", name, flags=re.IGNORECASE)

    return name.strip() or "Untitled"


def notebook_name_from_dir(dir_path: Path) -> str:
    """Extract a notebook name from a directory path."""
    return dir_path.name or "Untitled"
