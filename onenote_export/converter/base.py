"""Base converter with shared file I/O logic for all output formats.

Subclasses implement render_page() to produce format-specific content.
"""

import logging
import re
from pathlib import Path

from onenote_export.model.content import EmbeddedFile, ImageElement
from onenote_export.model.notebook import Notebook
from onenote_export.model.page import Page
from onenote_export.model.section import Section

logger = logging.getLogger(__name__)


class BaseConverter:
    """Abstract base converter for OneNote content.

    Handles directory creation, file writing, images, and attachments.
    Subclasses must set FILE_EXTENSION and implement render_page().
    """

    FILE_EXTENSION: str = ""

    def __init__(self, output_dir: str | Path) -> None:
        self.output_dir = Path(output_dir)

    def convert_notebook(self, notebook: Notebook) -> list[Path]:
        """Convert an entire notebook to output files.

        Returns list of created file paths.
        """
        created: list[Path] = []
        notebook_dir = self.output_dir
        if notebook.relative_path:
            for part in Path(notebook.relative_path).parts:
                notebook_dir /= _sanitize_filename(part)
        elif notebook.name:
            notebook_dir /= _sanitize_filename(notebook.name)

        for section in notebook.sections:
            files = self.convert_section(section, notebook_dir)
            created.extend(files)

        return created

    def convert_section(
        self, section: Section, parent_dir: Path | None = None
    ) -> list[Path]:
        """Convert a section to output files.

        Returns list of created file paths.
        """
        created: list[Path] = []
        base_dir = parent_dir or self.output_dir
        section_dir = base_dir / _sanitize_filename(section.name)
        section_dir.mkdir(parents=True, exist_ok=True)

        seen_titles: dict[str, int] = {}

        for page in section.pages:
            content = self.render_page(page)
            filename = _page_filename(page.title, seen_titles, self.FILE_EXTENSION)
            file_path = section_dir / filename

            file_path.write_text(content, encoding="utf-8")
            created.append(file_path)
            logger.info("Wrote %s", file_path)

            image_files = self._write_images(page, section_dir)
            created.extend(image_files)

            embedded_files = self._write_embedded_files(page, section_dir)
            created.extend(embedded_files)

        return created

    def render_page(self, page: Page) -> str:
        """Render a single page to the target format.

        Subclasses must override this method.
        """
        raise NotImplementedError

    def _write_images(self, page: Page, section_dir: Path) -> list[Path]:
        """Write image data to files."""
        created: list[Path] = []
        images_dir = section_dir / "images"
        img_count = 0

        for element in page.elements:
            if isinstance(element, ImageElement) and element.data:
                images_dir.mkdir(exist_ok=True)
                img_count += 1
                filename = _sanitize_filename(
                    element.filename
                    or f"image_{img_count:03d}.{element.format or 'bin'}"
                )
                img_path = images_dir / filename
                img_path.write_bytes(element.data)
                created.append(img_path)
                logger.info("Wrote image %s", img_path)

        return created

    def _write_embedded_files(self, page: Page, section_dir: Path) -> list[Path]:
        """Write embedded file data to files."""
        created: list[Path] = []
        attachments_dir = section_dir / "attachments"

        for element in page.elements:
            if isinstance(element, EmbeddedFile) and element.data:
                attachments_dir.mkdir(exist_ok=True)
                filename = _sanitize_filename(element.filename or "attachment")
                file_path = attachments_dir / filename
                file_path.write_bytes(element.data)
                created.append(file_path)
                logger.info("Wrote attachment %s", file_path)

        return created


def _sanitize_filename(name: str) -> str:
    """Sanitize a string for use as a filename."""
    sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    sanitized = re.sub(r"[_\s]+", " ", sanitized).strip()
    if len(sanitized) > 200:
        sanitized = sanitized[:200]
    return sanitized or "unnamed"


def _page_filename(title: str, seen: dict[str, int], extension: str = ".md") -> str:
    """Generate a unique filename for a page title."""
    base = _sanitize_filename(title or "Untitled")
    if not base.endswith(extension):
        base = f"{base}{extension}"

    key = base.lower()
    if key in seen:
        seen[key] += 1
        stem = base[: -len(extension)]
        base = f"{stem} ({seen[key]}){extension}"
    else:
        seen[key] = 1

    return base
