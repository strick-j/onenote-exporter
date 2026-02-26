"""Markdown converter for OneNote content model.

Converts Page objects with ContentElements into Markdown text,
and writes the files preserving directory structure.
"""

import logging
import re
from pathlib import Path

from onenote_export.model.content import (
    ContentElement,
    EmbeddedFile,
    ImageElement,
    RichText,
    TableElement,
)
from onenote_export.model.notebook import Notebook
from onenote_export.model.page import Page
from onenote_export.model.section import Section

logger = logging.getLogger(__name__)


class MarkdownConverter:
    """Converts OneNote content model to Markdown files."""

    def __init__(self, output_dir: str | Path) -> None:
        self.output_dir = Path(output_dir)

    def convert_notebook(self, notebook: Notebook) -> list[Path]:
        """Convert an entire notebook to Markdown files.

        Returns list of created file paths.
        """
        created: list[Path] = []
        notebook_dir = self.output_dir / _sanitize_filename(notebook.name)

        for section in notebook.sections:
            files = self.convert_section(section, notebook_dir)
            created.extend(files)

        return created

    def convert_section(
        self, section: Section, parent_dir: Path | None = None
    ) -> list[Path]:
        """Convert a section to Markdown files.

        Returns list of created file paths.
        """
        created: list[Path] = []
        base_dir = parent_dir or self.output_dir
        section_dir = base_dir / _sanitize_filename(section.name)
        section_dir.mkdir(parents=True, exist_ok=True)

        seen_titles: dict[str, int] = {}

        for page in section.pages:
            md_content = self.render_page(page)
            filename = _page_filename(page.title, seen_titles)
            file_path = section_dir / filename

            file_path.write_text(md_content, encoding="utf-8")
            created.append(file_path)
            logger.info("Wrote %s", file_path)

            # Write images
            image_files = self._write_images(page, section_dir)
            created.extend(image_files)

            # Write embedded files
            embedded_files = self._write_embedded_files(page, section_dir)
            created.extend(embedded_files)

        return created

    def render_page(self, page: Page) -> str:
        """Render a single page to Markdown text."""
        lines: list[str] = []

        # Page title as H1
        if page.title:
            lines.append(f"# {page.title}")
            lines.append("")

        # Render each content element, tracking ordered list counters
        # per indent level so numbered lists increment correctly.
        ordered_counters: dict[int, int] = {}

        for element in page.elements:
            if isinstance(element, RichText) and element.list_type == "ordered":
                level = element.indent_level
                # Initialize counter for new indent levels
                if level not in ordered_counters:
                    ordered_counters[level] = 0
                # Clear deeper level counters when returning to a
                # shallower level (they restart on re-entry)
                for k in list(ordered_counters):
                    if k > level:
                        del ordered_counters[k]
                ordered_counters[level] += 1
                md = self._render_rich_text(
                    element,
                    ordered_number=ordered_counters[level],
                )
            else:
                if not (isinstance(element, RichText) and element.list_type):
                    ordered_counters.clear()
                md = self._render_element(element)

            if md:
                lines.append(md)
                lines.append("")

        # Add metadata footer if available
        if page.author:
            lines.append("---")
            lines.append(f"*Author: {page.author}*")
            lines.append("")

        return "\n".join(lines)

    def _render_element(self, element: ContentElement) -> str:
        """Render a single content element to Markdown."""
        if isinstance(element, RichText):
            return self._render_rich_text(element)
        elif isinstance(element, ImageElement):
            return self._render_image(element)
        elif isinstance(element, TableElement):
            return self._render_table(element)
        elif isinstance(element, EmbeddedFile):
            return self._render_embedded_file(element)
        return ""

    def _render_rich_text(
        self,
        rt: RichText,
        ordered_number: int = 0,
    ) -> str:
        """Render rich text to Markdown."""
        parts: list[str] = []

        for run in rt.runs:
            text = run.text
            if not text:
                continue

            # Apply formatting (skip inline formatting for headings)
            if not rt.heading_level:
                if run.strikethrough:
                    text = f"~~{text}~~"
                if run.bold and run.italic:
                    text = f"***{text}***"
                elif run.bold:
                    text = f"**{text}**"
                elif run.italic:
                    text = f"*{text}*"
                if run.underline and not run.hyperlink_url:
                    text = f"*{text}*"

            if run.hyperlink_url:
                text = f"[{run.text}]({run.hyperlink_url})"

            if not rt.heading_level:
                if run.superscript:
                    text = f"<sup>{text}</sup>"
                if run.subscript:
                    text = f"<sub>{text}</sub>"

            parts.append(text)

        result = "".join(parts)

        # Apply heading prefix
        if rt.heading_level:
            prefix = "#" * rt.heading_level
            result = f"{prefix} {result}"
        elif rt.list_type:
            indent = "   " * rt.indent_level
            if rt.list_type == "ordered":
                num = ordered_number if ordered_number > 0 else 1
                marker = f"{num}."
            else:
                marker = "-"
            result = f"{indent}{marker} {result}"
        elif rt.indent_level > 0:
            indent = "   " * rt.indent_level
            result = f"{indent}- {result}"

        return result

    def _render_image(self, img: ImageElement) -> str:
        """Render image reference in Markdown."""
        alt = img.alt_text or img.filename or "image"
        if img.data:
            # Image will be saved to images/ subdirectory
            filename = _sanitize_filename(
                img.filename or f"image.{img.format or 'bin'}"
            )
            return f"![{alt}](./images/{filename})"
        return f"![{alt}]({img.filename})"

    def _render_table(self, table: TableElement) -> str:
        """Render table in Markdown."""
        if not table.rows:
            return ""

        lines: list[str] = []

        for i, row in enumerate(table.rows):
            cells = []
            for cell_elements in row:
                # Render cell contents
                cell_text = " ".join(
                    self._render_element(e).strip()
                    for e in cell_elements
                    if self._render_element(e).strip()
                )
                cells.append(cell_text or " ")

            lines.append("| " + " | ".join(cells) + " |")

            # Add header separator after first row
            if i == 0:
                lines.append("| " + " | ".join("---" for _ in cells) + " |")

        return "\n".join(lines)

    def _render_embedded_file(self, ef: EmbeddedFile) -> str:
        """Render embedded file reference in Markdown."""
        name = ef.filename or "attachment"
        if ef.data:
            filename = _sanitize_filename(name)
            return f"[{name}](./attachments/{filename})"
        return f"[{name}]"

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
    # Remove or replace problematic characters
    sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    # Collapse multiple underscores/spaces
    sanitized = re.sub(r"[_\s]+", " ", sanitized).strip()
    # Limit length
    if len(sanitized) > 200:
        sanitized = sanitized[:200]
    return sanitized or "unnamed"


def _page_filename(title: str, seen: dict[str, int]) -> str:
    """Generate a unique .md filename for a page title."""
    base = _sanitize_filename(title or "Untitled")
    if not base.endswith(".md"):
        base = f"{base}.md"

    key = base.lower()
    if key in seen:
        seen[key] += 1
        stem = base[:-3]  # Remove .md
        base = f"{stem} ({seen[key]}).md"
    else:
        seen[key] = 1

    return base
