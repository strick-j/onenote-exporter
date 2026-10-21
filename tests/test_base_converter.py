"""Tests for onenote_export.converter.base module."""

import tempfile

from onenote_export.converter.base import (
    BaseConverter,
    _page_filename,
    _sanitize_filename,
)
from onenote_export.model.content import (
    EmbeddedFile,
    ImageElement,
)
from onenote_export.model.notebook import Notebook
from onenote_export.model.page import Page
from onenote_export.model.section import Section


class TestSanitizeFilename:
    """Tests for _sanitize_filename."""

    def test_clean_name(self):
        assert _sanitize_filename("hello") == "hello"

    def test_removes_special_chars(self):
        result = _sanitize_filename('file<>:"/\\|?*name')
        assert "<" not in result
        assert ">" not in result
        assert ":" not in result
        assert '"' not in result

    def test_collapses_underscores(self):
        result = _sanitize_filename("a___b")
        assert result == "a b"

    def test_empty_returns_unnamed(self):
        assert _sanitize_filename("") == "unnamed"

    def test_truncates_long_names(self):
        long_name = "a" * 300
        result = _sanitize_filename(long_name)
        assert len(result) <= 200


class TestPageFilename:
    """Tests for _page_filename with various extensions."""

    def test_simple_title_md(self):
        seen: dict[str, int] = {}
        result = _page_filename("My Page", seen, ".md")
        assert result == "My Page.md"

    def test_simple_title_html(self):
        seen: dict[str, int] = {}
        result = _page_filename("My Page", seen, ".html")
        assert result == "My Page.html"

    def test_duplicate_title_gets_numbered(self):
        seen: dict[str, int] = {}
        first = _page_filename("Notes", seen)
        second = _page_filename("Notes", seen)
        assert first == "Notes.md"
        assert second == "Notes (2).md"

    def test_duplicate_title_html(self):
        seen: dict[str, int] = {}
        first = _page_filename("Notes", seen, ".html")
        second = _page_filename("Notes", seen, ".html")
        assert first == "Notes.html"
        assert second == "Notes (2).html"

    def test_untitled_page(self):
        seen: dict[str, int] = {}
        result = _page_filename("", seen)
        assert result == "Untitled.md"

    def test_default_extension_is_md(self):
        seen: dict[str, int] = {}
        result = _page_filename("Test", seen)
        assert result.endswith(".md")


class _StubConverter(BaseConverter):
    """Minimal subclass for testing base class file I/O."""

    FILE_EXTENSION = ".txt"

    def render_page(self, page: Page) -> str:
        return f"TITLE:{page.title}"


class TestBaseConverterConvertSection:
    """Tests for BaseConverter.convert_section via a stub subclass."""

    def test_creates_files_with_correct_extension(self, tmp_path):
        converter = _StubConverter(tmp_path)
        section = Section(
            name="Test Section",
            pages=[
                Page(title="Page 1"),
                Page(title="Page 2"),
            ],
        )
        created = converter.convert_section(section)
        assert len(created) == 2
        assert (tmp_path / "Test Section" / "Page 1.txt").exists()
        assert (tmp_path / "Test Section" / "Page 2.txt").exists()

    def test_file_content_from_render_page(self, tmp_path):
        converter = _StubConverter(tmp_path)
        section = Section(
            name="Test",
            pages=[Page(title="Hello")],
        )
        converter.convert_section(section)
        content = (tmp_path / "Test" / "Hello.txt").read_text()
        assert content == "TITLE:Hello"


class TestBaseConverterConvertNotebook:
    """Tests for BaseConverter.convert_notebook."""

    def test_creates_nested_dirs(self, tmp_path):
        converter = _StubConverter(tmp_path)
        notebook = Notebook(
            name="My Notebook",
            sections=[
                Section(name="Section A", pages=[Page(title="Page 1")]),
                Section(name="Section B", pages=[Page(title="Page 2")]),
            ],
        )
        converter.convert_notebook(notebook)
        assert (tmp_path / "My Notebook" / "Section A" / "Page 1.txt").exists()
        assert (tmp_path / "My Notebook" / "Section B" / "Page 2.txt").exists()

    def test_preserves_relative_path(self, tmp_path):
        converter = _StubConverter(tmp_path)
        notebook = Notebook(
            name="Cloud Accounts",
            relative_path="RIPTiDE A4Me/Cloud Accounts",
            sections=[Section(name="GCP", pages=[Page(title="Project Notes")])],
        )

        converter.convert_notebook(notebook)

        assert (
            tmp_path / "RIPTiDE A4Me" / "Cloud Accounts" / "GCP" / "Project Notes.txt"
        ).exists()


class TestBaseConverterRenderPageAbstract:
    """Test that BaseConverter.render_page raises NotImplementedError."""

    def test_raises_not_implemented(self):
        converter = BaseConverter(tempfile.gettempdir())
        import pytest

        with pytest.raises(NotImplementedError):
            converter.render_page(Page(title="Test"))


class TestBaseConverterWriteImages:
    """Tests for _write_images."""

    def test_writes_images(self, tmp_path):
        converter = _StubConverter(tmp_path)
        png_header = b"\x89PNG\r\n\x1a\n" + b"\x00" * 100
        section = Section(
            name="Test",
            pages=[
                Page(
                    title="With Image",
                    elements=[
                        ImageElement(data=png_header, filename="pic.png", format="png")
                    ],
                ),
            ],
        )
        converter.convert_section(section)
        assert (tmp_path / "Test" / "images" / "pic.png").exists()


class TestBaseConverterWriteEmbeddedFiles:
    """Tests for _write_embedded_files."""

    def test_writes_attachments(self, tmp_path):
        converter = _StubConverter(tmp_path)
        section = Section(
            name="Test",
            pages=[
                Page(
                    title="With File",
                    elements=[EmbeddedFile(data=b"pdf content", filename="doc.pdf")],
                ),
            ],
        )
        converter.convert_section(section)
        assert (tmp_path / "Test" / "attachments" / "doc.pdf").exists()
