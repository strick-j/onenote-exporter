"""Tests for onenote_export.converter.markdown module."""

import tempfile

from onenote_export.converter.markdown import (
    MarkdownConverter,
    _page_filename,
    _sanitize_filename,
)
from onenote_export.model.content import (
    EmbeddedFile,
    ImageElement,
    RichText,
    TableElement,
    TextRun,
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
    """Tests for _page_filename."""

    def test_simple_title(self):
        seen = {}
        result = _page_filename("My Page", seen)
        assert result == "My Page.md"

    def test_duplicate_title_gets_numbered(self):
        seen = {}
        first = _page_filename("Notes", seen)
        second = _page_filename("Notes", seen)
        assert first == "Notes.md"
        assert second == "Notes (2).md"

    def test_untitled_page(self):
        seen = {}
        result = _page_filename("", seen)
        assert result == "Untitled.md"


class TestMarkdownConverterRenderPage:
    """Tests for MarkdownConverter.render_page."""

    def setup_method(self):
        self.converter = MarkdownConverter(tempfile.gettempdir() + "/test_output")

    def test_page_with_title(self):
        page = Page(title="My Title")
        result = self.converter.render_page(page)
        assert result.startswith("# My Title")

    def test_page_with_author(self):
        page = Page(title="Test", author="John Doe")
        result = self.converter.render_page(page)
        assert "*Author: John Doe*" in result
        assert "---" in result

    def test_page_with_plain_text(self):
        page = Page(
            title="Test",
            elements=[RichText(runs=[TextRun(text="Hello world")])],
        )
        result = self.converter.render_page(page)
        assert "Hello world" in result

    def test_page_with_bold_text(self):
        page = Page(
            title="Test",
            elements=[RichText(runs=[TextRun(text="important", bold=True)])],
        )
        result = self.converter.render_page(page)
        assert "**important**" in result

    def test_page_with_italic_text(self):
        page = Page(
            title="Test",
            elements=[RichText(runs=[TextRun(text="emphasis", italic=True)])],
        )
        result = self.converter.render_page(page)
        assert "*emphasis*" in result

    def test_page_with_bold_italic_text(self):
        page = Page(
            title="Test",
            elements=[RichText(runs=[TextRun(text="strong", bold=True, italic=True)])],
        )
        result = self.converter.render_page(page)
        assert "***strong***" in result

    def test_page_with_strikethrough(self):
        page = Page(
            title="Test",
            elements=[RichText(runs=[TextRun(text="removed", strikethrough=True)])],
        )
        result = self.converter.render_page(page)
        assert "~~removed~~" in result

    def test_page_with_hyperlink(self):
        page = Page(
            title="Test",
            elements=[
                RichText(
                    runs=[
                        TextRun(text="click here", hyperlink_url="https://example.com")
                    ]
                )
            ],
        )
        result = self.converter.render_page(page)
        assert "[click here](https://example.com)" in result

    def test_page_with_superscript(self):
        page = Page(
            title="Test",
            elements=[RichText(runs=[TextRun(text="2", superscript=True)])],
        )
        result = self.converter.render_page(page)
        assert "<sup>2</sup>" in result

    def test_page_with_subscript(self):
        page = Page(
            title="Test",
            elements=[RichText(runs=[TextRun(text="2", subscript=True)])],
        )
        result = self.converter.render_page(page)
        assert "<sub>2</sub>" in result

    def test_page_with_indented_text(self):
        page = Page(
            title="Test",
            elements=[RichText(runs=[TextRun(text="list item")], indent_level=1)],
        )
        result = self.converter.render_page(page)
        assert "- list item" in result

    def test_page_with_nested_indent(self):
        page = Page(
            title="Test",
            elements=[RichText(runs=[TextRun(text="nested")], indent_level=3)],
        )
        result = self.converter.render_page(page)
        assert "    - nested" in result

    def test_page_with_image(self):
        page = Page(
            title="Test",
            elements=[
                ImageElement(data=b"\x89PNG", filename="screenshot.png", format="png")
            ],
        )
        result = self.converter.render_page(page)
        assert "![screenshot.png](./images/screenshot.png)" in result

    def test_page_with_image_no_data(self):
        page = Page(
            title="Test",
            elements=[ImageElement(filename="remote.png")],
        )
        result = self.converter.render_page(page)
        assert "![remote.png](remote.png)" in result

    def test_page_with_embedded_file(self):
        page = Page(
            title="Test",
            elements=[EmbeddedFile(data=b"content", filename="report.pdf")],
        )
        result = self.converter.render_page(page)
        assert "[report.pdf](./attachments/report.pdf)" in result

    def test_page_with_embedded_file_no_data(self):
        page = Page(
            title="Test",
            elements=[EmbeddedFile(filename="missing.pdf")],
        )
        result = self.converter.render_page(page)
        assert "[missing.pdf]" in result

    def test_empty_page(self):
        page = Page()
        result = self.converter.render_page(page)
        assert result.strip() == ""


class TestMarkdownConverterRenderTable:
    """Tests for table rendering."""

    def setup_method(self):
        self.converter = MarkdownConverter(tempfile.gettempdir() + "/test_output")

    def test_simple_table(self):
        table = TableElement(
            rows=[
                [
                    [RichText(runs=[TextRun(text="Header 1")])],
                    [RichText(runs=[TextRun(text="Header 2")])],
                ],
                [
                    [RichText(runs=[TextRun(text="Cell 1")])],
                    [RichText(runs=[TextRun(text="Cell 2")])],
                ],
            ]
        )
        result = self.converter._render_table(table)
        assert "| Header 1 | Header 2 |" in result
        assert "| --- | --- |" in result
        assert "| Cell 1 | Cell 2 |" in result

    def test_empty_table(self):
        table = TableElement(rows=[])
        result = self.converter._render_table(table)
        assert result == ""


class TestMarkdownConverterWriteFiles:
    """Tests for file writing operations."""

    def test_convert_section_creates_files(self, tmp_path):
        converter = MarkdownConverter(tmp_path)
        section = Section(
            name="Test Section",
            pages=[
                Page(
                    title="Page 1",
                    elements=[RichText(runs=[TextRun(text="Content 1")])],
                ),
                Page(
                    title="Page 2",
                    elements=[RichText(runs=[TextRun(text="Content 2")])],
                ),
            ],
        )
        created = converter.convert_section(section)
        assert len(created) == 2
        assert (tmp_path / "Test Section" / "Page 1.md").exists()
        assert (tmp_path / "Test Section" / "Page 2.md").exists()

    def test_convert_section_content(self, tmp_path):
        converter = MarkdownConverter(tmp_path)
        section = Section(
            name="Test",
            pages=[
                Page(title="Hello", elements=[RichText(runs=[TextRun(text="world")])])
            ],
        )
        converter.convert_section(section)
        content = (tmp_path / "Test" / "Hello.md").read_text()
        assert "# Hello" in content
        assert "world" in content

    def test_convert_notebook_creates_nested_dirs(self, tmp_path):
        converter = MarkdownConverter(tmp_path)
        notebook = Notebook(
            name="My Notebook",
            sections=[
                Section(name="Section A", pages=[Page(title="Page 1")]),
                Section(name="Section B", pages=[Page(title="Page 2")]),
            ],
        )
        converter.convert_notebook(notebook)
        assert (tmp_path / "My Notebook" / "Section A" / "Page 1.md").exists()
        assert (tmp_path / "My Notebook" / "Section B" / "Page 2.md").exists()

    def test_writes_images(self, tmp_path):
        converter = MarkdownConverter(tmp_path)
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

    def test_writes_attachments(self, tmp_path):
        converter = MarkdownConverter(tmp_path)
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
