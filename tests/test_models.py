"""Tests for onenote_export.model module."""

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


class TestTextRun:
    """Tests for TextRun dataclass."""

    def test_defaults(self):
        run = TextRun(text="hello")
        assert run.text == "hello"
        assert run.bold is False
        assert run.italic is False
        assert run.underline is False
        assert run.strikethrough is False
        assert run.superscript is False
        assert run.subscript is False
        assert run.font == ""
        assert run.font_size == 0
        assert run.hyperlink_url == ""

    def test_frozen(self):
        run = TextRun(text="hello")
        try:
            run.text = "world"
            assert False, "TextRun should be frozen"
        except AttributeError:
            pass

    def test_formatting_flags(self):
        run = TextRun(text="test", bold=True, italic=True, strikethrough=True)
        assert run.bold is True
        assert run.italic is True
        assert run.strikethrough is True


class TestRichText:
    """Tests for RichText dataclass."""

    def test_defaults(self):
        rt = RichText()
        assert rt.runs == []
        assert rt.indent_level == 0
        assert rt.is_title is False
        assert rt.alignment == ""

    def test_with_runs(self):
        runs = [TextRun(text="hello"), TextRun(text="world", bold=True)]
        rt = RichText(runs=runs, indent_level=1)
        assert len(rt.runs) == 2
        assert rt.indent_level == 1


class TestImageElement:
    """Tests for ImageElement dataclass."""

    def test_defaults(self):
        img = ImageElement()
        assert img.data == b""
        assert img.filename == ""
        assert img.format == ""

    def test_with_data(self):
        img = ImageElement(data=b"\x89PNG", filename="test.png", format="png")
        assert img.data == b"\x89PNG"
        assert img.format == "png"


class TestTableElement:
    """Tests for TableElement dataclass."""

    def test_defaults(self):
        table = TableElement()
        assert table.rows == []
        assert table.column_widths == []
        assert table.borders_visible is True


class TestEmbeddedFile:
    """Tests for EmbeddedFile dataclass."""

    def test_defaults(self):
        ef = EmbeddedFile()
        assert ef.data == b""
        assert ef.filename == ""
        assert ef.source_path == ""


class TestPage:
    """Tests for Page dataclass."""

    def test_defaults(self):
        page = Page()
        assert page.title == ""
        assert page.level == 0
        assert page.elements == []
        assert page.author == ""

    def test_with_elements(self):
        elements = [RichText(runs=[TextRun(text="hello")])]
        page = Page(title="Test Page", elements=elements)
        assert page.title == "Test Page"
        assert len(page.elements) == 1


class TestSection:
    """Tests for Section dataclass."""

    def test_defaults(self):
        section = Section()
        assert section.name == ""
        assert section.pages == []

    def test_with_pages(self):
        pages = [Page(title="Page 1"), Page(title="Page 2")]
        section = Section(name="My Section", pages=pages)
        assert len(section.pages) == 2


class TestNotebook:
    """Tests for Notebook dataclass."""

    def test_defaults(self):
        nb = Notebook()
        assert nb.name == ""
        assert nb.sections == []

    def test_with_sections(self):
        sections = [Section(name="S1"), Section(name="S2")]
        nb = Notebook(name="My Notebook", sections=sections)
        assert len(nb.sections) == 2
