"""Integration tests using real .one files from test_data/.

Parses actual OneNote files through the full pipeline
(parse → extract → convert) and verifies table export.
"""

import re
from pathlib import Path

import pytest

from onenote_export.parser.one_store import OneStoreParser
from onenote_export.parser.content_extractor import extract_section
from onenote_export.converter.markdown import MarkdownConverter
from onenote_export.model.content import ContentElement, TableElement, RichText


def _cell_texts(row: list[list[ContentElement]]) -> list[str]:
    """Extract the concatenated text from each cell in a row."""
    texts = []
    for cell in row:
        parts = []
        for elem in cell:
            if isinstance(elem, RichText):
                parts.extend(run.text for run in elem.runs)
        texts.append(" ".join(parts) if parts else "")
    return texts


TEST_DATA = Path(__file__).parent / "test_data"
NOTEBOOK_DIR = TEST_DATA / "Example-NoteBook-1"

# The file with the table is the latest version of Section 2
TABLE_FILE = NOTEBOOK_DIR / "Example-Section-2 (On 2-25-26).one"


def _find_latest_section2():
    """Find the latest version of Example-Section-2."""
    candidates = sorted(
        NOTEBOOK_DIR.glob("Example-Section-2*.one"),
        key=lambda p: p.stat().st_mtime,
    )
    return candidates[-1] if candidates else None


@pytest.fixture(scope="module")
def parsed_section():
    """Parse the .one file containing the table."""
    path = TABLE_FILE if TABLE_FILE.exists() else _find_latest_section2()
    assert path is not None, "No Example-Section-2 .one file found in test_data"
    parser = OneStoreParser(path)
    return parser.parse()


@pytest.fixture(scope="module")
def extracted_section(parsed_section):
    """Extract structured content from the parsed section."""
    return extract_section(parsed_section)


class TestParserFindsTableObjects:
    """Verify the parser returns table-related objects from the .one file."""

    def test_section_has_pages(self, parsed_section):
        assert len(parsed_section.pages) >= 1

    def test_note3_exists(self, parsed_section):
        titles = [p.title for p in parsed_section.pages]
        assert "Note 3" in titles, f"Expected 'Note 3' in {titles}"

    def test_note3_has_table_node(self, parsed_section):
        page = next(p for p in parsed_section.pages if p.title == "Note 3")
        table_nodes = [o for o in page.objects if o.obj_type == "jcidTableNode"]
        assert len(table_nodes) >= 1, "No jcidTableNode found on Note 3"

    def test_note3_has_table_row_nodes(self, parsed_section):
        page = next(p for p in parsed_section.pages if p.title == "Note 3")
        row_nodes = [o for o in page.objects if o.obj_type == "jcidTableRowNode"]
        assert len(row_nodes) == 4, f"Expected 4 rows, got {len(row_nodes)}"

    def test_note3_has_table_cell_nodes(self, parsed_section):
        page = next(p for p in parsed_section.pages if p.title == "Note 3")
        cell_nodes = [o for o in page.objects if o.obj_type == "jcidTableCellNode"]
        assert len(cell_nodes) == 16, f"Expected 16 cells, got {len(cell_nodes)}"

    def test_table_node_has_row_and_column_count(self, parsed_section):
        page = next(p for p in parsed_section.pages if p.title == "Note 3")
        table_node = next(o for o in page.objects if o.obj_type == "jcidTableNode")
        assert "RowCount" in table_node.properties
        assert "ColumnCount" in table_node.properties


class TestContentExtractorHandlesTables:
    """Verify content extraction produces TableElement with rows populated."""

    def test_note3_has_table_element(self, extracted_section):
        page = next(p for p in extracted_section.pages if p.title == "Note 3")
        tables = [e for e in page.elements if isinstance(e, TableElement)]
        assert len(tables) >= 1, (
            f"No TableElement found. Element types: "
            f"{[type(e).__name__ for e in page.elements]}"
        )

    def test_table_has_4_rows(self, extracted_section):
        page = next(p for p in extracted_section.pages if p.title == "Note 3")
        table = next(e for e in page.elements if isinstance(e, TableElement))
        assert len(table.rows) == 4, (
            f"Expected 4 rows, got {len(table.rows)}. "
            f"Table rows are empty — cell content is not being linked."
        )

    def test_each_row_has_4_cells(self, extracted_section):
        page = next(p for p in extracted_section.pages if p.title == "Note 3")
        table = next(e for e in page.elements if isinstance(e, TableElement))
        for i, row in enumerate(table.rows):
            assert len(row) == 4, f"Row {i} has {len(row)} cells, expected 4"

    def test_header_row_has_column_1_through_4_in_order(self, extracted_section):
        """First row should be headers: Column 1, Column 2, Column 3, Column 4."""
        page = next(p for p in extracted_section.pages if p.title == "Note 3")
        table = next(e for e in page.elements if isinstance(e, TableElement))
        if not table.rows:
            pytest.skip("Table rows not populated")

        header_texts = _cell_texts(table.rows[0])
        for i in range(1, 5):
            assert f"Column {i}" in header_texts[i - 1], (
                f"Expected 'Column {i}' in cell {i - 1}, got header row: {header_texts}"
            )

    def test_body_row_cells_in_correct_order(self, extracted_section):
        """Body rows should read left-to-right: Row N Column 1 .. Column 4."""
        page = next(p for p in extracted_section.pages if p.title == "Note 3")
        table = next(e for e in page.elements if isinstance(e, TableElement))
        if len(table.rows) < 2:
            pytest.skip("Table rows not populated")

        # All body rows should have correct content in order
        for row_num in range(1, 4):
            row_texts = _cell_texts(table.rows[row_num])
            for col in range(1, 5):
                assert f"Row {row_num} Column {col}" in row_texts[col - 1], (
                    f"Expected 'Row {row_num} Column {col}' in cell "
                    f"{col - 1}, got row: {row_texts}"
                )

    def test_rows_in_correct_order(self, extracted_section):
        """Rows should be top-to-bottom: header, Row 1, Row 2, Row 3."""
        page = next(p for p in extracted_section.pages if p.title == "Note 3")
        table = next(e for e in page.elements if isinstance(e, TableElement))
        if len(table.rows) < 4:
            pytest.skip("Table rows not populated")

        row0 = " ".join(_cell_texts(table.rows[0]))
        row3 = " ".join(_cell_texts(table.rows[3]))

        assert "Column 1" in row0 and "Row" not in row0, (
            f"First row should be headers, got: {row0}"
        )
        assert "Row 3" in row3, f"Last row should be Row 3, got: {row3}"

    def test_no_table_text_rendered_as_standalone(self, extracted_section):
        """Table cell content should not appear as standalone RichText."""
        page = next(p for p in extracted_section.pages if p.title == "Note 3")
        standalone_cell_texts = []
        for elem in page.elements:
            if isinstance(elem, RichText):
                for run in elem.runs:
                    if "Row" in run.text and "Column" in run.text:
                        standalone_cell_texts.append(run.text)

        assert not standalone_cell_texts, (
            f"Table cell text leaked as standalone: {standalone_cell_texts}"
        )


class TestMarkdownTableOutput:
    """Verify the markdown converter produces proper table syntax."""

    def test_markdown_contains_table_syntax(self, extracted_section, tmp_path):
        converter = MarkdownConverter(tmp_path)
        page = next(p for p in extracted_section.pages if p.title == "Note 3")
        md = converter.render_page(page)

        # Check for markdown table pipe syntax
        assert "|" in md, f"No table pipe syntax found in markdown:\n{md}"

    def test_markdown_table_has_header_separator(self, extracted_section, tmp_path):
        converter = MarkdownConverter(tmp_path)
        page = next(p for p in extracted_section.pages if p.title == "Note 3")
        md = converter.render_page(page)

        assert "| ---" in md, f"No header separator found in markdown:\n{md}"

    def test_markdown_table_has_all_rows(self, extracted_section, tmp_path):
        converter = MarkdownConverter(tmp_path)
        page = next(p for p in extracted_section.pages if p.title == "Note 3")
        md = converter.render_page(page)

        # Count table rows (lines starting with |, excluding separator)
        table_rows = [
            line
            for line in md.splitlines()
            if line.strip().startswith("|") and "---" not in line
        ]
        assert len(table_rows) == 4, (
            f"Expected 4 table rows, got {len(table_rows)}:\n{md}"
        )

    def test_full_section_write(self, extracted_section, tmp_path):
        """Write the full section and verify output files exist."""
        converter = MarkdownConverter(tmp_path)
        created = converter.convert_section(extracted_section)
        assert len(created) >= 1, "No files created"

        # Find the Note 3 markdown file
        note3_files = [f for f in created if "Note 3" in f.name and f.suffix == ".md"]
        assert len(note3_files) == 1, (
            f"Expected 1 Note 3 file, got: {[f.name for f in created]}"
        )

        content = note3_files[0].read_text()
        assert "# Note 3" in content


class TestHeadingPreservation:
    """Verify OneNote heading styles are preserved in the content model."""

    def test_note1_has_h2_heading(self, extracted_section):
        """Note 1 should have 'What is Lorem Ipsum?' as heading level 2."""
        page = next(p for p in extracted_section.pages if p.title == "Note 1")
        h2_elements = [
            e for e in page.elements if isinstance(e, RichText) and e.heading_level == 2
        ]
        h2_texts = [r.text for e in h2_elements for r in e.runs]
        assert any("What is Lorem Ipsum" in t for t in h2_texts), (
            f"Expected h2 'What is Lorem Ipsum?', got h2 texts: {h2_texts}"
        )

    def test_note1_has_h1_heading(self, extracted_section):
        """Note 1 should have 'Why do we use it?' as heading level 1."""
        page = next(p for p in extracted_section.pages if p.title == "Note 1")
        h1_elements = [
            e for e in page.elements if isinstance(e, RichText) and e.heading_level == 1
        ]
        h1_texts = [r.text for e in h1_elements for r in e.runs]
        assert any("Why do we use it" in t for t in h1_texts), (
            f"Expected h1 'Why do we use it?', got h1 texts: {h1_texts}"
        )

    def test_note1_has_h3_heading(self, extracted_section):
        """Note 1 should have 'Where can I get some?' as heading level 3."""
        page = next(p for p in extracted_section.pages if p.title == "Note 1")
        h3_elements = [
            e for e in page.elements if isinstance(e, RichText) and e.heading_level == 3
        ]
        h3_texts = [r.text for e in h3_elements for r in e.runs]
        assert any("Where can I get some" in t for t in h3_texts), (
            f"Expected h3 'Where can I get some?', got h3 texts: {h3_texts}"
        )

    def test_note1_has_h4_heading(self, extracted_section):
        """Note 1 should have 'Where does it come from?' as heading level 4."""
        page = next(p for p in extracted_section.pages if p.title == "Note 1")
        h4_elements = [
            e for e in page.elements if isinstance(e, RichText) and e.heading_level == 4
        ]
        h4_texts = [r.text for e in h4_elements for r in e.runs]
        assert any("Where does it come from" in t for t in h4_texts), (
            f"Expected h4 'Where does it come from?', got h4 texts: {h4_texts}"
        )

    def test_note3_inserted_image_is_heading(self, extracted_section):
        """Note 3 'Inserted Image:' should be a heading."""
        page = next(p for p in extracted_section.pages if p.title == "Note 3")
        heading_elements = [
            e for e in page.elements if isinstance(e, RichText) and e.heading_level > 0
        ]
        heading_texts = [r.text for e in heading_elements for r in e.runs]
        assert any("Inserted Image" in t for t in heading_texts), (
            f"Expected heading 'Inserted Image:', got: {heading_texts}"
        )

    def test_normal_text_has_no_heading(self, extracted_section):
        """Body paragraphs should have heading_level=0."""
        page = next(p for p in extracted_section.pages if p.title == "Note 1")
        body_elements = [
            e
            for e in page.elements
            if isinstance(e, RichText) and e.heading_level == 0 and not e.is_title
        ]
        # Should have at least the body paragraphs
        assert len(body_elements) >= 4, (
            f"Expected at least 4 body paragraphs, got {len(body_elements)}"
        )

    def test_markdown_renders_headings(self, extracted_section, tmp_path):
        """Headings should render as # in Markdown."""
        converter = MarkdownConverter(tmp_path)
        page = next(p for p in extracted_section.pages if p.title == "Note 1")
        md = converter.render_page(page)
        assert "## What is Lorem Ipsum?" in md, f"h2 heading not rendered:\n{md}"
        assert "### Where can I get some?" in md, f"h3 heading not rendered:\n{md}"

    def test_note3_markdown_has_heading_for_inserted_image(
        self, extracted_section, tmp_path
    ):
        """Note 3 should render 'Inserted Image:' as a heading in Markdown."""
        converter = MarkdownConverter(tmp_path)
        page = next(p for p in extracted_section.pages if p.title == "Note 3")
        md = converter.render_page(page)
        assert "# Inserted Image:" in md, (
            f"Heading not rendered for 'Inserted Image:':\n{md}"
        )


class TestListPreservation:
    """Verify OneNote lists are preserved with correct type and nesting."""

    def test_note2_has_unordered_list(self, extracted_section):
        """Note 2 should have unordered (bullet) list items."""
        page = next(p for p in extracted_section.pages if p.title == "Note 2")
        bullets = [
            e
            for e in page.elements
            if isinstance(e, RichText)
            and e.list_type == "unordered"
            and e.heading_level == 0
        ]
        assert len(bullets) == 8, f"Expected 8 bullet items, got {len(bullets)}"

    def test_note2_has_ordered_list(self, extracted_section):
        """Note 2 should have ordered (numbered) list items."""
        page = next(p for p in extracted_section.pages if p.title == "Note 2")
        numbered = [
            e
            for e in page.elements
            if isinstance(e, RichText)
            and e.list_type == "ordered"
            and e.heading_level == 0
        ]
        assert len(numbered) == 8, f"Expected 8 numbered items, got {len(numbered)}"

    def test_bullet_list_has_4_nesting_levels(self, extracted_section):
        """Bullet list should have items at indent levels 0, 1, 2, 3."""
        page = next(p for p in extracted_section.pages if p.title == "Note 2")
        levels = {
            e.indent_level
            for e in page.elements
            if isinstance(e, RichText) and e.list_type == "unordered"
        }
        assert levels == {0, 1, 2, 3}, (
            f"Expected indent levels {{0, 1, 2, 3}}, got {levels}"
        )

    def test_numbered_list_has_4_nesting_levels(self, extracted_section):
        """Numbered list should have items at indent levels 0, 1, 2, 3."""
        page = next(p for p in extracted_section.pages if p.title == "Note 2")
        levels = {
            e.indent_level
            for e in page.elements
            if isinstance(e, RichText) and e.list_type == "ordered"
        }
        assert levels == {0, 1, 2, 3}, (
            f"Expected indent levels {{0, 1, 2, 3}}, got {levels}"
        )

    def test_bullet_top_level_items(self, extracted_section):
        """Top-level bullet items should be at indent 0."""
        page = next(p for p in extracted_section.pages if p.title == "Note 2")
        top_bullets = [
            e
            for e in page.elements
            if isinstance(e, RichText)
            and e.list_type == "unordered"
            and e.indent_level == 0
        ]
        texts = [r.text for e in top_bullets for r in e.runs]
        assert any("Lorem ipsum" in t for t in texts)
        assert any("Proin ut dui" in t for t in texts)

    def test_numbered_top_level_items(self, extracted_section):
        """Top-level numbered items should be at indent 0."""
        page = next(p for p in extracted_section.pages if p.title == "Note 2")
        top_numbered = [
            e
            for e in page.elements
            if isinstance(e, RichText)
            and e.list_type == "ordered"
            and e.indent_level == 0
        ]
        texts = [r.text for e in top_numbered for r in e.runs]
        assert any("Lorem ipsum" in t for t in texts)
        assert any("Proin ut dui" in t for t in texts)

    def test_markdown_bullet_list_syntax(self, extracted_section, tmp_path):
        """Bullet list items should render with '-' prefix."""
        converter = MarkdownConverter(tmp_path)
        page = next(p for p in extracted_section.pages if p.title == "Note 2")
        md = converter.render_page(page)
        bullet_lines = [
            line for line in md.splitlines() if line.strip().startswith("- ")
        ]
        assert len(bullet_lines) == 8, (
            f"Expected 8 bullet lines, got {len(bullet_lines)}"
        )

    def test_markdown_numbered_list_syntax(self, extracted_section, tmp_path):
        """Numbered list items should render with incrementing numbers."""
        converter = MarkdownConverter(tmp_path)
        page = next(p for p in extracted_section.pages if p.title == "Note 2")
        md = converter.render_page(page)
        numbered_lines = [
            line for line in md.splitlines() if re.match(r"\s*\d+\.", line)
        ]
        assert len(numbered_lines) == 8, (
            f"Expected 8 numbered lines, got {len(numbered_lines)}"
        )
        # Top-level items should increment
        top_level = [line for line in numbered_lines if not line.startswith(" ")]
        assert top_level[0].startswith("1.")
        assert top_level[1].startswith("2.")

    def test_markdown_nested_bullet_indentation(self, extracted_section, tmp_path):
        """Nested bullet items should be indented with 3 spaces per level."""
        converter = MarkdownConverter(tmp_path)
        page = next(p for p in extracted_section.pages if p.title == "Note 2")
        md = converter.render_page(page)
        # Find a level-3 bullet item
        deep_bullets = [
            line
            for line in md.splitlines()
            if line.startswith("         - ")  # 9 spaces = 3 levels
        ]
        assert len(deep_bullets) >= 1, "No level-3 bullet items found in markdown"

    def test_no_list_text_as_plain(self, extracted_section):
        """List item text should not appear as non-list RichText."""
        page = next(p for p in extracted_section.pages if p.title == "Note 2")
        # Collect text from list items
        list_texts = set()
        for elem in page.elements:
            if isinstance(elem, RichText) and elem.list_type:
                for run in elem.runs:
                    if "Lorem ipsum" in run.text or "Donec ornare" in run.text:
                        list_texts.add(run.text.strip())

        # Check no plain elements have the same text
        plain_dupes = []
        for elem in page.elements:
            if isinstance(elem, RichText) and not elem.list_type:
                for run in elem.runs:
                    if run.text.strip() in list_texts:
                        plain_dupes.append(run.text)

        assert not plain_dupes, f"List text leaked as plain text: {plain_dupes}"


class TestAllTestDataFiles:
    """Verify all .one files in test_data can be parsed without errors."""

    @pytest.fixture(scope="class")
    def all_one_files(self):
        return sorted(TEST_DATA.rglob("*.one"))

    def test_test_data_has_files(self, all_one_files):
        assert len(all_one_files) >= 1, "No .one files in test_data"

    @pytest.mark.parametrize(
        "filename",
        [
            "Example-Section-1 (On 2-25-26 - 3).one",
            "Example-Section-2 (On 2-25-26).one",
        ],
    )
    def test_parse_without_errors(self, filename):
        path = NOTEBOOK_DIR / filename
        if not path.exists():
            pytest.skip(f"{filename} not found")
        parser = OneStoreParser(path)
        section = parser.parse()
        assert len(section.pages) >= 1

    @pytest.mark.parametrize(
        "filename",
        [
            "Example-Section-1 (On 2-25-26 - 3).one",
            "Example-Section-2 (On 2-25-26).one",
        ],
    )
    def test_extract_without_errors(self, filename):
        path = NOTEBOOK_DIR / filename
        if not path.exists():
            pytest.skip(f"{filename} not found")
        parser = OneStoreParser(path)
        parsed = parser.parse()
        section = extract_section(parsed)
        assert len(section.pages) >= 1
        for page in section.pages:
            assert page.title
