"""Tests for onenote_export.utils module."""

from pathlib import Path


from onenote_export.utils import (
    discover_one_files,
    notebook_name_from_dir,
    section_name_from_filename,
)


class TestSectionNameFromFilename:
    """Tests for section_name_from_filename."""

    def test_simple_filename(self):
        assert section_name_from_filename("Notes.one") == "Notes"

    def test_filename_with_date_suffix(self):
        assert section_name_from_filename("ADI (On 2-25-26).one") == "ADI"

    def test_filename_with_dotone_and_date(self):
        assert section_name_from_filename("ADP.one (On 8-24-25).one") == "ADP"

    def test_filename_with_spaces(self):
        assert section_name_from_filename("Meeting Notes.one") == "Meeting Notes"

    def test_filename_with_spaces_and_date(self):
        assert (
            section_name_from_filename("Meeting Notes (On 10-3-22).one")
            == "Meeting Notes"
        )

    def test_empty_filename_returns_untitled(self):
        assert section_name_from_filename(".one") == "Untitled"

    def test_only_date_suffix(self):
        assert section_name_from_filename("(On 1-1-25).one") == "Untitled"

    def test_multi_digit_date(self):
        assert section_name_from_filename("BMS (On 12-31-99).one") == "BMS"


class TestNotebookNameFromDir:
    """Tests for notebook_name_from_dir."""

    def test_named_directory(self):
        assert notebook_name_from_dir(Path("/some/path/My Notebook")) == "My Notebook"

    def test_root_directory(self):
        # Path("/").name is "" but the function returns "Untitled" for empty names
        assert notebook_name_from_dir(Path("/")) == "Untitled"


class TestDiscoverOneFiles:
    """Tests for discover_one_files."""

    def test_finds_one_files(self, tmp_path):
        (tmp_path / "section1.one").touch()
        (tmp_path / "section2.one").touch()
        result = discover_one_files(tmp_path)
        assert len(result) == 2

    def test_excludes_onetoc2_files(self, tmp_path):
        (tmp_path / "section.one").touch()
        (tmp_path / "Open Notebook.onetoc2").touch()
        result = discover_one_files(tmp_path)
        assert len(result) == 1
        assert result[0].name == "section.one"

    def test_finds_files_recursively(self, tmp_path):
        sub = tmp_path / "notebook"
        sub.mkdir()
        (sub / "section.one").touch()
        result = discover_one_files(tmp_path)
        assert len(result) == 1

    def test_empty_directory(self, tmp_path):
        result = discover_one_files(tmp_path)
        assert result == []

    def test_returns_sorted(self, tmp_path):
        (tmp_path / "b_section.one").touch()
        (tmp_path / "a_section.one").touch()
        result = discover_one_files(tmp_path)
        assert result[0].name == "a_section.one"
        assert result[1].name == "b_section.one"
