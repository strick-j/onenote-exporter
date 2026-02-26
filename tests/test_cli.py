"""Tests for onenote_export.cli module."""

from onenote_export.cli import _deduplicate_sections


class TestDeduplicateSections:
    """Tests for _deduplicate_sections."""

    def test_no_duplicates(self, tmp_path):
        files = [
            tmp_path / "Notes.one",
            tmp_path / "Tasks.one",
        ]
        for f in files:
            f.touch()
        result = _deduplicate_sections(files)
        assert len(result) == 2

    def test_keeps_latest_version(self, tmp_path):
        old = tmp_path / "ADI (On 10-3-22).one"
        new = tmp_path / "ADI (On 2-25-26).one"
        old.touch()
        new.touch()
        result = _deduplicate_sections([old, new])
        assert len(result) == 1
        assert result[0] == new

    def test_dotone_date_pattern(self, tmp_path):
        old = tmp_path / "ADP.one (On 8-24-22).one"
        new = tmp_path / "ADP.one (On 8-24-25).one"
        old.touch()
        new.touch()
        result = _deduplicate_sections([old, new])
        assert len(result) == 1
        assert result[0] == new

    def test_undated_file_kept_when_no_dated_version(self, tmp_path):
        f = tmp_path / "Notes.one"
        f.touch()
        result = _deduplicate_sections([f])
        assert len(result) == 1
        assert result[0] == f

    def test_dated_wins_over_undated(self, tmp_path):
        undated = tmp_path / "Notes.one"
        dated = tmp_path / "Notes (On 2-25-26).one"
        undated.touch()
        dated.touch()
        result = _deduplicate_sections([undated, dated])
        assert len(result) == 1
        assert result[0] == dated

    def test_empty_list(self):
        assert _deduplicate_sections([]) == []

    def test_two_digit_year_normalization(self, tmp_path):
        old = tmp_path / "Test (On 1-1-99).one"  # 1999
        new = tmp_path / "Test (On 1-1-24).one"  # 2024
        old.touch()
        new.touch()
        result = _deduplicate_sections([old, new])
        assert len(result) == 1
        assert result[0] == new
