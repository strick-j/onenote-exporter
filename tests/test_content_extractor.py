"""Tests for onenote_export.parser.content_extractor module."""

from onenote_export.parser.content_extractor import (
    _as_bool,
    _clean_text,
    _decode_text_value,
    _detect_image_format,
    _looks_garbled,
    _parse_font_size,
    _parse_int_prop,
)


class TestDecodeTextValue:
    """Tests for _decode_text_value."""

    def test_plain_string(self):
        assert _decode_text_value("Hello world") == "Hello world"

    def test_empty_string(self):
        assert _decode_text_value("") == ""

    def test_none_value(self):
        assert _decode_text_value(None) == ""

    def test_bytes_unicode(self):
        text = "Hello"
        raw = text.encode("utf-16-le")
        result = _decode_text_value(raw, encoding="unicode")
        assert result == "Hello"

    def test_bytes_ascii(self):
        raw = b"Hello\x00"
        result = _decode_text_value(raw, encoding="ascii")
        assert result == "Hello"

    def test_hex_string_ascii(self):
        raw = "48656c6c6f"  # "Hello" in hex
        result = _decode_text_value(raw, encoding="ascii")
        assert result == "Hello"

    def test_string_with_null_bytes(self):
        result = _decode_text_value("Hello\x00World")
        assert "\x00" not in result


class TestLooksGarbled:
    """Tests for _looks_garbled."""

    def test_normal_text(self):
        assert _looks_garbled("Hello world") is False

    def test_empty_text(self):
        assert _looks_garbled("") is False

    def test_short_text(self):
        assert _looks_garbled("ab") is False

    def test_garbled_text(self):
        # Simulate CJK characters from misinterpreted ASCII
        garbled = "\u4e48\u5f00\u53d1"
        assert _looks_garbled(garbled) is True


class TestCleanText:
    """Tests for _clean_text."""

    def test_removes_null_bytes(self):
        assert _clean_text("hello\x00world") == "helloworld"

    def test_strips_whitespace(self):
        assert _clean_text("  hello  ") == "hello"

    def test_empty_string(self):
        assert _clean_text("") == ""


class TestAsBool:
    """Tests for _as_bool."""

    def test_true_bool(self):
        assert _as_bool(True) is True

    def test_false_bool(self):
        assert _as_bool(False) is False

    def test_string_true(self):
        assert _as_bool("true") is True
        assert _as_bool("True") is True
        assert _as_bool("1") is True
        assert _as_bool("yes") is True

    def test_string_false(self):
        assert _as_bool("false") is False
        assert _as_bool("no") is False
        assert _as_bool("") is False

    def test_integer(self):
        assert _as_bool(1) is True
        assert _as_bool(0) is False

    def test_none(self):
        assert _as_bool(None) is False


class TestParseIntProp:
    """Tests for _parse_int_prop."""

    def test_int_value(self):
        assert _parse_int_prop(42) == 42

    def test_zero(self):
        assert _parse_int_prop(0) == 0

    def test_bytes_value(self):
        raw = (100).to_bytes(4, "little")
        assert _parse_int_prop(raw) == 100

    def test_string_with_number(self):
        assert _parse_int_prop("size: 12pt") == 12

    def test_string_no_number(self):
        assert _parse_int_prop("none") == 0

    def test_none(self):
        assert _parse_int_prop(None) == 0


class TestParseFontSize:
    """Tests for _parse_font_size."""

    def test_int_value(self):
        assert _parse_font_size(11) == 11

    def test_bytes_value(self):
        raw = (14).to_bytes(2, "little")
        assert _parse_font_size(raw) == 14

    def test_string_value(self):
        assert _parse_font_size("12") == 12

    def test_zero(self):
        assert _parse_font_size(0) == 0


class TestDetectImageFormat:
    """Tests for _detect_image_format."""

    def test_png(self):
        data = b"\x89PNG\r\n\x1a\n" + b"\x00" * 100
        assert _detect_image_format(data) == "png"

    def test_jpeg(self):
        data = b"\xff\xd8\xff\xe0" + b"\x00" * 100
        assert _detect_image_format(data) == "jpeg"

    def test_gif87a(self):
        data = b"GIF87a" + b"\x00" * 100
        assert _detect_image_format(data) == "gif"

    def test_gif89a(self):
        data = b"GIF89a" + b"\x00" * 100
        assert _detect_image_format(data) == "gif"

    def test_bmp(self):
        data = b"BM" + b"\x00" * 100
        assert _detect_image_format(data) == "bmp"

    def test_webp(self):
        data = b"RIFF\x00\x00\x00\x00WEBP" + b"\x00" * 100
        assert _detect_image_format(data) == "webp"

    def test_unknown_format(self):
        data = b"\x00\x01\x02\x03" + b"\x00" * 100
        assert _detect_image_format(data) == ""

    def test_empty_data(self):
        assert _detect_image_format(b"") == ""

    def test_too_short(self):
        assert _detect_image_format(b"\x89PN") == ""
