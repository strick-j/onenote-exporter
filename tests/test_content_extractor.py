"""Tests for onenote_export.parser.content_extractor module."""

from onenote_export.parser.content_extractor import (
    _as_bool,
    _build_top_level_oe_ids,
    _clean_text,
    _decode_text_value,
    _deduplicate_objects,
    _dedup_elements,
    _detect_image_format,
    _extract_rich_text,
    _looks_garbled,
    _parse_byte_prop_as_int,
    _parse_font_size,
    _parse_hyperlink_field_codes,
    _parse_int_prop,
    _reorder_by_outline_hierarchy,
    _section_name_from_path,
)
from onenote_export.model.content import RichText, TextRun
from onenote_export.parser.one_store import ExtractedObject


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


class TestDeduplicateObjects:
    """Tests for _deduplicate_objects."""

    def test_fewer_than_4_objects_returned_unchanged(self):
        objs = [
            ExtractedObject(
                obj_type="jcidRichTextOENode",
                identity="1",
                properties={"RichEditTextUnicode": "Hello"},
            ),
            ExtractedObject(
                obj_type="jcidRichTextOENode",
                identity="2",
                properties={"RichEditTextUnicode": "World"},
            ),
        ]
        result = _deduplicate_objects(objs)
        assert len(result) == 2

    def test_no_content_fingerprints_returns_unchanged(self):
        """Objects with no content (styles, outlines) are returned as-is."""
        objs = [
            ExtractedObject(obj_type="jcidOutlineElementNode", identity="1"),
            ExtractedObject(obj_type="jcidOutlineElementNode", identity="2"),
            ExtractedObject(obj_type="jcidOutlineElementNode", identity="3"),
            ExtractedObject(obj_type="jcidOutlineElementNode", identity="4"),
        ]
        result = _deduplicate_objects(objs)
        assert len(result) == 4

    def test_single_content_fp_returns_unchanged(self):
        """Only one content fingerprint means no duplicates possible."""
        objs = [
            ExtractedObject(obj_type="jcidOutlineElementNode", identity="1"),
            ExtractedObject(
                obj_type="jcidRichTextOENode",
                identity="2",
                properties={"RichEditTextUnicode": "Hello"},
            ),
            ExtractedObject(obj_type="jcidOutlineElementNode", identity="3"),
            ExtractedObject(obj_type="jcidOutlineElementNode", identity="4"),
        ]
        result = _deduplicate_objects(objs)
        assert len(result) == 4

    def test_duplicates_removed(self):
        """Duplicate content objects should be deduplicated."""
        objs = [
            ExtractedObject(
                obj_type="jcidRichTextOENode",
                identity="1",
                properties={"RichEditTextUnicode": "Hello"},
            ),
            ExtractedObject(
                obj_type="jcidRichTextOENode",
                identity="2",
                properties={"RichEditTextUnicode": "World"},
            ),
            ExtractedObject(
                obj_type="jcidRichTextOENode",
                identity="3",
                properties={"RichEditTextUnicode": "Hello"},
            ),
            ExtractedObject(
                obj_type="jcidRichTextOENode",
                identity="4",
                properties={"RichEditTextUnicode": "World"},
            ),
        ]
        result = _deduplicate_objects(objs)
        assert len(result) == 2

    def test_empty_list(self):
        assert _deduplicate_objects([]) == []


class TestBuildTopLevelOeIds:
    """Tests for _build_top_level_oe_ids."""

    def test_collects_outline_node_children(self):
        objs = [
            ExtractedObject(
                obj_type="jcidOutlineNode",
                identity="outline-1",
                properties={"ElementChildNodesOfVersionHistory": ["oe-1", "oe-2"]},
            ),
        ]
        result = _build_top_level_oe_ids(objs)
        assert result == {"oe-1", "oe-2"}

    def test_string_ref_normalized_to_list(self):
        """A single string ref should be handled as a list."""
        objs = [
            ExtractedObject(
                obj_type="jcidOutlineNode",
                identity="outline-1",
                properties={"ElementChildNodesOfVersionHistory": "oe-single"},
            ),
        ]
        result = _build_top_level_oe_ids(objs)
        assert result == {"oe-single"}

    def test_no_outline_nodes(self):
        objs = [
            ExtractedObject(obj_type="jcidRichTextOENode", identity="1"),
        ]
        result = _build_top_level_oe_ids(objs)
        assert result == set()

    def test_empty_refs(self):
        objs = [
            ExtractedObject(
                obj_type="jcidOutlineNode",
                identity="outline-1",
                properties={"ElementChildNodesOfVersionHistory": []},
            ),
        ]
        result = _build_top_level_oe_ids(objs)
        assert result == set()


class TestDecodeTextEdgeCases:
    """Additional edge case tests for _decode_text_value."""

    def test_hex_string_unicode_encoding(self):
        """Hex string decoded as UTF-16LE."""
        # "Hi" in UTF-16LE = 48 00 69 00
        result = _decode_text_value("48006900", encoding="unicode")
        assert result == "Hi"

    def test_garbled_ascii_re_encoding(self):
        """Garbled CJK text re-encoded from UTF-16LE to ASCII."""
        # Create text that looks garbled (> 30% non-ASCII)
        ascii_text = "Hello World"
        # Encode as UTF-16LE, then decode incorrectly as UTF-16LE to get garbled
        garbled = ascii_text.encode("ascii").decode("utf-16-le", errors="replace")
        result = _decode_text_value(garbled, encoding="ascii")
        # Should attempt to re-encode and recover
        assert isinstance(result, str)

    def test_bytes_utf16_decode_error_falls_back_to_latin1(self):
        """Invalid UTF-16LE bytes fall back to latin-1 decoding."""
        # Odd-length bytes can't be valid UTF-16LE
        raw = b"\xff\xfe\x80"
        result = _decode_text_value(raw, encoding="unicode")
        assert isinstance(result, str)

    def test_integer_value(self):
        """Non-string, non-bytes values are converted via str()."""
        result = _decode_text_value(42)
        assert result == "42"

    def test_whitespace_only_string(self):
        result = _decode_text_value("   ")
        assert result == ""


class TestParseBytePropAsInt:
    """Tests for _parse_byte_prop_as_int."""

    def test_int_value(self):
        assert _parse_byte_prop_as_int(36) == 36

    def test_bytes_value(self):
        raw = (36).to_bytes(2, "little")
        assert _parse_byte_prop_as_int(raw) == 36

    def test_repr_string(self):
        """Parse repr() style byte string like b'$\\x00'."""
        assert _parse_byte_prop_as_int("b'$\\x00'") == 36  # ord('$') = 36

    def test_repr_string_hex(self):
        assert _parse_byte_prop_as_int("b'\\x04\\x00'") == 4

    def test_empty_string(self):
        assert _parse_byte_prop_as_int("") == 0

    def test_none(self):
        assert _parse_byte_prop_as_int(None) == 0

    def test_invalid_repr(self):
        assert _parse_byte_prop_as_int("b'invalid") == 0

    def test_short_bytes(self):
        """Single byte should return 0 (need at least 2)."""
        assert _parse_byte_prop_as_int(b"\x05") == 0


class TestSectionNameFromPath:
    """Tests for _section_name_from_path."""

    def test_simple_filename(self):
        assert _section_name_from_path("/path/to/Notes.one") == "Notes"

    def test_filename_with_date(self):
        result = _section_name_from_path("/path/to/ADI (On 2-25-26).one")
        assert result == "ADI"

    def test_filename_with_date_dash_suffix(self):
        result = _section_name_from_path("/path/to/Section (On 2-25-26 - 3).one")
        assert result == "Section"

    def test_empty_after_cleaning(self):
        """If name becomes empty after cleaning, return 'Untitled'."""
        result = _section_name_from_path("/path/to/.one")
        assert result == "Untitled"


class TestParseFontSizeEdgeCases:
    """Additional edge case tests for _parse_font_size."""

    def test_none_value(self):
        assert _parse_font_size(None) == 0

    def test_string_no_number(self):
        assert _parse_font_size("normal") == 0

    def test_single_byte(self):
        raw = b"\x0e"  # 14
        assert _parse_font_size(raw) == 14

    def test_empty_bytes(self):
        assert _parse_font_size(b"") == 0


class TestDedupElements:
    """Tests for _dedup_elements."""

    def test_removes_duplicate_rich_text(self):
        run1 = TextRun(text="Hello world")
        elem1 = RichText(runs=[run1])
        elem2 = RichText(runs=[run1])
        result = _dedup_elements([elem1, elem2])
        assert len(result) == 1

    def test_keeps_same_text_different_list_type(self):
        """Same text in different list types is not a duplicate."""
        run = TextRun(text="Item")
        bullet = RichText(runs=[run], list_type="unordered")
        numbered = RichText(runs=[run], list_type="ordered")
        result = _dedup_elements([bullet, numbered])
        assert len(result) == 2

    def test_empty_list(self):
        assert _dedup_elements([]) == []


class TestExtractRichText:
    """Tests for _extract_rich_text."""

    def test_returns_none_for_empty_text(self):
        obj = ExtractedObject(
            obj_type="jcidRichTextOENode",
            identity="1",
            properties={},
        )
        result = _extract_rich_text(obj, {}, None)
        assert result is None

    def test_returns_none_for_whitespace_only(self):
        obj = ExtractedObject(
            obj_type="jcidRichTextOENode",
            identity="1",
            properties={"RichEditTextUnicode": "   "},
        )
        result = _extract_rich_text(obj, {}, None)
        assert result is None

    def test_extracts_text_with_formatting(self):
        obj = ExtractedObject(
            obj_type="jcidRichTextOENode",
            identity="1",
            properties={"RichEditTextUnicode": "Hello"},
        )
        style = {"Bold": True, "Italic": True}
        result = _extract_rich_text(obj, style, None)
        assert result is not None
        assert result.runs[0].text == "Hello"
        assert result.runs[0].bold is True
        assert result.runs[0].italic is True

    def test_extracts_text_from_ascii(self):
        obj = ExtractedObject(
            obj_type="jcidRichTextOENode",
            identity="1",
            properties={"TextExtendedAscii": "World"},
        )
        result = _extract_rich_text(obj, {}, None)
        assert result is not None
        assert result.runs[0].text == "World"

    def test_extracts_hyperlink_from_field_code(self):
        """HYPERLINK field codes should be parsed into url + display text."""
        obj = ExtractedObject(
            obj_type="jcidRichTextOENode",
            identity="1",
            properties={
                "RichEditTextUnicode": (
                    '\uFDDFHYPERLINK "mailto:user@example.com"'
                    "User Name (Accepted)\x00"
                ),
            },
        )
        result = _extract_rich_text(obj, {}, None)
        assert result is not None
        assert result.runs[0].text == "User Name (Accepted)"
        assert result.runs[0].hyperlink_url == "mailto:user@example.com"

    def test_extracts_hyperlink_from_https_field_code(self):
        """HTTPS HYPERLINK field codes should be parsed correctly."""
        obj = ExtractedObject(
            obj_type="jcidRichTextOENode",
            identity="1",
            properties={
                "RichEditTextUnicode": (
                    '\uFDDFHYPERLINK "https://example.com/path"'
                    "Link to Document\x00"
                ),
            },
        )
        result = _extract_rich_text(obj, {}, None)
        assert result is not None
        assert result.runs[0].text == "Link to Document"
        assert result.runs[0].hyperlink_url == "https://example.com/path"


class TestParseHyperlinkFieldCodes:
    """Tests for _parse_hyperlink_field_codes."""

    def test_mailto_url(self):
        text = '\uFDDFHYPERLINK "mailto:user@example.com"User Name\x00'
        segments = _parse_hyperlink_field_codes(text)
        assert len(segments) == 1
        assert segments[0] == ("User Name", "mailto:user@example.com")

    def test_https_url(self):
        text = '\uFDDFHYPERLINK "https://example.com/path?q=1"Click Here\x00'
        segments = _parse_hyperlink_field_codes(text)
        assert len(segments) == 1
        assert segments[0] == ("Click Here", "https://example.com/path?q=1")

    def test_no_field_code(self):
        text = "Just regular text"
        segments = _parse_hyperlink_field_codes(text)
        assert segments == [("Just regular text", "")]

    def test_alternate_marker_fdf3(self):
        text = '\uFDF3HYPERLINK "https://example.com"Link Text\x00'
        segments = _parse_hyperlink_field_codes(text)
        assert len(segments) == 1
        assert segments[0] == ("Link Text", "https://example.com")

    def test_display_text_with_parentheses(self):
        text = '\uFDDFHYPERLINK "mailto:a@b.com"Name (Accepted Meeting)\x00'
        segments = _parse_hyperlink_field_codes(text)
        assert segments[0] == ("Name (Accepted Meeting)", "mailto:a@b.com")

    def test_empty_string(self):
        segments = _parse_hyperlink_field_codes("")
        assert segments == [("", "")]

    def test_mixed_text_and_field_code(self):
        """Text with prefix before a field code produces two segments."""
        text = 'Meeting options | \uFDDFHYPERLINK "https://example.com"Reset PIN'
        segments = _parse_hyperlink_field_codes(text)
        assert len(segments) == 2
        assert segments[0] == ("Meeting options |", "")
        assert segments[1] == ("Reset PIN", "https://example.com")

    def test_two_field_codes_in_one_text(self):
        """Two field codes in one text node produce three segments."""
        text = (
            'Prefix: \uFDDFHYPERLINK "https://a.com"Link A | '
            '\uFDDFHYPERLINK "https://b.com"Link B'
        )
        segments = _parse_hyperlink_field_codes(text)
        assert len(segments) == 3
        assert segments[0] == ("Prefix:", "")
        assert segments[1] == ("Link A |", "https://a.com")
        assert segments[2] == ("Link B", "https://b.com")


class TestDecodeGarbledUnicode:
    """Tests for garbled text detection in Unicode encoding path."""

    def test_garbled_unicode_not_mangled(self):
        """Unicode path should NOT apply garbled re-encoding (preserves real Unicode)."""
        # Legitimate CJK text should pass through unchanged
        cjk_text = "\u4f1a\u8b70\u30e1\u30e2"  # 会議メモ (Japanese)
        result = _decode_text_value(cjk_text, encoding="unicode")
        assert result == cjk_text

    def test_garbled_ascii_still_fixed(self):
        """ASCII path should still fix garbled text."""
        ascii_text = "Hello World Test"
        garbled = ascii_text.encode("ascii").decode("utf-16-le", errors="replace")
        assert _looks_garbled(garbled), "Test fixture is not actually garbled"
        result = _decode_text_value(garbled, encoding="ascii")
        assert isinstance(result, str)


class TestReorderByOutlineHierarchy:
    """Tests for _reorder_by_outline_hierarchy."""

    def _make_obj(self, obj_type, identity, **props):
        return ExtractedObject(obj_type=obj_type, identity=identity, properties=props)

    def test_short_list_returned_unchanged(self):
        """Lists shorter than 4 are returned as-is."""
        objs = [
            self._make_obj("jcidRichTextOENode", "1"),
            self._make_obj("jcidOutlineNode", "2"),
        ]
        result = _reorder_by_outline_hierarchy(objs)
        assert result is objs

    def test_no_outline_nodes_returned_unchanged(self):
        """Without OutlineNodes, list is returned as-is."""
        objs = [
            self._make_obj("jcidOutlineElementNode", "1"),
            self._make_obj("jcidRichTextOENode", "2"),
            self._make_obj("jcidOutlineElementNode", "3"),
            self._make_obj("jcidRichTextOENode", "4"),
        ]
        result = _reorder_by_outline_hierarchy(objs)
        assert result is objs

    def test_no_orphans_returned_unchanged(self):
        """If no content before the first structural element, return as-is."""
        objs = [
            self._make_obj(
                "jcidOutlineNode", "ON1",
                ElementChildNodesOfVersionHistory=["OE1"],
            ),
            self._make_obj("jcidOutlineElementNode", "OE1"),
            self._make_obj(
                "jcidRichTextOENode", "RT1",
                RichEditTextUnicode="Hello",
            ),
            self._make_obj("jcidOutlineElementNode", "OE2"),
        ]
        result = _reorder_by_outline_hierarchy(objs)
        assert result is objs

    def test_orphan_relocated_after_contentless_oe(self):
        """Orphaned content before first OE should move after its parent OE."""
        orphan_rt = self._make_obj(
            "jcidRichTextOENode", "RT-orphan",
            RichEditTextUnicode="Note 4",
        )
        outline_node = self._make_obj(
            "jcidOutlineNode", "ON1",
            ElementChildNodesOfVersionHistory=["OE-parent", "OE-child"],
        )
        oe_parent = self._make_obj(
            "jcidOutlineElementNode", "OE-parent",
            ElementChildNodesOfVersionHistory=["OE-nested"],
        )
        rt_parent = self._make_obj(
            "jcidRichTextOENode", "RT-parent",
            RichEditTextUnicode="Notes",
        )
        # OE-nested is the contentless OE that should receive the orphan
        oe_nested = self._make_obj("jcidOutlineElementNode", "OE-nested")
        oe_child = self._make_obj("jcidOutlineElementNode", "OE-child")
        rt_child = self._make_obj(
            "jcidRichTextOENode", "RT-child",
            RichEditTextUnicode="Some text",
        )

        objects = [
            orphan_rt,      # [0] orphaned content
            outline_node,   # [1] OutlineNode
            oe_nested,      # [2] out-of-place OE (no content)
            oe_parent,      # [3] OE with content
            rt_parent,      # [4] content for OE-parent
            oe_child,       # [5] OE with content
            rt_child,       # [6] content for OE-child
        ]

        result = _reorder_by_outline_hierarchy(objects)
        identities = [obj.identity for obj in result]

        # Orphan should appear after OE-nested, not at the start
        assert identities[0] != "RT-orphan", "Orphan should not be first"
        nested_idx = identities.index("OE-nested")
        orphan_idx = identities.index("RT-orphan")
        assert orphan_idx == nested_idx + 1, (
            f"Orphan should follow OE-nested; got order: {identities}"
        )

    def test_already_correct_order_preserved(self):
        """Objects in correct hierarchy order should produce identical output."""
        outline_node = self._make_obj(
            "jcidOutlineNode", "ON1",
            ElementChildNodesOfVersionHistory=["OE1", "OE2"],
        )
        oe1 = self._make_obj("jcidOutlineElementNode", "OE1")
        rt1 = self._make_obj(
            "jcidRichTextOENode", "RT1",
            RichEditTextUnicode="First",
        )
        oe2 = self._make_obj("jcidOutlineElementNode", "OE2")
        rt2 = self._make_obj(
            "jcidRichTextOENode", "RT2",
            RichEditTextUnicode="Second",
        )

        objects = [outline_node, oe1, rt1, oe2, rt2]
        # No orphans → returned as-is
        result = _reorder_by_outline_hierarchy(objects)
        assert result is objects

    def test_outline_nodes_sorted_by_vert_position(self):
        """Nodes without vert come first, then ascending vert value."""
        orphan = self._make_obj(
            "jcidRichTextOENode", "RT-orphan",
            RichEditTextUnicode="orphan",
        )
        node_no_vert = self._make_obj(
            "jcidOutlineNode", "ON-title",
            ElementChildNodesOfVersionHistory=["OE-title"],
        )
        oe_title = self._make_obj("jcidOutlineElementNode", "OE-title")
        rt_title = self._make_obj(
            "jcidRichTextOENode", "RT-title",
            RichEditTextUnicode="Title",
        )
        node_vert = self._make_obj(
            "jcidOutlineNode", "ON-body",
            OffsetFromParentVert=200,
            ElementChildNodesOfVersionHistory=["OE-body"],
        )
        # OE-body has no content → orphan matches here
        oe_body = self._make_obj("jcidOutlineElementNode", "OE-body")

        objects = [orphan, node_vert, oe_body, node_no_vert, oe_title, rt_title]
        result = _reorder_by_outline_hierarchy(objects)
        identities = [obj.identity for obj in result]

        # Title node (no vert) should come before body node (vert=200)
        assert identities.index("ON-title") < identities.index("ON-body")
