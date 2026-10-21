"""Tests for onenote_export.parser.one_store module."""

from pathlib import Path
from types import SimpleNamespace

from pyOneNote.OneDocument import OneDocment

from onenote_export.parser.one_store import (
    ExtractedObject,
    ExtractedPage,
    ExtractedSection,
    OneStoreParser,
    _clean_text,
    _extract_guid,
    _parse_int,
)


class TestExtractGuid:
    """Tests for _extract_guid."""

    def test_standard_identity_string(self):
        identity = "<ExtendedGUID> (abc-def-123, 42)"
        assert _extract_guid(identity) == "abc-def-123"

    def test_identity_with_extra_spaces(self):
        identity = "<ExtendedGUID> ( some-guid , 99)"
        assert _extract_guid(identity) == "some-guid"

    def test_empty_string(self):
        assert _extract_guid("") == ""

    def test_no_parentheses(self):
        assert _extract_guid("no-parens-here") == ""

    def test_missing_comma(self):
        assert _extract_guid("(no-comma)") == ""

    def test_complex_guid(self):
        identity = "<ExtendedGUID> ({12345678-abcd-ef01-2345-6789abcdef01}, 138)"
        assert _extract_guid(identity) == "{12345678-abcd-ef01-2345-6789abcdef01}"


class TestCleanText:
    """Tests for _clean_text in one_store module."""

    def test_removes_null_bytes(self):
        assert _clean_text("hello\x00world") == "helloworld"

    def test_strips_whitespace(self):
        assert _clean_text("  hello  ") == "hello"

    def test_empty_string(self):
        assert _clean_text("") == ""

    def test_only_null_bytes(self):
        assert _clean_text("\x00\x00\x00") == ""

    def test_mixed_content(self):
        assert _clean_text("  test\x00data  ") == "testdata"


class TestParseInt:
    """Tests for _parse_int."""

    def test_int_value(self):
        assert _parse_int(42) == 42

    def test_zero(self):
        assert _parse_int(0) == 0

    def test_bytes_value(self):
        raw = (100).to_bytes(4, "little")
        assert _parse_int(raw) == 100

    def test_bytes_short(self):
        raw = (5).to_bytes(2, "little")
        assert _parse_int(raw) == 5

    def test_string_with_number(self):
        assert _parse_int("level: 3") == 3

    def test_string_no_number(self):
        assert _parse_int("none") == 0

    def test_none(self):
        assert _parse_int(None) == 0

    def test_empty_string(self):
        assert _parse_int("") == 0

    def test_negative_int(self):
        assert _parse_int(-1) == -1


class TestExtractedDataclasses:
    """Tests for dataclass constructors."""

    def test_extracted_object_defaults(self):
        obj = ExtractedObject(obj_type="test", identity="id-1")
        assert obj.obj_type == "test"
        assert obj.identity == "id-1"
        assert obj.properties == {}

    def test_extracted_object_with_properties(self):
        props = {"Bold": True, "Font": "Arial"}
        obj = ExtractedObject(obj_type="text", identity="id-2", properties=props)
        assert obj.properties["Bold"] is True
        assert obj.properties["Font"] == "Arial"

    def test_extracted_page_defaults(self):
        page = ExtractedPage()
        assert page.title == ""
        assert page.level == 0
        assert page.author == ""
        assert page.objects == []

    def test_extracted_page_with_values(self):
        page = ExtractedPage(title="My Page", level=2, author="Test User")
        assert page.title == "My Page"
        assert page.level == 2
        assert page.author == "Test User"

    def test_extracted_section_defaults(self):
        section = ExtractedSection()
        assert section.file_path == ""
        assert section.display_name == ""
        assert section.pages == []
        assert section.file_data == {}
        assert section.paragraph_styles == {}

    def test_extracted_section_with_pages(self):
        page = ExtractedPage(title="Test")
        section = ExtractedSection(
            file_path="/test.one",
            display_name="Test Section",
            pages=[page],
        )
        assert section.file_path == "/test.one"
        assert section.display_name == "Test Section"
        assert len(section.pages) == 1


def _identity(guid: str) -> str:
    """Return the identity format used by pyOneNote for an object GUID."""
    return f"<ExtendedGUID> ({guid}, 1)"


def _metadata(
    guid: str,
    revision_id: str,
    title: str,
    *,
    level: int = 1,
) -> ExtractedObject:
    return ExtractedObject(
        obj_type="jcidPageMetaData",
        identity=_identity(guid),
        revision_id=revision_id,
        properties={
            "CachedTitleString": title,
            "PageLevel": level,
            "TopologyCreationTimeStamp": f"created-{revision_id}",
        },
    )


def _page_node(
    guid: str,
    revision_id: str,
    *,
    title: str = "",
    author: str = "",
    child_refs: list[str] | str | None = None,
) -> ExtractedObject:
    properties = {
        "CachedTitleStringFromPage": title,
        "Author": author,
        "LastModifiedTime": f"modified-{revision_id}",
    }
    if child_refs is not None:
        properties["ElementChildNodesOfVersionHistory"] = child_refs
    return ExtractedObject(
        obj_type="jcidPageNode",
        identity=_identity(guid),
        revision_id=revision_id,
        properties=properties,
    )


def _build_revision_page(
    objects: list[ExtractedObject],
    *,
    page_guid: str = "page-1",
    active_revision: str | None = "current",
    dependencies: dict[str, str] | None = None,
) -> ExtractedPage:
    parser = OneStoreParser("synthetic.one")
    parser._revision_dependencies = dependencies or {}
    parser._active_revisions = (
        {(0, ""): active_revision} if active_revision is not None else {}
    )
    pages = parser._build_pages(objects)
    assert len(pages) == 1
    return pages[0]


class TestRevisionAwarePages:
    """Page metadata follows the active revision, not traversal order."""

    def test_real_properties_keep_revision_context(self):
        path = Path(__file__).parent / "test_data" / "ExporterEval" / "Hierarchy.one"
        parser = OneStoreParser(path)
        with path.open("rb") as file:
            raw = parser._extract_raw_properties(OneDocment(file))

        assert raw
        assert all(item["revision_id"] for item in raw)
        assert parser._revision_dependencies
        assert parser._active_revisions

    def test_revision_state_uses_last_active_manifest_role(self):
        manifest_type = type("RevisionManifestStart6FND", (), {})

        def manifest(revision: str, dependent: str):
            data = manifest_type()
            data.rid = revision
            data.ridDependent = dependent
            data.RevisionRole = 1
            return SimpleNamespace(data=data, children=[])

        root = SimpleNamespace(
            fragments=[
                SimpleNamespace(
                    fileNodes=[
                        manifest("historical", ""),
                        manifest("current", "historical"),
                    ]
                )
            ]
        )
        parser = OneStoreParser("synthetic.one")

        parser._extract_revision_state(root)

        assert parser._revision_dependencies == {
            "historical": "",
            "current": "historical",
        }
        assert list(parser._active_revisions.values()) == ["current"]

    def test_first_declaration_wins_when_revision_ids_match(self):
        objects = [
            _metadata("page-1", "current", "First title"),
            _metadata("page-1", "current", "Later duplicate"),
            _page_node("page-1", "current"),
        ]

        page = _build_revision_page(objects)

        assert page.title == "First title"

    def test_default_context_ignores_unrelated_active_context(self):
        parser = OneStoreParser("synthetic.one")
        parser._active_revisions = {
            (0, ""): "current",
            (0, "alternate"): "foreign",
        }
        parser._revision_contexts = {
            "current": {(0, "")},
            "foreign": {(0, "alternate")},
        }
        candidates = [
            _metadata("page-1", "current", "Current title"),
            _metadata("page-1", "foreign", "Foreign title"),
        ]

        selected = parser._select_revisioned(candidates)

        assert selected is candidates[0]

    def test_current_title_beats_empty_historical_title(self):
        objects = [
            _metadata("page-1", "historical", ""),
            _metadata("page-1", "current", "Current title"),
            _page_node("page-1", "current"),
        ]

        page = _build_revision_page(objects, dependencies={"current": "historical"})

        assert page.title == "Current title"

    def test_current_shorter_rename_beats_longest_historical_title(self):
        objects = [
            _metadata("page-1", "historical", "A much longer historical title"),
            _metadata("page-1", "current", "Renamed"),
            _page_node("page-1", "current"),
        ]

        page = _build_revision_page(objects, dependencies={"current": "historical"})

        assert page.title == "Renamed"

    def test_current_revision_inherits_metadata_from_ancestor(self):
        objects = [
            _metadata("page-1", "historical", "Inherited title", level=3),
            _page_node("page-1", "current"),
        ]

        page = _build_revision_page(objects, dependencies={"current": "historical"})

        assert page.title == "Inherited title"
        assert page.level == 3

    def test_stale_candidates_after_current_do_not_replace_current_metadata(self):
        objects = [
            _metadata("page-1", "current", "Current title", level=2),
            _metadata("page-1", "stale", "Stale title", level=9),
            _page_node("page-1", "current", author="Current author"),
            _page_node("page-1", "stale", author="Stale author"),
        ]

        page = _build_revision_page(objects)

        assert page.title == "Current title"
        assert page.level == 2
        assert page.author == "Current author"
        assert page.last_modified == "modified-current"

    def test_page_node_title_is_used_when_metadata_title_is_empty(self):
        objects = [
            _metadata("page-1", "current", ""),
            _page_node("page-1", "current", title="Node title"),
        ]

        page = _build_revision_page(objects)

        assert page.title == "Node title"

    def test_missing_revision_metadata_preserves_first_candidate(self):
        objects = [
            _metadata("page-1", "first", "First title"),
            _metadata("page-1", "second", "Second title"),
            _page_node("page-1", "second"),
        ]

        page = _build_revision_page(objects, active_revision=None)

        assert page.title == "First title"
        assert page.level == 1

    def test_single_revision_remains_unchanged(self):
        objects = [
            _metadata("page-1", "only", "Only title", level=4),
            _page_node("page-1", "only", author="Author"),
        ]

        page = _build_revision_page(objects, active_revision="only")

        assert page.title == "Only title"
        assert page.level == 4
        assert page.author == "Author"

    def test_same_guid_body_remains_included(self):
        body = ExtractedObject(
            obj_type="jcidRichTextOENode",
            identity=_identity("page-1"),
            revision_id="current",
        )
        page = _build_revision_page(
            [
                _metadata("page-1", "current", "Page"),
                _page_node("page-1", "current"),
                body,
            ]
        )

        assert page.objects == [body]

    def test_active_page_node_includes_separate_guid_body(self):
        body = ExtractedObject(
            obj_type="jcidRichTextOENode",
            identity=_identity("body-1"),
            revision_id="current",
        )
        page = _build_revision_page(
            [
                _metadata("page-1", "current", "Page"),
                _page_node(
                    "page-1",
                    "current",
                    child_refs="body-1",
                ),
                body,
            ]
        )

        assert page.objects == [body]

    def test_inactive_page_node_reference_is_not_included(self):
        active_body = ExtractedObject(
            obj_type="jcidRichTextOENode",
            identity=_identity("active-body"),
            revision_id="current",
        )
        stale_body = ExtractedObject(
            obj_type="jcidRichTextOENode",
            identity=_identity("stale-body"),
            revision_id="stale",
        )
        page = _build_revision_page(
            [
                _metadata("page-1", "current", "Page"),
                _page_node(
                    "page-1",
                    "current",
                    child_refs=[_identity("active-body")],
                ),
                _page_node(
                    "page-1",
                    "stale",
                    child_refs=[_identity("stale-body")],
                ),
                active_body,
                stale_body,
            ]
        )

        assert page.objects == [active_body]

    def test_repeated_page_node_references_do_not_duplicate_body(self):
        body = ExtractedObject(
            obj_type="jcidRichTextOENode",
            identity=_identity("body-1"),
            revision_id="current",
        )
        page = _build_revision_page(
            [
                _metadata("page-1", "current", "Page"),
                _page_node(
                    "page-1",
                    "current",
                    child_refs=[_identity("body-1"), _identity("body-1")],
                ),
                body,
            ]
        )

        assert page.objects == [body]

    def test_referenced_body_uses_active_revision(self):
        current_body = ExtractedObject(
            obj_type="jcidRichTextOENode",
            identity=_identity("body-1"),
            revision_id="current",
            properties={"RichEditTextUnicode": "current"},
        )
        stale_body = ExtractedObject(
            obj_type="jcidRichTextOENode",
            identity=_identity("body-1"),
            revision_id="stale",
            properties={"RichEditTextUnicode": "stale"},
        )
        page = _build_revision_page(
            [
                _metadata("page-1", "current", "Page"),
                _page_node(
                    "page-1",
                    "current",
                    child_refs=[_identity("body-1")],
                ),
                current_body,
                stale_body,
            ],
            dependencies={"current": "stale"},
        )

        assert page.objects == [current_body]
