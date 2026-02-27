"""Extracts structured content from pyOneNote parsed objects.

Bridges the parser output (ExtractedSection) to the high-level
content model (Section with Pages and ContentElements).
"""

import ast
import logging
import re
from dataclasses import dataclass
from pathlib import Path

from onenote_export.model.content import (
    ContentElement,
    EmbeddedFile,
    ImageElement,
    RichText,
    TableElement,
    TextRun,
)
from onenote_export.model.page import Page
from onenote_export.model.section import Section
from onenote_export.parser.one_store import (
    ExtractedObject,
    ExtractedPage,
    ExtractedSection,
)

logger = logging.getLogger(__name__)

# JCID type constants
_RICH_TEXT = "jcidRichTextOENode"
_IMAGE_NODE = "jcidImageNode"
_TABLE_NODE = "jcidTableNode"
_TABLE_ROW = "jcidTableRowNode"
_TABLE_CELL = "jcidTableCellNode"
_EMBEDDED_FILE = "jcidEmbeddedFileNode"
_OUTLINE_ELEMENT = "jcidOutlineElementNode"
_OUTLINE_NODE = "jcidOutlineNode"
_NUMBER_LIST = "jcidNumberListNode"
_STYLE_CONTAINER = "jcidPersistablePropertyContainerForTOCSection"

# ParagraphStyleId → heading level mapping
_HEADING_STYLE_MAP: dict[str, int] = {
    "h1": 1,
    "h2": 2,
    "h3": 3,
    "h4": 4,
    "h5": 5,
    "h6": 6,
}

# HYPERLINK field code pattern: OneNote collapsible sections embed
# hyperlinks as RTF-style field codes with U+FDDF or U+FDF3 marker.
# Format: <marker>HYPERLINK "url"display_text
# Display text captures up to (not including) the next field code marker.
_HYPERLINK_FIELD_RE = re.compile(
    r"[\uFDDF\uFDF3]HYPERLINK\s+\"([^\"]+)\"([^\uFDDF\uFDF3]+)",
)


def _deduplicate_objects(
    objects: list[ExtractedObject],
) -> list[ExtractedObject]:
    """Remove duplicate content objects caused by OneNote revision history.

    OneNote stores full copies of all page content for each revision.
    This creates repeated blocks of identical text/images.  We detect
    blocks that repeat and keep only the first copy.

    The detection finds the first content fingerprint that appears
    more than once.  If no duplicate is found, the objects are
    returned unchanged (no revision copies present).
    """
    if len(objects) < 4:
        return objects

    # Build a fingerprint sequence for content-bearing objects
    fingerprints: list[str] = []
    for obj in objects:
        fp = _object_fingerprint(obj)
        fingerprints.append(fp)

    # Find the repeat pattern: look for the first content fingerprint
    # appearing again later in the sequence
    content_fps = [
        (i, fp)
        for i, fp in enumerate(fingerprints)
        if fp and objects[i].obj_type in (_RICH_TEXT, _IMAGE_NODE, _EMBEDDED_FILE)
    ]

    if len(content_fps) < 2:
        return objects

    # Find where the first content element repeats
    first_fp = content_fps[0][1]
    repeat_idx = None

    for i, fp in content_fps[1:]:
        if fp == first_fp:
            repeat_idx = i
            break

    if repeat_idx is None:
        return objects

    # Take objects up to the repeat point, plus any non-content objects
    # (styles, outline elements) that follow
    result: list[ExtractedObject] = []
    seen_content: set[str] = set()

    for i, obj in enumerate(objects):
        fp = fingerprints[i]

        # Non-content objects (styles, outlines) - always include
        if not fp:
            result.append(obj)
            continue

        # Content object - include only the first occurrence
        if fp not in seen_content:
            seen_content.add(fp)
            result.append(obj)

    return result


def _object_fingerprint(obj: ExtractedObject) -> str:
    """Create a content-based fingerprint for deduplication.

    Normalises text by decoding from both Unicode and ASCII property
    fields so that different encoding representations of the same
    content produce the same fingerprint.
    """
    if obj.obj_type == _RICH_TEXT:
        raw_unicode = obj.properties.get("RichEditTextUnicode", "")
        raw_ascii = obj.properties.get("TextExtendedAscii", "")
        if raw_unicode:
            decoded = _decode_text_value(raw_unicode, encoding="unicode")
        elif raw_ascii:
            decoded = _decode_text_value(raw_ascii, encoding="ascii")
        else:
            decoded = ""
        return f"text:{decoded}" if decoded.strip() else ""
    elif obj.obj_type == _IMAGE_NODE:
        filename = str(obj.properties.get("ImageFilename", ""))
        alt = str(obj.properties.get("ImageAltText", ""))
        return f"img:{filename}:{alt}" if (filename or alt) else ""
    elif obj.obj_type == _EMBEDDED_FILE:
        name = str(obj.properties.get("EmbeddedFileName", ""))
        return f"file:{name}" if name else ""
    return ""


def _reorder_by_outline_hierarchy(
    objects: list[ExtractedObject],
) -> list[ExtractedObject]:
    """Reorder page objects so content follows its parent OE in hierarchy order.

    OneNote stores recently-edited content earlier in the flat object list
    than its parent outline structure.  This function walks the outline
    hierarchy (via ``ElementChildNodesOfVersionHistory``) to reconstruct
    the correct order.

    When no orphaned content objects are detected (content before the first
    structural element), the list is returned unchanged.
    """
    if len(objects) < 4:
        return objects

    has_outline = any(obj.obj_type == _OUTLINE_NODE for obj in objects)
    if not has_outline:
        return objects

    _BOUNDARY = {_OUTLINE_ELEMENT, _OUTLINE_NODE}

    # Collect orphaned content: content objects before the first structural element.
    orphans: list[ExtractedObject] = []
    for obj in objects:
        if obj.obj_type in _BOUNDARY:
            break
        if obj.obj_type in (_RICH_TEXT, _IMAGE_NODE, _EMBEDDED_FILE):
            orphans.append(obj)

    # If no orphaned content, objects are already in usable order.
    if not orphans:
        return objects

    # Build identity → object lookup
    id_to_obj: dict[str, ExtractedObject] = {}
    for obj in objects:
        if obj.identity:
            id_to_obj[obj.identity] = obj

    # Build OE → content group (non-structural objects following the OE).
    oe_content: dict[str, list[ExtractedObject]] = {}
    for i, obj in enumerate(objects):
        if obj.obj_type != _OUTLINE_ELEMENT:
            continue
        group: list[ExtractedObject] = []
        j = i + 1
        while j < len(objects):
            nxt = objects[j]
            if nxt.obj_type in _BOUNDARY:
                break
            group.append(nxt)
            j += 1
        oe_content[obj.identity] = group

    visited: set[str] = set()
    result: list[ExtractedObject] = []

    def _emit(obj: ExtractedObject) -> None:
        ident = obj.identity
        if ident and ident in visited:
            return
        if ident:
            visited.add(ident)
        result.append(obj)

    def _walk_oe(oe_id: str) -> None:
        if oe_id in visited:
            return
        oe_obj = id_to_obj.get(oe_id)
        if oe_obj is None:
            return

        _emit(oe_obj)

        group = oe_content.get(oe_id, [])
        if group:
            for g in group:
                _emit(g)
        elif orphans:
            _emit(orphans.pop(0))

        child_refs = oe_obj.properties.get(
            "ElementChildNodesOfVersionHistory", []
        )
        if isinstance(child_refs, str):
            child_refs = [child_refs]
        if isinstance(child_refs, list):
            for ref in child_refs:
                if isinstance(ref, str):
                    _walk_oe(ref)

    # Process outline nodes sorted by vertical position.
    # Nodes without OffsetFromParentVert (title/date blocks) come first,
    # ordered by original index.  Nodes with a vert value follow, sorted
    # ascending (top-to-bottom on the page).
    outline_nodes = [
        (i, obj)
        for i, obj in enumerate(objects)
        if obj.obj_type == _OUTLINE_NODE
    ]

    def _node_sort_key(
        item: tuple[int, ExtractedObject],
    ) -> tuple[int, int]:
        idx, node = item
        vert = node.properties.get("OffsetFromParentVert")
        if vert is None:
            return (0, idx)
        return (1, _parse_int_prop(vert))

    outline_nodes.sort(key=_node_sort_key)

    for _, node in outline_nodes:
        _emit(node)
        child_refs = node.properties.get(
            "ElementChildNodesOfVersionHistory", []
        )
        if isinstance(child_refs, str):
            child_refs = [child_refs]
        if isinstance(child_refs, list):
            for ref in child_refs:
                if isinstance(ref, str):
                    _walk_oe(ref)

    # Append remaining unvisited objects in original order.
    for obj in objects:
        _emit(obj)

    return result


def extract_section(parsed: ExtractedSection) -> Section:
    """Convert an ExtractedSection into a high-level Section model."""
    section = Section(
        name=parsed.display_name or _section_name_from_path(parsed.file_path),
        file_path=parsed.file_path,
    )

    for extracted_page in parsed.pages:
        page = _build_page(
            extracted_page,
            parsed.file_data,
            parsed.paragraph_styles,
        )
        section.pages.append(page)

    return section


def _find_out_of_line_table_refs(objects: list[ExtractedObject]) -> set[int]:
    """Find objects that are referenced as out-of-line table cell content.

    When a cell is recently edited, OneNote may store its content under a
    new revision GUID that lands earlier in the flat object list.  The
    cell's ``ElementChildNodesOfVersionHistory`` references that GUID.

    This pre-scan identifies those referenced objects so the main loop
    can skip them (they will be pulled in by ``_extract_table`` instead).
    """
    _TABLE_CELL_TYPE = "jcidTableCellNode"

    # Build identity → index map
    identity_map: dict[str, int] = {}
    for idx, obj in enumerate(objects):
        if obj.identity:
            identity_map[obj.identity] = idx

    # Collect all cell child-ref GUIDs and the index range of inline
    # content that immediately follows each cell.
    inline_ranges: set[int] = set()
    cell_refs: list[str] = []

    for idx, obj in enumerate(objects):
        if obj.obj_type != _TABLE_CELL_TYPE:
            continue
        refs = obj.properties.get("ElementChildNodesOfVersionHistory", [])
        if isinstance(refs, list):
            cell_refs.extend(refs)
        # Mark the inline content region (objects right after the cell)
        j = idx + 1
        while j < len(objects):
            inner = objects[j]
            if inner.obj_type in (
                _TABLE_CELL_TYPE,
                "jcidTableRowNode",
                "jcidTableNode",
            ):
                break
            inline_ranges.add(j)
            j += 1

    # For each ref, if its target is OUTSIDE the inline table region,
    # it's out-of-line content that should be skipped in the main loop.
    out_of_line: set[int] = set()
    for ref in cell_refs:
        ref_idx = identity_map.get(ref)
        if ref_idx is None or ref_idx in inline_ranges:
            continue
        # Walk forward from the ref, collecting the outline group
        j = ref_idx
        while j < len(objects):
            robj = objects[j]
            if robj.obj_type in (
                "jcidOutlineElementNode",
                "jcidRichTextOENode",
                "jcidImageNode",
                "jcidEmbeddedFileNode",
            ):
                out_of_line.add(j)
                j += 1
            else:
                break

    return out_of_line


@dataclass(frozen=True)
class _ListInfo:
    """Resolved list information for a single list item."""

    list_type: str  # "ordered" or "unordered"
    indent_level: int  # 0-based nesting depth


# NumberListFormat first-byte values (MS-ONESTORE NumberListNode).
_NUMBER_LIST_FORMAT_BULLET = 0x01
_NUMBER_LIST_FORMAT_NUMBERED = 0x03

# ListMSAAIndex → indent level for non-top-level list items.
# Values observed empirically from OneNote test files; top-level items
# (those in OutlineNode.ElementChildNodesOfVersionHistory) are always
# level 0 regardless of msaa.
_BULLET_MSAA_LEVEL: dict[int, int] = {1: 1, 4: 2, 9: 3}
_NUMBERED_MSAA_LEVEL: dict[int, int] = {36: 1, 53: 2, 45: 3}


def _build_list_node_map(
    objects: list[ExtractedObject],
) -> dict[str, dict[str, object]]:
    """Build a mapping from NumberListNode identity to its properties."""
    result: dict[str, dict[str, object]] = {}
    for obj in objects:
        if obj.obj_type == _NUMBER_LIST:
            result[obj.identity] = dict(obj.properties)
    return result


def _build_top_level_oe_ids(
    objects: list[ExtractedObject],
) -> set[str]:
    """Identify OutlineElement identities that are direct children of OutlineNodes.

    An OutlineNode's ``ElementChildNodesOfVersionHistory`` lists its
    top-level children.  Elements in that set are at indent level 0;
    elements NOT in the set are nested deeper.
    """
    top_ids: set[str] = set()
    for obj in objects:
        if obj.obj_type != _OUTLINE_NODE:
            continue
        refs = obj.properties.get("ElementChildNodesOfVersionHistory", [])
        if isinstance(refs, str):
            refs = [refs]
        if isinstance(refs, list):
            for ref in refs:
                if isinstance(ref, str):
                    top_ids.add(ref)
    return top_ids


def _resolve_list_info(
    oe_obj: ExtractedObject,
    list_node_map: dict[str, dict[str, object]],
    top_level_oe_ids: set[str],
) -> _ListInfo | None:
    """Resolve list type and indent level for an OutlineElement.

    Returns None if the element is not a list item.
    """
    list_refs = oe_obj.properties.get("ListNodes")
    if not list_refs:
        return None

    # Normalize to list
    if isinstance(list_refs, str):
        list_refs = [list_refs]
    if not isinstance(list_refs, list) or not list_refs:
        return None

    # Look up the NumberListNode
    ref = str(list_refs[0])
    node_props = list_node_map.get(ref)
    if not node_props:
        logger.debug(
            "ListNodes ref %r not found in list_node_map; defaulting to unordered", ref
        )
        return _ListInfo(list_type="unordered", indent_level=0)

    # Determine bullet vs numbered from NumberListFormat
    fmt = node_props.get("NumberListFormat", "")
    if isinstance(fmt, str) and fmt:
        fmt_byte = ord(fmt[0])
    else:
        fmt_byte = 0

    is_ordered = fmt_byte == _NUMBER_LIST_FORMAT_NUMBERED
    list_type = "ordered" if is_ordered else "unordered"

    # Parse msaa value — pyOneNote stores 2-byte properties as either
    # actual bytes or as their repr() string (e.g. "b'$\\x00'").
    msaa_raw = node_props.get("ListMSAAIndex", b"")
    msaa_val = _parse_byte_prop_as_int(msaa_raw)

    # Determine indent level: top-level OEs are at level 0,
    # non-top OEs use the msaa value to determine depth.
    is_top = oe_obj.identity in top_level_oe_ids
    if is_top:
        indent_level = 0
    else:
        level_map = _NUMBERED_MSAA_LEVEL if is_ordered else _BULLET_MSAA_LEVEL
        indent_level = level_map.get(msaa_val, 1)

    return _ListInfo(list_type=list_type, indent_level=indent_level)


def _build_page(
    extracted: ExtractedPage,
    file_data: dict[str, bytes],
    paragraph_styles: dict[str, str] | None = None,
) -> Page:
    """Build a Page model from an ExtractedPage."""
    page = Page(
        title=extracted.title or "Untitled",
        level=extracted.level,
        author=extracted.author,
    )

    # Deduplicate objects: OneNote revisions repeat the full content.
    # Remove objects that are exact duplicates (same type + same text content).
    deduped_objects = _deduplicate_objects(extracted.objects)

    # Reorder objects so that recently-edited content appears after
    # its parent OutlineElement rather than at the top of the list.
    deduped_objects = _reorder_by_outline_hierarchy(deduped_objects)

    # Pre-scan: identify objects that are out-of-line table cell content.
    # These are stored earlier in the list but referenced by a cell's
    # ElementChildNodesOfVersionHistory GUID.  They must be skipped in
    # the main loop so they only appear inside the table.
    skip_indices = _find_out_of_line_table_refs(deduped_objects)

    # Pre-scan: build list resolution data structures.
    list_node_map = _build_list_node_map(deduped_objects)
    top_level_oe_ids = _build_top_level_oe_ids(deduped_objects)

    # Process objects in order, building content elements.
    # Use index-based iteration so table processing can consume
    # subsequent row/cell/content objects.
    current_style: dict[str, object] = {}
    list_info: _ListInfo | None = None
    list_info_used = False  # True once a non-empty RT used list_info
    i = 0

    while i < len(deduped_objects):
        if i in skip_indices:
            i += 1
            continue

        obj = deduped_objects[i]

        if obj.obj_type == _STYLE_CONTAINER:
            current_style = dict(obj.properties)
            i += 1
            continue

        if obj.obj_type == _NUMBER_LIST:
            i += 1
            continue

        if obj.obj_type == _OUTLINE_ELEMENT:
            resolved = _resolve_list_info(
                obj,
                list_node_map,
                top_level_oe_ids,
            )
            if resolved is not None:
                # OE has its own list marker — always use it.
                list_info = resolved
                list_info_used = False
            elif list_info_used:
                # Previous list_info already produced content.
                # This non-list OE is NOT a wrapper — reset.
                list_info = None
            # Otherwise: list_info carries forward (wrapper OE for
            # a recently-edited list item whose text is here).
            i += 1
            continue

        if obj.obj_type == _RICH_TEXT:
            element = _extract_rich_text(
                obj,
                current_style,
                list_info,
                paragraph_styles,
            )
            if element:
                page.elements.append(element)
                if list_info is not None:
                    list_info_used = True
            i += 1
            continue

        if obj.obj_type == _IMAGE_NODE:
            element = _extract_image(obj, file_data)
            if element:
                page.elements.append(element)
                if list_info is not None:
                    list_info_used = True
            i += 1
            continue

        if obj.obj_type == _TABLE_NODE:
            element, consumed, _out_of_line = _extract_table(
                obj,
                deduped_objects,
                i,
                current_style,
                file_data,
            )
            if element:
                page.elements.append(element)
                if list_info is not None:
                    list_info_used = True
            i += 1 + consumed
            continue

        if obj.obj_type == _EMBEDDED_FILE:
            element = _extract_embedded_file(obj, file_data)
            if element:
                page.elements.append(element)
                if list_info is not None:
                    list_info_used = True
            i += 1
            continue

        # Skip table row/cell nodes that weren't consumed by a table
        # (shouldn't happen, but be defensive)
        if obj.obj_type in (_TABLE_ROW, _TABLE_CELL):
            i += 1
            continue

        i += 1

    page.elements = _dedup_elements(page.elements)
    return page


def _dedup_elements(elements: list[ContentElement]) -> list[ContentElement]:
    """Remove revision-duplicate elements from a page's content list.

    The same text may legitimately appear in different list contexts
    (e.g. bullet and numbered), so the dedup key includes list_type.
    Within the same list type, duplicate text at any indent level is
    treated as a revision artifact.
    """
    seen: set[tuple[str, str]] = set()
    result: list[ContentElement] = []
    for elem in elements:
        if isinstance(elem, RichText):
            text = " ".join(r.text for r in elem.runs).strip()
            key = (text, elem.list_type)
            if key in seen:
                continue
            seen.add(key)
        result.append(elem)
    return result


def _extract_rich_text(
    obj: ExtractedObject,
    style: dict[str, object],
    list_info: _ListInfo | None,
    paragraph_styles: dict[str, str] | None = None,
) -> RichText | None:
    """Extract rich text from a RichTextOENode."""
    props = obj.properties

    # Get text content - try RichEditTextUnicode first, then TextExtendedAscii
    text = ""
    raw_unicode = props.get("RichEditTextUnicode", "")
    raw_ascii = props.get("TextExtendedAscii", "")

    if raw_unicode:
        text = _decode_text_value(raw_unicode, encoding="unicode")
    elif raw_ascii:
        text = _decode_text_value(raw_ascii, encoding="ascii")

    if not text or not text.strip():
        return None

    # Get formatting from the style context
    bold = _as_bool(style.get("Bold", False))
    italic = _as_bool(style.get("Italic", False))
    underline = _as_bool(style.get("Underline", False))
    strikethrough = _as_bool(style.get("Strikethrough", False))
    superscript = _as_bool(style.get("Superscript", False))
    subscript = _as_bool(style.get("Subscript", False))
    font = _clean_text(str(style.get("Font", "")))
    font_size = _parse_font_size(style.get("FontSize", 0))

    # Check for HYPERLINK field codes embedded in the text
    # (collapsible sections store URLs as field codes in the text itself)
    has_field_code = "\uFDDF" in text or "\uFDF3" in text
    wz_hyperlink = _clean_text(str(props.get("WzHyperlinkUrl", "")))

    # Check if this is title text
    is_title = _as_bool(props.get("IsTitleText", False))

    # Resolve heading level from ParagraphStyle OSID reference
    heading_level = _resolve_heading_level(props, paragraph_styles)

    # Build text runs — may produce multiple runs when field codes are mixed
    # with regular text that has its own WzHyperlinkUrl.
    runs: list[TextRun] = []
    if has_field_code:
        segments = _parse_hyperlink_field_codes(text)
        for i, (seg_text, seg_url) in enumerate(segments):
            # First segment without a field-code URL inherits WzHyperlinkUrl
            url = seg_url if seg_url else (wz_hyperlink if i == 0 else "")
            runs.append(TextRun(
                text=seg_text,
                bold=bold,
                italic=italic,
                underline=underline,
                strikethrough=strikethrough,
                superscript=superscript,
                subscript=subscript,
                font=font,
                font_size=font_size,
                hyperlink_url=url,
            ))
    else:
        runs.append(TextRun(
            text=text,
            bold=bold,
            italic=italic,
            underline=underline,
            strikethrough=strikethrough,
            superscript=superscript,
            subscript=subscript,
            font=font,
            font_size=font_size,
            hyperlink_url=wz_hyperlink,
        ))

    indent_level = 0
    list_type = ""
    if list_info is not None:
        indent_level = list_info.indent_level
        list_type = list_info.list_type

    return RichText(
        runs=runs,
        indent_level=indent_level,
        is_title=is_title,
        heading_level=heading_level,
        list_type=list_type,
    )


def _resolve_heading_level(
    props: dict[str, object],
    paragraph_styles: dict[str, str] | None,
) -> int:
    """Resolve the heading level from a RichTextOENode's ParagraphStyle.

    The ParagraphStyle property is an OSID reference (list of CompactID
    strings).  Each CompactID maps to a ``jcidParagraphStyleObjectForText``
    whose ``ParagraphStyleId`` is "h1"–"h6", "p", "PageTitle", etc.

    Returns 1–6 for headings, 0 for normal text.
    """
    if not paragraph_styles:
        return 0

    para_style = props.get("ParagraphStyle")
    if not para_style or not isinstance(para_style, list):
        return 0

    # The first entry is the CompactID string for the paragraph style
    style_ref = str(para_style[0])
    style_id = paragraph_styles.get(style_ref, "")

    return _HEADING_STYLE_MAP.get(style_id, 0)


def _extract_image(
    obj: ExtractedObject, file_data: dict[str, bytes]
) -> ImageElement | None:
    """Extract image from an ImageNode."""
    props = obj.properties

    filename = _clean_text(str(props.get("ImageFilename", "")))
    alt_text = _clean_text(str(props.get("ImageAltText", "")))
    width = _parse_int_prop(props.get("PictureWidth", 0))
    height = _parse_int_prop(props.get("PictureHeight", 0))

    # Try to find image data
    data = b""
    pic_container = props.get("PictureContainer")
    if isinstance(pic_container, bytes):
        data = pic_container
    elif isinstance(pic_container, list) and file_data:
        # PictureContainer is a list of identity strings referencing
        # the file data store.  Look up by identity first.
        for ref in pic_container:
            ref_str = str(ref)
            if ref_str in file_data:
                data = file_data[ref_str]
                break

    if not data and not filename:
        return None

    fmt = _detect_image_format(data) if data else ""

    return ImageElement(
        data=data,
        filename=filename or f"image.{fmt or 'bin'}",
        alt_text=alt_text,
        width=width,
        height=height,
        format=fmt,
    )


def _extract_table(
    obj: ExtractedObject,
    objects: list[ExtractedObject],
    table_idx: int,
    style: dict[str, object],
    file_data: dict[str, bytes],
) -> tuple[TableElement | None, int, set[int]]:
    """Extract table with rows and cell content from a TableNode.

    After the TableNode, the flat object list contains:
        TableRowNode → TableCellNode → content* → TableCellNode → content* →
        ... → TableRowNode → ...

    OneNote stores rows bottom-to-top and cells right-to-left, so both
    are reversed to produce the natural reading order.

    Recently edited cells may have their content stored out-of-line
    (earlier in the object list) and referenced by GUID.  A lookup map
    is built so that cells with no inline content can resolve their
    referenced objects.

    Returns (element, consumed) where *consumed* is the number of objects
    after the TableNode that were part of this table.
    """
    props = obj.properties
    row_count = _parse_int_prop(props.get("RowCount", 0))
    col_count = _parse_int_prop(props.get("ColumnCount", 0))
    borders = _as_bool(props.get("TableBordersVisible", True))

    if row_count == 0 or col_count == 0:
        return None, 0, set()

    # Build identity → index lookup so out-of-line cell content can be
    # resolved.  Identity strings look like:
    #   '<ExtendedGUID> (guid-here, 138)'
    identity_map: dict[str, int] = {}
    for idx, o in enumerate(objects):
        if o.identity:
            identity_map[o.identity] = idx

    # Track which objects have been consumed as out-of-line cell content
    # so _build_page can skip them later.
    out_of_line_indices: set[int] = set()

    rows: list[list[list[ContentElement]]] = []
    consumed = 0
    i = table_idx + 1  # start after the TableNode

    while i < len(objects) and len(rows) < row_count:
        row_obj = objects[i]
        if row_obj.obj_type != _TABLE_ROW:
            break

        consumed += 1
        i += 1

        # Collect cells for this row
        row_cells: list[list[ContentElement]] = []
        while i < len(objects) and len(row_cells) < col_count:
            cell_obj = objects[i]
            if cell_obj.obj_type != _TABLE_CELL:
                break

            consumed += 1
            i += 1

            # Each cell's ElementChildNodesOfVersionHistory tells us
            # how many outline groups belong to it and provides GUIDs
            # for out-of-line content resolution.
            child_refs = cell_obj.properties.get(
                "ElementChildNodesOfVersionHistory", []
            )
            max_outlines = len(child_refs) if isinstance(child_refs, list) else 1

            # --- Inline content (objects immediately after the cell) ---
            cell_elements: list[ContentElement] = []
            outlines_seen = 0
            while i < len(objects):
                inner = objects[i]
                if inner.obj_type in (_TABLE_CELL, _TABLE_ROW):
                    break

                if inner.obj_type == _OUTLINE_ELEMENT:
                    outlines_seen += 1
                    if outlines_seen > max_outlines:
                        break

                consumed += 1
                i += 1

                if inner.obj_type == _OUTLINE_ELEMENT:
                    continue
                if inner.obj_type == _RICH_TEXT:
                    elem = _extract_rich_text(inner, style, None)
                    if elem:
                        cell_elements.append(elem)
                elif inner.obj_type == _IMAGE_NODE:
                    elem = _extract_image(inner, file_data)
                    if elem:
                        cell_elements.append(elem)
                elif inner.obj_type == _EMBEDDED_FILE:
                    elem = _extract_embedded_file(inner, file_data)
                    if elem:
                        cell_elements.append(elem)

            # --- Out-of-line content (referenced by GUID) ---
            # If the cell had no inline content, look up its child refs
            # in the identity map to find content stored elsewhere.
            if not cell_elements and isinstance(child_refs, list):
                for ref in child_refs:
                    ref_idx = identity_map.get(ref)
                    if ref_idx is None:
                        continue
                    # Walk forward from the referenced object, collecting
                    # its content (same pattern: OutlineElement → content).
                    j = ref_idx
                    while j < len(objects):
                        robj = objects[j]
                        if robj.obj_type == _OUTLINE_ELEMENT:
                            out_of_line_indices.add(j)
                            j += 1
                            continue
                        if robj.obj_type == _RICH_TEXT:
                            out_of_line_indices.add(j)
                            elem = _extract_rich_text(robj, style, None)
                            if elem:
                                cell_elements.append(elem)
                            j += 1
                            continue
                        if robj.obj_type == _IMAGE_NODE:
                            out_of_line_indices.add(j)
                            elem = _extract_image(robj, file_data)
                            if elem:
                                cell_elements.append(elem)
                            j += 1
                            continue
                        break

            row_cells.append(cell_elements)

        # Cells are stored right-to-left; reverse to natural order
        row_cells.reverse()
        rows.append(row_cells)

    # Rows are stored bottom-to-top; reverse to natural order
    rows.reverse()

    table = TableElement(rows=rows, borders_visible=borders)
    return table, consumed, out_of_line_indices


def _extract_embedded_file(
    obj: ExtractedObject, file_data: dict[str, bytes]
) -> EmbeddedFile | None:
    """Extract embedded file from an EmbeddedFileNode."""
    props = obj.properties
    filename = _clean_text(str(props.get("EmbeddedFileName", "")))
    source_path = _clean_text(str(props.get("SourceFilepath", "")))

    data = b""
    container = props.get("EmbeddedFileContainer")
    if isinstance(container, bytes):
        data = container

    if not filename and not data:
        return None

    return EmbeddedFile(
        data=data,
        filename=filename,
        source_path=source_path,
    )


def _decode_text_value(value: object, encoding: str = "unicode") -> str:
    """Decode text from pyOneNote property values.

    pyOneNote returns text in various formats:
    - Direct string for RichEditTextUnicode
    - Hex string for TextExtendedAscii
    - Garbled UTF-16 decoded strings for TextExtendedAscii (needs re-encoding)
    - bytes for raw data
    """
    if isinstance(value, str):
        cleaned = value.strip()
        if not cleaned:
            return ""

        # Check if it's a hex string (common for TextExtendedAscii)
        if all(c in "0123456789abcdefABCDEF" for c in cleaned):
            try:
                raw = bytes.fromhex(cleaned)
                if encoding == "ascii":
                    return _clean_text(
                        raw.decode("ascii", errors="replace").rstrip("\x00")
                    )
                else:
                    return _clean_text(
                        raw.decode("utf-16-le", errors="replace").rstrip("\x00")
                    )
            except (ValueError, UnicodeDecodeError):
                pass

        # Check for garbled text: pyOneNote sometimes decodes ASCII bytes
        # as UTF-16LE, producing CJK/symbol characters. Detect and fix this.
        # Only apply to ASCII path — Unicode text legitimately contains
        # CJK/Arabic/etc. characters that would be destroyed by re-encoding.
        if encoding == "ascii" and _looks_garbled(cleaned):
            try:
                raw = cleaned.encode("utf-16-le")
                return _clean_text(raw.decode("ascii", errors="replace").rstrip("\x00"))
            except (UnicodeEncodeError, UnicodeDecodeError):
                pass

        return _clean_text(cleaned)

    if isinstance(value, bytes):
        if encoding == "ascii":
            return _clean_text(value.decode("ascii", errors="replace").rstrip("\x00"))
        try:
            return _clean_text(value.decode("utf-16-le").rstrip("\x00"))
        except UnicodeDecodeError:
            return _clean_text(value.decode("latin-1").rstrip("\x00"))

    return _clean_text(str(value)) if value else ""


def _looks_garbled(text: str) -> bool:
    """Detect if text looks like ASCII bytes misinterpreted as UTF-16LE.

    This happens when pyOneNote decodes TextExtendedAscii raw bytes as
    UTF-16 strings instead of ASCII. The result contains CJK characters,
    unusual symbols, and zero-width spaces for normal English text.
    """
    if not text:
        return False
    # Count characters outside normal ASCII+extended range
    non_ascii = sum(1 for c in text if ord(c) > 0xFF)
    # If more than 30% of characters are non-ASCII, it's likely garbled
    return len(text) > 2 and non_ascii / len(text) > 0.3


def _parse_hyperlink_field_codes(text: str) -> list[tuple[str, str]]:
    """Parse text that may contain embedded HYPERLINK field codes.

    OneNote collapsible/outline sections embed hyperlinks as RTF-style
    field codes: U+FDDF HYPERLINK "url" display_text

    Returns a list of (text, url) segments.  Segments without a URL
    have url="".  Plain text returns ``[(text, "")]``.

    Example with mixed content::

        "Meeting options | \\uFDDFHYPERLINK \\"url\\"Reset PIN"
        → [("Meeting options |", ""), ("Reset PIN", "url")]
    """
    if not text:
        return [(text, "")]

    segments: list[tuple[str, str]] = []
    remaining = text

    while remaining:
        match = _HYPERLINK_FIELD_RE.search(remaining)
        if not match:
            cleaned = _clean_text(remaining)
            if cleaned:
                segments.append((cleaned, ""))
            break

        # Text before the field code marker
        prefix = remaining[: match.start()]
        prefix_clean = _clean_text(prefix)
        if prefix_clean:
            segments.append((prefix_clean, ""))

        url = _clean_text(match.group(1))
        display = _clean_text(match.group(2))
        if display:
            segments.append((display, url))
        elif url:
            segments.append((url, url))

        remaining = remaining[match.end() :]

    return segments if segments else [(text, "")]


def _section_name_from_path(file_path: str) -> str:
    """Extract a clean section name from a file path."""
    name = Path(file_path).stem
    name = re.sub(r"\s*\(On\s+\d+-\d+-\d+(?:\s*-\s*\d+)?\)", "", name)
    name = re.sub(r"\.one$", "", name, flags=re.IGNORECASE)
    return name.strip() or "Untitled"


def _clean_text(text: str) -> str:
    """Clean text by removing null bytes, control characters, and replacement chars."""
    # Remove null bytes and vertical tabs (common OneNote artifacts)
    text = text.replace("\x00", "").replace("\x0b", "")
    # Replace narrow no-break space (U+202F) with regular space
    text = text.replace("\u202f", " ")
    # Remove Unicode replacement character (U+FFFD)
    text = text.replace("\ufffd", "")
    return text.strip()


def _as_bool(value: object) -> bool:
    """Convert a property value to bool."""
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.lower() in ("true", "1", "yes")
    return bool(value)


def _parse_int_prop(value: object) -> int:
    """Parse an integer property."""
    if isinstance(value, int):
        return value
    if isinstance(value, bytes) and len(value) >= 2:
        return int.from_bytes(value[:4].ljust(4, b"\x00"), "little")
    if isinstance(value, str):
        match = re.search(r"\d+", value)
        if match:
            return int(match.group())
    return 0


def _parse_byte_prop_as_int(value: object) -> int:
    """Parse a 2-byte property value that may be bytes or a repr() string.

    pyOneNote stores short binary properties as either actual ``bytes``
    or as their ``repr()`` string (e.g. ``"b'$\\x00'"``, ``"b'\\x04\\x00'"``).
    """
    if isinstance(value, int):
        return value
    if isinstance(value, bytes) and len(value) >= 2:
        return int.from_bytes(value[:2], "little")
    if isinstance(value, str):
        # Try to recover bytes from repr string like "b'$\x00'"
        if value.startswith("b'") or value.startswith('b"'):
            try:
                raw = ast.literal_eval(value)
                if isinstance(raw, bytes) and len(raw) >= 2:
                    return int.from_bytes(raw[:2], "little")
            except (ValueError, SyntaxError):
                pass
    return 0


def _parse_font_size(value: object) -> int:
    """Parse font size from pyOneNote format."""
    if isinstance(value, int):
        return value
    if isinstance(value, bytes):
        return int.from_bytes(value[:2].ljust(2, b"\x00"), "little")
    if isinstance(value, str):
        match = re.search(r"\d+", value)
        if match:
            return int(match.group())
    return 0


def _detect_image_format(data: bytes) -> str:
    """Detect image format from magic bytes."""
    if not data or len(data) < 4:
        return ""
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return "png"
    if data[:3] == b"\xff\xd8\xff":
        return "jpeg"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "gif"
    if data[:2] == b"BM":
        return "bmp"
    if data[:4] == b"RIFF" and len(data) > 8 and data[8:12] == b"WEBP":
        return "webp"
    return ""
