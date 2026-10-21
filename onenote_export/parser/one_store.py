"""High-level OneNote parser using pyOneNote as the binary parsing engine.

Extracts structured content (text, images, formatting) from .one files
and organizes it by page.
"""

import logging
import re
import struct
import uuid
from dataclasses import dataclass, field
from pathlib import Path

from pyOneNote.FileNode import ObjectSpaceObjectPropSet, PropertyID, PropertySet
from pyOneNote.Header import Header
from pyOneNote.OneDocument import OneDocment

logger = logging.getLogger(__name__)


def _patch_pyonenote() -> None:
    """Monkey-patch pyOneNote to fix bugs and add missing features.

    Fixes applied:
    1. ``ObjectSpaceObjectStreamOfIDs.read()`` never increments its
       ``head`` pointer, so all OSID references within the same
       PropertySet resolve to the first entry.  This breaks properties
       like ``ListNodes`` when multiple OSIDs are present.

    2. ``PropertySet.__init__`` raises ``NotImplementedError`` for
       property type 0x10 (prtArrayOfPropertyValues).  Per the
       MS-ONESTORE spec (section 2.6.9), the format is:

           cProperties : uint32  — number of child PropertySets
           prid        : PropertyID (4 bytes) — only if cProperties > 0
           Data        : cProperties consecutive PropertySet structures

    3. ``FileNode.__init__`` crashes on a FileNodeListFragment
       *terminator* node.  A terminator has ``file_node_id == 0`` (or
       255), which is not in pyOneNote's ``_FileNodeIDs`` map, so its
       type resolves to ``"Invalid"`` and the dispatch chain's silent
       ``else`` branch never sets ``self.data``.  The constructor then
       unconditionally recurses for reference nodes
       (``if baseType == 2: FileNodeList(..., self.data.ref)``) and
       dereferences the missing ``self.data``.  The fragment loop *does*
       break on ``id in (0, 255)`` — but only *after* the node is
       constructed, so the crash beats the guard.  Files whose root
       fragment ends on a "dirty" terminator (id 0 with non-zero upper
       bits, e.g. large/heavily-revised sections) fail to parse at all.

       We wrap the constructor so a terminator node that raises during
       construction is treated as empty (``data = None``, no children);
       the existing fragment-loop break then ends the fragment cleanly.
    """
    from pyOneNote.FileNode import FileNode as _FileNode
    from pyOneNote.FileNode import (
        ObjectSpaceObjectStreamOfIDs,
        PrtFourBytesOfLengthFollowedByData,
    )

    _original_read = ObjectSpaceObjectStreamOfIDs.read

    def _patched_read(self):
        res = None
        if self.head < len(self.body):
            res = self.body[self.head]
            self.head += 1
        return res

    ObjectSpaceObjectStreamOfIDs.read = _patched_read
    _original_init = PropertySet.__init__

    def _patched_init(
        self, file, OIDs=None, OSIDs=None, ContextIDs=None, document=None
    ):
        self.current = file.tell()
        (self.cProperties,) = struct.unpack("<H", file.read(2))
        self.rgPrids = []
        self.indent = ""
        self.document = document
        self.current_revision = document.cur_revision if document else None
        self._formated_properties = None

        for _i in range(self.cProperties):
            self.rgPrids.append(PropertyID(file))

        self.rgData = []
        for i in range(self.cProperties):
            ptype = self.rgPrids[i].type
            if ptype == 0x1:
                self.rgData.append(None)
            elif ptype == 0x2:
                self.rgData.append(self.rgPrids[i].boolValue)
            elif ptype == 0x3:
                self.rgData.append(struct.unpack("c", file.read(1))[0])
            elif ptype == 0x4:
                self.rgData.append(struct.unpack("2s", file.read(2))[0])
            elif ptype == 0x5:
                self.rgData.append(struct.unpack("4s", file.read(4))[0])
            elif ptype == 0x6:
                self.rgData.append(struct.unpack("8s", file.read(8))[0])
            elif ptype == 0x7:
                from pyOneNote.FileNode import PrtFourBytesOfLengthFollowedByData

                self.rgData.append(PrtFourBytesOfLengthFollowedByData(file, self))
            elif ptype in (0x8, 0x9):
                count = 1
                if ptype == 0x9:
                    (count,) = struct.unpack("<I", file.read(4))
                self.rgData.append(PropertySet.get_compact_ids(OIDs, count))
            elif ptype in (0xA, 0xB):
                count = 1
                if ptype == 0xB:
                    (count,) = struct.unpack("<I", file.read(4))
                self.rgData.append(PropertySet.get_compact_ids(OSIDs, count))
            elif ptype in (0xC, 0xD):
                count = 1
                if ptype == 0xD:
                    (count,) = struct.unpack("<I", file.read(4))
                self.rgData.append(PropertySet.get_compact_ids(ContextIDs, count))
            elif ptype == 0x10:
                # ArrayOfPropertyValues (MS-ONESTORE section 2.6.9)
                (arr_count,) = struct.unpack("<I", file.read(4))
                child_sets = []
                if arr_count > 0:
                    # Read prid (must have type 0x11); we validate but
                    # don't use it — each child is a full PropertySet.
                    _arr_prid = PropertyID(file)
                    for _j in range(arr_count):
                        child = PropertySet(
                            file,
                            OIDs,
                            OSIDs,
                            ContextIDs,
                            document,
                        )
                        child_sets.append(child)
                self.rgData.append(child_sets)
            elif ptype == 0x11:
                self.rgData.append(PropertySet(file, OIDs, OSIDs, ContextIDs, document))
            else:
                raise ValueError(f"rgPrids[{i}].type 0x{ptype:x} is not valid")

    PropertySet.__init__ = _patched_init

    def _patched_get_properties(self):
        if self._formated_properties is not None:
            return self._formated_properties

        self._formated_properties = {}
        for i in range(self.cProperties):
            property_name = str(self.rgPrids[i])
            if property_name == "Unknown":
                continue

            raw_value = self.rgData[i]
            property_value = ""
            if isinstance(raw_value, PrtFourBytesOfLengthFollowedByData):
                if "guid" in property_name.lower():
                    property_value = uuid.UUID(bytes_le=raw_value.Data).hex
                else:
                    try:
                        property_value = raw_value.Data.decode("utf-16")
                    except UnicodeError:
                        property_value = raw_value.Data.hex()
            elif getattr(self.rgPrids[i], "type", None) == 0x02 or isinstance(
                raw_value, bool
            ):
                property_value = str(raw_value)
            else:
                property_name_lower = property_name.lower()
                if "time" in property_name_lower:
                    if isinstance(raw_value, (bytes, bytearray, memoryview)):
                        if len(raw_value) == 8:
                            (timestamp_in_nano,) = struct.unpack("<Q", raw_value)
                            property_value = str(
                                PropertySet.parse_filetime(timestamp_in_nano)
                            )
                        elif len(raw_value) == 4:
                            (timestamp_in_sec,) = struct.unpack("<I", raw_value)
                            property_value = str(
                                PropertySet.time32_to_datetime(timestamp_in_sec)
                            )
                        else:
                            property_value = bytes(raw_value).hex()
                    else:
                        property_value = str(raw_value)
                elif (
                    "height" in property_name_lower
                    or "width" in property_name_lower
                    or "offset" in property_name_lower
                    or "margin" in property_name_lower
                ):
                    (size,) = struct.unpack("<f", raw_value)
                    property_value = PropertySet.half_inch_size_to_pixels(size)
                elif "langid" in property_name_lower:
                    (lcid,) = struct.unpack("<H", raw_value)
                    property_value = f"{PropertySet.lcid_to_string(lcid)}({lcid})"
                elif "languageid" in property_name_lower:
                    (lcid,) = struct.unpack("<I", raw_value)
                    property_value = f"{PropertySet.lcid_to_string(lcid)}({lcid})"
                else:
                    if isinstance(raw_value, list):
                        property_value = [str(item) for item in raw_value]
                    else:
                        property_value = str(raw_value)

            self._formated_properties[property_name] = property_value
        return self._formated_properties

    PropertySet.get_properties = _patched_get_properties

    _original_filenode_init = _FileNode.__init__

    def _patched_filenode_init(self, file, document):
        try:
            _original_filenode_init(self, file, document)
        except AttributeError:
            # Only swallow the terminator case (id 0/255); anything else
            # is a genuine parse error and must propagate.
            header = getattr(self, "file_node_header", None)
            fnid = getattr(header, "file_node_id", None)
            if fnid in (0, 255):
                if not hasattr(self, "data"):
                    self.data = None
                if not hasattr(self, "children"):
                    self.children = []
            else:
                raise

    _FileNode.__init__ = _patched_filenode_init


# Apply patches at import time
_patch_pyonenote()

# JCID type names from the OneNote spec
_PAGE_META = "jcidPageMetaData"
_SECTION_NODE = "jcidSectionNode"
_SECTION_META = "jcidSectionMetaData"
_PAGE_SERIES = "jcidPageSeriesNode"
_PAGE_MANIFEST = "jcidPageManifestNode"
_PAGE_NODE = "jcidPageNode"
_TITLE_NODE = "jcidTitleNode"
_OUTLINE_NODE = "jcidOutlineNode"
_OUTLINE_ELEMENT = "jcidOutlineElementNode"
_RICH_TEXT = "jcidRichTextOENode"
_IMAGE_NODE = "jcidImageNode"
_TABLE_NODE = "jcidTableNode"
_TABLE_ROW = "jcidTableRowNode"
_TABLE_CELL = "jcidTableCellNode"
_EMBEDDED_FILE = "jcidEmbeddedFileNode"
_NUMBER_LIST = "jcidNumberListNode"
_STYLE_CONTAINER = "jcidPersistablePropertyContainerForTOCSection"
_REVISION_META = "jcidRevisionMetaData"


@dataclass
class ExtractedProperty:
    """A single property from a OneNote object."""

    name: str
    value: object  # str, bytes, int, bool, list, etc.


@dataclass
class ExtractedObject:
    """A parsed object from the OneNote file."""

    obj_type: str
    identity: str
    properties: dict[str, object] = field(default_factory=dict)


@dataclass
class ExtractedPage:
    """A page with its title and content objects."""

    title: str = ""
    level: int = 0
    author: str = ""
    creation_time: str = ""
    last_modified: str = ""
    objects: list[ExtractedObject] = field(default_factory=list)


@dataclass
class ExtractedSection:
    """All pages extracted from a single .one file."""

    file_path: str = ""
    display_name: str = ""
    pages: list[ExtractedPage] = field(default_factory=list)
    file_data: dict[str, bytes] = field(default_factory=dict)
    paragraph_styles: dict[str, str] = field(default_factory=dict)


class OneStoreParser:
    """Parses a MS-ONESTORE (.one) file using pyOneNote."""

    def __init__(self, file_path: str | Path) -> None:
        self.file_path = Path(file_path)

    def parse(self) -> ExtractedSection:
        """Parse the .one file and return structured content."""
        section = ExtractedSection(file_path=str(self.file_path))

        with open(self.file_path, "rb") as f:
            doc = OneDocment(f)

        # Validate it's a .one file
        if doc.header.guidFileType != Header.ONE_UUID:
            raise ValueError(f"{self.file_path} is not a .one file")

        # Get all properties (objects with their property sets)
        raw_props = doc.get_properties()

        # Get embedded files — store by both GUID and identity string
        # so that PictureContainer references (which use identity) can
        # resolve to the correct file data.
        raw_files = doc.get_files()
        for guid, finfo in raw_files.items():
            content = finfo.get("content", b"")
            if content:
                section.file_data[guid] = content
                identity = finfo.get("identity", "")
                if identity:
                    section.file_data[identity] = content

        # Extract paragraph styles from ReadOnly object declarations
        section.paragraph_styles = self._extract_paragraph_styles(doc)

        # Convert raw properties to ExtractedObjects
        all_objects = []
        for raw in raw_props:
            obj = ExtractedObject(
                obj_type=raw["type"],
                identity=raw["identity"],
                properties=dict(raw["val"]),
            )
            all_objects.append(obj)

        # Build page structure
        section.pages = self._build_pages(all_objects)
        section.display_name = self._extract_section_name(all_objects)

        return section

    def _extract_paragraph_styles(
        self,
        doc: OneDocment,
    ) -> dict[str, str]:
        """Extract paragraph style IDs from ReadOnly object declarations.

        pyOneNote traverses ``ObjectDeclaration2RefCountFND`` nodes but
        skips ``ReadOnlyObjectDeclaration2RefCountFND`` nodes which
        contain the ``jcidParagraphStyleObjectForText`` property sets.

        This method finds those ReadOnly nodes, seeks to their data in
        the file, parses the PropertySet, and builds a mapping from the
        object's identity string to its ``ParagraphStyleId`` value.
        """
        all_nodes: list[object] = []
        OneDocment.traverse_nodes(doc.root_file_node_list, all_nodes, [])

        styles: dict[str, str] = {}
        for node in all_nodes:
            data = getattr(node, "data", None)
            if data is None:
                continue
            if type(data).__name__ != "ReadOnlyObjectDeclaration2RefCountFND":
                continue

            base = data.base
            ref_stp = base.ref.stp
            ref_cb = base.ref.cb
            if ref_stp == 0 or ref_cb == 0:
                continue

            try:
                with open(self.file_path, "rb") as f:
                    f.seek(ref_stp)
                    prop_set = ObjectSpaceObjectPropSet(f, doc)
                    props = dict(prop_set.body.get_properties())
            except Exception:
                continue

            style_id = props.get("ParagraphStyleId", "")
            if not style_id:
                continue

            oid = str(base.body.oid)
            # Clean the null terminator from the style ID
            clean_id = style_id.replace("\x00", "").strip()
            styles[oid] = clean_id

        return styles

    def _extract_section_name(self, objects: list[ExtractedObject]) -> str:
        """Extract section display name from section metadata."""
        for obj in objects:
            if obj.obj_type == _SECTION_META:
                name = obj.properties.get("SectionDisplayName", "")
                if name:
                    return str(name).strip()
        return ""

    def _build_pages(self, objects: list[ExtractedObject]) -> list[ExtractedPage]:
        """Group objects into pages based on the document structure.

        OneNote stores multiple revisions per page, each sharing a GUID.
        Content objects (text, images, etc.) share the GUID of their
        owning page node.  Orphan page metadata entries (from older
        revisions) have a different GUID with no associated content.

        Strategy:
        1. Identify "content GUIDs" — GUIDs that own a jcidPageNode
           (these are the actual page revisions with content objects).
        2. Build one page per content GUID using its metadata + content.
        3. If a content GUID has no metadata, fall back to matching
           orphan metadata by title.
        """
        pages: list[ExtractedPage] = []

        # Classify every object by type, indexed by GUID
        guid_objects: dict[str, list[ExtractedObject]] = {}
        page_metas: list[ExtractedObject] = []
        page_node_guids: list[str] = []

        for obj in objects:
            guid = _extract_guid(obj.identity)
            guid_objects.setdefault(guid, []).append(obj)

            if obj.obj_type == _PAGE_META:
                page_metas.append(obj)
            elif obj.obj_type == _PAGE_NODE:
                if guid not in page_node_guids:
                    page_node_guids.append(guid)

        # No page metadata at all — single unnamed page
        if not page_metas:
            all_content = [
                o
                for o in objects
                if o.obj_type
                in (
                    _RICH_TEXT,
                    _IMAGE_NODE,
                    _TABLE_NODE,
                    _TABLE_ROW,
                    _TABLE_CELL,
                    _EMBEDDED_FILE,
                    _OUTLINE_ELEMENT,
                    _OUTLINE_NODE,
                    _NUMBER_LIST,
                )
            ]
            if all_content:
                pages.append(ExtractedPage(objects=all_content))
            return pages

        # Build a lookup: GUID -> page metadata.
        #
        # A page's object space contains MULTIPLE metadata revisions.  The
        # first one encountered in traversal order is the current revision;
        # later ones are stale copies carrying an outdated PageLevel and/or
        # title (from before the page was renamed or promoted/demoted to a
        # subpage).  Keep the first and never overwrite — overwriting with a
        # stale copy corrupts both the subpage level and the title.
        meta_by_guid: dict[str, ExtractedObject] = {}
        for meta in page_metas:
            guid = _extract_guid(meta.identity)
            if guid not in meta_by_guid:
                meta_by_guid[guid] = meta

        # Orphan metas: metadata GUIDs with no page node (old revisions)
        orphan_metas: dict[str, ExtractedObject] = {
            g: m for g, m in meta_by_guid.items() if g not in page_node_guids
        }

        # Content types we care about
        _CONTENT_TYPES = {
            _RICH_TEXT,
            _IMAGE_NODE,
            _TABLE_NODE,
            _TABLE_ROW,
            _TABLE_CELL,
            _EMBEDDED_FILE,
            _OUTLINE_ELEMENT,
            _OUTLINE_NODE,
            _NUMBER_LIST,
        }

        # Build one page per content GUID (GUID that has a PageNode).
        # Each unique page-node GUID is a distinct page; revisions of the
        # *same* page already share a GUID and are collapsed by the
        # ``page_node_guids`` uniqueness above.  We deliberately do NOT
        # deduplicate by title — distinct pages legitimately share titles
        # (e.g. repeated "Example Output" subpages), and collapsing them
        # silently drops real content.
        for content_guid in page_node_guids:
            objs = guid_objects.get(content_guid, [])

            # Find metadata — prefer same GUID, fall back to orphan
            meta = meta_by_guid.get(content_guid)
            if not meta:
                # Match orphan metadata by scanning (first available)
                for og, om in list(orphan_metas.items()):
                    meta = om
                    del orphan_metas[og]
                    break

            title = ""
            level = 0
            creation = ""
            if meta:
                title = _clean_text(str(meta.properties.get("CachedTitleString", "")))
                level = _parse_int(meta.properties.get("PageLevel", 0))
                creation = str(meta.properties.get("TopologyCreationTimeStamp", ""))

            # Extract author from the page node
            author = ""
            last_modified = ""
            for o in objs:
                if o.obj_type == _PAGE_NODE:
                    author = _clean_text(str(o.properties.get("Author", "")))
                    last_modified = str(o.properties.get("LastModifiedTime", ""))
                    break

            # Collect content objects for this GUID only
            content = [o for o in objs if o.obj_type in _CONTENT_TYPES]

            page = ExtractedPage(
                title=title or "Untitled",
                level=level,
                author=author,
                creation_time=creation,
                last_modified=last_modified,
                objects=content,
            )

            pages.append(page)

        return pages


def _extract_guid(identity_str: str) -> str:
    """Extract the GUID from an ExtendedGUID identity string.

    Input format: '<ExtendedGUID> (guid-string, n)'
    Returns just the guid-string part.
    """
    match = re.search(r"\(([^,]+),", identity_str)
    if match:
        return match.group(1).strip()
    return ""


def _clean_text(text: str) -> str:
    """Clean up text by stripping null bytes and extra whitespace."""
    text = text.replace("\x00", "").strip()
    return text


def _parse_int(value: object) -> int:
    """Parse an integer from various formats pyOneNote returns."""
    if isinstance(value, int):
        return value
    if isinstance(value, bytes):
        try:
            return int.from_bytes(value[:4], "little")
        except Exception:
            return 0
    if isinstance(value, str):
        # Try to extract numeric value
        match = re.search(r"\d+", value)
        if match:
            return int(match.group())
    return 0
