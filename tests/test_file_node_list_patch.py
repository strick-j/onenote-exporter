"""Regression tests for the pyOneNote FileNodeList compatibility patch."""

import struct
from io import BytesIO
from types import SimpleNamespace

import pytest
from pyOneNote.FileNode import FileNode, FileNodeList

import onenote_export.parser.one_store  # noqa: F401 - installs compatibility patch


def _reference(stp: int, cb: int) -> SimpleNamespace:
    return SimpleNamespace(stp=stp, cb=cb)


def _empty_fragment() -> bytes:
    """Build a valid empty FileNodeListFragment with an fcrNil successor."""
    return b"\0" * 16 + struct.pack("<QI", 0xFFFFFFFFFFFFFFFF, 0) + b"\0" * 8


class _SeekRecorder:
    def __init__(self) -> None:
        self.offsets: list[int] = []

    def seek(self, offset: int) -> None:
        self.offsets.append(offset)
        raise RuntimeError("delegated FileNodeList parse")


def test_fcr_zero_is_an_empty_child_file_node_list():
    file = _SeekRecorder()

    child_list = FileNodeList(file, object(), _reference(0, 0))

    assert child_list.end == 0
    assert child_list.fragments == []
    assert file.offsets == []


def test_non_empty_file_node_list_still_delegates_to_pyonenote():
    file = BytesIO(_empty_fragment())

    child_list = FileNodeList(file, object(), _reference(0, 36))

    assert child_list.end == 36
    assert len(child_list.fragments) == 1


@pytest.mark.parametrize(
    "reference",
    [
        _reference(0xFFFFFFFFFFFFFFFF, 0),  # fcrNil
        _reference(0, 1),
        _reference(1, 0),
    ],
)
def test_non_fcr_zero_references_are_delegated(reference):
    file = _SeekRecorder()

    with pytest.raises(RuntimeError, match="delegated"):
        FileNodeList(file, object(), reference)

    assert file.offsets == [reference.stp]


def test_malformed_non_empty_reference_still_raises():
    with pytest.raises(struct.error):
        FileNodeList(BytesIO(b""), object(), _reference(0, 1))


def test_malformed_terminator_node_remains_empty():
    # BaseType 2 makes the dependency attempt to recurse through missing data;
    # the existing compatibility patch must continue to absorb that terminator.
    node = FileNode(BytesIO(struct.pack("<I", 2 << 27)), object())

    assert node.data is None
    assert node.children == []
