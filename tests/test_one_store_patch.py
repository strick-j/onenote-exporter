"""Regression tests for the pyOneNote property compatibility patch."""

import struct

import pytest

from onenote_export.parser.one_store import PropertySet


class _Property:
    """Minimal PropertyID stand-in for PropertySet.get_properties()."""

    def __init__(self, name: str, property_type: int) -> None:
        self.name = name
        self.type = property_type

    def __str__(self) -> str:
        return self.name


def _property_set(name: str, value: object, property_type: int) -> PropertySet:
    prop_set = object.__new__(PropertySet)
    prop_set._formated_properties = None
    prop_set.cProperties = 1
    prop_set.rgPrids = [_Property(name, property_type)]
    prop_set.rgData = [value]
    return prop_set


@pytest.mark.parametrize("value", [True, False])
def test_boolean_time_property_is_formatted_as_boolean(value):
    prop_set = _property_set("IsTitleTime", value, 0x02)

    assert prop_set.get_properties() == {"IsTitleTime": str(value)}


def test_ordinary_boolean_property_is_formatted_as_boolean():
    prop_set = _property_set("IsTitleText", True, 0x02)

    assert prop_set.get_properties() == {"IsTitleText": "True"}


def test_four_byte_time_property_keeps_existing_conversion():
    value = struct.pack("<I", 0)
    prop_set = _property_set("CreationTimeStamp", value, 0x05)

    assert prop_set.get_properties() == {
        "CreationTimeStamp": str(PropertySet.time32_to_datetime(0))
    }


def test_eight_byte_time_property_keeps_existing_conversion():
    value = struct.pack("<Q", 0)
    prop_set = _property_set("LastModifiedTime", value, 0x06)

    assert prop_set.get_properties() == {
        "LastModifiedTime": str(PropertySet.parse_filetime(0))
    }


def test_unsupported_time_bytes_are_preserved_as_hex():
    value = b"\x01\x02"
    prop_set = _property_set("CreationTimeStamp", value, 0x04)

    assert prop_set.get_properties() == {"CreationTimeStamp": value.hex()}


def test_non_byte_time_value_is_stringified():
    prop_set = _property_set("CreationTimeStamp", "unknown", 0x07)

    assert prop_set.get_properties() == {"CreationTimeStamp": "unknown"}


def test_formatted_properties_are_cached():
    prop_set = _property_set("IsTitleTime", True, 0x02)

    first = prop_set.get_properties()
    prop_set.rgData[0] = False
    second = prop_set.get_properties()

    assert first is second
    assert second == {"IsTitleTime": "True"}
