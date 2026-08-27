from __future__ import annotations

import ctypes
import random
import string

import numpy as np
import pytest
from xlsxwriter.sharedstrings import SharedStringTable
from xlsxwriter.xmlwriter import XMLwriter

from mojo_xlsxwriter._lib import deduplicate, deduplicate_packed, escape_xml, lib


def upstream_escape(values):
    writer = XMLwriter()
    return [
        writer._escape_data(writer._escape_control_characters(value))
        for value in values
    ]


def test_deduplicate_matches_upstream_insertion_order():
    values = [
        "alpha",
        "beta",
        "alpha",
        "",
        "βeta",
        "Unicode café",
        "beta",
        "",
    ]
    table = SharedStringTable()
    expected = [table._get_shared_string_index(value) for value in values]
    ids, uniques = deduplicate(values)
    assert ids.tolist() == expected
    assert uniques == ["alpha", "beta", "", "βeta", "Unicode café"]


def test_deduplicate_random_collision_heavy_data():
    rng = random.Random(19)
    vocabulary = [
        "".join(rng.choices(string.ascii_letters + "β中", k=rng.randrange(0, 25)))
        for _ in range(300)
    ]
    values = [rng.choice(vocabulary) for _ in range(20_003)]
    expected_table = {}
    expected_ids = []
    expected_unique = []
    for value in values:
        if value not in expected_table:
            expected_table[value] = len(expected_unique)
            expected_unique.append(value)
        expected_ids.append(expected_table[value])
    ids, uniques = deduplicate(values)
    assert ids.tolist() == expected_ids
    assert uniques == expected_unique


def test_packed_deduplicate_is_zero_copy_and_handles_simd_tail():
    values = ["abcdefghijk", "abcdefghijx", "abcdefghijk", "z", "z"]
    encoded = [value.encode() for value in values]
    data = np.frombuffer(b"".join(encoded), dtype=np.uint8)
    offsets = np.asarray(
        [0, *np.cumsum([len(value) for value in encoded])], dtype=np.int64
    )
    data_address = data.ctypes.data
    offsets_address = offsets.ctypes.data
    ids, unique_indices = deduplicate_packed(data, offsets)
    assert data.ctypes.data == data_address
    assert offsets.ctypes.data == offsets_address
    assert ids.tolist() == [0, 1, 0, 2, 2]
    assert unique_indices.tolist() == [0, 1, 3]


def test_xml_escape_exact_upstream_parity():
    values = [
        "",
        "ordinary UTF-8 β中 café",
        "& < > \" stays",
        " leading",
        "trailing ",
        "\x00\x01\x08\x09\x0a\x0b\x0c\x0d\x0e\x1f",
        "_x0000_ _xABcd_ _X0000_",
        "\ufffe\uffff",
        "a&<b\x01_x0000_",
    ]
    assert escape_xml(values) == upstream_escape(values)


def test_xml_escape_random_scalar_and_vector_tails():
    rng = random.Random(37)
    alphabet = "abcXYZ&<>_x09AF" + "".join(chr(i) for i in range(32)) + "β中"
    values = [
        "".join(rng.choice(alphabet) for _ in range(length))
        for length in range(65)
    ]
    assert escape_xml(values) == upstream_escape(values)


@pytest.mark.parametrize("count", [16_383, 16_384])
def test_xml_escape_parallel_threshold_matches_upstream(count):
    values = [
        f"{index}&<\x01_x0000_β中" if index % 11 == 0 else f"value-{index}"
        for index in range(count)
    ]
    assert escape_xml(values) == upstream_escape(values)


def test_empty_batches_do_not_cross_null_pointers():
    ids, unique = deduplicate([])
    assert ids.dtype == np.int64
    assert ids.size == 0
    assert unique == []
    assert escape_xml([]) == []
    assert escape_xml(["", ""]) == ["", ""]


def test_ffi_exports_have_int64_contract():
    library = lib()
    assert library.mxw_dedup_utf8.restype is ctypes.c_int64
    assert library.mxw_escape_xml.restype is ctypes.c_int64
    assert library.mxw_dedup_utf8.argtypes == [ctypes.c_int64] * 10
    assert library.mxw_escape_xml.argtypes == [ctypes.c_int64] * 9


def test_packed_deduplicate_rejects_invalid_layouts_before_ffi():
    data = np.arange(8, dtype=np.uint8)
    valid = np.array([0, 4, 8], dtype=np.int64)
    for offsets in (
        np.array([0, 5, 4, 8], dtype=np.int64),
        np.array([0, -1, 8], dtype=np.int64),
        np.array([0, 4, 7], dtype=np.int64),
    ):
        with pytest.raises(ValueError):
            deduplicate_packed(data, offsets)
    with pytest.raises(TypeError):
        deduplicate_packed(data[::2], valid)
    with pytest.raises(TypeError):
        deduplicate_packed(data, valid.astype(np.int32))


def test_ffi_rejects_bad_lengths_and_output_capacity():
    data, offsets = np.frombuffer(b"&", dtype=np.uint8), np.array([0, 1], dtype=np.int64)
    escaped_offsets = np.empty(2, dtype=np.int64)
    destination = np.empty(4, dtype=np.uint8)
    result = lib().mxw_escape_xml(
        data.ctypes.data,
        data.size,
        offsets.ctypes.data,
        offsets.size,
        1,
        escaped_offsets.ctypes.data,
        escaped_offsets.size,
        destination.ctypes.data,
        destination.size,
    )
    assert result == -3
