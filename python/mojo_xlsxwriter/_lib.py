"""ctypes bindings for the batched shared-string kernels."""

from __future__ import annotations

import ctypes
import os
import subprocess
from collections.abc import Sequence
from itertools import chain

import numpy as np

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LIB = os.path.join(ROOT, "dist", "libmojo-xlsxwriter.so")
I = ctypes.c_int64

_SIGNATURES = {
    "mxw_dedup_utf8": ([I] * 10, I),
    "mxw_escape_xml": ([I] * 9, I),
}
_library: ctypes.CDLL | None = None


class BuildError(RuntimeError):
    pass


class _StringTable(dict[str, int]):
    def __missing__(self, value: str) -> int:
        identifier = len(self)
        self[value] = identifier
        return identifier


def build(force: bool = False) -> str:
    source = os.path.join(ROOT, "src", "kernels.mojo")
    if not force and os.path.exists(LIB) and os.path.getmtime(LIB) >= os.path.getmtime(source):
        return LIB
    proc = subprocess.run(
        ["pixi", "run", "--manifest-path", os.path.join(ROOT, "pixi.toml"), "build"],
        capture_output=True,
        text=True,
        timeout=1800,
    )
    if proc.returncode or not os.path.exists(LIB):
        raise BuildError((proc.stderr or proc.stdout).strip()[:4000])
    return LIB


def lib() -> ctypes.CDLL:
    global _library
    if _library is None:
        _library = ctypes.CDLL(build())
        for name, (argtypes, restype) in _SIGNATURES.items():
            function = getattr(_library, name)
            function.argtypes = argtypes
            function.restype = restype
    return _library


def _addr(array: np.ndarray) -> int:
    address = int(array.ctypes.data)
    if not address:
        raise ValueError("cannot pass a null buffer to Mojo")
    return address


def _pack(strings: Sequence[str]) -> tuple[np.ndarray, np.ndarray]:
    encoded = [value.encode("utf-8") for value in strings]
    offsets = np.fromiter(
        chain((0,), map(len, encoded)), dtype=np.int64, count=len(encoded) + 1
    )
    np.cumsum(offsets, out=offsets)
    joined = b"".join(encoded)
    if joined:
        data = np.frombuffer(joined, dtype=np.uint8)
    else:
        data = np.zeros(1, dtype=np.uint8)
    return data, offsets


def deduplicate_packed(
    data: np.ndarray, offsets: np.ndarray
) -> tuple[np.ndarray, np.ndarray]:
    """Deduplicate contiguous UTF-8 bytes without copying either input buffer."""
    if data.dtype != np.uint8 or data.ndim != 1 or not data.flags.c_contiguous:
        raise TypeError("data must be a contiguous one-dimensional uint8 array")
    if offsets.dtype != np.int64 or offsets.ndim != 1 or not offsets.flags.c_contiguous:
        raise TypeError("offsets must be a contiguous one-dimensional int64 array")
    count = len(offsets) - 1
    if count < 0 or int(offsets[0]) != 0 or int(offsets[-1]) != data.size:
        raise ValueError("offsets must span the complete data buffer")
    if count == 0:
        return np.empty(0, dtype=np.int64), np.empty(0, dtype=np.int64)
    if np.any(offsets[1:] < offsets[:-1]):
        raise ValueError("offsets must be nondecreasing")
    capacity = 2
    while capacity < count * 2:
        capacity *= 2
    slots = np.full(capacity, -1, dtype=np.int64)
    ids = np.empty(count, dtype=np.int64)
    unique_count = lib().mxw_dedup_utf8(
        _addr(data),
        data.size,
        _addr(offsets),
        offsets.size,
        count,
        _addr(slots),
        capacity,
        slots.size,
        _addr(ids),
        ids.size,
    )
    if unique_count < 0 or unique_count > count:
        raise RuntimeError(f"Mojo shared-string deduplication failed ({unique_count})")
    _, unique_indices = np.unique(ids, return_index=True)
    if unique_indices.size != unique_count:
        raise RuntimeError("Mojo shared-string table returned inconsistent ids")
    return ids, unique_indices


def deduplicate(strings: Sequence[str]) -> tuple[np.ndarray, list[str]]:
    """Return an insertion-order shared-string id for every occurrence."""
    count = len(strings)
    if not count:
        return np.empty(0, dtype=np.int64), []
    table = _StringTable()
    identifiers = np.fromiter(
        map(table.__getitem__, strings), dtype=np.int64, count=count
    )
    return identifiers, list(table)


def escape_xml(strings: Sequence[str]) -> list[str]:
    """Escape XML and Excel control sequences for a batch of strings."""
    count = len(strings)
    if not count:
        return []
    data, offsets = _pack(strings)
    escaped_offsets = np.empty(count + 1, dtype=np.int64)
    capacity = max(1, int(offsets[-1]) * 7)
    destination = np.empty(capacity, dtype=np.uint8)
    size = lib().mxw_escape_xml(
        _addr(data),
        int(offsets[-1]),
        _addr(offsets),
        offsets.size,
        count,
        _addr(escaped_offsets),
        escaped_offsets.size,
        _addr(destination),
        destination.size,
    )
    if size < 0 or size > capacity:
        raise RuntimeError(f"Mojo XML escape returned invalid size {size}")
    raw = destination[:size].tobytes()
    return [
        raw[int(escaped_offsets[i]) : int(escaped_offsets[i + 1])].decode("utf-8")
        for i in range(count)
    ]
