"""
Minimal OLE Compound File Binary builder.

olefile can only overwrite existing streams at the same size,
so we need our own builder to create/modify VBA projects where
stream sizes change after transformation.
"""

from __future__ import annotations

import struct
import io
from dataclasses import dataclass, field

SECTOR_SIZE = 512
MINI_SECTOR_SIZE = 64
MINI_STREAM_CUTOFF = 0x1000
ENDOFCHAIN = -2  # 0xFFFFFFFE
FREESECT = -1  # 0xFFFFFFFF
FATSECT = -3  # 0xFFFFFFFD
DIFSECT = -4  # 0xFFFFFFFC


@dataclass
class _DirEntry:
    name: str
    entry_type: int  # 1=storage, 2=stream, 5=root
    children: list = field(default_factory=list)
    data: bytes = b""
    sid: int = -1


class OLEBuilder:
    """Build an OLE Compound File from a tree of storages and streams."""

    def __init__(self):
        self._root = _DirEntry(name="Root Entry", entry_type=5)
        self._entries: list[_DirEntry] = [self._root]

    def add_stream(self, path: str, data: bytes):
        parts = path.split("/")
        parent = self._root
        for part in parts[:-1]:
            child = self._find_child(parent, part)
            if child is None:
                child = _DirEntry(name=part, entry_type=1)
                parent.children.append(child)
                self._entries.append(child)
            parent = child
        stream = _DirEntry(name=parts[-1], entry_type=2, data=data)
        parent.children.append(stream)
        self._entries.append(stream)

    def _find_child(self, parent: _DirEntry, name: str) -> _DirEntry | None:
        for child in parent.children:
            if child.name.lower() == name.lower():
                return child
        return None

    def build(self) -> bytes:
        self._assign_sids()
        sectors: list[bytearray] = []

        mini_stream_data = bytearray()
        mini_fat_entries: list[int] = []

        for entry in self._entries:
            if entry.entry_type == 2 and len(entry.data) < MINI_STREAM_CUTOFF:
                entry._mini_start = len(mini_stream_data) // MINI_SECTOR_SIZE
                num_mini = (len(entry.data) + MINI_SECTOR_SIZE - 1) // MINI_SECTOR_SIZE
                padded = entry.data + b"\x00" * (num_mini * MINI_SECTOR_SIZE - len(entry.data))
                mini_stream_data += padded
                for j in range(num_mini):
                    idx = entry._mini_start + j
                    if j < num_mini - 1:
                        mini_fat_entries.append(idx + 1)
                    else:
                        mini_fat_entries.append(ENDOFCHAIN)
            else:
                entry._mini_start = -1

        fat: list[int] = []

        dir_entries_data = self._build_dir_entries()
        num_dir_sectors = (len(dir_entries_data) + SECTOR_SIZE - 1) // SECTOR_SIZE
        dir_start = len(sectors)
        for i in range(num_dir_sectors):
            chunk = dir_entries_data[i * SECTOR_SIZE:(i + 1) * SECTOR_SIZE]
            chunk = chunk + b"\x00" * (SECTOR_SIZE - len(chunk))
            sectors.append(bytearray(chunk))
            fat.append(dir_start + i + 1 if i < num_dir_sectors - 1 else ENDOFCHAIN)

        for entry in self._entries:
            if entry.entry_type == 2 and entry._mini_start == -1 and len(entry.data) > 0:
                num_sectors_needed = (len(entry.data) + SECTOR_SIZE - 1) // SECTOR_SIZE
                entry._sector_start = len(sectors)
                for j in range(num_sectors_needed):
                    chunk = entry.data[j * SECTOR_SIZE:(j + 1) * SECTOR_SIZE]
                    chunk = chunk + b"\x00" * (SECTOR_SIZE - len(chunk))
                    sectors.append(bytearray(chunk))
                    fat.append(entry._sector_start + j + 1 if j < num_sectors_needed - 1 else ENDOFCHAIN)
            elif entry.entry_type != 2:
                entry._sector_start = -1

        if mini_stream_data:
            mini_stream_start = len(sectors)
            num_ms_sectors = (len(mini_stream_data) + SECTOR_SIZE - 1) // SECTOR_SIZE
            for j in range(num_ms_sectors):
                chunk = mini_stream_data[j * SECTOR_SIZE:(j + 1) * SECTOR_SIZE]
                chunk = chunk + b"\x00" * (SECTOR_SIZE - len(chunk))
                sectors.append(bytearray(chunk))
                fat.append(mini_stream_start + j + 1 if j < num_ms_sectors - 1 else ENDOFCHAIN)
            self._root._sector_start = mini_stream_start
            self._root._size = len(mini_stream_data)
        else:
            self._root._sector_start = ENDOFCHAIN
            self._root._size = 0

        if mini_fat_entries:
            num_mf_sectors = (len(mini_fat_entries) * 4 + SECTOR_SIZE - 1) // SECTOR_SIZE
            mini_fat_start = len(sectors)
            total_mf_slots = num_mf_sectors * (SECTOR_SIZE // 4)
            padded_mf = mini_fat_entries + [FREESECT] * (total_mf_slots - len(mini_fat_entries))
            for j in range(num_mf_sectors):
                sec = bytearray(SECTOR_SIZE)
                for k in range(SECTOR_SIZE // 4):
                    idx = j * (SECTOR_SIZE // 4) + k
                    struct.pack_into("<i", sec, k * 4, padded_mf[idx])
                sectors.append(sec)
                fat.append(mini_fat_start + j + 1 if j < num_mf_sectors - 1 else ENDOFCHAIN)
            num_mini_fat_sectors = num_mf_sectors
        else:
            mini_fat_start = ENDOFCHAIN
            num_mini_fat_sectors = 0

        num_fat_sectors = max(1, (len(fat) + 2 + SECTOR_SIZE // 4 - 1) // (SECTOR_SIZE // 4))
        while True:
            total_entries = len(fat) + num_fat_sectors
            needed = max(1, (total_entries + SECTOR_SIZE // 4 - 1) // (SECTOR_SIZE // 4))
            if needed == num_fat_sectors:
                break
            num_fat_sectors = needed

        fat_start = len(sectors)
        for j in range(num_fat_sectors):
            fat.append(FATSECT)
            sectors.append(bytearray(SECTOR_SIZE))

        remaining_free = num_fat_sectors * (SECTOR_SIZE // 4) - len(fat)
        fat.extend([FREESECT] * remaining_free)

        for j in range(num_fat_sectors):
            sec = sectors[fat_start + j]
            for k in range(SECTOR_SIZE // 4):
                idx = j * (SECTOR_SIZE // 4) + k
                if idx < len(fat):
                    struct.pack_into("<i", sec, k * 4, fat[idx])

        self._update_dir_entries(sectors, dir_start, mini_stream_data)

        header = self._build_header(
            dir_start=dir_start,
            fat_start=fat_start,
            num_fat_sectors=num_fat_sectors,
            mini_fat_start=mini_fat_start,
            num_mini_fat_sectors=num_mini_fat_sectors,
        )

        result = bytearray(header)
        for sec in sectors:
            result.extend(sec)
        return bytes(result)

    def _assign_sids(self):
        for i, entry in enumerate(self._entries):
            entry.sid = i

    def _build_dir_entries(self) -> bytes:
        """Build initial dir entries (start sectors filled in later)."""
        data = bytearray()
        for entry in self._entries:
            data.extend(self._make_dir_entry(entry))
        return bytes(data)

    def _make_dir_entry(self, entry: _DirEntry) -> bytes:
        buf = bytearray(128)
        name_bytes = entry.name.encode("utf-16-le")
        name_len = min(len(name_bytes), 62)
        buf[0:name_len] = name_bytes[:name_len]
        buf[name_len:name_len + 2] = b"\x00\x00"
        struct.pack_into("<H", buf, 0x40, name_len + 2)
        buf[0x42] = entry.entry_type
        buf[0x43] = 1  # black node

        left_sid = -1
        right_sid = -1
        child_sid = -1

        if entry.children:
            child_sid = self._build_tree(entry.children)

        struct.pack_into("<i", buf, 0x44, left_sid)
        struct.pack_into("<i", buf, 0x48, right_sid)
        struct.pack_into("<i", buf, 0x4C, child_sid)

        struct.pack_into("<i", buf, 0x74, ENDOFCHAIN)
        struct.pack_into("<I", buf, 0x78, 0)

        return bytes(buf)

    def _build_tree(self, children: list[_DirEntry]) -> int:
        if not children:
            return -1
        sorted_children = sorted(children, key=lambda c: c.name.lower())
        mid = len(sorted_children) // 2
        root = sorted_children[mid]
        if mid > 0:
            root._left_sid = self._build_tree(sorted_children[:mid])
        else:
            root._left_sid = -1
        if mid + 1 < len(sorted_children):
            root._right_sid = self._build_tree(sorted_children[mid + 1:])
        else:
            root._right_sid = -1
        return root.sid

    def _update_dir_entries(self, sectors: list[bytearray], dir_start: int,
                            mini_stream_data: bytes):
        for entry in self._entries:
            offset_in_dir = entry.sid * 128
            sec_idx = dir_start + (offset_in_dir // SECTOR_SIZE)
            sec_offset = offset_in_dir % SECTOR_SIZE
            buf = sectors[sec_idx]

            name_bytes = entry.name.encode("utf-16-le")
            name_len = min(len(name_bytes), 62)
            for i in range(64):
                buf[sec_offset + i] = 0
            buf[sec_offset:sec_offset + name_len] = name_bytes[:name_len]
            struct.pack_into("<H", buf, sec_offset + 0x40, name_len + 2)
            buf[sec_offset + 0x42] = entry.entry_type
            buf[sec_offset + 0x43] = 1

            left_sid = getattr(entry, "_left_sid", -1)
            right_sid = getattr(entry, "_right_sid", -1)

            child_sid = -1
            if entry.children:
                child_sid = self._build_tree(entry.children)

            struct.pack_into("<i", buf, sec_offset + 0x44, left_sid)
            struct.pack_into("<i", buf, sec_offset + 0x48, right_sid)
            struct.pack_into("<i", buf, sec_offset + 0x4C, child_sid)

            if entry.entry_type == 5:
                start = getattr(entry, "_sector_start", ENDOFCHAIN)
                size = getattr(entry, "_size", 0)
                if start == -1:
                    start = ENDOFCHAIN
                struct.pack_into("<i", buf, sec_offset + 0x74, start)
                struct.pack_into("<I", buf, sec_offset + 0x78, size)
            elif entry.entry_type == 2:
                if entry._mini_start >= 0:
                    struct.pack_into("<i", buf, sec_offset + 0x74, entry._mini_start)
                else:
                    start = getattr(entry, "_sector_start", ENDOFCHAIN)
                    if start == -1:
                        start = ENDOFCHAIN
                    struct.pack_into("<i", buf, sec_offset + 0x74, start)
                struct.pack_into("<I", buf, sec_offset + 0x78, len(entry.data))
            else:
                struct.pack_into("<i", buf, sec_offset + 0x74, ENDOFCHAIN)
                struct.pack_into("<I", buf, sec_offset + 0x78, 0)

    def _build_header(self, dir_start: int, fat_start: int,
                      num_fat_sectors: int, mini_fat_start: int,
                      num_mini_fat_sectors: int) -> bytes:
        header = bytearray(SECTOR_SIZE)
        header[0:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
        struct.pack_into("<H", header, 0x18, 0x003E)
        struct.pack_into("<H", header, 0x1A, 0x0003)
        struct.pack_into("<H", header, 0x1C, 0xFFFE)
        struct.pack_into("<H", header, 0x1E, 0x0009)
        struct.pack_into("<H", header, 0x20, 0x0006)
        struct.pack_into("<I", header, 0x2C, num_fat_sectors)
        struct.pack_into("<I", header, 0x30, dir_start)
        struct.pack_into("<I", header, 0x38, MINI_STREAM_CUTOFF)
        struct.pack_into("<i", header, 0x3C, mini_fat_start)
        struct.pack_into("<I", header, 0x40, num_mini_fat_sectors)
        struct.pack_into("<i", header, 0x44, ENDOFCHAIN)
        struct.pack_into("<I", header, 0x48, 0)

        for i in range(109):
            offset = 0x4C + i * 4
            if i < num_fat_sectors:
                struct.pack_into("<I", header, offset, fat_start + i)
            else:
                struct.pack_into("<i", header, offset, FREESECT)

        return bytes(header)
