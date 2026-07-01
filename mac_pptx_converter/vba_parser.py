"""
Parse and modify VBA projects inside OLE compound files (vbaProject.bin).

Handles reading module source code from the dir stream and module streams,
and writing modified source code back into the OLE container.
"""

from __future__ import annotations

import struct
import io
import os
import sys

try:
    import olefile
except ImportError:
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), "_vendor"))
    import olefile

from .vba_compress import decompress_stream, compress_stream
from .ole_builder import OLEBuilder


class VBAModule:
    def __init__(self, name: str, stream_name: str, text_offset: int, source: str):
        self.name = name
        self.stream_name = stream_name
        self.text_offset = text_offset
        self.source = source
        self.original_source = source


class VBAProject:
    def __init__(self, ole_data: bytes):
        self.ole_data = ole_data
        self.modules: list[VBAModule] = []
        self._all_streams: dict[str, bytes] = {}
        self._parse()

    def _parse(self):
        ole = olefile.OleFileIO(io.BytesIO(self.ole_data))
        try:
            for entry in ole.listdir():
                path = "/".join(entry)
                try:
                    self._all_streams[path] = ole.openstream(path).read()
                except Exception:
                    pass
            self._parse_dir(ole)
            self._read_module_sources(ole)
        finally:
            ole.close()

    def _parse_dir(self, ole: olefile.OleFileIO):
        dir_stream = ole.openstream("VBA/dir").read()
        decompressed = decompress_stream(dir_stream)
        pos = 0
        modules_to_read = []
        while pos < len(decompressed) - 5:
            record_id = struct.unpack_from("<H", decompressed, pos)[0]
            record_size = struct.unpack_from("<I", decompressed, pos + 2)[0]
            pos += 6
            if pos + record_size > len(decompressed):
                break
            if record_id == 0x0019:
                module_name = decompressed[pos:pos + record_size].decode("ascii", errors="replace")
                modules_to_read.append({"name": module_name, "stream_name": module_name})
            elif record_id == 0x001A:
                if modules_to_read:
                    modules_to_read[-1]["stream_name"] = decompressed[pos:pos + record_size].decode(
                        "ascii", errors="replace"
                    )
            elif record_id == 0x0031:
                text_offset = struct.unpack_from("<I", decompressed, pos)[0]
                if modules_to_read:
                    modules_to_read[-1]["text_offset"] = text_offset
            pos += record_size

        for mod_info in modules_to_read:
            self.modules.append(VBAModule(
                name=mod_info["name"],
                stream_name=mod_info.get("stream_name", mod_info["name"]),
                text_offset=mod_info.get("text_offset", 0),
                source="",
            ))

    def _read_module_sources(self, ole: olefile.OleFileIO):
        for module in self.modules:
            stream_path = f"VBA/{module.stream_name}"
            try:
                stream_data = ole.openstream(stream_path).read()
            except Exception:
                continue
            compressed_source = stream_data[module.text_offset:]
            try:
                source_bytes = decompress_stream(compressed_source)
                module.source = source_bytes.decode("utf-8", errors="replace")
                module.original_source = module.source
            except Exception:
                module.source = ""
                module.original_source = ""

    def save(self) -> bytes:
        builder = OLEBuilder()
        for path, data in self._all_streams.items():
            is_module = False
            for module in self.modules:
                if path == f"VBA/{module.stream_name}":
                    is_module = True
                    if module.source != module.original_source:
                        original_data = data
                        performance_cache = original_data[:module.text_offset]
                        new_source_bytes = module.source.encode("utf-8")
                        new_compressed = compress_stream(new_source_bytes)
                        data = performance_cache + new_compressed
                    break
            if "__SRP_" in path:
                continue
            builder.add_stream(path, data)
        return builder.build()
