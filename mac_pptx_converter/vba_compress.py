"""
VBA compression/decompression per MS-OVBA 2.4.1 spec.

The VBA source code in OLE streams is stored using a custom compression
algorithm. This module handles both decompression (reading existing code)
and compression (writing modified code back).
"""

import struct


def decompress_stream(compressed_data: bytes) -> bytes:
    if len(compressed_data) < 1:
        return b""
    signature = compressed_data[0]
    if signature != 0x01:
        raise ValueError(f"Invalid compression signature: 0x{signature:02x}")
    pos = 1
    result = bytearray()
    while pos < len(compressed_data):
        if pos + 2 > len(compressed_data):
            break
        chunk_header = struct.unpack_from("<H", compressed_data, pos)[0]
        pos += 2
        chunk_size = (chunk_header & 0x0FFF) + 3
        chunk_is_compressed = (chunk_header >> 15) & 1
        chunk_data = compressed_data[pos:pos + chunk_size - 2]
        if chunk_is_compressed:
            _decompress_chunk(chunk_data, result)
        else:
            result.extend(chunk_data)
        pos += chunk_size - 2
    return bytes(result)


def _decompress_chunk(chunk_data: bytes, result: bytearray):
    chunk_start = len(result)
    pos = 0
    while pos < len(chunk_data):
        flag_byte = chunk_data[pos]
        pos += 1
        for bit_index in range(8):
            if pos >= len(chunk_data):
                return
            if (flag_byte >> bit_index) & 1 == 0:
                result.append(chunk_data[pos])
                pos += 1
            else:
                if pos + 1 >= len(chunk_data):
                    return
                copy_token = struct.unpack_from("<H", chunk_data, pos)[0]
                pos += 2
                decompressed_current = len(result)
                bit_count = _max_bit_count(decompressed_current, chunk_start)
                length_mask = 0xFFFF >> bit_count
                offset_mask = ~length_mask & 0xFFFF
                length = (copy_token & length_mask) + 3
                offset = ((copy_token & offset_mask) >> (16 - bit_count)) + 1
                for _ in range(length):
                    src_pos = len(result) - offset
                    if src_pos < 0:
                        result.append(0)
                    else:
                        result.append(result[src_pos])


def _max_bit_count(decompressed_current: int, decompressed_chunk_start: int) -> int:
    diff = decompressed_current - decompressed_chunk_start
    if diff <= 0:
        return 4
    bit_count = diff.bit_length()
    if bit_count < 4:
        return 4
    if bit_count > 12:
        return 12
    return bit_count


def compress_stream(data: bytes) -> bytes:
    if len(data) == 0:
        return b"\x01"
    result = bytearray([0x01])
    pos = 0
    while pos < len(data):
        chunk_start = pos
        chunk_end = min(pos + 4096, len(data))
        chunk = data[chunk_start:chunk_end]
        compressed_chunk = _compress_chunk(chunk)
        chunk_size = len(compressed_chunk) + 2
        chunk_header = ((chunk_size - 3) & 0x0FFF) | 0xB000
        result.extend(struct.pack("<H", chunk_header))
        result.extend(compressed_chunk)
        pos = chunk_end
    return bytes(result)


def _compress_chunk(data: bytes) -> bytes:
    result = bytearray()
    pos = 0
    while pos < len(data):
        flag_byte = 0
        flag_pos = len(result)
        result.append(0)
        for bit_index in range(8):
            if pos >= len(data):
                break
            best_len, best_offset = _find_match(data, pos)
            if best_len >= 3:
                flag_byte |= (1 << bit_index)
                bit_count = max(4, (pos).bit_length()) if pos > 0 else 4
                if bit_count > 12:
                    bit_count = 12
                length_mask = 0xFFFF >> bit_count
                offset_shifted = (best_offset - 1) << (16 - bit_count)
                length_field = best_len - 3
                if length_field > length_mask:
                    length_field = length_mask
                    best_len = length_field + 3
                copy_token = offset_shifted | length_field
                result.extend(struct.pack("<H", copy_token))
                pos += best_len
            else:
                result.append(data[pos])
                pos += 1
        result[flag_pos] = flag_byte
    return bytes(result)


def _find_match(data: bytes, pos: int) -> tuple:
    best_len = 0
    best_offset = 0
    max_offset = min(pos, 4096)
    for offset in range(1, max_offset + 1):
        match_len = 0
        while (pos + match_len < len(data) and
               match_len < 4096 and
               data[pos + match_len] == data[pos - offset + match_len]):
            match_len += 1
        if match_len > best_len:
            best_len = match_len
            best_offset = offset
            if match_len > 100:
                break
    return best_len, best_offset
