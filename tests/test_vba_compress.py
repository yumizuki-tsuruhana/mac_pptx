"""Tests for VBA compression/decompression round-trip."""

import unittest
from mac_pptx_converter.vba_compress import compress_stream, decompress_stream


class TestVBACompression(unittest.TestCase):
    def test_round_trip_simple(self):
        original = b"Hello World"
        compressed = compress_stream(original)
        decompressed = decompress_stream(compressed)
        self.assertEqual(decompressed, original)

    def test_round_trip_vba_code(self):
        original = (
            b"Sub Hello()\r\n"
            b"    MsgBox \"Hello World\"\r\n"
            b"End Sub\r\n"
        )
        compressed = compress_stream(original)
        decompressed = decompress_stream(compressed)
        self.assertEqual(decompressed, original)

    def test_round_trip_repeated_content(self):
        original = b"AAAA" * 100
        compressed = compress_stream(original)
        decompressed = decompress_stream(compressed)
        self.assertEqual(decompressed, original)
        self.assertLess(len(compressed), len(original))

    def test_round_trip_large(self):
        original = b"Sub Test()\r\n" * 500
        compressed = compress_stream(original)
        decompressed = decompress_stream(compressed)
        self.assertEqual(decompressed, original)

    def test_empty(self):
        original = b""
        compressed = compress_stream(original)
        decompressed = decompress_stream(compressed)
        self.assertEqual(decompressed, original)

    def test_signature_byte(self):
        compressed = compress_stream(b"test")
        self.assertEqual(compressed[0], 0x01)

    def test_round_trip_unicode_like(self):
        original = "日本語テスト".encode("utf-8")
        compressed = compress_stream(original)
        decompressed = decompress_stream(compressed)
        self.assertEqual(decompressed, original)


if __name__ == "__main__":
    unittest.main()
