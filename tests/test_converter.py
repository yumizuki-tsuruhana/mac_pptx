"""Tests for the end-to-end converter with synthetic PPTM files."""

import os
import struct
import tempfile
import unittest
import zipfile

from mac_pptx_converter.converter import convert_file, format_report
from mac_pptx_converter.vba_compress import compress_stream
from mac_pptx_converter.ole_builder import OLEBuilder


def _build_dir_stream(modules: list[dict]) -> bytes:
    """Build a minimal VBA dir stream with module definitions."""
    data = bytearray()

    data.extend(struct.pack("<HI", 0x0001, 4))
    data.extend(struct.pack("<I", 0x00000001))

    data.extend(struct.pack("<HI", 0x0002, 4))
    data.extend(struct.pack("<I", 0x0409))

    data.extend(struct.pack("<HI", 0x0014, 4))
    data.extend(struct.pack("<I", 0x0409))

    data.extend(struct.pack("<HI", 0x0003, 2))
    data.extend(struct.pack("<H", 65001))

    name = b"VBAProject"
    data.extend(struct.pack("<HI", 0x0004, len(name)))
    data.extend(name)

    data.extend(struct.pack("<HI", 0x000F, 2))
    data.extend(struct.pack("<H", len(modules)))

    data.extend(struct.pack("<HI", 0x0013, 2))
    data.extend(struct.pack("<H", 0xFFFF))

    for mod in modules:
        mod_name = mod["name"].encode("ascii")
        data.extend(struct.pack("<HI", 0x0019, len(mod_name)))
        data.extend(mod_name)

        mod_name_uni = mod["name"].encode("utf-16-le")
        data.extend(struct.pack("<HI", 0x0047, len(mod_name_uni)))
        data.extend(mod_name_uni)

        data.extend(struct.pack("<HI", 0x001A, len(mod_name)))
        data.extend(mod_name)

        data.extend(struct.pack("<HI", 0x0032, len(mod_name_uni)))
        data.extend(mod_name_uni)

        text_offset = mod.get("text_offset", 0)
        data.extend(struct.pack("<HI", 0x0031, 4))
        data.extend(struct.pack("<I", text_offset))

        data.extend(struct.pack("<HI", 0x002B, 0))

    data.extend(struct.pack("<HI", 0x0010, 0))

    return bytes(data)


def _create_test_pptm(path: str, modules: dict[str, str]):
    """Create a synthetic .pptm file with VBA modules for testing."""
    ole_buf = _create_vba_ole(modules)

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '</Types>'
        ))
        zf.writestr("_rels/.rels", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
            '</Relationships>'
        ))
        zf.writestr("ppt/presentation.xml", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
        ))
        zf.writestr("ppt/vbaProject.bin", ole_buf)


def _create_vba_ole(modules: dict[str, str]) -> bytes:
    """Create a minimal OLE file containing VBA modules."""
    mod_list = [{"name": name, "text_offset": 0} for name in modules]
    dir_data = _build_dir_stream(mod_list)
    compressed_dir = compress_stream(dir_data)

    builder = OLEBuilder()
    builder.add_stream("VBA/dir", compressed_dir)
    builder.add_stream("VBA/_VBA_PROJECT", b"\xCC\x61\xFF\xFF\x00\x00\x00")

    for name, source in modules.items():
        source_bytes = source.encode("utf-8")
        compressed_source = compress_stream(source_bytes)
        builder.add_stream(f"VBA/{name}", compressed_source)

    return builder.build()


class TestConverterEndToEnd(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.mkdtemp()

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    def test_simple_conversion(self):
        input_path = os.path.join(self.tmpdir, "test.pptm")
        output_path = os.path.join(self.tmpdir, "test_mac.pptm")

        _create_test_pptm(input_path, {
            "Module1": (
                'Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)\r\n'
                "\r\n"
                "Sub WaitABit()\r\n"
                "    Sleep 1000\r\n"
                "End Sub\r\n"
            ),
        })

        result = convert_file(input_path, output_path)
        self.assertTrue(result.vba_found)
        self.assertTrue(os.path.exists(output_path))
        self.assertGreater(result.total_issues, 0)
        self.assertGreater(result.fixed_count, 0)

    def test_no_vba(self):
        input_path = os.path.join(self.tmpdir, "test.pptx")
        output_path = os.path.join(self.tmpdir, "test_mac.pptx")

        with zipfile.ZipFile(input_path, "w") as zf:
            zf.writestr("[Content_Types].xml", "<Types/>")
            zf.writestr("ppt/presentation.xml", "<p/>")

        result = convert_file(input_path, output_path)
        self.assertFalse(result.vba_found)

    def test_clean_code_no_changes(self):
        input_path = os.path.join(self.tmpdir, "test.pptm")
        output_path = os.path.join(self.tmpdir, "test_mac.pptm")

        _create_test_pptm(input_path, {
            "Module1": (
                "Sub Hello()\r\n"
                "    MsgBox \"Hello!\"\r\n"
                "End Sub\r\n"
            ),
        })

        result = convert_file(input_path, output_path)
        self.assertTrue(result.vba_found)
        self.assertEqual(result.total_issues, 0)

    def test_multiple_issues(self):
        input_path = os.path.join(self.tmpdir, "test.pptm")
        output_path = os.path.join(self.tmpdir, "test_mac.pptm")

        _create_test_pptm(input_path, {
            "Module1": (
                'Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)\r\n'
                'path = Environ("USERPROFILE") & "\\file.txt"\r\n'
                'Set fso = CreateObject("Scripting.FileSystemObject")\r\n'
                'Shell "cmd.exe /c dir"\r\n'
            ),
        })

        result = convert_file(input_path, output_path)
        self.assertGreater(result.total_issues, 3)

    def test_report_format(self):
        input_path = os.path.join(self.tmpdir, "test.pptm")
        _create_test_pptm(input_path, {
            "Module1": 'Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)\r\n',
        })
        result = convert_file(input_path)
        report_text = format_report(result)
        self.assertIn("Conversion Report", report_text)
        self.assertIn("Module1", report_text)

    def test_file_not_found(self):
        result = convert_file("/nonexistent/file.pptm")
        self.assertGreater(len(result.errors), 0)

    def test_unsupported_extension(self):
        result = convert_file("test.txt")
        self.assertGreater(len(result.errors), 0)


if __name__ == "__main__":
    unittest.main()
