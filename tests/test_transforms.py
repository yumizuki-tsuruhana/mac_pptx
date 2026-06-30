"""Tests for VBA source code transformations."""

import unittest
from mac_pptx_converter.transforms import apply_all_transforms


class TestAPIDeclarations(unittest.TestCase):
    def test_sleep_declaration(self):
        source = 'Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)'
        result, report = apply_all_transforms(source, "Module1")
        self.assertIn("#If Mac Then", result)
        self.assertIn("#Else", result)
        self.assertIn("#End If", result)
        self.assertIn("Timer", result)
        self.assertTrue(any(r.category == "API Declaration" for r in report.results))

    def test_gettickcount_declaration(self):
        source = 'Declare Function GetTickCount Lib "kernel32" () As Long'
        result, report = apply_all_transforms(source, "Module1")
        self.assertIn("#If Mac Then", result)
        self.assertIn("Timer", result)

    def test_unknown_api_wrapped(self):
        source = 'Declare Function SomeFunc Lib "user32" (ByVal x As Long) As Long'
        result, report = apply_all_transforms(source, "Module1")
        self.assertIn("#If Not Mac Then", result)
        self.assertIn("no Mac equivalent", result)

    def test_private_declare(self):
        source = 'Private Declare Function GetTickCount Lib "kernel32" () As Long'
        result, report = apply_all_transforms(source, "Module1")
        self.assertIn("#If Mac Then", result)

    def test_ptrsafe_declare(self):
        source = 'Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long'
        result, report = apply_all_transforms(source, "Module1")
        self.assertIn("#If Mac Then", result)


class TestFilePaths(unittest.TestCase):
    def test_windows_path_conversion(self):
        source = 'path = "C:\\Users\\foo\\file.txt"'
        result, report = apply_all_transforms(source, "Module1")
        self.assertNotIn("C:\\", result)
        self.assertIn("/Users/foo/file.txt", result)

    def test_backslash_separator(self):
        source = 'fullPath = folder & "\\" & filename'
        result, report = apply_all_transforms(source, "Module1")
        self.assertIn("Application.PathSeparator", result)

    def test_comments_not_touched(self):
        source = "' This is a comment with C:\\path"
        result, _ = apply_all_transforms(source, "Module1")
        self.assertEqual(source, result)


class TestCreateObject(unittest.TestCase):
    def test_fso_detected(self):
        source = 'Set fso = CreateObject("Scripting.FileSystemObject")'
        _, report = apply_all_transforms(source, "Module1")
        self.assertTrue(any(
            r.category == "COM Object" and "FileSystemObject" in r.description
            for r in report.results
        ))

    def test_xmlhttp_detected(self):
        source = 'Set http = CreateObject("MSXML2.XMLHTTP")'
        _, report = apply_all_transforms(source, "Module1")
        self.assertTrue(any(r.category == "COM Object" for r in report.results))

    def test_wscript_shell_detected(self):
        source = 'Set sh = CreateObject("WScript.Shell")'
        _, report = apply_all_transforms(source, "Module1")
        self.assertTrue(any(
            "WScript.Shell" in r.description for r in report.results
        ))


class TestShellCalls(unittest.TestCase):
    def test_shell_function(self):
        source = 'Shell "cmd.exe /c dir", vbHide'
        _, report = apply_all_transforms(source, "Module1")
        self.assertTrue(any(r.category == "Shell Command" for r in report.results))

    def test_shell_paren(self):
        source = 'x = Shell("cmd.exe /c echo hello")'
        _, report = apply_all_transforms(source, "Module1")
        self.assertTrue(any(r.category == "Shell Command" for r in report.results))


class TestSendKeys(unittest.TestCase):
    def test_sendkeys_detected(self):
        source = 'SendKeys "%{F4}", True'
        _, report = apply_all_transforms(source, "Module1")
        self.assertTrue(any(r.category == "SendKeys" for r in report.results))


class TestEnviron(unittest.TestCase):
    def test_userprofile_mapped(self):
        source = 'path = Environ("USERPROFILE")'
        result, report = apply_all_transforms(source, "Module1")
        self.assertIn('"HOME"', result)

    def test_temp_mapped(self):
        source = 'path = Environ$("TEMP")'
        result, report = apply_all_transforms(source, "Module1")
        self.assertIn('"TMPDIR"', result)

    def test_windir_flagged(self):
        source = 'path = Environ("WINDIR")'
        _, report = apply_all_transforms(source, "Module1")
        self.assertTrue(any(r.severity == "error" for r in report.results))


class TestRegistry(unittest.TestCase):
    def test_registry_read(self):
        source = 'val = sh.RegRead("HKEY_CURRENT_USER\\Software\\MyApp\\Setting")'
        _, report = apply_all_transforms(source, "Module1")
        self.assertTrue(any(r.category == "Registry Access" for r in report.results))

    def test_getsetting(self):
        source = 'val = GetSetting("MyApp", "Settings", "Value")'
        _, report = apply_all_transforms(source, "Module1")
        self.assertTrue(any(r.category == "Registry Access" for r in report.results))


class TestClipboard(unittest.TestCase):
    def test_dataobject_detected(self):
        source = "Dim cb As New MSForms.DataObject"
        _, report = apply_all_transforms(source, "Module1")
        self.assertTrue(any(r.category == "Clipboard" for r in report.results))


class TestActiveX(unittest.TestCase):
    def test_treeview_detected(self):
        source = "Dim tv As MSComctlLib.TreeView"
        _, report = apply_all_transforms(source, "Module1")
        self.assertTrue(any(r.category == "ActiveX Control" for r in report.results))


class TestNoFalsePositives(unittest.TestCase):
    def test_clean_code(self):
        source = (
            "Sub Hello()\n"
            "    MsgBox \"Hello from Mac!\"\n"
            "End Sub"
        )
        result, report = apply_all_transforms(source, "Module1")
        self.assertEqual(result, source)
        self.assertEqual(len(report.results), 0)


if __name__ == "__main__":
    unittest.main()
