"""
VBA source code transformations for Windows → Mac compatibility.

Each transform function takes VBA source code and returns (modified_source, list_of_changes).
Changes are reported as TransformResult objects for the conversion report.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


@dataclass
class TransformResult:
    category: str
    line_number: int
    original: str
    replacement: str
    description: str
    severity: str = "warning"


@dataclass
class TransformReport:
    module_name: str
    results: list[TransformResult] = field(default_factory=list)

    @property
    def has_changes(self) -> bool:
        return len(self.results) > 0


def apply_all_transforms(source: str, module_name: str) -> tuple[str, TransformReport]:
    report = TransformReport(module_name=module_name)
    source = _transform_api_declares(source, report)
    source = _transform_file_paths(source, report)
    source = _transform_createobject(source, report)
    source = _transform_shell_calls(source, report)
    source = _transform_sendkeys(source, report)
    source = _transform_clipboard(source, report)
    source = _transform_environ(source, report)
    source = _transform_registry(source, report)
    source = _transform_activex_references(source, report)
    return source, report


# --- Windows API Declare transforms ---

_WIN_API_PATTERN = re.compile(
    r'^(\s*)((?:Public|Private)\s+)?Declare\s+(PtrSafe\s+)?'
    r'(Function|Sub)\s+(\w+)\s+Lib\s+"([^"]+)"'
    r'(?:\s+Alias\s+"([^"]+)")?\s*(\([^)]*\))(.*)$',
    re.MULTILINE | re.IGNORECASE,
)

_MAC_ALTERNATIVES = {
    "sleep": (
        "Public Sub Sleep(ByVal ms As Long)\n"
        "    Dim endTime As Double\n"
        "    endTime = Timer + ms / 1000#\n"
        "    Do While Timer < endTime\n"
        "        DoEvents\n"
        "    Loop\n"
        "End Sub"
    ),
    "gettickcount": (
        "Public Function GetTickCount() As Long\n"
        "    GetTickCount = CLng(Timer * 1000)\n"
        "End Function"
    ),
    "gettemppatha": (
        "Public Function GetTempPathA(ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long\n"
        "    Dim tmpPath As String\n"
        "    tmpPath = Environ$(\"TMPDIR\")\n"
        "    If tmpPath = \"\" Then tmpPath = \"/tmp/\"\n"
        "    lpBuffer = tmpPath\n"
        "    GetTempPathA = Len(tmpPath)\n"
        "End Function"
    ),
    "gettempfilenamea": (
        "Public Function GetTempFileNameA(ByVal lpszPath As String, ByVal lpPrefixString As String, "
        "ByVal wUnique As Long, ByVal lpTempFileName As String) As Long\n"
        "    lpTempFileName = lpszPath & lpPrefixString & Format(Int(Rnd * 99999), \"00000\") & \".tmp\"\n"
        "    GetTempFileNameA = 1\n"
        "End Function"
    ),
    "shgetfolderpath": (
        "Public Function SHGetFolderPath(ByVal hwnd As Long, ByVal csidl As Long, "
        "ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long\n"
        "    pszPath = Environ$(\"HOME\")\n"
        "    SHGetFolderPath = 0\n"
        "End Function"
    ),
}


def _transform_api_declares(source: str, report: TransformReport) -> str:
    lines = source.split("\n")
    new_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        match = _WIN_API_PATTERN.match(line)
        if match:
            indent = match.group(1)
            func_name = match.group(5).lower()
            lib_name = match.group(6).lower()

            mac_alt_key = func_name
            mac_alt = _MAC_ALTERNATIVES.get(mac_alt_key)

            if mac_alt:
                replacement = (
                    f"{indent}#If Mac Then\n"
                    + "\n".join(f"{indent}{l}" for l in mac_alt.split("\n"))
                    + f"\n{indent}#Else\n"
                    f"{indent}{line.strip()}\n"
                    f"{indent}#End If"
                )
                report.results.append(TransformResult(
                    category="API Declaration",
                    line_number=i + 1,
                    original=line.strip(),
                    replacement=f"Conditional compilation with Mac alternative for {func_name}",
                    description=f"Added Mac-compatible implementation for {func_name} (Lib \"{lib_name}\")",
                    severity="fixed",
                ))
            else:
                replacement = (
                    f"{indent}#If Not Mac Then\n"
                    f"{indent}{line.strip()}\n"
                    f"{indent}#End If\n"
                    f"{indent}' NOTE: {func_name} from {lib_name} has no Mac equivalent - calls will fail on Mac"
                )
                report.results.append(TransformResult(
                    category="API Declaration",
                    line_number=i + 1,
                    original=line.strip(),
                    replacement=f"Wrapped in #If Not Mac (no Mac alternative for {func_name})",
                    description=(
                        f"{func_name} from \"{lib_name}\" is Windows-only. "
                        "Wrapped in conditional compilation. Manual Mac alternative needed."
                    ),
                    severity="error",
                ))
            new_lines.append(replacement)
        else:
            new_lines.append(line)
        i += 1
    return "\n".join(new_lines)


# --- File path transforms ---

_HARDCODED_WIN_PATH = re.compile(
    r'("([A-Za-z]:\\[^"]*)")',
)

_BACKSLASH_SEPARATOR = re.compile(
    r'(\s*&\s*"\\"\s*&\s*)',
)


def _transform_file_paths(source: str, report: TransformReport) -> str:
    lines = source.split("\n")
    new_lines = []
    for i, line in enumerate(lines):
        original_line = line
        stripped = line.lstrip()
        if stripped.startswith("'") or stripped.upper().startswith("REM "):
            new_lines.append(line)
            continue

        for match in _HARDCODED_WIN_PATH.finditer(line):
            win_path = match.group(2)
            mac_path = _convert_win_path(win_path)
            old_val = match.group(1)
            new_val = f'"{mac_path}"'

            report.results.append(TransformResult(
                category="File Path",
                line_number=i + 1,
                original=old_val,
                replacement=new_val,
                description=f"Converted Windows path to Unix-style path",
                severity="fixed",
            ))
            line = line.replace(old_val, new_val, 1)

        for match in _BACKSLASH_SEPARATOR.finditer(line):
            old_val = match.group(1)
            new_val = ' & Application.PathSeparator & '
            if old_val != new_val:
                report.results.append(TransformResult(
                    category="File Path",
                    line_number=i + 1,
                    original=old_val.strip(),
                    replacement=new_val.strip(),
                    description="Replaced hardcoded backslash separator with Application.PathSeparator",
                    severity="fixed",
                ))
                line = line.replace(old_val, new_val, 1)

        new_lines.append(line)
    return "\n".join(new_lines)


def _convert_win_path(win_path: str) -> str:
    path = win_path.replace("\\", "/")
    if len(path) >= 2 and path[1] == ":":
        drive = path[0].upper()
        rest = path[2:]
        if drive == "C":
            known_mappings = {
                "/Users/": "/Users/",
                "/Temp": "${TMPDIR}",
                "/Windows/Temp": "${TMPDIR}",
            }
            for win_prefix, mac_prefix in known_mappings.items():
                if rest.startswith(win_prefix):
                    return mac_prefix + rest[len(win_prefix):]
            return rest
        return f"/Volumes/{drive}{rest}"
    return path


# --- CreateObject transforms ---

_CREATEOBJECT_PATTERN = re.compile(
    r'CreateObject\(\s*"([^"]+)"\s*\)',
    re.IGNORECASE,
)

_COM_REPLACEMENTS = {
    "scripting.filesystemobject": {
        "description": "FileSystemObject is not available on Mac. Use VBA native file I/O (Dir, Open, Kill, etc.)",
        "severity": "error",
    },
    "msxml2.xmlhttp": {
        "description": "MSXML2.XMLHTTP not available on Mac. Use 'MacScript \"do shell script ...curl...\"' instead",
        "severity": "error",
    },
    "msxml2.xmlhttp.6.0": {
        "description": "MSXML2.XMLHTTP not available on Mac. Use 'MacScript \"do shell script ...curl...\"' instead",
        "severity": "error",
    },
    "msxml2.domdocument": {
        "description": "MSXML2.DOMDocument not available on Mac. Consider using string parsing or shell xmllint",
        "severity": "error",
    },
    "msxml2.domdocument.6.0": {
        "description": "MSXML2.DOMDocument not available on Mac. Consider using string parsing or shell xmllint",
        "severity": "error",
    },
    "wscript.shell": {
        "description": "WScript.Shell not available on Mac. Use MacScript/AppleScriptTask for shell commands",
        "severity": "error",
    },
    "shell.application": {
        "description": "Shell.Application not available on Mac. Use MacScript for Finder operations",
        "severity": "error",
    },
    "vbscript.regexp": {
        "description": "VBScript.RegExp not available on Mac. Use Like operator or manual string parsing",
        "severity": "error",
    },
    "adodb.connection": {
        "description": "ADODB is not available on Mac Office. Consider alternative data access methods",
        "severity": "error",
    },
    "adodb.recordset": {
        "description": "ADODB is not available on Mac Office. Consider alternative data access methods",
        "severity": "error",
    },
    "outlook.application": {
        "description": "Outlook COM automation works differently on Mac. Review Outlook Mac VBA docs",
        "severity": "warning",
    },
    "internetexplorer.application": {
        "description": "Internet Explorer does not exist on Mac. Use MacScript with Safari or curl",
        "severity": "error",
    },
}


def _transform_createobject(source: str, report: TransformReport) -> str:
    lines = source.split("\n")
    new_lines = []
    for i, line in enumerate(lines):
        stripped = line.lstrip()
        if stripped.startswith("'") or stripped.upper().startswith("REM "):
            new_lines.append(line)
            continue

        for match in _CREATEOBJECT_PATTERN.finditer(line):
            prog_id = match.group(1)
            key = prog_id.lower()
            if key in _COM_REPLACEMENTS:
                info = _COM_REPLACEMENTS[key]
                report.results.append(TransformResult(
                    category="COM Object",
                    line_number=i + 1,
                    original=match.group(0),
                    replacement="(manual fix required)",
                    description=info["description"],
                    severity=info["severity"],
                ))
        new_lines.append(line)
    return "\n".join(new_lines)


# --- Shell command transforms ---

_SHELL_PATTERN = re.compile(
    r'\bShell\s*\(\s*"([^"]*)"',
    re.IGNORECASE,
)

_SHELL_PATTERN2 = re.compile(
    r'\bShell\s+"([^"]*)"',
    re.IGNORECASE,
)


def _transform_shell_calls(source: str, report: TransformReport) -> str:
    lines = source.split("\n")
    new_lines = []
    for i, line in enumerate(lines):
        stripped = line.lstrip()
        if stripped.startswith("'") or stripped.upper().startswith("REM "):
            new_lines.append(line)
            continue

        for pattern in [_SHELL_PATTERN, _SHELL_PATTERN2]:
            for match in pattern.finditer(line):
                cmd = match.group(1)
                report.results.append(TransformResult(
                    category="Shell Command",
                    line_number=i + 1,
                    original=match.group(0),
                    replacement='MacScript "do shell script ""...""" or AppleScriptTask',
                    description=(
                        f"Windows Shell command detected: \"{cmd}\". "
                        "On Mac, use MacScript/AppleScriptTask with 'do shell script' instead."
                    ),
                    severity="error",
                ))
        new_lines.append(line)
    return "\n".join(new_lines)


# --- SendKeys transforms ---

_SENDKEYS_PATTERN = re.compile(
    r'\bSendKeys\b',
    re.IGNORECASE,
)


def _transform_sendkeys(source: str, report: TransformReport) -> str:
    lines = source.split("\n")
    for i, line in enumerate(lines):
        stripped = line.lstrip()
        if stripped.startswith("'") or stripped.upper().startswith("REM "):
            continue
        if _SENDKEYS_PATTERN.search(line):
            report.results.append(TransformResult(
                category="SendKeys",
                line_number=i + 1,
                original=line.strip(),
                replacement="(manual fix required)",
                description=(
                    "SendKeys is unreliable on Mac and may not work at all. "
                    "Consider using Application-level methods or AppleScript instead."
                ),
                severity="warning",
            ))
    return source


# --- Clipboard transforms ---

_CLIPBOARD_PATTERN = re.compile(
    r'\bMSForms\.DataObject\b|\bNew\s+DataObject\b',
    re.IGNORECASE,
)


def _transform_clipboard(source: str, report: TransformReport) -> str:
    lines = source.split("\n")
    for i, line in enumerate(lines):
        stripped = line.lstrip()
        if stripped.startswith("'") or stripped.upper().startswith("REM "):
            continue
        if _CLIPBOARD_PATTERN.search(line):
            report.results.append(TransformResult(
                category="Clipboard",
                line_number=i + 1,
                original=line.strip(),
                replacement="(manual fix required)",
                description=(
                    "MSForms.DataObject clipboard access doesn't work on Mac. "
                    "Use MacScript with 'the clipboard' AppleScript command instead."
                ),
                severity="error",
            ))
    return source


# --- Environ transforms ---

_ENVIRON_PATTERN = re.compile(
    r'Environ\$?\(\s*"([^"]+)"\s*\)',
    re.IGNORECASE,
)

_WIN_ENVIRON_MAP = {
    "userprofile": "HOME",
    "appdata": "HOME",
    "localappdata": "HOME",
    "temp": "TMPDIR",
    "tmp": "TMPDIR",
    "username": "USER",
    "computername": "HOSTNAME",
    "homedrive": None,
    "homepath": "HOME",
    "systemroot": None,
    "windir": None,
    "programfiles": None,
    "programfiles(x86)": None,
    "commonprogramfiles": None,
}


def _transform_environ(source: str, report: TransformReport) -> str:
    lines = source.split("\n")
    new_lines = []
    for i, line in enumerate(lines):
        stripped = line.lstrip()
        if stripped.startswith("'") or stripped.upper().startswith("REM "):
            new_lines.append(line)
            continue

        original_line = line
        for match in _ENVIRON_PATTERN.finditer(line):
            var_name = match.group(1)
            key = var_name.lower()
            if key in _WIN_ENVIRON_MAP:
                mac_var = _WIN_ENVIRON_MAP[key]
                if mac_var:
                    old_val = match.group(0)
                    func = "Environ$" if "$" in match.group(0) else "Environ"
                    new_val = f'{func}("{mac_var}")'
                    line = line.replace(old_val, new_val, 1)
                    report.results.append(TransformResult(
                        category="Environment Variable",
                        line_number=i + 1,
                        original=old_val,
                        replacement=new_val,
                        description=f"Mapped Windows env var %{var_name}% to Mac ${mac_var}",
                        severity="fixed",
                    ))
                else:
                    report.results.append(TransformResult(
                        category="Environment Variable",
                        line_number=i + 1,
                        original=match.group(0),
                        replacement="(no Mac equivalent)",
                        description=f"Windows env var %{var_name}% has no Mac equivalent",
                        severity="error",
                    ))
        new_lines.append(line)
    return "\n".join(new_lines)


# --- Registry transforms ---

_REGISTRY_PATTERNS = [
    re.compile(r'\b(?:GetSetting|SaveSetting|DeleteSetting|GetAllSettings)\b', re.IGNORECASE),
    re.compile(r'\.RegRead\b|\.RegWrite\b|\.RegDelete\b', re.IGNORECASE),
    re.compile(r'"HKEY_[^"]*"', re.IGNORECASE),
]


def _transform_registry(source: str, report: TransformReport) -> str:
    lines = source.split("\n")
    for i, line in enumerate(lines):
        stripped = line.lstrip()
        if stripped.startswith("'") or stripped.upper().startswith("REM "):
            continue
        for pattern in _REGISTRY_PATTERNS:
            if pattern.search(line):
                report.results.append(TransformResult(
                    category="Registry Access",
                    line_number=i + 1,
                    original=line.strip(),
                    replacement="(manual fix required)",
                    description=(
                        "Windows Registry access is not available on Mac. "
                        "Consider using plist files, Environ, or application settings instead."
                    ),
                    severity="error",
                ))
                break
    return source


# --- ActiveX reference transforms ---

_ACTIVEX_PATTERNS = [
    re.compile(r'\bMSComctlLib\b', re.IGNORECASE),
    re.compile(r'\bMSComDlg\b', re.IGNORECASE),
    re.compile(r'\bMSFlexGridLib\b', re.IGNORECASE),
    re.compile(r'\bProgressBar\b', re.IGNORECASE),
    re.compile(r'\bTreeView\b', re.IGNORECASE),
    re.compile(r'\bListView\b', re.IGNORECASE),
    re.compile(r'\bStatusBar\b', re.IGNORECASE),
    re.compile(r'\bToolbar\b', re.IGNORECASE),
    re.compile(r'\bImageList\b', re.IGNORECASE),
    re.compile(r'\bTabStrip\b', re.IGNORECASE),
]


def _transform_activex_references(source: str, report: TransformReport) -> str:
    lines = source.split("\n")
    reported = set()
    for i, line in enumerate(lines):
        stripped = line.lstrip()
        if stripped.startswith("'") or stripped.upper().startswith("REM "):
            continue
        for pattern in _ACTIVEX_PATTERNS:
            match = pattern.search(line)
            if match and match.group(0).lower() not in reported:
                reported.add(match.group(0).lower())
                report.results.append(TransformResult(
                    category="ActiveX Control",
                    line_number=i + 1,
                    original=match.group(0),
                    replacement="(not available on Mac)",
                    description=(
                        f"ActiveX control/library '{match.group(0)}' is not supported on Mac Office. "
                        "These controls must be replaced with UserForm controls or removed."
                    ),
                    severity="error",
                ))
    return source
