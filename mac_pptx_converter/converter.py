"""
Main converter: takes a .pptm/.pptx file, extracts VBA, transforms it
for Mac compatibility, and writes the converted file.
"""

import zipfile
import io
import os
import shutil
from dataclasses import dataclass, field

from .vba_parser import VBAProject
from .transforms import apply_all_transforms, TransformReport


@dataclass
class ConversionResult:
    input_path: str
    output_path: str
    module_reports: list[TransformReport] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    vba_found: bool = False

    @property
    def total_issues(self) -> int:
        return sum(len(r.results) for r in self.module_reports)

    @property
    def fixed_count(self) -> int:
        return sum(
            1 for r in self.module_reports
            for t in r.results if t.severity == "fixed"
        )

    @property
    def warning_count(self) -> int:
        return sum(
            1 for r in self.module_reports
            for t in r.results if t.severity == "warning"
        )

    @property
    def error_count(self) -> int:
        return sum(
            1 for r in self.module_reports
            for t in r.results if t.severity == "error"
        )


def convert_file(input_path: str, output_path: str | None = None) -> ConversionResult:
    if output_path is None:
        base, ext = os.path.splitext(input_path)
        output_path = f"{base}_mac{ext}"

    result = ConversionResult(input_path=input_path, output_path=output_path)

    if not os.path.exists(input_path):
        result.errors.append(f"Input file not found: {input_path}")
        return result

    ext = os.path.splitext(input_path)[1].lower()
    if ext not in (".pptm", ".pptx", ".xlsm", ".xlsx", ".docm", ".docx"):
        result.errors.append(
            f"Unsupported file type: {ext}. "
            "Supported: .pptm, .pptx, .xlsm, .xlsx, .docm, .docx"
        )
        return result

    try:
        _convert(input_path, output_path, result)
    except Exception as e:
        result.errors.append(f"Conversion failed: {e}")

    return result


def _convert(input_path: str, output_path: str, result: ConversionResult):
    shutil.copy2(input_path, output_path)

    with zipfile.ZipFile(output_path, "r") as zf:
        vba_path = _find_vba_project(zf)
        if vba_path is None:
            result.vba_found = False
            return
        result.vba_found = True
        vba_data = zf.read(vba_path)

    project = VBAProject(vba_data)

    if not project.modules:
        result.errors.append("No VBA modules found in the project")
        return

    has_changes = False
    for module in project.modules:
        if not module.source.strip():
            continue
        new_source, report = apply_all_transforms(module.source, module.name)
        result.module_reports.append(report)
        if new_source != module.source:
            module.source = new_source
            has_changes = True

    if has_changes:
        new_vba_data = project.save()
        _replace_in_zip(output_path, vba_path, new_vba_data)


def _find_vba_project(zf: zipfile.ZipFile) -> str | None:
    for name in zf.namelist():
        if name.lower().endswith("vbaproject.bin"):
            return name
    return None


def _replace_in_zip(zip_path: str, target_name: str, new_data: bytes):
    temp_path = zip_path + ".tmp"
    with zipfile.ZipFile(zip_path, "r") as zin:
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == target_name:
                    zout.writestr(item, new_data)
                else:
                    zout.writestr(item, zin.read(item.filename))
    os.replace(temp_path, zip_path)


def format_report(result: ConversionResult) -> str:
    lines = []
    lines.append("=" * 70)
    lines.append("  Mac PPTX Converter - Conversion Report")
    lines.append("=" * 70)
    lines.append(f"  Input:  {result.input_path}")
    lines.append(f"  Output: {result.output_path}")
    lines.append("")

    if result.errors:
        lines.append("  ERRORS:")
        for err in result.errors:
            lines.append(f"    ✗ {err}")
        lines.append("")

    if not result.vba_found:
        lines.append("  No VBA project found in this file.")
        lines.append("  The file was copied as-is (no macros to convert).")
        lines.append("=" * 70)
        return "\n".join(lines)

    if result.total_issues == 0:
        lines.append("  ✓ No compatibility issues detected!")
        lines.append("  The VBA code appears to be Mac-compatible.")
        lines.append("=" * 70)
        return "\n".join(lines)

    lines.append(f"  Summary: {result.fixed_count} auto-fixed, "
                 f"{result.warning_count} warnings, "
                 f"{result.error_count} need manual fixes")
    lines.append("-" * 70)

    for report in result.module_reports:
        if not report.has_changes:
            continue
        lines.append(f"\n  Module: {report.module_name}")
        lines.append("  " + "-" * 40)

        for tr in report.results:
            icon = {"fixed": "✓", "warning": "⚠", "error": "✗"}[tr.severity]
            lines.append(f"    {icon} [{tr.category}] Line {tr.line_number}")
            lines.append(f"      {tr.description}")
            if tr.severity == "fixed":
                lines.append(f"      Before: {tr.original}")
                lines.append(f"      After:  {tr.replacement}")
            elif tr.severity in ("error", "warning"):
                lines.append(f"      Code: {tr.original}")
            lines.append("")

    lines.append("=" * 70)

    if result.error_count > 0:
        lines.append("")
        lines.append("  ✗ Items marked with ✗ require manual attention.")
        lines.append("    Open the converted file in Mac PowerPoint's VBA editor")
        lines.append("    and apply the suggested changes.")

    lines.append("")
    return "\n".join(lines)
