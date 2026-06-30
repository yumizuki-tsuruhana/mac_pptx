"""
Command-line interface for the Mac PPTX Converter.

Usage:
    mac-pptx input.pptm [-o output.pptm]
    mac-pptx input.pptm --report-only
    mac-pptx *.pptm
"""

import argparse
import glob
import sys
import os

from .converter import convert_file, format_report


def main():
    parser = argparse.ArgumentParser(
        prog="mac-pptx",
        description="Convert Windows VBA macro-enabled Office files for Mac compatibility",
    )
    parser.add_argument(
        "input",
        nargs="+",
        help="Input file(s) (.pptm, .pptx, .xlsm, .xlsx, .docm, .docx). Glob patterns supported.",
    )
    parser.add_argument(
        "-o", "--output",
        help="Output file path (only valid with single input file)",
    )
    parser.add_argument(
        "--report-only",
        action="store_true",
        help="Only analyze and report issues without modifying the file",
    )
    parser.add_argument(
        "--suffix",
        default="_mac",
        help="Suffix to add to output filename (default: _mac)",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Only show errors and summary",
    )

    args = parser.parse_args()

    input_files = []
    for pattern in args.input:
        expanded = glob.glob(pattern)
        if expanded:
            input_files.extend(expanded)
        else:
            input_files.append(pattern)

    if args.output and len(input_files) > 1:
        print("Error: --output can only be used with a single input file", file=sys.stderr)
        sys.exit(1)

    total_fixed = 0
    total_warnings = 0
    total_errors = 0
    conversion_errors = 0

    for input_path in input_files:
        if len(input_files) > 1 and not args.quiet:
            print(f"\nProcessing: {input_path}")

        if args.output:
            output_path = args.output
        else:
            base, ext = os.path.splitext(input_path)
            output_path = f"{base}{args.suffix}{ext}"

        if args.report_only:
            output_path = input_path
            result = convert_file(input_path, output_path=None)
            result.output_path = "(report only - no file written)"
            if os.path.exists(f"{base}{args.suffix}{ext}"):
                os.remove(f"{base}{args.suffix}{ext}")
        else:
            result = convert_file(input_path, output_path)

        if not args.quiet:
            print(format_report(result))

        total_fixed += result.fixed_count
        total_warnings += result.warning_count
        total_errors += result.error_count
        if result.errors:
            conversion_errors += 1

    if len(input_files) > 1:
        print(f"\n{'=' * 70}")
        print(f"  Total: {len(input_files)} files processed")
        print(f"  {total_fixed} auto-fixed, {total_warnings} warnings, "
              f"{total_errors} need manual fixes")
        if conversion_errors:
            print(f"  {conversion_errors} file(s) had errors")
        print(f"{'=' * 70}")

    sys.exit(1 if conversion_errors else 0)


if __name__ == "__main__":
    main()
