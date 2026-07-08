#!/usr/bin/env python3
"""
run.py -- the command you actually type.

Usage:
    python3 run.py <wireframe_filename> [extra wireframe_analyzer.py flags]

Looks for <wireframe_filename> inside source/, runs the full analysis via
wireframe_analyzer.py, and writes the design spec document into output/.

If a document with that name already exists in output/, this does NOT
overwrite it -- it writes a new version instead (_v2, _v3, ...). Delete
old versions from output/ manually whenever you want to clean up.

Examples:
    python3 run.py homepage.html
    python3 run.py dashboard_v2.png
    python3 run.py funnel_report.html --canvas 1920x1080
    python3 run.py funnel_report.html --xlsx
"""
import sys
import os
import subprocess

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_DIR = os.path.join(ROOT_DIR, "source")
OUTPUT_DIR = os.path.join(ROOT_DIR, "output")


def next_available_path(base_path):
    """If base_path doesn't exist, use it as-is. Otherwise, find the next
    free _v2, _v3, ... version instead of overwriting anything."""
    if not os.path.exists(base_path):
        return base_path
    root, ext = os.path.splitext(base_path)
    n = 2
    while True:
        candidate = f"{root}_v{n}{ext}"
        if not os.path.exists(candidate):
            return candidate
        n += 1


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 run.py <wireframe_filename> [extra flags]")
        print(f"\nThe file must already be inside: {SOURCE_DIR}")
        print("\nExamples:")
        print("  python3 run.py homepage.html")
        print("  python3 run.py dashboard.png")
        print("  python3 run.py report.html --canvas 1920x1080")
        sys.exit(1)

    filename = sys.argv[1]
    extra_args = sys.argv[2:]  # passed straight through to wireframe_analyzer.py
    input_path = os.path.join(SOURCE_DIR, filename)

    if not os.path.isfile(input_path):
        print(f"File not found: {input_path}")
        print(f"Make sure '{filename}' is sitting directly inside the source/ folder.")
        sys.exit(1)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    base_name = os.path.splitext(filename)[0] + "_design_spec.docx"
    out_path = next_available_path(os.path.join(OUTPUT_DIR, base_name))

    script_path = os.path.join(ROOT_DIR, "wireframe_analyzer.py")
    cmd = [sys.executable, script_path, input_path, "--out", out_path] + extra_args

    print(f"Analyzing:  {input_path}")
    print(f"Output ->   {out_path}\n")

    result = subprocess.run(cmd)
    sys.exit(result.returncode)


if __name__ == "__main__":
    main()
