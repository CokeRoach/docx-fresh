"""Command line: ``docx-fresh FILE.docx [-o OUT.docx] [--author NAME]``."""

from __future__ import annotations

import argparse
import sys

import docx

from docx_fresh import __version__, fresh


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="docx-fresh",
        description="Replace the python-docx default template's metadata in a .docx file.",
    )
    parser.add_argument("file", help=".docx file to fix")
    parser.add_argument("-o", "--output", help="write here instead of overwriting FILE")
    parser.add_argument("--author", default="", help="author to set (default: empty)")
    parser.add_argument("--version", action="version", version=f"docx-fresh {__version__}")
    args = parser.parse_args(argv)

    document = docx.Document(args.file)
    fresh(document, author=args.author)
    out = args.output or args.file
    document.save(out)
    print(f"wrote {out}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
