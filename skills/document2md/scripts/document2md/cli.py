from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .converter import ConversionOptions, SUPPORTED_SUFFIXES, convert_files


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="document2md",
        description="Convert PDF, DOCX, PPTX, and XLSX files to readable Markdown.",
    )
    parser.add_argument("inputs", nargs="+", help="PDF or Office files to convert.")
    parser.add_argument("-o", "--output", help="Exact .md path for one input, or an output directory.")
    parser.add_argument("--assets-dir", help="Asset directory for one input.")
    parser.add_argument("--values-only", action="store_true", help="XLSX: export cached values instead of formulas.")
    parser.add_argument("--include-hidden", action="store_true", help="XLSX: include hidden worksheets.")
    parser.add_argument("--include-notes", action="store_true", help="PPTX: include speaker notes.")
    parser.add_argument(
        "--images",
        choices=["placeholder", "extract", "embed"],
        default="placeholder",
        help="Image handling: placeholders, extracted files, or PDF-only base64 embedding.",
    )
    parser.add_argument("--force", action="store_true", help="Overwrite existing Markdown outputs.")
    return parser


def _input_paths(raw_inputs: list[str]) -> list[Path]:
    paths = []
    for raw in raw_inputs:
        path = Path(raw).expanduser()
        if path.is_dir():
            paths.extend(
                child for child in sorted(path.iterdir())
                if child.is_file() and child.suffix.lower() in SUPPORTED_SUFFIXES
            )
        else:
            paths.append(path)
    if not paths:
        raise FileNotFoundError("No PDF, DOCX, PPTX, or XLSX inputs found.")
    return paths


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        inputs = _input_paths(args.inputs)
        output = Path(args.output).expanduser() if args.output else None
        if len(inputs) > 1 and output and output.suffix.lower() == ".md":
            raise ValueError("For multiple inputs, --output must be a directory.")
        if len(inputs) > 1 and args.assets_dir:
            raise ValueError("--assets-dir can only be used with one input.")

        targets = []
        for source in inputs:
            if output and len(inputs) == 1 and output.suffix.lower() == ".md":
                target = output
            elif output:
                output.mkdir(parents=True, exist_ok=True)
                target = output / f"{source.stem}.md"
            else:
                target = source.with_suffix(".md")
            targets.append(target)

        options = ConversionOptions(
            values_only=args.values_only,
            include_hidden=args.include_hidden,
            include_notes=args.include_notes,
            images=args.images,
            force=args.force,
            assets_dir=Path(args.assets_dir).expanduser().resolve() if args.assets_dir else None,
        )
        for result in convert_files(inputs, targets, options):
            print(result.output_path)
            for asset in result.assets:
                print(asset)
            for warning in result.warnings:
                print(f"WARNING: {warning}", file=sys.stderr)
        return 0
    except (FileNotFoundError, RuntimeError, ValueError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
