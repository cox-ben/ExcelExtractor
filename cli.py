"""
Batch Quote Extractor CLI.

Processes directories of legacy .xlsx and .xlsm quote files in batch mode
or extracts single quoting files, outputs structured JSON and CSV datasets,
and provides exit codes and summary logging.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import sys
from typing import List, Optional, Sequence

# Ensure src is in sys.path if invoked directly
_SRC_DIR = Path(__file__).resolve().parent / "src"
if str(_SRC_DIR) not in sys.path:
    sys.path.insert(0, str(_SRC_DIR))

from excel_extractor.engine import QuoteExtractorEngine
from excel_extractor.exporters import (
    export_to_csv,
    export_to_excel,
    export_to_json,
)
from excel_extractor.models import QuoteDocument


class CLIArgumentParser(argparse.ArgumentParser):
    """Custom argument parser that exits with code 1 on invocation errors."""

    def error(self, message: str) -> None:
        sys.stderr.write(f"Error: {message}\n")
        sys.exit(1)


def build_parser() -> CLIArgumentParser:
    """Constructs and returns the CLI argument parser."""
    parser = CLIArgumentParser(
        prog="cli.py",
        description="Batch Quote Extractor CLI - parse legacy Excel quotes to structured JSON and CSV.",
        add_help=True,
    )

    # Input: accepts -i, --input, -d, --dir, or positional input
    parser.add_argument(
        "-i",
        "--input",
        dest="input",
        type=str,
        default=None,
        help="Path to a single .xlsx/.xlsm file OR a directory of files",
    )
    parser.add_argument(
        "-d",
        "--dir",
        dest="dir",
        type=str,
        default=None,
        help="Directory containing quote files to scan (alias for --input)",
    )
    parser.add_argument(
        "positional_input",
        nargs="?",
        type=str,
        default=None,
        help="Positional path to input file or directory if -i is omitted",
    )

    # Output directory
    parser.add_argument(
        "-o",
        "--output-dir",
        dest="output_dir",
        type=str,
        default="./output",
        help="Directory to write output files (default: ./output)",
    )

    # Custom configuration file
    parser.add_argument(
        "-c",
        "--config",
        dest="config",
        type=str,
        default=None,
        help="Optional path to custom config.yaml",
    )

    # Export format
    parser.add_argument(
        "-f",
        "--format",
        dest="format",
        type=str,
        default="both",
        choices=["json", "csv", "both", "all", "excel", "xlsx"],
        help="Export format: json, csv, both, or all (default: both)",
    )

    # Strict mode
    parser.add_argument(
        "--strict",
        dest="strict",
        action="store_true",
        default=False,
        help="If set, fail with exit code 2 if any warnings occur",
    )

    # Recursive directory scanning
    parser.add_argument(
        "-r",
        "--recursive",
        dest="recursive",
        action="store_true",
        default=True,
        help="Recursively scan subdirectories (default: True)",
    )
    parser.add_argument(
        "--no-recursive",
        dest="recursive",
        action="store_false",
        help="Do not scan subdirectories recursively",
    )

    # Verbose logging
    parser.add_argument(
        "-v",
        "--verbose",
        dest="verbose",
        action="store_true",
        default=False,
        help="Enable verbose debug logging",
    )

    return parser


def main(args: Optional[Sequence[str]] = None) -> int:
    """
    Main entrypoint for CLI execution.

    Parameters:
        args: Sequence of CLI arguments. If None, sys.argv[1:] is used.

    Returns:
        Exit code:
            0: Success, all files parsed cleanly with no fatal errors.
            1: Invocation error (missing arguments, non-existent path) or fatal failure.
            2: Partial success with warnings when in strict mode or files skipped due to errors.
    """
    parser = build_parser()

    if args is None:
        args = sys.argv[1:]

    parsed_args = parser.parse_args(args)

    # 1. Resolve input path
    raw_input = parsed_args.input or parsed_args.dir or parsed_args.positional_input
    if not raw_input:
        sys.stderr.write(
            "Error: Missing required input path. Specify via --input, --dir, or positional argument.\n"
        )
        return 1

    input_path = Path(raw_input)
    if not input_path.exists():
        sys.stderr.write(f"Error: Input path does not exist: {input_path}\n")
        return 1

    # 2. Check custom config if specified
    if parsed_args.config:
        cfg_path = Path(parsed_args.config)
        if not cfg_path.exists():
            sys.stderr.write(f"Error: Custom config file does not exist: {cfg_path}\n")
            return 1

    # 3. Instantiate extraction engine
    try:
        engine = QuoteExtractorEngine(config=parsed_args.config)
    except Exception as exc:
        sys.stderr.write(f"Error initializing extraction engine: {exc}\n")
        return 1

    # 4. Ingest and extract quotes
    documents: List[QuoteDocument] = []
    files_scanned = 0

    if input_path.is_file():
        if input_path.suffix.lower() not in (".xlsx", ".xlsm"):
            sys.stderr.write(
                f"Error: Unsupported file format '{input_path.suffix}'. Only .xlsx and .xlsm are supported.\n"
            )
            return 1

        files_scanned = 1
        try:
            doc = engine.extract_file(input_path)
            documents.append(doc)
        except Exception as exc:
            sys.stderr.write(f"Fatal error extracting file {input_path.name}: {exc}\n")
            return 1
    elif input_path.is_dir():
        try:
            documents = engine.extract_directory(input_path, recursive=parsed_args.recursive)
            files_scanned = len(documents)
        except Exception as exc:
            sys.stderr.write(f"Fatal error scanning directory {input_path}: {exc}\n")
            return 1
    else:
        sys.stderr.write(f"Error: Input path is neither a file nor a directory: {input_path}\n")
        return 1

    # 5. Export structured outputs
    output_dir = Path(parsed_args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    export_fmt = parsed_args.format.lower()

    if export_fmt in ("json", "both", "all"):
        json_file = output_dir / "extracted_quotes.json"
        export_to_json(documents, file_path=json_file, indent=2, as_list=True)
        # Also export individual quote documents if quote_id exists
        for doc in documents:
            if doc.quote_id:
                indiv_file = output_dir / f"{doc.quote_id}.json"
                export_to_json(doc, file_path=indiv_file, indent=2)

    if export_fmt in ("csv", "both", "all"):
        csv_file = output_dir / "extracted_quotes.csv"
        export_to_csv(documents, file_path=csv_file)

    if export_fmt in ("excel", "xlsx", "all"):
        xlsx_file = output_dir / "extracted_quotes.xlsx"
        export_to_excel(documents, file_path=xlsx_file)

    # 6. Collate metrics & warnings
    quotes_extracted = sum(1 for d in documents if d.total_items > 0)
    total_items = sum(d.total_items for d in documents)

    all_warnings: List[str] = []
    failed_files = 0

    for doc in documents:
        has_failure = False
        for w in doc.warnings:
            all_warnings.append(f"[{doc.source_file}] {w}")
            if "failed" in w.lower() or "error" in w.lower():
                has_failure = True
        if has_failure:
            failed_files += 1
        for item in doc.items:
            for w in item.warnings:
                all_warnings.append(f"[{doc.source_file} / {item.quote_item_number}] {w}")

    # 7. Print console summary
    print("=" * 50)
    print("Batch Extraction Complete")
    print(f"Files Scanned:       {files_scanned}")
    print(f"Quotes Extracted:    {quotes_extracted} (Total Items: {total_items})")
    print(f"Warnings Detected:   {len(all_warnings)}")
    print(f"Errors / Failed:     {failed_files}")
    print(f"Output written to:   {output_dir.resolve()}")
    print("=" * 50)

    if parsed_args.verbose and all_warnings:
        print("\nWarnings Details:")
        for w in all_warnings:
            print(f"  - {w}")

    # 8. Determine exit code
    if failed_files > 0:
        return 2

    if parsed_args.strict and len(all_warnings) > 0:
        return 2

    return 0


if __name__ == "__main__":
    sys.exit(main())
