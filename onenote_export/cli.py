"""CLI entry point for the OneNote to Markdown exporter."""

import argparse
import logging
import re
import sys
from pathlib import Path

from onenote_export.converter.markdown import MarkdownConverter
from onenote_export.model.notebook import Notebook
from onenote_export.parser.content_extractor import extract_section
from onenote_export.parser.one_store import OneStoreParser
from onenote_export.utils import (
    discover_one_files,
    notebook_name_from_dir,
    section_name_from_filename,
)


def main(argv: list[str] | None = None) -> int:
    """Main entry point for the onenote-export CLI."""
    parser = argparse.ArgumentParser(
        prog="onenote-export",
        description="Export OneNote (.one) files to Markdown",
    )
    parser.add_argument(
        "-i",
        "--input",
        required=True,
        help="Input directory containing .one files",
    )
    parser.add_argument(
        "-o",
        "--output",
        required=True,
        help="Output directory for Markdown files",
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Enable debug logging (very verbose)",
    )
    parser.add_argument(
        "--flat",
        action="store_true",
        help="Output flat directory structure (no notebook subdirectories)",
    )

    args = parser.parse_args(argv)

    # Configure logging
    if args.debug:
        log_level = logging.DEBUG
    elif args.verbose:
        log_level = logging.INFO
    else:
        log_level = logging.WARNING

    logging.basicConfig(
        level=log_level,
        format="%(levelname)s: %(message)s",
    )

    input_dir = Path(args.input).resolve()
    output_dir = Path(args.output).resolve()

    if not input_dir.is_dir():
        print(f"Error: Input directory does not exist: {input_dir}", file=sys.stderr)
        return 1

    # Discover .one files
    one_files = discover_one_files(input_dir)
    if not one_files:
        print(f"No .one files found in {input_dir}", file=sys.stderr)
        return 1

    print(f"Found {len(one_files)} .one file(s) in {input_dir}")

    # Group files by parent directory (notebook)
    notebooks: dict[Path, list[Path]] = {}
    for f in one_files:
        parent = f.parent
        notebooks.setdefault(parent, []).append(f)

    # Deduplicate: keep only the latest version of each section
    # Files like "ADI (On 2-25-26).one" and "ADI.one (On 10-3-22).one"
    # are the same section at different dates
    for notebook_dir in notebooks:
        notebooks[notebook_dir] = _deduplicate_sections(notebooks[notebook_dir])

    # Process each notebook
    converter = MarkdownConverter(output_dir)
    total_files = 0
    total_pages = 0
    errors: list[str] = []

    for notebook_dir, section_files in sorted(notebooks.items()):
        notebook_name = notebook_name_from_dir(notebook_dir)
        print(f"\nProcessing notebook: {notebook_name}")

        notebook = Notebook(
            name=notebook_name,
            dir_path=str(notebook_dir),
        )

        for section_file in sorted(section_files):
            section_name = section_name_from_filename(section_file.name)
            print(f"  Section: {section_name} ({section_file.name})")

            try:
                # Parse the .one file
                store_parser = OneStoreParser(section_file)
                parsed = store_parser.parse()

                # Extract structured content
                section = extract_section(parsed)
                section.name = section_name

                notebook.sections.append(section)
                page_count = len(section.pages)
                total_pages += page_count
                print(f"    -> {page_count} page(s) extracted")

            except Exception as e:
                error_msg = f"    ERROR: {section_file.name}: {e}"
                print(error_msg, file=sys.stderr)
                errors.append(error_msg)
                logging.debug("Full traceback:", exc_info=True)

        # Write the notebook
        if notebook.sections:
            try:
                if args.flat:
                    for section in notebook.sections:
                        files = converter.convert_section(section)
                        total_files += len(files)
                else:
                    files = converter.convert_notebook(notebook)
                    total_files += len(files)
            except Exception as e:
                error_msg = f"  ERROR writing notebook {notebook_name}: {e}"
                print(error_msg, file=sys.stderr)
                errors.append(error_msg)

    # Summary
    print(f"\n{'=' * 50}")
    print("Export complete:")
    print(f"  Pages extracted: {total_pages}")
    print(f"  Files written:   {total_files}")
    print(f"  Output:          {output_dir}")

    if errors:
        print(f"\n  Errors ({len(errors)}):")
        for err in errors:
            print(f"    {err}")
        return 2 if total_files == 0 else 0

    return 0


def _deduplicate_sections(files: list[Path]) -> list[Path]:
    """Keep only the latest version of each section.

    Files follow patterns like:
      'ADI (On 2-25-26).one'      -> section 'ADI', date 2026-02-25
      'ADI.one (On 10-3-22).one'  -> section 'ADI', date 2022-10-03

    Groups by section name and keeps the file with the latest date.
    """
    date_pattern = re.compile(r"\(On\s+(\d+)-(\d+)-(\d+)(?:\s*-\s*\d+)?\)")

    section_versions: dict[str, list[tuple[Path, tuple[int, int, int]]]] = {}

    for f in files:
        section_name = section_name_from_filename(f.name)
        match = date_pattern.search(f.name)
        if match:
            month, day, year = (
                int(match.group(1)),
                int(match.group(2)),
                int(match.group(3)),
            )
            # Normalize 2-digit year
            if year < 100:
                year += 2000 if year < 50 else 1900
            section_versions.setdefault(section_name, []).append(
                (f, (year, month, day))
            )
        else:
            section_versions.setdefault(section_name, []).append((f, (0, 0, 0)))

    result: list[Path] = []
    for section_name, versions in sorted(section_versions.items()):
        # Sort by date descending and take the latest
        versions.sort(key=lambda x: x[1], reverse=True)
        latest = versions[0][0]
        result.append(latest)
        if len(versions) > 1:
            skipped = [v[0].name for v in versions[1:]]
            logging.info(
                "Section '%s': using %s, skipping older: %s",
                section_name,
                latest.name,
                ", ".join(skipped),
            )

    return result


if __name__ == "__main__":
    sys.exit(main())
