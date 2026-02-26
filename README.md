# markdown-onenote-export

[![GitHub Release](https://img.shields.io/github/v/release/strick-j/markdown-onenote-export)](https://github.com/strick-j/markdown-onenote-export/releases)
[![Tests](https://github.com/strick-j/markdown-onenote-export/actions/workflows/test.yml/badge.svg)](https://github.com/strick-j/markdown-onenote-export/actions/workflows/test.yml)
[![Security](https://github.com/strick-j/markdown-onenote-export/actions/workflows/security.yml/badge.svg)](https://github.com/strick-j/markdown-onenote-export/actions/workflows/security.yml)
[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](https://www.python.org/downloads/)

A command-line tool that exports Microsoft OneNote `.one` files to Markdown.

OneNote stores notebooks in a proprietary binary format (MS-ONESTORE). This tool parses those files and converts pages — including text formatting, images, tables, and attachments — into clean, version-control-friendly Markdown.

## Features

- Converts `.one` files to Markdown with formatting preserved (bold, italic, strikethrough, links)
- Extracts images and embedded file attachments alongside Markdown output
- Renders tables in standard Markdown table syntax
- Automatically deduplicates section versions (e.g. `Notes.one` vs `Notes (On 2-25-26).one`)
- Preserves notebook/section/page hierarchy in the output directory
- Optional flat output mode (no notebook subdirectories)

## Requirements

- **Python 3.11+**
- **macOS, Linux, or Windows**

## Installation

```bash
git clone https://github.com/strick-j/markdown-onenote-export.git
cd markdown-onenote-export

# Create and activate a virtual environment (recommended)
python3 -m venv .venv
source .venv/bin/activate          # macOS/Linux
# .venv\Scripts\activate.bat       # Windows (CMD)
# .venv\Scripts\Activate.ps1       # Windows (PowerShell)

# Install the package
pip install -e .
```

## Usage

```bash
onenote-export -i <input_dir> -o <output_dir> [options]
```

### Arguments

| Flag | Long Form | Required | Description |
|------|-----------|----------|-------------|
| `-i` | `--input` | Yes | Directory containing `.one` files (searched recursively) |
| `-o` | `--output` | Yes | Directory where Markdown files will be written |
| `-v` | `--verbose` | No | Show progress details (INFO level logging) |
| | `--debug` | No | Show detailed diagnostic output (DEBUG level logging) |
| | `--flat` | No | Write all sections to the output root without notebook subdirectories |

### Examples

**Basic export:**

```bash
onenote-export -i ~/OneDrive/OneNote_Notebooks -o ~/Documents/exported_notes
```

**Verbose output with flat structure:**

```bash
onenote-export -i ~/OneDrive/OneNote_Notebooks -o ~/Documents/exported_notes --verbose --flat
```

**Debug mode for troubleshooting:**

```bash
onenote-export -i ~/OneDrive/OneNote_Notebooks -o ~/Documents/exported_notes --debug
```

## Input Structure

The tool expects `.one` files organized into directories, where each directory represents a notebook:

```
input_dir/
├── Work Notebook/
│   ├── Meeting Notes.one
│   ├── Projects.one
│   └── Projects (On 2-25-26).one   # older version, auto-skipped
└── Personal/
    ├── Recipes.one
    └── Travel.one
```

Files with date suffixes like `(On 2-25-26)` are recognized as older versions of the same section. Only the most recent version is exported.

## Output Structure

**Default (nested):**

```
output_dir/
├── Work Notebook/
│   ├── Meeting Notes/
│   │   ├── Standup 2024-01-15.md
│   │   ├── images/
│   │   │   └── image_001.png
│   │   └── attachments/
│   │       └── report.pdf
│   └── Projects/
│       └── Project Alpha.md
└── Personal/
    ├── Recipes/
    │   └── Pasta Carbonara.md
    └── Travel/
        └── Japan Trip.md
```

**With `--flat`:**

```
output_dir/
├── Meeting Notes/
│   └── Standup 2024-01-15.md
├── Projects/
│   └── Project Alpha.md
├── Recipes/
│   └── Pasta Carbonara.md
└── Travel/
    └── Japan Trip.md
```

## Supported Content

| Content Type | How It's Exported |
|---|---|
| Rich text | Markdown with **bold**, *italic*, ~~strikethrough~~, superscript, subscript |
| Hyperlinks | `[text](url)` |
| Images | Saved to `images/` subdirectory, referenced in Markdown |
| Tables | Standard Markdown table syntax |
| Embedded files | Saved to `attachments/` subdirectory, linked in Markdown |

Image formats detected automatically: PNG, JPEG, GIF, BMP, WEBP.

> **Note:** Underline is rendered as italic since Markdown has no native underline syntax.

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Export completed successfully |
| 1 | Fatal error (input directory missing or no `.one` files found) |
| 2 | All sections failed to process |

Errors are printed to stderr. Use `-v` or `--debug` for additional detail.

## Testing

The project includes a test suite covering utilities, data models, Markdown rendering, content extraction, and CLI logic.

### Setup

```bash
# Install with dev dependencies (pytest + pytest-cov)
pip install -e ".[dev]"
```

### Running Tests

```bash
# Run all tests
pytest

# Run tests with verbose output
pytest -v

# Run tests with coverage report
pytest --cov=onenote_export

# Run a specific test file
pytest tests/test_utils.py

# Run a specific test class or method
pytest tests/test_markdown_converter.py::TestMarkdownConverterRenderPage::test_page_with_bold_text
```

### Test Structure

```
tests/
├── Example-NoteBook-1/           # Sample .one files for testing
│   ├── Example-Section-1 (On 2-25-26 - 2).one
│   ├── Example-Section-1 (On 2-25-26 - 3).one
│   ├── Example-Section-2 (On 2-25-26 - 3).one
│   └── Example-Section-2 (On 2-25-26).one
├── test_cli.py                   # CLI argument parsing and section deduplication
├── test_content_extractor.py     # Text decoding, image format detection, type parsing
├── test_markdown_converter.py    # Markdown rendering, file/image/attachment writing
├── test_models.py                # Data model construction and immutability
└── test_utils.py                 # File discovery and name extraction
```

### What's Tested

| Module | Tests | Coverage |
|--------|-------|----------|
| `utils.py` | File discovery, section name parsing, notebook name extraction | File glob, `.onetoc2` exclusion, date suffix stripping |
| `model/` | All dataclass defaults, immutability (`TextRun` is frozen), element construction | All content types: `RichText`, `ImageElement`, `TableElement`, `EmbeddedFile` |
| `converter/markdown.py` | Rendering for every content type, file writing, filename sanitization | Bold, italic, strikethrough, links, super/subscript, tables, images, attachments |
| `parser/content_extractor.py` | Text decoding (hex, UTF-16, ASCII), garbled text detection, image format magic bytes | All image formats (PNG, JPEG, GIF, BMP, WEBP), edge cases |
| `cli.py` | Section deduplication with date parsing, 2-digit year normalization | Dated vs undated files, multiple versions, empty input |

## How It Works

1. **Discover** — Recursively finds all `.one` files in the input directory (skips `.onetoc2` files)
2. **Deduplicate** — Groups files by section name and keeps only the latest version
3. **Parse** — Reads each `.one` file using the [pyOneNote](https://pypi.org/project/pyOneNote/) library to extract raw objects from the MS-ONESTORE binary format
4. **Extract** — Converts raw objects into a structured model (notebooks, sections, pages, content elements)
5. **Convert** — Renders each page as Markdown, writing images and attachments to disk

## Limitations

- Requires binary `.one` files — does not work with OneNote for Web exports or `.onetoc2` table-of-contents files
- Some complex OneNote formatting may not have a direct Markdown equivalent
- OneNote revision history may produce duplicate content in edge cases
- Tested on macOS, Linux, and Windows

## License

MIT
