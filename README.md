# Mathematical Equation Extractor

This tool extracts mathematical equations from Microsoft Word documents (.docx and .doc files).

## Features

- Extracts equations from .docx files using multiple methods:
  - Direct XML parsing of Office Math Markup Language (OMML) elements
  - Identification of equation objects in the document structure
  - Pattern matching for LaTeX-style equation formats
- Options for raw XML output or cleaned text extraction
- Outputs equations to console or a specified file
- Detailed logging for troubleshooting

## Installation

1. Clone this repository or download the script files
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Basic usage:

```bash
python extract_equations.py path/to/document.docx
```

With options:

```bash
# Output to file with verbose logging
python extract_equations.py path/to/document.docx -o output.txt -v

# Extract raw XML without cleaning
python extract_equations.py path/to/document.docx -r
```

### Command-line Arguments

- `file_path`: Path to the Word document (required)
- `-o, --output`: Output file path (optional, default: print to console)
- `-v, --verbose`: Enable verbose output (optional)
- `-r, --raw`: Return raw XML without cleaning (optional)

## Debugging

The repository includes VSCode launch configurations for debugging:

1. **Extract Equations**: Runs the script with standard output cleaning
2. **Extract Equations (Raw XML)**: Runs the script preserving the raw XML structure

## How It Works

The script uses multiple approaches to extract equations:

1. **Direct XML Parsing**: Opens the .docx file (which is a ZIP archive) and parses the XML content to find OMML equation elements
2. **Namespace Handling**: Searches through different XML namespaces to locate equation elements
3. **Text Processing**: Optionally cleans up the extracted XML to provide more readable equation text

## Limitations

- For .doc files (older Word format), direct processing is not implemented. Convert to .docx first using LibreOffice/OpenOffice.
- Complex equations or those using custom fonts may not be extracted correctly.
- The script primarily identifies equations based on common patterns and markup; some custom or unusual equation formats might be missed.

## Requirements

- Python 3.6+
- Dependencies listed in requirements.txt:
  - python-docx: For reading .docx files
  - mammoth: For HTML conversion
  - beautifulsoup4: For HTML parsing
  - lxml: XML processing backend
  - docx2python: Alternative .docx parsing
