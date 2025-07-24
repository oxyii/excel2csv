# Excel2CSV Converter

A Go-based tool for converting Excel files (.xlsx, .xls, .ods) to CSV format with intelligent table boundary detection.

## Features

- **Smart Table Detection**: Automatically identifies the start and end of tabular data
- **Header Preservation**: Detects and preserves column headers
- **Footer Exclusion**: Automatically excludes footer rows and summary data
- **Multiple Format Support**: Supports .xlsx, .xls, and .ods files
- **Configurable CSV Separator**: Choose between comma, semicolon, or tab separators
- **Line Break Cleaning**: Replaces line breaks within cell data with spaces
- **Force Row Options**: Override automatic detection with manual row specification
- **LibreOffice Integration**: Uses LibreOffice headless mode for reliable conversion

## Installation

### Prerequisites

- Go 1.19 or later
- LibreOffice (for file conversion)

### Install LibreOffice

**Ubuntu/Debian:**
```bash
sudo apt-get install libreoffice
```

**macOS:**
```bash
brew install --cask libreoffice
```

**Windows:**
Download and install from [LibreOffice website](https://www.libreoffice.org/download/download/)

### Build from Source

```bash
git clone https://github.com/oxyii/excel2csv.git
cd excel2csv
go build -o excel2csv ./cmd/excel2csv
```

## Usage

### Basic Usage

```bash
./excel2csv -input input.xlsx -output output.csv
```

### Advanced Options

```bash
./excel2csv \
  -input input.xlsx \
  -output output.csv \
  -separator ";" \
  -start-row 5 \
  -end-row 100 \
  -no-clean-breaks
```

### Command Line Options

| Option | Description | Default |
|--------|-------------|---------|
| `-input` | Input Excel file path (required) | - |
| `-output` | Output CSV file path (required) | - |
| `-separator` | CSV separator: comma, semicolon, tab | comma |
| `-start-row` | Force table start row (0-based, optional) | auto-detect |
| `-end-row` | Force table end row (0-based, optional) | auto-detect |
| `-no-clean-breaks` | Disable line break cleaning | false |

### Examples

**Convert with semicolon separator:**
```bash
./excel2csv -input data.xlsx -output data.csv -separator "semicolon"
```

**Force specific table boundaries:**
```bash
./excel2csv -input data.xlsx -output data.csv -start-row 3 -end-row 50
```

**Use tab separator and preserve line breaks:**
```bash
./excel2csv -input data.xlsx -output data.csv -separator "tab" -no-clean-breaks
```

## How It Works

### Automatic Table Detection

The converter uses intelligent algorithms to:

1. **Find Header Row**: Identifies the row with the most non-empty cells as the table header
2. **Detect Table Start**: Locates where actual tabular data begins (skipping logos, contact info, etc.)
3. **Identify Table End**: Stops at footer rows, summaries, or significant column count changes
4. **Preserve Structure**: Maintains column headers and data integrity

### Supported File Formats

- **Excel 2007+** (.xlsx)
- **Excel 97-2003** (.xls)  
- **OpenDocument Spreadsheet** (.ods)

## Example Conversions

The tool has been tested with various file types:

- **Price Lists**: Automatically skips supplier information headers
- **Invoices**: Excludes totals and footer information
- **Stock Lists**: Handles large datasets (tested with 900k+ rows)
- **Reports**: Processes files with mixed content layouts

## API Usage

```go
package main

import (
    "github.com/oxyii/excel2csv"
)

func main() {
    converter := excel2csv.NewExcelConverter()
    
    // Configure options
    converter.CSVSeparator = ';'
    converter.CleanLineBreaks = true
    
    // Optional: force specific rows
    startRow := 5
    converter.ForceDataStartRow = &startRow
    
    // Convert file
    err := converter.ConvertFile("input.xlsx", "output.csv")
    if err != nil {
        panic(err)
    }
}
```

## Performance

- **Small files** (< 1MB): Near-instant conversion
- **Medium files** (1-10MB): Typically under 5 seconds
- **Large files** (> 100MB): Scales linearly, tested with 900k+ rows

## Troubleshooting

### Common Issues

**LibreOffice not found:**
```
Error: LibreOffice is not available. Please install LibreOffice
```
*Solution*: Install LibreOffice using your system's package manager

**Permission denied:**
```
Error: permission denied
```
*Solution*: Ensure read access to input file and write access to output directory

**Memory issues with large files:**
*Solution*: Use `-start-row` and `-end-row` to process file in chunks

## License

MIT License - see LICENSE file for details

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## Changelog

### v1.0.0
- Initial release
- Smart table boundary detection
- Multiple separator support
- LibreOffice integration
- Configurable options
