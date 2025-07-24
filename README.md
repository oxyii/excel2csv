# Excel2CSV Converter

A Go-based tool for converting Excel files (.xlsx, .xls, .ods) to CSV format with intelligent table boundary detection, multi-sheet support, and HTTP API.

## Features

- **Smart Table Detection**: Automatically identifies the start and end of tabular data
- **Multi-Sheet Support**: Convert specific sheets by name/index or all sheets at once
- **HTTP API Server**: Web service for integration with other applications
- **Sheet Management**: List all available sheets in Excel files
- **Header Preservation**: Detects and preserves column headers
- **Footer Exclusion**: Automatically excludes footer rows and summary data
- **Multiple Format Support**: Supports .xlsx, .xls, and .ods files
- **Configurable CSV Separator**: Choose between comma, semicolon, or tab separators
- **Line Break Cleaning**: Replaces line breaks within cell data with spaces
- **Force Row Options**: Override automatic detection with manual row specification
- **LibreOffice Integration**: Uses LibreOffice headless mode for reliable conversion
- **Snap Compatibility**: Automatic handling of LibreOffice snap container limitations

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
go build -o excel2csv-server ./cmd/excel2csv-server
```

## Usage

### Basic Usage

```bash
# Convert first sheet (default)
./excel2csv -input input.xlsx -output output.csv

# List all sheets in the file
./excel2csv -input input.xlsx -list-sheets

# Convert specific sheet by name
./excel2csv -input input.xlsx -sheet-name "Sales Data"

# Convert specific sheet by index (0-based)
./excel2csv -input input.xlsx -sheet-index 1

# Convert all sheets to separate files
./excel2csv -input input.xlsx -all-sheets
```

### Advanced Options

```bash
./excel2csv \
  -input input.xlsx \
  -output output.csv \
  -sheet-name "Report" \
  -separator ";" \
  -start-row 5

# Convert all sheets with custom separator
./excel2csv -input input.xlsx -all-sheets -separator "tab"
```

### Command Line Options

| Option | Description | Default |
|--------|-------------|---------|
| `-input` | Input Excel file path (required) | - |
| `-output` | Output CSV file path (optional) | auto-generated |
| `-separator` | CSV separator: comma, semicolon, tab | comma |
| `-start-row` | Force table start row (0-based, optional) | auto-detect |
| **Sheet Selection** | | |
| `-list-sheets` | List all sheets in the Excel file and exit | false |
| `-sheet-name` | Convert specific sheet by name | first sheet |
| `-sheet-index` | Convert specific sheet by index (0-based) | first sheet |
| `-all-sheets` | Convert all sheets to separate CSV files | false |

### Examples

**List available sheets:**
```bash
./excel2csv -input report.xlsx -list-sheets
```

**Convert specific sheet by name:**
```bash
./excel2csv -input data.xlsx -sheet-name "Quarterly Report" -separator ";"
```

**Convert specific sheet by index:**
```bash
./excel2csv -input data.xlsx -sheet-index 2 -start-row 5
```

**Convert all sheets to separate files:**
```bash
./excel2csv -input workbook.xlsx -all-sheets
# Creates: workbook_sheet_1_Sheet1.csv, workbook_sheet_2_Data.csv, etc.
```

**Force specific table boundaries on specific sheet:**
```bash
./excel2csv -input data.xlsx -sheet-name "Summary" -start-row 3
```

**Use tab separator with specific sheet:**
```bash
./excel2csv -input data.xlsx -sheet-index 1 -separator "tab"
```

## HTTP API Server

The project includes a web server that provides HTTP API for Excel to CSV conversion, perfect for integration with web applications and microservices.

### Starting the Server

```bash
# Start on default port 8080
./excel2csv-server

# Start on custom port
PORT=8082 ./excel2csv-server
```

### API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/health` | GET | Server health check and LibreOffice status |
| `/convert` | POST | Convert Excel file to CSV |
| `/info` | GET | API information and supported features |
| `/` | GET | Web interface for file upload |

### API Examples

**Health Check:**
```bash
curl http://localhost:8080/health
```

**Basic Conversion:**
```bash
curl -X POST -F "file=@input.xlsx" -F "separator=comma" \
  -o result.csv http://localhost:8080/convert
```

**Convert All Sheets (returns ZIP):**
```bash
curl -X POST -F "file=@input.xlsx" -F "all_sheets=true" \
  -o results.zip http://localhost:8080/convert
```

**Advanced Options:**
```bash
curl -X POST \
  -F "file=@input.xlsx" \
  -F "separator=semicolon" \
  -F "start_row=3" \
  -F "sheet_name=Data" \
  -o result.csv http://localhost:8080/convert
```

### API Parameters

| Parameter | Type | Description | Values |
|-----------|------|-------------|--------|
| `file` | file | Excel file (required) | .xlsx, .xls, .ods |
| `separator` | string | CSV separator | `comma`, `semicolon`, `tab` |
| `start_row` | integer | Force start row (0-based) | 0, 1, 2, ... |
| `sheet_name` | string | Specific sheet name | Sheet name |
| `sheet_index` | integer | Specific sheet index (0-based) | 0, 1, 2, ... |
| `all_sheets` | boolean | Convert all sheets | `true`, `false` |

### Web Interface

The server provides a simple web interface at `http://localhost:8080/` for:
- File upload via drag & drop or file picker
- Configuration of conversion parameters
- Direct download of converted files
- Multi-sheet conversion with ZIP download

### Docker Support

```bash
# Build image
docker build -t excel2csv .

# Run HTTP server
docker run -p 8080:8080 excel2csv

# Test the containerized API
curl -X POST -F "file=@test.xlsx" http://localhost:8080/convert
```

## How It Works

### Multi-Sheet Support

The converter provides flexible sheet handling:

1. **Sheet Discovery**: Automatically detects all available sheets in Excel files
2. **Sheet Selection**: Choose sheets by name or zero-based index
3. **Batch Processing**: Convert all sheets at once with descriptive filenames
4. **Fallback Detection**: Uses multiple methods to identify sheet names and count

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

The tool has been tested with various file types and sheet configurations:

- **Multi-Sheet Reports**: Process workbooks with data, charts, and summary sheets
- **Price Lists**: Automatically skips supplier information headers across multiple sheets
- **Invoices**: Excludes totals and footer information from specific sheets
- **Stock Lists**: Handles large datasets (tested with 900k+ rows) across multiple worksheets
- **Financial Reports**: Processes quarterly/monthly data from different sheets

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
    
    // Sheet selection options
    converter.SheetName = "Sales Data"        // Convert specific sheet by name
    // OR
    sheetIndex := 1
    converter.SheetIndex = &sheetIndex        // Convert specific sheet by index
    // OR
    converter.AllSheetsMode = true            // Convert all sheets
    
    // Optional: force specific rows
    startRow := 5
    converter.ForceDataStartRow = &startRow
    
    // Convert file
    err := converter.ConvertFile("input.xlsx", "output.csv")
    if err != nil {
        panic(err)
    }
    
    // List sheets programmatically
    sheets, err := converter.ListSheets("input.xlsx")
    if err != nil {
        panic(err)
    }
    
    for _, sheet := range sheets {
        fmt.Printf("Sheet %d: %s\n", sheet.Index, sheet.Name)
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

**LibreOffice Snap Compatibility:**
When using LibreOffice installed via snap, you might encounter file access issues. The application automatically handles this by:
- Using home directory for temporary files instead of `/tmp/`
- Adjusting file paths for snap container compatibility
- Setting appropriate environment variables

If you see warnings like:
```
/snap/libreoffice/355/javasettings.py:44: SyntaxWarning: invalid escape sequence '\d'
```
These are normal and don't affect functionality.

**HTTP Server Issues:**
```
Error: source file could not be loaded
```
*Solution*: This is automatically resolved in the latest version. The HTTP server now uses home directory for temporary files, ensuring LibreOffice snap compatibility.

**Permission denied:**
```
Error: permission denied
```
*Solution*: Ensure read access to input file and write access to output directory

**Memory issues with large files:**
*Solution*: Use `-start-row` and `-end-row` to process file in chunks

## License

MIT License
