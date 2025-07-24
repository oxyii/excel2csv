package excel2csv

import (
	"encoding/csv"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strconv"
	"strings"
	"time"
)

// ExcelConverter handles Excel to CSV conversion using LibreOffice
type ExcelConverter struct {
	CSVSeparator      rune // CSV separator (comma, semicolon, tab)
	CleanLineBreaks   bool // replace line breaks with spaces
	ForceDataStartRow *int // force data start from specific row (0-based), nil for auto-detection
	ForceDataEndRow   *int // force data end at specific row (0-based), nil for auto-detection
}

// NewExcelConverter creates a new converter with default settings
func NewExcelConverter() *ExcelConverter {
	return &ExcelConverter{
		CSVSeparator:    ',',  // comma separator by default
		CleanLineBreaks: true, // clean line breaks by default
	}
}

// ConvertFile converts an Excel file to CSV using LibreOffice
func (ec *ExcelConverter) ConvertFile(inputPath, outputPath string) error {
	ext := strings.ToLower(filepath.Ext(inputPath))

	// Check if the file is a supported Excel format
	switch ext {
	case ".xlsx", ".xls", ".ods":
		return ec.convertViaLibreOffice(inputPath, outputPath)
	default:
		return fmt.Errorf("unsupported file format: %s. Supported formats: .xlsx, .xls, .ods", ext)
	}
}

// convertViaLibreOffice converts Excel files using LibreOffice headless mode
func (ec *ExcelConverter) convertViaLibreOffice(inputPath, outputPath string) error {
	// Check if LibreOffice is available
	_, err := exec.LookPath("libreoffice")
	if err != nil {
		return fmt.Errorf("LibreOffice is not available. Please install LibreOffice")
	}

	// Create temp directory
	homeDir, _ := os.UserHomeDir()
	tempDir := filepath.Join(homeDir, "excel2csv_temp")
	os.MkdirAll(tempDir, 0755)
	defer os.RemoveAll(tempDir)

	// Convert using LibreOffice
	absInputPath, _ := filepath.Abs(inputPath)
	cmd := exec.Command("libreoffice", "--headless", "--convert-to", "csv", "--outdir", tempDir, absInputPath)

	output, err := cmd.CombinedOutput()
	fmt.Printf("LibreOffice output: %s\n", string(output))

	if err != nil {
		return fmt.Errorf("LibreOffice conversion failed: %w", err)
	}

	time.Sleep(200 * time.Millisecond)

	// Find generated CSV file
	files, _ := os.ReadDir(tempDir)
	var tempCSVPath string
	for _, file := range files {
		if strings.HasSuffix(strings.ToLower(file.Name()), ".csv") {
			tempCSVPath = filepath.Join(tempDir, file.Name())
			break
		}
	}

	if tempCSVPath == "" {
		return fmt.Errorf("LibreOffice did not generate CSV file")
	}

	// Read and copy CSV file
	return ec.copyCSVFile(tempCSVPath, outputPath)
}

func (ec *ExcelConverter) copyCSVFile(srcPath, dstPath string) error {
	srcFile, err := os.Open(srcPath)
	if err != nil {
		return err
	}
	defer srcFile.Close()

	dstFile, err := os.Create(dstPath)
	if err != nil {
		return err
	}
	defer dstFile.Close()

	reader := csv.NewReader(srcFile)
	writer := csv.NewWriter(dstFile)
	defer writer.Flush()

	// Set CSV separator
	writer.Comma = ec.CSVSeparator

	records, err := reader.ReadAll()
	if err != nil {
		return err
	}

	// Apply intelligent processing to detect table boundaries
	processedRecords := ec.processTableData(records)

	for _, record := range processedRecords {
		// Clean line breaks if needed
		if ec.CleanLineBreaks {
			for i, cell := range record {
				record[i] = ec.cleanCellData(cell)
			}
		}
		if err := writer.Write(record); err != nil {
			return err
		}
	}

	return nil
}

// processTableData intelligently processes table data based on structure analysis
func (ec *ExcelConverter) processTableData(records [][]string) [][]string {
	if len(records) == 0 {
		return records
	}

	// If manual boundaries are specified, use them
	if ec.ForceDataStartRow != nil && ec.ForceDataEndRow != nil {
		start := *ec.ForceDataStartRow
		end := *ec.ForceDataEndRow
		if start >= 0 && end >= start && start < len(records) && end < len(records) {
			fmt.Printf("Using manual boundaries: rows %d to %d\n", start+1, end+1)
			return records[start : end+1]
		}
	}

	// Use only the improved boundary detection
	tableStart, tableEnd := ec.detectTableBoundariesImproved(records)

	fmt.Printf("Detected table boundaries: start row %d, end row %d\n", tableStart+1, tableEnd+1)

	if tableStart >= 0 && tableEnd >= tableStart && tableEnd < len(records) {
		result := records[tableStart : tableEnd+1]
		fmt.Printf("Returning %d rows from the table\n", len(result))
		return result
	}

	// Fallback: return all records
	fmt.Printf("Fallback: returning all %d records\n", len(records))
	return records
}

// detectTableBoundariesImproved uses the insights from structure analysis
func (ec *ExcelConverter) detectTableBoundariesImproved(records [][]string) (int, int) {
	if len(records) == 0 {
		return 0, 0
	}

	// Find the row with maximum non-empty cells and minimal numeric content (likely headers)
	headerRow := -1
	maxNonEmpty := 0

	for i, record := range records {
		nonEmpty := ec.countNonEmptyCells(record)
		numeric := ec.countNumericCells(record)

		// Good header candidate: many non-empty cells, few numbers
		if nonEmpty >= 5 && numeric <= 1 && nonEmpty > maxNonEmpty {
			maxNonEmpty = nonEmpty
			headerRow = i
		}
	}

	if headerRow == -1 {
		// Fallback: first row with data
		for i, record := range records {
			if ec.hasData(record) {
				return i, len(records) - 1
			}
		}
		return 0, 0
	}

	fmt.Printf("Found header row at %d with %d non-empty cells\n", headerRow+1, maxNonEmpty)

	// Find the end: look for rows that maintain similar structure
	tableEnd := headerRow
	expectedCols := maxNonEmpty

	for i := headerRow + 1; i < len(records); i++ {
		nonEmpty := ec.countNonEmptyCells(records[i])

		// If row has significantly fewer cells, it's likely a footer/total
		if nonEmpty > 0 && nonEmpty < expectedCols/3 {
			fmt.Printf("Stopping at row %d - footer detected (%d cols vs expected %d)\n", i+1, nonEmpty, expectedCols)
			break
		}

		// If row has reasonable number of cells, include it
		if nonEmpty >= expectedCols/2 {
			tableEnd = i
		} else if nonEmpty == 0 {
			// Empty row - could be end or separator
			break
		}
	}

	return headerRow, tableEnd
}

// detectTableBoundaries detects table boundaries based on data structure analysis
func (ec *ExcelConverter) detectTableBoundaries(records [][]string) (int, int) {
	if len(records) == 0 {
		return 0, 0
	}

	// Step 1: Find the most consistent table structure
	tableStart := ec.findTableStart(records)
	tableEnd := ec.findTableEnd(records, tableStart)

	// Step 2: Check if there's a header row just before table data
	if tableStart > 0 {
		headerCandidate := tableStart - 1
		if ec.looksLikeHeaderRow(records[headerCandidate], records[tableStart]) {
			fmt.Printf("Found header row at %d\n", headerCandidate+1)
			tableStart = headerCandidate
		}
	}

	return tableStart, tableEnd
}

// findTableStart finds the start of consistent tabular data
func (ec *ExcelConverter) findTableStart(records [][]string) int {
	if ec.ForceDataStartRow != nil {
		return *ec.ForceDataStartRow
	}

	// Look for rows with consistent structure and data types
	for i := 0; i < len(records)-2; i++ { // Need at least 2 more rows to check consistency
		if ec.isDataRow(records[i]) {
			// Check if next few rows have similar structure
			consistency := ec.checkStructuralConsistency(records, i, 3)
			fmt.Printf("Row %d: data=%v, consistency=%.2f\n", i+1, ec.isDataRow(records[i]), consistency)

			if consistency > 0.6 { // Lower threshold but with stricter isDataRow
				return i
			}
		}
	}

	// Fallback: look for any data row in the second half of the file
	for i := len(records) / 2; i < len(records); i++ {
		if ec.isDataRow(records[i]) {
			return i
		}
	}

	// Final fallback: first non-empty row
	for i, record := range records {
		if ec.hasData(record) {
			return i
		}
	}

	return 0
}

// findTableEnd finds the end of consistent tabular data
func (ec *ExcelConverter) findTableEnd(records [][]string, startRow int) int {
	if ec.ForceDataEndRow != nil {
		return *ec.ForceDataEndRow
	}

	if startRow >= len(records) {
		return len(records) - 1
	}

	// Determine expected column count from start area
	expectedCols := ec.getExpectedColumnCount(records, startRow)
	lastGoodRow := startRow

	fmt.Printf("Expected columns: %d, starting from row %d\n", expectedCols, startRow+1)

	for i := startRow; i < len(records); i++ {
		record := records[i]
		cols := ec.countNonEmptyCells(record)
		isData := ec.isDataRow(record) || ec.looksLikeHeaderRow(record, records[min(i+1, len(records)-1)])
		isPartOfTable := ec.isPartOfTable(record, expectedCols)

		fmt.Printf("Row %d: cols=%d, isData=%v, isPartOfTable=%v\n", i+1, cols, isData, isPartOfTable)

		// Check if row maintains table structure
		if isPartOfTable && (isData || i == startRow) {
			lastGoodRow = i
		} else {
			// Special case: if this looks like a summary/total row with fewer columns, stop here
			if cols > 0 && cols < expectedCols/2 {
				fmt.Printf("Stopping at row %d - looks like summary/total\n", i+1)
				break
			}
			// If row is completely empty or very different structure, stop
			if cols == 0 || abs(cols-expectedCols) > 3 {
				break
			}
		}
	}

	return lastGoodRow
}

// isDataRow checks if a row contains structured data
func (ec *ExcelConverter) isDataRow(record []string) bool {
	nonEmptyCount := 0
	numericCount := 0

	for _, cell := range record {
		cell = strings.TrimSpace(cell)
		if cell != "" {
			nonEmptyCount++
			if ec.looksLikeNumber(cell) {
				numericCount++
			}
		}
	}

	// Data row should have multiple cells (at least 3) and at least one numeric value
	// This helps distinguish table data from contact info or single-value rows
	return nonEmptyCount >= 3 && numericCount >= 1
}

// looksLikeHeaderRow checks if a row could be headers for the data row
func (ec *ExcelConverter) looksLikeHeaderRow(headerRow, dataRow []string) bool {
	// Headers should have similar column count to data
	headerCols := ec.countNonEmptyCells(headerRow)
	dataCols := ec.countNonEmptyCells(dataRow)

	if headerCols < 2 || abs(headerCols-dataCols) > 2 {
		return false
	}

	// Headers should be mostly text, data should have numbers
	headerNumeric := ec.countNumericCells(headerRow)
	dataNumeric := ec.countNumericCells(dataRow)

	// Headers should have less numeric content than data
	return headerNumeric < dataNumeric || (headerNumeric == 0 && dataNumeric > 0)
}

// checkStructuralConsistency checks how consistent the structure is across rows
func (ec *ExcelConverter) checkStructuralConsistency(records [][]string, startRow, checkCount int) float64 {
	if startRow+checkCount > len(records) {
		checkCount = len(records) - startRow
	}

	if checkCount < 1 {
		return 0.0
	}

	referenceCols := ec.countNonEmptyCells(records[startRow])
	if referenceCols < 2 {
		return 0.0
	}

	matches := 0
	totalRows := 0

	for i := 0; i < checkCount; i++ {
		row := records[startRow+i]
		cols := ec.countNonEmptyCells(row)
		totalRows++

		// More flexible matching - allow headers and data rows
		if abs(cols-referenceCols) <= 2 { // Allow more variation
			// Either it's a data row, or it's the first row (could be header)
			if ec.isDataRow(row) || i == 0 {
				matches++
			}
		}
	}

	return float64(matches) / float64(totalRows)
}

// Helper functions
func (ec *ExcelConverter) hasData(record []string) bool {
	for _, cell := range record {
		if strings.TrimSpace(cell) != "" {
			return true
		}
	}
	return false
}

func (ec *ExcelConverter) countNonEmptyCells(record []string) int {
	count := 0
	for _, cell := range record {
		if strings.TrimSpace(cell) != "" {
			count++
		}
	}
	return count
}

func (ec *ExcelConverter) countNumericCells(record []string) int {
	count := 0
	for _, cell := range record {
		if ec.looksLikeNumber(strings.TrimSpace(cell)) {
			count++
		}
	}
	return count
}

func (ec *ExcelConverter) looksLikeNumber(value string) bool {
	if value == "" {
		return false
	}

	// Remove common number formatting
	value = strings.ReplaceAll(value, ",", "")
	value = strings.ReplaceAll(value, " ", "")

	_, err := strconv.ParseFloat(value, 64)
	return err == nil
}

func (ec *ExcelConverter) getExpectedColumnCount(records [][]string, startRow int) int {
	maxCols := 0
	for i := startRow; i < startRow+3 && i < len(records); i++ {
		cols := ec.countNonEmptyCells(records[i])
		if cols > maxCols {
			maxCols = cols
		}
	}
	return maxCols
}

func (ec *ExcelConverter) isPartOfTable(record []string, expectedCols int) bool {
	cols := ec.countNonEmptyCells(record)
	// Allow some variation but not too much
	return cols > 0 && abs(cols-expectedCols) <= 2
}

func abs(x int) int {
	if x < 0 {
		return -x
	}
	return x
}

// cleanCellData cleans problematic characters from cell data
func (ec *ExcelConverter) cleanCellData(text string) string {
	if !ec.CleanLineBreaks {
		return text
	}

	// Replace line breaks with spaces
	text = strings.ReplaceAll(text, "\n", " ")
	text = strings.ReplaceAll(text, "\r", " ")
	text = strings.ReplaceAll(text, "\r\n", " ")

	// Clean up multiple spaces
	for strings.Contains(text, "  ") {
		text = strings.ReplaceAll(text, "  ", " ")
	}

	return strings.TrimSpace(text)
}

// Helper function for min
func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
