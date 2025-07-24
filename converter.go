package excel2csv

import (
	"context"
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
	CSVSeparator      rune   // CSV separator (comma, semicolon, tab)
	CleanLineBreaks   bool   // replace line breaks with spaces
	ForceDataStartRow *int   // force data start from specific row (0-based), nil for auto-detection
	ForceDataEndRow   *int   // force data end at specific row (0-based), nil for auto-detection
	SheetName         string // specific sheet name to convert
	SheetIndex        *int   // specific sheet index to convert (0-based)
	AllSheetsMode     bool   // convert all sheets to separate CSV files
	TempDir           string // custom temp directory (if empty, uses default)
}

// SheetInfo contains information about a worksheet
type SheetInfo struct {
	Index int
	Name  string
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

	// Handle ConvertAllSheets mode
	if ec.AllSheetsMode {
		outputDir := filepath.Dir(outputPath)
		return ec.ConvertAllSheetsToFiles(inputPath, outputDir)
	}

	// Create temp directory with better permissions for HTTP context
	homeDir, _ := os.UserHomeDir()
	tempDir := ec.TempDir
	if tempDir == "" {
		tempDir = filepath.Join(homeDir, "excel2csv_temp")
	}

	// For HTTP context, ensure we use a subdirectory in home dir for better LibreOffice compatibility
	if strings.HasPrefix(tempDir, "/tmp/") {
		fmt.Printf("Warning: Using /tmp directory may cause LibreOffice issues, switching to home directory\n")
		tempDir = filepath.Join(homeDir, "excel2csv_temp_http")
	}

	_ = os.MkdirAll(tempDir, 0755)
	defer func() {
		if !strings.Contains(tempDir, "excel2csv_temp") {
			// Only remove if it's our temp directory
			_ = os.RemoveAll(tempDir)
		}
	}()

	// Convert using LibreOffice - improved for HTTP context
	absInputPath, _ := filepath.Abs(inputPath)

	// Check if input file exists and is readable
	if stat, err := os.Stat(absInputPath); err != nil {
		return fmt.Errorf("input file not accessible: %w", err)
	} else {
		fmt.Printf("Input file: %s (size: %d bytes, mode: %v)\n", absInputPath, stat.Size(), stat.Mode())
	}

	// For now, we'll only convert the first/default sheet since --sheet parameter is not supported
	// TODO: Implement proper multi-sheet support using LibreOffice UNO API or other methods
	if ec.SheetName != "" {
		fmt.Printf("Warning: sheet selection by name '%s' is not fully supported yet, converting default sheet\n", ec.SheetName)
	}
	if ec.SheetIndex != nil {
		fmt.Printf("Warning: sheet selection by index %d is not fully supported yet, converting default sheet\n", *ec.SheetIndex)
	}

	cmd := exec.Command("libreoffice", "--headless", "--convert-to", "csv", "--outdir", tempDir, absInputPath)

	// Set environment variables to fix LibreOffice issues in HTTP context
	cmd.Env = append(os.Environ(),
		"HOME="+homeDir,
		"TMPDIR="+tempDir,
		"DISPLAY=", // Empty DISPLAY for headless mode
		"LANG=en_US.UTF-8",
	)

	output, err := cmd.CombinedOutput()
	fmt.Printf("LibreOffice output: %s\n", string(output))

	if err != nil {
		return fmt.Errorf("LibreOffice conversion failed: %w", err)
	}

	time.Sleep(200 * time.Millisecond)

	// Find generated CSV file
	files, err := os.ReadDir(tempDir)
	if err != nil {
		fmt.Printf("Error reading temp directory %s: %v\n", tempDir, err)
		return fmt.Errorf("failed to read temp directory: %w", err)
	}

	fmt.Printf("Files in temp directory %s: %d files\n", tempDir, len(files))
	for _, file := range files {
		fmt.Printf("  - %s (isDir: %v)\n", file.Name(), file.IsDir())
	}

	var tempCSVPath string
	for _, file := range files {
		if strings.HasSuffix(strings.ToLower(file.Name()), ".csv") {
			tempCSVPath = filepath.Join(tempDir, file.Name())
			fmt.Printf("Found CSV file: %s\n", tempCSVPath)
			break
		}
	}

	if tempCSVPath == "" {
		fmt.Printf("No CSV files found in temp directory %s\n", tempDir)
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
	defer func() { _ = srcFile.Close() }()

	dstFile, err := os.Create(dstPath)
	if err != nil {
		return err
	}
	defer func() { _ = dstFile.Close() }()

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
		isData := ec.isDataRow(record) || ec.looksLikeHeaderRow(record, records[minInt(i+1, len(records)-1)])
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

// ListSheets returns information about all sheets in the Excel file
func (ec *ExcelConverter) ListSheets(inputPath string) ([]SheetInfo, error) {
	// Check if LibreOffice is available
	_, err := exec.LookPath("libreoffice")
	if err != nil {
		return nil, fmt.Errorf("LibreOffice is not available. Please install LibreOffice")
	}

	// Create temp directory
	homeDir, _ := os.UserHomeDir()
	tempDir := filepath.Join(homeDir, "excel2csv_temp_sheets")
	_ = os.MkdirAll(tempDir, 0755)
	defer func() { _ = os.RemoveAll(tempDir) }()

	// Use simpler fallback method by default (more reliable)
	return ec.fallbackListSheets(inputPath, tempDir)
}

// fallbackListSheets tries to detect sheets by attempting conversions
func (ec *ExcelConverter) fallbackListSheets(inputPath, tempDir string) ([]SheetInfo, error) {
	var sheets []SheetInfo
	absInputPath, _ := filepath.Abs(inputPath)

	fmt.Printf("Detecting sheets in %s...\n", filepath.Base(inputPath))

	// Since --sheet parameter is not supported, we can only reliably detect the first sheet
	// For now, just try to convert the default sheet and assume it exists
	fmt.Printf("Checking sheet 0... ")

	cmd := exec.Command("libreoffice", "--headless", "--convert-to", "csv",
		"--outdir", tempDir, absInputPath)

	// Set a timeout to avoid hanging
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	cmd = exec.CommandContext(ctx, cmd.Args[0], cmd.Args[1:]...)

	_, err := cmd.CombinedOutput()
	if err == nil {
		// Check if a CSV file was actually created
		files, _ := os.ReadDir(tempDir)
		csvFound := false
		for _, file := range files {
			if strings.HasSuffix(strings.ToLower(file.Name()), ".csv") {
				csvFound = true
				// Clean up the CSV file
				os.Remove(filepath.Join(tempDir, file.Name()))
				break
			}
		}

		if csvFound {
			sheets = append(sheets, SheetInfo{
				Index: 0,
				Name:  "Sheet1",
			})
			fmt.Printf("✓ found\n")
		} else {
			fmt.Printf("✗ no output\n")
		}
	} else {
		fmt.Printf("✗ error\n")
	}

	if len(sheets) == 0 {
		// Fallback - assume at least one sheet exists
		sheets = append(sheets, SheetInfo{
			Index: 0,
			Name:  "Sheet1",
		})
	}

	fmt.Printf("Note: Advanced multi-sheet detection requires LibreOffice version with --sheet support\n")
	return sheets, nil
}

// ConvertAllSheetsToFiles converts all sheets to separate CSV files
func (ec *ExcelConverter) ConvertAllSheetsToFiles(inputPath, outputDir string) error {
	sheets, err := ec.ListSheets(inputPath)
	if err != nil {
		return fmt.Errorf("failed to list sheets: %w", err)
	}

	if len(sheets) == 0 {
		return fmt.Errorf("no sheets found in file")
	}

	// Create output directory if it doesn't exist
	err = os.MkdirAll(outputDir, 0755)
	if err != nil {
		return fmt.Errorf("failed to create output directory: %w", err)
	}

	// Convert each sheet
	for _, sheet := range sheets {
		// Generate output filename
		baseName := strings.TrimSuffix(filepath.Base(inputPath), filepath.Ext(inputPath))
		outputFile := filepath.Join(outputDir, fmt.Sprintf("%s_sheet_%d_%s.csv", baseName, sheet.Index+1, sheet.Name))

		// Clean filename
		outputFile = strings.ReplaceAll(outputFile, " ", "_")
		outputFile = strings.ReplaceAll(outputFile, "/", "_")
		outputFile = strings.ReplaceAll(outputFile, "\\", "_")

		fmt.Printf("Converting sheet %d (%s) to %s\n", sheet.Index+1, sheet.Name, outputFile)

		// Create a temporary converter for this sheet
		tempConverter := *ec
		tempConverter.SheetIndex = &sheet.Index
		tempConverter.AllSheetsMode = false

		err = tempConverter.ConvertFile(inputPath, outputFile)
		if err != nil {
			fmt.Printf("Warning: failed to convert sheet %s: %v\n", sheet.Name, err)
		}
	}

	return nil
}

// convertSpecificSheet converts a specific sheet by index or name
func (ec *ExcelConverter) convertSpecificSheet(inputPath, tempDir string, sheetIndex int, sheetName string) error {
	absInputPath, _ := filepath.Abs(inputPath)

	var cmd *exec.Cmd
	if sheetName != "" {
		// Convert by sheet name
		cmd = exec.Command("libreoffice", "--headless", "--convert-to", "csv",
			"--outdir", tempDir, "--sheet", sheetName, absInputPath)
	} else {
		// Convert by sheet index
		cmd = exec.Command("libreoffice", "--headless", "--convert-to", "csv",
			"--outdir", tempDir, "--sheet", fmt.Sprintf("%d", sheetIndex), absInputPath)
	}

	output, err := cmd.CombinedOutput()
	if err != nil {
		return fmt.Errorf("LibreOffice conversion failed: %w, output: %s", err, string(output))
	}

	return nil
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

// Helper function for min (renamed to avoid collision with builtin)
func minInt(a, b int) int {
	if a < b {
		return a
	}
	return b
}
