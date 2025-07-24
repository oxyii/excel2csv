package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/oxyii/excel2csv"
)

func main() {
	var (
		inputFile     = flag.String("input", "", "Path to input Excel file (.xls, .xlsx, .ods)")
		outputFile    = flag.String("output", "", "Path to output CSV file (optional)")
		separatorFlag = flag.String("separator", ",", "CSV separator: ',' (comma), ';' (semicolon), 'tab' (tab)")
		startRowFlag  = flag.Int("start-row", -1, "Force data start from specific row (0-based), -1 for auto-detection")
		helpFlag      = flag.Bool("help", false, "Show help")
	)

	flag.Parse()

	if *helpFlag {
		showHelp()
		return
	}

	if *inputFile == "" {
		fmt.Println("Error: input file must be specified")
		showHelp()
		os.Exit(1)
	}

	// Check if input file exists
	if _, err := os.Stat(*inputFile); os.IsNotExist(err) {
		log.Fatalf("Input file does not exist: %s", *inputFile)
	}

	// Generate output file name if not specified
	if *outputFile == "" {
		ext := filepath.Ext(*inputFile)
		baseName := strings.TrimSuffix(*inputFile, ext)
		*outputFile = baseName + ".csv"
	}

	// Create converter
	converter := excel2csv.NewExcelConverter()

	// Set forced data start row if specified
	if *startRowFlag >= 0 {
		converter.ForceDataStartRow = startRowFlag
	}

	// Set CSV separator
	switch *separatorFlag {
	case ",":
		converter.CSVSeparator = ','
	case ";":
		converter.CSVSeparator = ';'
	case "tab":
		converter.CSVSeparator = '\t'
	default:
		if len(*separatorFlag) == 1 {
			converter.CSVSeparator = rune((*separatorFlag)[0])
		} else {
			log.Fatalf("Invalid separator: %s", *separatorFlag)
		}
	}

	// Print configuration
	fmt.Printf("Converting file: %s\n", *inputFile)
	fmt.Printf("Output file: %s\n", *outputFile)
	fmt.Printf("CSV separator: %s\n", getSeparatorName(*separatorFlag))

	// Convert file
	err := converter.ConvertFile(*inputFile, *outputFile)
	if err != nil {
		log.Fatalf("Conversion error: %v", err)
	}

	fmt.Println("Conversion completed successfully!")
}

func showHelp() {
	fmt.Println("Excel to CSV Converter (LibreOffice-based)")
	fmt.Println("Convert Excel files (.xls/.xlsx/.ods) to CSV")
	fmt.Println()
	fmt.Println("Usage:")
	fmt.Println("  go run . -input <excel_file_path> [-output <csv_file_path>] [-separator <separator>] [-start-row <row_number>]")
	fmt.Println()
	fmt.Println("Flags:")
	fmt.Println("  -help")
	fmt.Println("        Show help")
	fmt.Println("  -input string")
	fmt.Println("        Path to input Excel file (.xls, .xlsx, or .ods)")
	fmt.Println("  -output string")
	fmt.Println("        Path to output CSV file (optional)")
	fmt.Println("  -separator string")
	fmt.Println("        CSV separator: ',' (comma), ';' (semicolon), 'tab' (tab) (default \",\")")
	fmt.Println("  -start-row int")
	fmt.Println("        Force data start from specific row (0-based), -1 for auto-detection (default -1)")
	fmt.Println()
	fmt.Println("Examples:")
	fmt.Println("  go run . -input data.xlsx")
	fmt.Println("  go run . -input data.xls -output result.csv")
	fmt.Println("  go run . -input data.xlsx -separator ';'")
	fmt.Println("  go run . -input data.xlsx -start-row 5")
	fmt.Println("  go run . -input data.xlsx -start-row 3 -separator tab")
	fmt.Println()
	fmt.Println("Features:")
	fmt.Println("- 🔧 LibreOffice-powered conversion (reliable for all Excel formats)")
	fmt.Println("- 📋 Support for .xls, .xlsx, and .ods formats")
	fmt.Println("- ⚙️ Configurable CSV separator")
	fmt.Println("- 🧹 Automatic cleanup of line breaks in data")
	fmt.Println("- 🎯 Manual override for data start row when needed")
	fmt.Println()
	fmt.Println("Requirements:")
	fmt.Println("- LibreOffice must be installed and available in PATH")
}

func getSeparatorName(sep string) string {
	switch sep {
	case ",":
		return "comma (,)"
	case ";":
		return "semicolon (;)"
	case "tab":
		return "tab (\\t)"
	default:
		return fmt.Sprintf("custom (%s)", sep)
	}
}
