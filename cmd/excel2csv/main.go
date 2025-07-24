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
		sheetName     = flag.String("sheet-name", "", "Convert specific sheet by name")
		sheetIndex    = flag.Int("sheet-index", -1, "Convert specific sheet by index (0-based), -1 for first sheet")
		listSheets    = flag.Bool("list-sheets", false, "List all sheets in the Excel file and exit")
		allSheets     = flag.Bool("all-sheets", false, "Convert all sheets to separate CSV files")
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

	// Create converter
	converter := excel2csv.NewExcelConverter()

	// Handle list sheets command
	if *listSheets {
		sheets, err := converter.ListSheets(*inputFile)
		if err != nil {
			log.Fatalf("Failed to list sheets: %v", err)
		}

		fmt.Printf("Sheets in file %s:\n", *inputFile)
		for _, sheet := range sheets {
			fmt.Printf("  %d: %s\n", sheet.Index, sheet.Name)
		}
		return
	}

	// Set sheet selection
	if *sheetName != "" && *sheetIndex >= 0 {
		log.Fatalf("Cannot specify both -sheet-name and -sheet-index")
	}

	if *sheetName != "" {
		converter.SheetName = *sheetName
	} else if *sheetIndex >= 0 {
		converter.SheetIndex = sheetIndex
	}

	// Set convert all sheets mode
	converter.AllSheetsMode = *allSheets

	// Generate output file name if not specified
	if *outputFile == "" {
		if *allSheets {
			// For all sheets mode, use input directory
			*outputFile = filepath.Dir(*inputFile)
			if *outputFile == "" {
				*outputFile = "."
			}
		} else {
			ext := filepath.Ext(*inputFile)
			baseName := strings.TrimSuffix(*inputFile, ext)
			if *sheetName != "" {
				*outputFile = baseName + "_" + *sheetName + ".csv"
			} else if *sheetIndex >= 0 {
				*outputFile = fmt.Sprintf("%s_sheet_%d.csv", baseName, *sheetIndex+1)
			} else {
				*outputFile = baseName + ".csv"
			}
		}
	}

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
	if *allSheets {
		fmt.Printf("Converting all sheets to directory: %s\n", *outputFile)
	} else {
		fmt.Printf("Output file: %s\n", *outputFile)
		if *sheetName != "" {
			fmt.Printf("Sheet: %s\n", *sheetName)
		} else if *sheetIndex >= 0 {
			fmt.Printf("Sheet index: %d\n", *sheetIndex)
		} else {
			fmt.Printf("Sheet: first sheet (default)\n")
		}
	}
	fmt.Printf("CSV separator: %s\n", getSeparatorName(*separatorFlag))

	// Convert file
	err := converter.ConvertFile(*inputFile, *outputFile)
	if err != nil {
		log.Fatalf("Conversion error: %v", err)
	}

	if *allSheets {
		fmt.Println("All sheets converted successfully!")
	} else {
		fmt.Println("Conversion completed successfully!")
	}
}

func showHelp() {
	fmt.Println("Excel to CSV Converter (LibreOffice-based)")
	fmt.Println("Convert Excel files (.xls/.xlsx/.ods) to CSV with multi-sheet support")
	fmt.Println()
	fmt.Println("Usage:")
	fmt.Println("  go run . -input <excel_file_path> [options]")
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
	fmt.Println("Sheet Selection:")
	fmt.Println("  -list-sheets")
	fmt.Println("        List all sheets in the Excel file and exit")
	fmt.Println("  -sheet-name string")
	fmt.Println("        Convert specific sheet by name")
	fmt.Println("  -sheet-index int")
	fmt.Println("        Convert specific sheet by index (0-based), -1 for first sheet (default -1)")
	fmt.Println("  -all-sheets")
	fmt.Println("        Convert all sheets to separate CSV files")
	fmt.Println()
	fmt.Println("Examples:")
	fmt.Println("  # Convert first sheet (default)")
	fmt.Println("  go run . -input data.xlsx")
	fmt.Println()
	fmt.Println("  # List all sheets")
	fmt.Println("  go run . -input data.xlsx -list-sheets")
	fmt.Println()
	fmt.Println("  # Convert specific sheet by name")
	fmt.Println("  go run . -input data.xlsx -sheet-name \"Sales Data\"")
	fmt.Println()
	fmt.Println("  # Convert specific sheet by index (0-based)")
	fmt.Println("  go run . -input data.xlsx -sheet-index 1")
	fmt.Println()
	fmt.Println("  # Convert all sheets to separate files")
	fmt.Println("  go run . -input data.xlsx -all-sheets")
	fmt.Println()
	fmt.Println("  # Convert with custom separator")
	fmt.Println("  go run . -input data.xlsx -sheet-name \"Report\" -separator ';'")
	fmt.Println()
	fmt.Println("  # Force start row and convert specific sheet")
	fmt.Println("  go run . -input data.xlsx -sheet-index 2 -start-row 5")
	fmt.Println()
	fmt.Println("Features:")
	fmt.Println("- 🔧 LibreOffice-powered conversion (reliable for all Excel formats)")
	fmt.Println("- 📋 Support for .xls, .xlsx, and .ods formats")
	fmt.Println("- 📄 Multi-sheet support: select by name/index or convert all sheets")
	fmt.Println("- ⚙️ Configurable CSV separator")
	fmt.Println("- 🧹 Automatic cleanup of line breaks in data")
	fmt.Println("- 🎯 Manual override for data start row when needed")
	fmt.Println("- 📝 Sheet listing to see available worksheets")
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
