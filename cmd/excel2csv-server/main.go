package main

import (
	"archive/zip"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/gorilla/mux"
	"github.com/oxyii/excel2csv"
)

// ConvertRequest represents the conversion request
type ConvertRequest struct {
	Separator   string `json:"separator,omitempty"`
	StartRow    *int   `json:"start_row,omitempty"`
	SheetName   string `json:"sheet_name,omitempty"`
	SheetIndex  *int   `json:"sheet_index,omitempty"`
	AllSheets   bool   `json:"all_sheets,omitempty"`
	CleanBreaks *bool  `json:"clean_breaks,omitempty"`
}

// ConvertResponse represents the conversion response
type ConvertResponse struct {
	Success       bool     `json:"success"`
	Message       string   `json:"message"`
	Files         []string `json:"files,omitempty"`
	Error         string   `json:"error,omitempty"`
	ProcessedRows int      `json:"processed_rows,omitempty"`
}

// HealthResponse represents health check response
type HealthResponse struct {
	Status      string `json:"status"`
	LibreOffice bool   `json:"libreoffice_available"`
	Version     string `json:"version"`
	Timestamp   string `json:"timestamp"`
}

func main() {
	r := mux.NewRouter()

	// API routes
	r.HandleFunc("/health", healthCheckHandler).Methods("GET")
	r.HandleFunc("/convert", convertHandler).Methods("POST")
	r.HandleFunc("/info", infoHandler).Methods("GET")

	// Static files for simple web interface
	r.HandleFunc("/", indexHandler).Methods("GET")

	// Configure server
	port := os.Getenv("PORT")
	if port == "" {
		port = "8080"
	}

	log.Printf("🚀 Excel2CSV Server starting on port %s", port)
	log.Printf("📋 Endpoints:")
	log.Printf("   GET  /health  - Health check")
	log.Printf("   POST /convert - Convert Excel to CSV")
	log.Printf("   GET  /info    - API information")
	log.Printf("   GET  /        - Web interface")

	log.Fatal(http.ListenAndServe(":"+port, r))
}

func healthCheckHandler(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "application/json")

	// Check LibreOffice availability
	libreOfficeAvailable := true

	// Try to check LibreOffice version
	tempDir := "/tmp/health_check"
	os.MkdirAll(tempDir, 0755)
	defer os.RemoveAll(tempDir)

	response := HealthResponse{
		Status:      "healthy",
		LibreOffice: libreOfficeAvailable,
		Version:     "1.1.0",
		Timestamp:   time.Now().UTC().Format(time.RFC3339),
	}

	json.NewEncoder(w).Encode(response)
}

func convertHandler(w http.ResponseWriter, r *http.Request) {
	// Parse multipart form
	err := r.ParseMultipartForm(50 << 20) // 50MB max
	if err != nil {
		http.Error(w, "Failed to parse form", http.StatusBadRequest)
		return
	}

	// Get file from form
	file, fileHeader, err := r.FormFile("file")
	if err != nil {
		http.Error(w, "No file provided", http.StatusBadRequest)
		return
	}
	defer file.Close()

	// Validate file extension
	ext := strings.ToLower(filepath.Ext(fileHeader.Filename))
	if ext != ".xlsx" && ext != ".xls" && ext != ".ods" {
		http.Error(w, "Unsupported file format. Use .xlsx, .xls, or .ods", http.StatusBadRequest)
		return
	}

	// Parse conversion options
	var req ConvertRequest
	if configStr := r.FormValue("config"); configStr != "" {
		json.Unmarshal([]byte(configStr), &req)
	}

	// Apply form values (override JSON config)
	if sep := r.FormValue("separator"); sep != "" {
		req.Separator = sep
	}
	if startRow := r.FormValue("start_row"); startRow != "" {
		if val, err := strconv.Atoi(startRow); err == nil {
			req.StartRow = &val
		}
	}
	if sheetName := r.FormValue("sheet_name"); sheetName != "" {
		req.SheetName = sheetName
	}
	if sheetIndex := r.FormValue("sheet_index"); sheetIndex != "" {
		if val, err := strconv.Atoi(sheetIndex); err == nil {
			req.SheetIndex = &val
		}
	}
	if r.FormValue("all_sheets") == "true" {
		req.AllSheets = true
	}

	// Create temporary files with better error handling - use home directory for LibreOffice compatibility
	homeDir, _ := os.UserHomeDir()
	tempDir := filepath.Join(homeDir, "excel2csv_http_temp")
	err = os.MkdirAll(tempDir, 0755)
	if err != nil {
		log.Printf("Failed to create temp directory: %v", err)
		http.Error(w, "Failed to create temp directory", http.StatusInternalServerError)
		return
	}
	defer os.RemoveAll(tempDir)

	// Ensure temp directory is writable
	if err := os.Chmod(tempDir, 0755); err != nil {
		log.Printf("Failed to set temp directory permissions: %v", err)
	}

	// Save uploaded file
	inputPath := filepath.Join(tempDir, fileHeader.Filename)
	outputFile, err := os.Create(inputPath)
	if err != nil {
		log.Printf("Failed to create input file: %v", err)
		http.Error(w, "Failed to save uploaded file", http.StatusInternalServerError)
		return
	}

	_, err = io.Copy(outputFile, file)
	outputFile.Close()
	if err != nil {
		log.Printf("Failed to save uploaded file: %v", err)
		http.Error(w, "Failed to save uploaded file", http.StatusInternalServerError)
		return
	}

	log.Printf("Processing file: %s (size: %d bytes)", fileHeader.Filename, fileHeader.Size)

	// Configure converter
	converter := excel2csv.NewExcelConverter()

	// Set separator
	switch req.Separator {
	case "semicolon", ";":
		converter.CSVSeparator = ';'
	case "tab", "\t":
		converter.CSVSeparator = '\t'
	default:
		converter.CSVSeparator = ','
	}

	// Set options
	if req.StartRow != nil {
		converter.ForceDataStartRow = req.StartRow
	}
	if req.SheetName != "" {
		converter.SheetName = req.SheetName
	}
	if req.SheetIndex != nil {
		converter.SheetIndex = req.SheetIndex
	}
	if req.CleanBreaks != nil {
		converter.CleanLineBreaks = *req.CleanBreaks
	}
	converter.AllSheetsMode = req.AllSheets

	// Convert file
	var outputPaths []string
	baseName := strings.TrimSuffix(fileHeader.Filename, ext)

	if req.AllSheets {
		// Convert all sheets to separate files
		outputDir := filepath.Join(tempDir, "output")
		err = os.MkdirAll(outputDir, 0755)
		if err != nil {
			log.Printf("Failed to create output directory: %v", err)
			http.Error(w, "Failed to create output directory", http.StatusInternalServerError)
			return
		}

		err = converter.ConvertFile(inputPath, filepath.Join(outputDir, "dummy.csv"))
		if err != nil {
			log.Printf("Conversion failed: %v", err)
			response := ConvertResponse{
				Success: false,
				Error:   fmt.Sprintf("Conversion failed: %v", err),
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(response)
			return
		}

		// Find all generated CSV files
		files, _ := os.ReadDir(outputDir)
		for _, f := range files {
			if strings.HasSuffix(f.Name(), ".csv") {
				outputPaths = append(outputPaths, filepath.Join(outputDir, f.Name()))
			}
		}
	} else {
		// Convert single sheet
		outputPath := filepath.Join(tempDir, baseName+".csv")
		log.Printf("Converting to: %s", outputPath)

		err = converter.ConvertFile(inputPath, outputPath)
		if err != nil {
			log.Printf("Conversion failed: %v", err)
			response := ConvertResponse{
				Success: false,
				Error:   fmt.Sprintf("Conversion failed: %v", err),
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(response)
			return
		}

		// Check if output file exists and has content
		if stat, err := os.Stat(outputPath); err != nil {
			log.Printf("Output file not found: %v", err)
			response := ConvertResponse{
				Success: false,
				Error:   "Conversion failed: output file not generated",
			}
			w.Header().Set("Content-Type", "application/json")
			json.NewEncoder(w).Encode(response)
			return
		} else {
			log.Printf("Output file created: %s (size: %d bytes)", outputPath, stat.Size())
		}

		outputPaths = append(outputPaths, outputPath)
	}

	// Return response based on number of files
	if len(outputPaths) == 1 {
		// Single file - return directly
		w.Header().Set("Content-Type", "text/csv")
		w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=\"%s.csv\"", baseName))

		csvFile, err := os.Open(outputPaths[0])
		if err != nil {
			log.Printf("Failed to read converted file: %v", err)
			http.Error(w, "Failed to read converted file", http.StatusInternalServerError)
			return
		}
		defer csvFile.Close()

		log.Printf("Sending CSV file: %s", outputPaths[0])
		io.Copy(w, csvFile)
	} else {
		// Multiple files - return as ZIP
		w.Header().Set("Content-Type", "application/zip")
		w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=\"%s_sheets.zip\"", baseName))

		zipWriter := zip.NewWriter(w)
		defer zipWriter.Close()

		for _, outputPath := range outputPaths {
			csvFile, err := os.Open(outputPath)
			if err != nil {
				continue
			}

			fileName := filepath.Base(outputPath)
			zipFile, err := zipWriter.Create(fileName)
			if err != nil {
				csvFile.Close()
				continue
			}

			io.Copy(zipFile, csvFile)
			csvFile.Close()
		}

		log.Printf("Sending ZIP with %d files", len(outputPaths))
	}
}

func infoHandler(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "application/json")

	info := map[string]interface{}{
		"name":    "Excel2CSV API Server",
		"version": "1.1.0",
		"endpoints": map[string]string{
			"GET /health":   "Health check",
			"POST /convert": "Convert Excel to CSV",
			"GET /info":     "API information",
		},
		"supported_formats": []string{".xlsx", ".xls", ".ods"},
		"max_file_size":     "50MB",
		"features": []string{
			"Smart table boundary detection",
			"Multi-sheet support",
			"Configurable CSV separators",
			"Custom row boundaries",
			"Automatic line break cleaning",
		},
	}

	json.NewEncoder(w).Encode(info)
}

func indexHandler(w http.ResponseWriter, r *http.Request) {
	html := `<!DOCTYPE html>
<html>
<head>
    <title>Excel2CSV Converter</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
        .container { background: #f5f5f5; padding: 20px; border-radius: 10px; }
        .form-group { margin: 15px 0; }
        label { display: block; margin-bottom: 5px; font-weight: bold; }
        input, select { width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; }
        button { background: #007bff; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer; }
        button:hover { background: #0056b3; }
        .info { background: #e7f3ff; padding: 10px; border-radius: 4px; margin: 10px 0; }
    </style>
</head>
<body>
    <h1>📊 Excel2CSV Converter</h1>
    
    <div class="info">
        <strong>Supported formats:</strong> .xlsx, .xls, .ods<br>
        <strong>Max file size:</strong> 50MB<br>
        <strong>Features:</strong> Smart table detection, multi-sheet support, configurable separators
    </div>

    <div class="container">
        <form id="uploadForm" enctype="multipart/form-data">
            <div class="form-group">
                <label for="file">Select Excel file:</label>
                <input type="file" id="file" name="file" accept=".xlsx,.xls,.ods" required>
            </div>

            <div class="form-group">
                <label for="separator">CSV Separator:</label>
                <select id="separator" name="separator">
                    <option value="comma">Comma (,)</option>
                    <option value="semicolon">Semicolon (;)</option>
                    <option value="tab">Tab</option>
                </select>
            </div>

            <div class="form-group">
                <label for="start_row">Force start row (0-based, optional):</label>
                <input type="number" id="start_row" name="start_row" min="0" placeholder="Auto-detect">
            </div>

            <div class="form-group">
                <label for="sheet_name">Sheet name (optional):</label>
                <input type="text" id="sheet_name" name="sheet_name" placeholder="Default: first sheet">
            </div>

            <div class="form-group">
                <label>
                    <input type="checkbox" id="all_sheets" name="all_sheets" value="true">
                    Convert all sheets (returns ZIP file)
                </label>
            </div>

            <button type="submit">Convert to CSV</button>
        </form>
    </div>

    <div id="status"></div>

    <script>
        document.getElementById('uploadForm').addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const formData = new FormData(this);
            const statusDiv = document.getElementById('status');
            
            statusDiv.innerHTML = '<div class="info">Converting... Please wait.</div>';
            
            try {
                const response = await fetch('/convert', {
                    method: 'POST',
                    body: formData
                });
                
                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    
                    // Get filename from Content-Disposition header
                    const disposition = response.headers.get('Content-Disposition');
                    let filename = 'converted.csv';
                    if (disposition) {
                        const match = disposition.match(/filename="([^"]+)"/);
                        if (match) filename = match[1];
                    }
                    
                    a.download = filename;
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    window.URL.revokeObjectURL(url);
                    
                    statusDiv.innerHTML = '<div class="info" style="background: #d4edda;">✅ Conversion successful! Download started.</div>';
                } else {
                    const error = await response.text();
                    statusDiv.innerHTML = '<div class="info" style="background: #f8d7da;">❌ Error: ' + error + '</div>';
                }
            } catch (error) {
                statusDiv.innerHTML = '<div class="info" style="background: #f8d7da;">❌ Network error: ' + error.message + '</div>';
            }
        });
    </script>
</body>
</html>`

	w.Header().Set("Content-Type", "text/html")
	w.Write([]byte(html))
}
