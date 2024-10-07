package excel2csv

import (
	"encoding/csv"
	"errors"
	"os"
	"strings"
)

var Comma = ';'

var (
	xlsFabric  = &XLS{}
	xlsxFabric = &XLSX{}
)

var SupportedTypes = []Excel{
	xlsFabric,
	xlsxFabric,
}

var (
	errRequiredHeaders = errors.New("all requiredHeaders must be in possibleHeaders")
	errMissedHeaders   = errors.New("no required headers found in the file")
	errNotSupported    = errors.New("file type not supported")
	errEmptyBook       = errors.New("no sheets found")
)

type Sheet struct {
	Name      string
	RowsCount int

	filename string
	excel    Excel

	index int

	requiredHeaders []string
	possibleHeaders map[string]string

	headersRow      int
	headers         []string
	matterIndexes   []int // columns indexes that contain any header
	requiredIndexes []int // columns indexes that contain required headers

	outputWriter *csv.Writer
}

func Open(filename string) ([]*Sheet, error) {
	var excel Excel

	for _, t := range SupportedTypes {
		excel = t.MayBeSupported(filename)
		if excel != nil {
			sheetNames := excel.GetSheets()
			if len(sheetNames) == 0 {
				return nil, errEmptyBook
			}

			var sheets []*Sheet

			for i, sheetName := range sheetNames {
				excel.UseSheetByIndex(i)
				sheets = append(sheets, &Sheet{
					Name:      sheetName,
					RowsCount: excel.GetRowsCount(),
					filename:  filename,
					excel:     excel,
					index:     i,
				})
			}

			excel.UseSheetByIndex(0)
			return sheets, nil
		}
	}

	return nil, errNotSupported
}

func (s *Sheet) Convert(dst *os.File, possibleHeaders map[string]string, requiredHeaders []string) error {
	s.excel.UseSheetByIndex(s.index)

	if err := s.parseIncomingHeadersInfo(possibleHeaders, requiredHeaders); err != nil {
		return err
	}

	if err := s.detectFileHeaders(); err != nil {
		return err
	}

	// create output file and write headers
	if err := s.createOutputWriter(dst); err != nil {
		return err
	}
	defer func(c *Sheet) {
		s.outputWriter.Flush()
	}(s)

	// write data
	for i := s.headersRow + 1; i < s.RowsCount; i++ {
		row := s.excel.GetRow(i)
		checkedRow := s.getMatterCells(row)
		if checkedRow != nil {
			_ = s.outputWriter.Write(s.getMatterCells(row))
		}
	}

	return nil
}

func (s *Sheet) parseIncomingHeadersInfo(possibleHeaders map[string]string, requiredHeaders []string) error {
	s.possibleHeaders = possibleHeaders
	s.requiredHeaders = requiredHeaders

	for _, possibleHeader := range possibleHeaders {
		for i, requiredHeader := range requiredHeaders {
			if requiredHeader == possibleHeader {
				requiredHeaders = s.remove(requiredHeaders, i)
				break
			}
		}
	}

	if len(requiredHeaders) > 0 {
		return errRequiredHeaders
	} else {
		return nil
	}
}

func (s *Sheet) remove(slice []string, i int) []string {
	return append(slice[:i], slice[i+1:]...)
}

func (s *Sheet) detectFileHeaders() error {
	for i := 0; i < s.RowsCount; i++ {
		row := s.excel.GetRow(i)
		if s.mayBeHeaders(row) {
			s.headersRow = i
			s.headers = row
			for j, cell := range row {
				if strings.Trim(cell, " ") != "" {
					s.matterIndexes = append(s.matterIndexes, j)
				}
			}
			return nil
		}
	}
	return errMissedHeaders
}

func (s *Sheet) mayBeHeaders(row []string) bool {
	requiredHeaders := s.requiredHeaders
	for k, cell := range row {
		for possibleHeader, resolveAs := range s.possibleHeaders {
			if strings.Contains(strings.ToLower(strings.Trim(cell, " ")), strings.ToLower(possibleHeader)) {
				for i, requiredHeader := range requiredHeaders {
					if requiredHeader == resolveAs {
						s.requiredIndexes = append(s.requiredIndexes, k)
						requiredHeaders = s.remove(requiredHeaders, i)
						break
					}
				}
			}
		}
	}

	if len(requiredHeaders) == 0 {
		return true
	} else {
		return false
	}
}

func (s *Sheet) checkRequiredCells(row []string) bool {
	for _, index := range s.requiredIndexes {
		if strings.Trim(row[index], " ") == "" {
			return false
		}
	}
	return true
}

func (s *Sheet) getMatterCells(row []string) []string {
	if !s.checkRequiredCells(row) {
		return nil
	}

	var matterCells []string
	for _, index := range s.matterIndexes {
		matterCells = append(matterCells, row[index])
	}
	return matterCells
}

func (s *Sheet) createOutputWriter(file *os.File) error {
	s.outputWriter = csv.NewWriter(file)
	s.outputWriter.Comma = Comma

	row2flash := s.getMatterCells(s.headers)
	if row2flash != nil {
		return s.outputWriter.Write(row2flash)
	}

	return errMissedHeaders
}
