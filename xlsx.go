package excel2csv

import (
	"github.com/tealeg/xlsx"
)

type XLSX struct {
	Excel // implement Excel interface

	file  *xlsx.File
	sheet *xlsx.Sheet
}

func (x *XLSX) MayBeSupported(filename string) Excel {
	if wb, err := xlsx.OpenFile(filename); err != nil {
		return nil
	} else {
		return &XLSX{file: wb}
	}
}

func (x *XLSX) GetSheets() []string {
	sheets := make([]string, len(x.file.Sheets))
	for i, sheet := range x.file.Sheets {
		sheets[i] = sheet.Name
	}
	return sheets
}

func (x *XLSX) UseSheetByIndex(index int) {
	x.sheet = x.file.Sheets[index]
}

func (x *XLSX) GetRowsCount() int {
	return x.sheet.MaxRow
}

func (x *XLSX) GetRow(rowIndex int) []string {
	row := x.sheet.Rows[rowIndex]
	cells := make([]string, len(row.Cells))
	for i, cell := range row.Cells {
		cells[i] = cell.String()
	}
	return cells
}
