package excel2csv

import (
	"fmt"

	"github.com/oxyii/xls"
)

type XLS struct {
	Excel // implement Excel interface

	file  *xls.XLS
	sheet *xls.Sheet
}

func (x *XLS) MayBeSupported(filename string) Excel {
	if xlFile, err := xls.Open(filename); err != nil {
		return nil
	} else {
		return &XLS{file: xlFile}
	}
}

func (x *XLS) GetSheets() []string {
	var ret []string
	sheets := x.file.Sheets()
	for i := 0; i < len(sheets); i++ {
		ret = append(ret, sheets[i].Name())
	}
	return ret
}

func (x *XLS) UseSheetByIndex(index int) {
	x.sheet = x.file.Sheets()[index]
}

func (x *XLS) GetRowsCount() int {
	return int(x.sheet.Rows()) // 0-based index
}

func (x *XLS) GetRow(rowIndex int) []string {
	row := x.sheet.Row(rowIndex)
	cells := make([]string, row.Cols())
	for i := 0; i < row.Cols(); i++ {
		cells[i] = fmt.Sprint(row.Cell(i).Value())
	}
	return cells
}
