package excel2csv

type Excel interface {
	MayBeSupported(string) Excel
	GetSheets() []string
	UseSheetByIndex(int)
	GetRowsCount() int
	GetRow(int) []string
	// TODO: add more methods
}
