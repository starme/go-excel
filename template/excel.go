package extemplate

import "github.com/xuri/excelize/v2"

type Excel struct {
	FileName string
	f        *excelize.File
	Sheets   []SheetExport

	DefaultColWidth  float64 // 默认列宽
	DefaultRowHeight float64 // 默认行高
}

type ExcelTag struct {
	Alias  string
	Column string
	Type   string
	Format string
	Width  float64

	Required bool
	Unique   bool
	Regexp   string
}
