package extemplate

import (
	"github.com/xuri/excelize/v2"
)

type SheetExport interface {
	FormCollection
}

type SheetConfig struct {
	Name             string
	DefaultColWidth  float64
	DefaultRowHeight float64
	SpecialColWidth  map[string]float64
	SpecialRowHeight map[int]float64
	MergeCell        map[string]string
}

func (a SheetConfig) Collection() [][]string {
	return [][]string{}
}

type WithMultipleSheets interface {
	Sheets() []SheetExport
}

type WithHeading interface {
	Header() []string
}

type WithTitle interface {
	Title() string
}

type WithStyle interface {
	Style(file *excelize.File) error
}

type WithColumnWidth interface {
	ColumnWidth() map[string]float64
}

type FormCollection interface {
	Collection() [][]string
}
