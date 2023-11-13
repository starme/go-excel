package extemplate

import (
	"github.com/starme/go-excel/style"
	"github.com/xuri/excelize/v2"
)

type HandleSheetStyle func(f *excelize.File, s *exstyle.Style) error

type Excel struct {
	Name string

	//DefaultFontSize   float64 // 默认字体大小
	//DefaultFontFamily string  // 默认字体

	DefaultColWidth  float64 // 默认列宽
	DefaultRowHeight float64 // 默认行高

	//DefaultVerticalAlign   string // 默认垂直对齐方式
	//DefaultHorizontalAlign string // 默认水平对齐方式

	Sheets []interface{} // 工作表

	f *excelize.File
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

//type SheetInterface interface {
//	formCollection
//}

type Sheet struct {
	Index int    // 工作表索引
	Name  string // 工作表名称

	//DefaultFontSize   float64 // 默认字体大小
	//DefaultFontFamily string  // 默认字体

	DefaultColWidth  float64 // 默认列宽
	DefaultRowHeight float64 // 默认行高
	IsCustomHigh     bool    // 是否自定义行高

	//DefaultVerticalAlign   string // 默认垂直对齐方式
	//DefaultHorizontalAlign string // 默认水平对齐方式

	MergeCell map[string]string // 需要合并的单元格(如：A1:D3)

	headers    []string           // 表头
	rows       [][]string         // 行数据
	customWith map[string]float64 // 自定义列宽

	styleHandle HandleSheetStyle // 工作表样式
}

//func (s Sheet) Collection() [][]string {
//	return [][]string{}
//}

type withHeading interface {
	Header() []string
}

type WithTitle interface {
	Title() string
}

type withStyle interface {
	Style() HandleSheetStyle
}

type withColumnWidth interface {
	ColumnWidth() map[string]float64
}

type formCollection interface {
	Collection() [][]string
}
