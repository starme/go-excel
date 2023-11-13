package extemplate

import (
	"github.com/starme/go-excel/style"
	"github.com/xuri/excelize/v2"
)

type HandleSheetStyle func(f *excelize.File, s *exstyle.Style) error

type Excel struct {
	Name string

	DefaultColWidth  float64 // 默认列宽
	DefaultRowHeight float64 // 默认行高

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

type Sheet struct {
	index int    // 工作表索引
	Name  string // 工作表名称

	DefaultColWidth  float64 // 默认列宽
	DefaultRowHeight float64 // 默认行高
	IsCustomHigh     bool    // 是否自定义行高

	MergeCell map[string]string // 需要合并的单元格(如：A1:D3)

	headers    []string           // 表头
	rows       [][]string         // 行数据
	customWith map[string]float64 // 自定义列宽

	styleHandle HandleSheetStyle // 工作表样式
}

// withTitle 工作表标题
type withTitle interface {
	Title() string
}

// withStyle 工作表样式
type withStyle interface {
	Style() HandleSheetStyle
}

// withColumnWidth 自定义列宽
type withColumnWidth interface {
	ColumnWidth() map[string]float64
}

type withMergeCell interface {
	MergeCell() map[string]string
}

// withHeading 表头
type withHeading interface {
	Header() []string
}

// formCollection 表数据
type formCollection interface {
	Collection() [][]string
}
