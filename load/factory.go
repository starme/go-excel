package exload

import (
	"errors"
	"mime/multipart"
	"path"
	"strings"
)

type Reader interface {
	Read() (rows [][]string, err error)
	ReadStream() (rows [][]string, err error)
}

const (
	// ExcelExt 文件扩展名（2007之后版本）
	ExcelExt = ".xlsx"
	// ExcelExt2003 文件扩展名(2007之前版本)
	ExcelExt2003 = ".xls"
	// ExcelExtCSV csv文件扩展名
	ExcelExtCSV = ".csv"
)

// NewReaderStream 创建Reader
func NewReaderStream(fh *multipart.FileHeader, sheet string, colCount int) (r Reader, err error) {
	file, err := fh.Open()
	if err != nil {
		return
	}
	switch getFileExt(fh.Filename) {
	case ExcelExt2003:
		r = &excel2003{file: file, sheetName: sheet, colCount: colCount}
	case ExcelExt:
		r = &excel{file: file, sheetName: sheet, colCount: colCount}
	//case ExcelExtCSV:
	//	return &csv{filePath: filePath, sheet: sheet, colCount: colCount}
	default:
		err = errors.New("不支持的文件类型")
	}
	return
}

// NewReader 创建Reader
func NewReader(filePath, sheet string, colCount int) (r Reader, err error) {
	switch getFileExt(filePath) {
	case ExcelExt2003:
		r = &excel2003{filePath: filePath, sheetName: sheet, colCount: colCount}
	case ExcelExt:
		r = &excel{filePath: filePath, sheetName: sheet, colCount: colCount}
	//case ExcelExtCSV:
	//	return &csv{filePath: filePath, sheet: sheet, colCount: colCount}
	default:
		err = errors.New("不支持的文件类型")
	}
	return
}

// getFileExt 获取文件后缀
func getFileExt(fileName string) (ext string) {
	// 获取文件后缀
	ext = path.Ext(fileName)
	return
}

// formatCellValue 格式化单元格值
// 去除换行符，空格
func formatCellValue(v string) string {
	v = strings.ReplaceAll(v, "\r\n", "")
	v = strings.ReplaceAll(v, "\n", "")
	v = strings.TrimSpace(v)
	return v
}
