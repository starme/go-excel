package exload

import (
	"errors"
	exerrors "github.com/starme/go-excel/errors"
	"github.com/xuri/excelize/v2"
	"mime/multipart"
)

type excel struct {
	file multipart.File

	filePath  string
	sheetName string
	colCount  int
}

func (e excel) ReadStream() (row [][]string, err error) {
	var f *excelize.File
	// 读取excel
	f, err = excelize.OpenReader(e.file)
	if err != nil {
		err = errors.New("读取Excel文件失败：" + err.Error())
		return
	}
	// 关闭文件流
	defer f.Close()
	// 读取Excel值
	var rowsGet [][]string
	rowsGet, err = f.GetRows(e.sheetName)
	if err != nil {
		err = exerrors.ErrSheetNotExist{SheetName: e.sheetName}
		return
	}
	for _, row := range rowsGet {
		if e.colCount == 0 {
			e.colCount = len(row)
		}
		rowData := make([]string, e.colCount, e.colCount)
		for i, cell := range row {
			if i+1 > e.colCount {
				continue
			}
			rowData[i] = formatCellValue(cell)
		}
		rows = append(rows, rowData)
	}
	return
}

func (e excel) Read() (rows [][]string, err error) {
	var f *excelize.File
	// 读取excel文件
	f, err = excelize.OpenFile(e.filePath)
	if err != nil {
		err = errors.New("读取Excel文件失败：" + err.Error())
		return
	}
	// 关闭文件流
	defer f.Close()
	// 读取Excel值
	var rowsGet [][]string
	rowsGet, err = f.GetRows(e.sheetName)
	if err != nil {
		err = exerrors.ErrSheetNotExist{SheetName: e.sheetName}
		return
	}
	for _, row := range rowsGet {
		if e.colCount == 0 {
			e.colCount = len(row)
		}
		rowData := make([]string, e.colCount, e.colCount)
		for i, cell := range row {
			if i+1 > e.colCount {
				continue
			}
			rowData[i] = formatCellValue(cell)
		}
		rows = append(rows, rowData)
	}
	return
}
