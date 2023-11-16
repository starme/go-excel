package exload

import (
	"errors"
	"github.com/shakinm/xlsReader/xls"
	"github.com/starme/go-excel/errors"
	"mime/multipart"
)

type excel2003 struct {
	file multipart.File

	filePath  string
	sheetName string
	colCount  int
}

func (e *excel2003) Read() (rows [][]string, err error) {
	var open xls.Workbook
	if e.filePath != "" {
		open, err = xls.OpenFile(e.filePath)
	}

	if e.file != nil {
		open, err = xls.OpenReader(e.file)
	}
	if err != nil {
		err = errors.New("读取Excel文件失败：" + err.Error())
		return
	}
	// 获取第一个工作表
	var sheet *xls.Sheet
	sheet, err = resolveXlsSheet(open.GetSheets(), e.sheetName)
	if err != nil {
		return
	}

	// 遍历xls文件
	for i := 0; i < sheet.GetNumberRows(); i++ {
		xlsRow, _ := sheet.GetRow(i)
		if e.colCount == 0 {
			e.colCount = len(xlsRow.GetCols())
		}
		rowData := make([]string, e.colCount, e.colCount)
		for j := 0; j < e.colCount; j++ {
			col, _ := xlsRow.GetCol(j)
			rowData[j] = formatCellValue(col.GetString())
		}
		rows = append(rows, rowData)
	}
	return
}

func resolveXlsSheet(sheets []xls.Sheet, sheetName string) (*xls.Sheet, error) {
	for _, sheet := range sheets {
		if sheet.GetName() == sheetName {
			return &sheet, nil
		}
	}
	return nil, exerrors.ErrSheetNotExist{SheetName: sheetName}
}
