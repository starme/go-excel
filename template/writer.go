package extemplate

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"os"
	"path"
	"time"
)

const (
	// DefaultSheetName 默认工作表名
	DefaultSheetName = "Sheet1"
	// DefaultColWidth 默认列宽
	DefaultColWidth = 25.00
	// DefaultRowHeight 默认行高
	DefaultRowHeight = 20.00
	// DefaultFontFamily 默认字体
	DefaultFontFamily = "宋体"
	// DefaultFontSize 默认字号
	DefaultFontSize = 20.00
	// DefaultHorizontalAlign 默认水平对齐方式
	DefaultHorizontalAlign = "center"
	// DefaultVerticalAlign 默认垂直对齐方式
	DefaultVerticalAlign = "center"
)

const (
	// ExcelExt 文件扩展名（2007之后版本）
	ExcelExt = ".xlsx"
	// ExcelExt2003 文件扩展名(2007之前版本)
	ExcelExt2003 = ".xls"
	// ExcelExtCSV csv文件扩展名
	ExcelExtCSV = ".csv"
)

var (
	// DefaultExcelFileName 导出的默认的Excel文件名
	DefaultExcelFileName = time.Now().Format("20060102150405.xlsx")
)

func (e *Excel) Export() (f *excelize.File, err error) {
	f = excelize.NewFile()
	for _, sheet := range e.Sheets {
		var sheetName = DefaultSheetName
		if withTitle, ok := sheet.(WithTitle); ok {
			sheetName = withTitle.Title()
			if err = f.SetSheetName(DefaultSheetName, sheetName); err != nil {
				return
			}
		}

		// 设置默认列宽
		if e.DefaultColWidth == 0 {
			e.DefaultColWidth = DefaultColWidth
		}
		// 设置默认行高
		if e.DefaultRowHeight == 0 {
			e.DefaultRowHeight = DefaultRowHeight
		}

		if heading, ok := sheet.(WithHeading); ok {
			// 获取有效的最后一列的列名
			// 获取列对应的列名
			tableHeader := heading.Header()
			//_, err = excelize.ColumnNumberToName(len(tableHeader))
			//if err != nil {
			//	err = errors.New(fmt.Sprintf("获取第%d列对应的列名失败：%s", len(tableHeader), err.Error()))
			//	return
			//}

			// 设置表头
			if err = f.SetSheetRow(sheetName, "A1", &tableHeader); err != nil {
				err = errors.New(fmt.Sprintf("设置表头失败：%s", err.Error()))
				return
			}
		}

		var ok = true
		err = f.SetSheetProps(sheetName, &excelize.SheetPropsOptions{
			// 设置默认列宽
			DefaultColWidth: &e.DefaultColWidth,
			// 设置默认行高
			DefaultRowHeight: &e.DefaultRowHeight,
			CustomHeight:     &ok, // 是否自定义行高
		})
		if err != nil {
			return
		}

		if withColumn, ok := sheet.(WithColumnWidth); ok {
			for c, w := range withColumn.ColumnWidth() {
				err = f.SetColWidth(sheetName, c, c, w)
				if err != nil {
					return
				}
			}
		}

		if withStyle, ok := sheet.(WithStyle); ok {
			// 设置格式
			if err = withStyle.Style(f); err != nil {
				return
			}
		}

		if collection, ok := sheet.(FormCollection); ok {
			dimension, _ := f.GetRows(sheetName)
			for i, item := range collection.Collection() {
				// 获取行号
				row := i + 1 + len(dimension)
				fmt.Println(item)
				if err = f.SetSheetRow(sheetName, fmt.Sprintf("A%d", row), &item); err != nil {
					return
				}
			}
		}

	}
	e.f = f
	return
}

func (e *Excel) ExportFile(filePath string) (err error) {
	if e.f == nil {
		err = errors.New("excelize.File 为 nil")
		return
	}
	// 判断保存路径对应的文件夹是否存在
	ok, _ := PathExists(filePath)
	if !ok {
		// 创建多次文件夹
		err = os.MkdirAll(filePath, os.ModePerm)
		if err != nil {
			err = errors.New("创建文件夹失败：" + err.Error())
			return err
		}
	}
	// 根据指定路径保存文件
	if e.FileName == "" {
		e.FileName = DefaultExcelFileName
	} else {
		// 扩展名
		ext := path.Ext(e.FileName)
		if ext == "" {
			e.FileName += ExcelExt
		} else if ext != ExcelExt {
			err = errors.New("错误的文件扩展名：" + ext)
			return
		}
	}
	// 导出文件
	err = e.f.SaveAs(path.Join(filePath, e.FileName))
	return
}

// PathExists
/**
 *  @Description: 判断路径是否存在
 *  @param path
 *  @return bool
 *  @return error
 */
func PathExists(path string) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	}
	// IsNotExist来判断，是不是不存在的错误
	if os.IsNotExist(err) { //如果返回的错误类型使用os.isNotExist()判断为true，说明文件或者文件夹不存在
		return false, nil
	}
	return false, err //如果有错误了，但是不是不存在的错误，所以把这个错误原封不动的返回
}
