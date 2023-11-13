package extemplate

import (
	"errors"
	"fmt"
	"github.com/starme/go-excel/style"
	"github.com/xuri/excelize/v2"
	"os"
	"path"
	"strconv"
	"strings"
	"time"
)

const (
	// DefaultSheetName 默认工作表名
	DefaultSheetName = "Sheet1"
	// DefaultColWidth 默认列宽
	DefaultColWidth = 25.00
	// DefaultRowHeight 默认行高
	DefaultRowHeight = 20.00
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

// Export Excel导出
func (e *Excel) Export() error {
	e.f = excelize.NewFile()
	for i, sheet := range e.Sheets {
		s, err := e.newSheet(sheet, i+1)
		if err != nil {
			return err
		}
		// 设置默认列宽
		if err = e.setSheetColumnWidth(s); err != nil {
			return err
		}
		// 设置样式
		if err = e.setSheetStyle(s); err != nil {
			return err
		}
		// 设置表头
		if err = e.setSheetHeader(s); err != nil {
			return err
		}
		// 设置数据
		if err = e.setSheetData(s); err != nil {
			return err
		}
	}
	return nil
}

// newSheet 创建工作表
func (e *Excel) newSheet(s interface{}, i int) (s1 Sheet, err error) {
	s1 = e.getDefaultSheet(s, i)

	if _, ok := s.(withTitle); ok {
		s1.Name = s.(withTitle).Title()
	}

	if _, ok := s.(withHeading); ok {
		s1.headers = s.(withHeading).Header()
	}

	if _, ok := s.(withColumnWidth); ok {
		s1.customWith = s.(withColumnWidth).ColumnWidth()
	}

	if _, ok := s.(withStyle); ok {
		s1.styleHandle = s.(withStyle).Style()
	}

	if _, ok := s.(formCollection); ok {
		s1.rows = s.(formCollection).Collection()
	}

	if err = e.f.SetSheetName(DefaultSheetName, s1.Name); err != nil {
		err = errors.New(fmt.Sprintf("设置工作表失败：%s", err.Error()))
		return
	}

	var spo = excelize.SheetPropsOptions{
		DefaultRowHeight: &s1.DefaultRowHeight,
		CustomHeight:     &s1.IsCustomHigh,
	}

	if s1.customWith == nil {
		spo.DefaultColWidth = &s1.DefaultColWidth
	}

	if err = e.f.SetSheetProps(s1.Name, &spo); err != nil {
		return
	}

	return
}

// getDefaultSheet 获取默认的工作表
func (e *Excel) getDefaultSheet(s interface{}, i int) Sheet {
	var s1 Sheet
	if _, ok := s.(Sheet); ok {
		s1 = s.(Sheet)
	}

	if s1.Name == "" {
		s1.Name = strings.ReplaceAll(DefaultSheetName, "1", strconv.Itoa(i))
	}

	s1.DefaultColWidth = DefaultColWidth
	if e.DefaultColWidth != 0 {
		s1.DefaultColWidth = e.DefaultColWidth
	}

	s1.DefaultRowHeight = DefaultRowHeight
	if e.DefaultRowHeight != 0 {
		s1.DefaultRowHeight = e.DefaultRowHeight
	}
	s1.IsCustomHigh = false
	return s1
}

// setSheetColumnWidth 设置列宽
func (e *Excel) setSheetColumnWidth(s Sheet) error {
	if s.customWith == nil {
		return nil
	}

	for col, w := range s.customWith {
		if err := e.f.SetColWidth(s.Name, col, col, w); err != nil {
			return errors.New(fmt.Sprintf("设置列宽失败：%s", err.Error()))
		}
	}
	return nil
}

// setSheetStyle 设置样式
func (e *Excel) setSheetStyle(s Sheet) error {
	if s.styleHandle == nil {
		return nil
	}
	return s.styleHandle(e.f, &exstyle.Style{})
}

// setSheetHeader 设置表头
func (e *Excel) setSheetHeader(s Sheet) error {
	if len(s.headers) == 0 {
		return nil
	}

	if err := e.f.SetSheetRow(s.Name, "A1", &s.headers); err != nil {
		return errors.New(fmt.Sprintf("设置表头失败：%s", err.Error()))
	}
	return nil
}

// setSheetRow 设置行数据
func (e *Excel) setSheetData(s Sheet) error {
	if len(s.rows) == 0 {
		return nil
	}

	start := 0
	if s.headers != nil {
		start = 2
	}

	for i, row := range s.rows {
		rowI := fmt.Sprintf("%s%d", "A", i+start)
		if err := e.f.SetSheetRow(s.Name, rowI, &row); err != nil {
			return errors.New(fmt.Sprintf("设置行数据失败：%s", err.Error()))
		}
	}
	return nil
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
	if e.Name == "" {
		e.Name = DefaultExcelFileName
	} else {
		// 扩展名
		ext := path.Ext(e.Name)
		if ext == "" {
			e.Name += ExcelExt
		} else if ext != ExcelExt {
			err = errors.New("错误的文件扩展名：" + ext)
			return
		}
	}
	// 导出文件
	err = e.f.SaveAs(path.Join(filePath, e.Name))
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
