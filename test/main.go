package main

import (
	"fmt"
	extemplate "github.com/starme/go-excel/template"
	"github.com/xuri/excelize/v2"
	"time"
)

var (
	// DefaultExcelFileName 导出的默认的Excel文件名
	DefaultExcelFileName = time.Now().Format("20060102150405.xlsx")
)

type Person struct {
	Name     string    `json:"name" ex:"column:姓名"`
	Age      int       `json:"age" ex:"column:年龄;len:20"`
	BirthDay time.Time `json:"birth_day" ex:"column:生日;format:01-02-06"`
	//Alise    string    `json:"alise" ex:"column:别名;required;unique"`
}

type NameList Person

type User struct {
	Id       string `json:"id" ex:"column:ID"`
	Name     string `json:"name" ex:"column:姓名"`
	NickName string `json:"nick_name" ex:"column:昵称"`
	Phone    string `json:"phone" ex:"column:手机号"`
	Age      int    `json:"age" ex:"column:年龄"`
	Sex      string `json:"sex" ex:"column:性别"`
}

type Sheet1 extemplate.Sheet

func (s Sheet1) Header() []string {
	return []string{"ID", "姓名", "昵称", "手机号", "年龄", "性别"}
}

func (s Sheet1) Title() string {
	return "车辆导入模板"
}

func (s Sheet1) ColumnWidth() map[string]float64 {
	return map[string]float64{
		"A": 10,
		"B": 20,
		"C": 30,
		"D": 40,
		"E": 50,
		"F": 60,
	}
}

func (s Sheet1) Style() extemplate.HandleSheetStyle {
	return func(f *excelize.File) error {
		style, err := f.NewStyle(&excelize.Style{
			Alignment: &excelize.Alignment{
				Horizontal: "center",
			},
			Font: &excelize.Font{
				Bold: true,
			},
		})
		if err != nil {
			return err
		}
		err = f.SetCellStyle(s.Name, "A1", "F1", style)
		if err != nil {
			return err
		}
		return nil
	}
}

func (s Sheet1) Collection() [][]string {
	return [][]string{
		{
			"001",
			"张三",
			"别人家的孩子",
			"123456789",
			"30",
			"男",
		},
		{
			"002",
			"张四",
			"淘气的孩子",
			"987654321",
			"31",
			"男",
		},
	}
}

func main() {
	//rows, err := extemplate.LoadExcelByStruct[NameList]("./", "a.xlsx", "Sheet1", NameList{}, 1)
	//if err != nil {
	//	fmt.Println(err)
	//}
	//
	//fmt.Println(rows)

	//创建一个excelConfig（每个Excel文件需要一个）
	exTemp3 := extemplate.Excel{
		Name: "车辆导入模板" + DefaultExcelFileName, // 导出后的文件名
		Sheets: []interface{}{
			Sheet1{
				Name: "车辆导入模板",
				MergeCell: map[string]string{ // 需要合并的单元格
					"A1": "D3",
				},
			},
		},
		DefaultColWidth:  10, // 默认列宽
		DefaultRowHeight: 20, // 默认行高（无效）
	}

	err := exTemp3.Export()
	if err != nil {
		fmt.Printf("导出失败: %s", err.Error())
		return
	}

	if err = exTemp3.ExportFile("./"); err != nil {
		fmt.Printf("导出失败: %s", err.Error())
		return
	}

	//// 获取Excel导入模板
	//file, err := exTemp3.GetTemplateByStruct("车辆导入模板", user)
	//if err != nil {
	//	panic(err)
	//}
}
