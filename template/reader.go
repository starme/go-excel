package extemplate

import (
	"encoding/json"
	"errors"
	"fmt"
	"github.com/starme/go-excel/errors"
	"github.com/starme/go-excel/load"
	"mime/multipart"
	"path"
	"regexp"
	"strconv"
	"strings"
	"time"
)

// LoadHttpExcelByStruct 通过结构体解析上传的excel
func LoadHttpExcelByStruct[T interface{}](fh *multipart.FileHeader, sheetName string, data T, skipRows int) (rows []T, err error) {
	var tableHead, tableAlias []string
	var exTagMap map[string]ExcelTag
	tableHead, tableAlias, exTagMap, err = parseTag(data, 1)
	if err != nil {
		return
	}

	exrows, err := LoadHttpExcel(fh, sheetName, len(tableHead))
	if err != nil {
		return
	}

	rows, err = readerToStruct[T](exrows, tableHead, tableAlias, exTagMap, skipRows)
	return
}

// LoadExcelByStruct 通过结构体解析excel
func LoadExcelByStruct[T interface{}](filePath, fileName, sheetName string, data T, skipRows int) (rows []T, err error) {
	var tableHead, tableAlias []string
	var exTagMap map[string]ExcelTag
	tableHead, tableAlias, exTagMap, err = parseTag(data, 1)
	if err != nil {
		return
	}

	exrows, err := LoadExcel(filePath, fileName, sheetName, len(tableHead))
	if err != nil {
		return
	}

	rows, err = readerToStruct[T](exrows, tableHead, tableAlias, exTagMap, skipRows)
	return
}

// readerToStruct 将excel内容读到结构体中
func readerToStruct[T interface{}](exrows [][]string, tableHead, tableAlias []string, exTagMap map[string]ExcelTag, skipRows int) (rows []T, err error) {
	validated := &exerrors.Validated{}
	uniqueMap := make(map[string]map[string]struct{}, 0)
	for i, row := range exrows {
		// 表头校验
		if i == 0 {
			headStr0 := strings.Join(tableHead, ",")
			headStr1 := strings.Join(row, ",")
			if headStr0 != headStr1 {
				err = errors.New(fmt.Sprintf("表头不匹配，期望值：%s，实际值：%s", headStr0, headStr1))
				return
			}
		}

		// 忽略行
		if i < skipRows {
			fmt.Println(fmt.Sprintf("【忽略】第 %d 行,值：%s", i+1, row))
			continue
		}

		fmt.Println(fmt.Sprintf("第 %d 行,值：%s", i+1, row))
		var mapRow = make(map[string]interface{}, len(row))
		for j, col := range row {
			tag, ok := exTagMap[tableAlias[j]]
			if !ok {
				continue
			}
			// 必填校验
			if col == "" {
				if tag.Required {
					validated.Append(fmt.Sprintf("第 %d 行,第 %d 列,值不能为空", i+1, j+1))
				}
				continue
			}
			// 唯一校验
			if tag.Unique {
				_, exists := uniqueMap[tag.Alias][col]
				if exists {
					validated.Append(fmt.Sprintf("第 %d 行,第 %d 列,值重复", i+1, j+1))
					continue
				} else {
					_, exists = uniqueMap[tag.Alias]
					if !exists {
						uniqueMap[tag.Alias] = map[string]struct{}{}
					}
					uniqueMap[tag.Alias][col] = struct{}{}
				}
			}
			// 正则校验
			if tag.Regexp != "" {
				if !regexp.MustCompile(tag.Regexp).MatchString(col) {
					validated.Append(fmt.Sprintf("第 %d 行,第 %d 列,值格式错误", i+1, j+1))
					continue
				}
			}
			mapRow[tableAlias[j]] = convertRowI(col, tag)
		}

		var t T
		marshal, _ := json.Marshal(mapRow)
		_ = json.Unmarshal(marshal, &t)
		rows = append(rows, t)
	}
	if validated != nil {
		err = validated
		return
	}
	return
}

// convertRowI 将excel中的数据转换为对应的类型
func convertRowI(data string, tag ExcelTag) interface{} {
	if strings.Contains(tag.Type, "time.Time") {
		if tag.Format == "" {
			tag.Format = "2006-01-02 15:04:05"
		}
		t, _ := time.Parse(tag.Format, data)
		return t
	}

	if strings.Contains(tag.Type, "int") {
		i, _ := strconv.Atoi(data)
		return i
	}

	if strings.Contains(tag.Type, "float") {
		f, _ := strconv.ParseFloat(data, 64)
		return f
	}

	if strings.Contains(tag.Type, "[]") {
		return strings.Split(data, ",")
	}

	return data
}

// LoadHttpExcel 读取上传的excel
func LoadHttpExcel(fh *multipart.FileHeader, sheetName string, colCount int) (rows [][]string, err error) {
	reader, err := exload.NewReaderStream(fh, sheetName, colCount)
	if err != nil {
		return
	}

	rows, err = reader.ReadSteam()
	return
}

// LoadExcel 读取指定的excel
func LoadExcel(filePath, fileName, sheetName string, colCount int) (rows [][]string, err error) {
	reader, err := exload.NewReader(path.Join(filePath, fileName), sheetName, colCount)
	if err != nil {
		return
	}

	rows, err = reader.Read()
	return
}
