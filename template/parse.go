package extemplate

import (
	"errors"
	"reflect"
	"strings"
)

// parse
/**
 *  @Description: 解析数据
 *  @param data
 *  @param index 递归索引
 *  @return tableHead
 *  @return exTagMap
 *  @return err
 */
func parseTag(data interface{}, index int) (tableHead, tableAlias []string, exTagMap map[string]ExcelTag, err error) {
	// 获取结构体实例的反射类型对象
	typeOf := reflect.TypeOf(data)
	if typeOf.Kind() != reflect.Struct {
		err = errors.New("非结构体")
		return
	}
	exTagMap = make(map[string]ExcelTag)
	//// 获取值
	valueOf := reflect.ValueOf(data)
	// 可以获取所有属性
	// 获取结构体字段个数：t.NumField()
	for i := 0; i < typeOf.NumField(); i++ {
		// 取每个字段
		f := typeOf.Field(i)
		// 获取字段的值信息
		switch valueOf.Field(i).Kind() {
		case reflect.Ptr:
			continue
		case reflect.String, reflect.Slice, reflect.Array:
			fallthrough
		case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
			fallthrough
		case reflect.Float32, reflect.Float64:
			fallthrough
		case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
			// 获取Tag
			tableHead, tableAlias, exTagMap = makeTag(f, tableHead, tableAlias, exTagMap)
			index += 1
		default:
			if f.Type.String() == "time.Time" {
				tableHead, tableAlias, exTagMap = makeTag(f, tableHead, tableAlias, exTagMap)
				index += 1
				continue
			}
			// Interface()：获取字段对应的值
			val := valueOf.Field(i).Interface()
			tableHead0, tableAlias0, exTagMap0, errDG := parseTag(val, index)
			if errDG != nil {
				return
			}
			for k, v := range exTagMap0 {
				exTagMap[k] = v
			}
			tableHead = append(tableHead, tableHead0...)
			tableAlias = append(tableAlias, tableAlias0...)
		}
	}
	return
}

func makeTag(f reflect.StructField, tableHead, tableAlias []string, exTagMap map[string]ExcelTag) ([]string, []string, map[string]ExcelTag) {
	exTag := f.Tag.Get("ex")
	if exTag == "" || exTag == "-" {
		return tableHead, tableAlias, exTagMap
	}
	var excelTag ExcelTag
	excelTag.Alias = f.Tag.Get("json")
	excelTag.Type = f.Type.String()

	exs := strings.Split(exTag, ";")
	for _, ex := range exs {
		if strings.Contains(ex, ":") {
			kv := strings.Split(ex, ":")
			if len(kv) >= 2 {
				k, v := kv[0], kv[1]
				switch k {
				case "alias":
					excelTag.Alias = v
				case "column":
					excelTag.Column = v
				case "format":
					excelTag.Format = v
				case "re":
					excelTag.Regexp = v
				}
			}
		} else if strings.Contains(ex, "required") {
			excelTag.Required = true
		} else if strings.Contains(ex, "unique") {
			excelTag.Unique = true
		}
	}
	if excelTag.Alias == "" {
		excelTag.Alias = snakeStr(f.Name)
	}
	if excelTag.Column == "" {
		excelTag.Column = snakeStr(f.Name)
	}

	tableHead = append(tableHead, excelTag.Column)
	tableAlias = append(tableAlias, excelTag.Alias)
	exTagMap[excelTag.Alias] = excelTag
	return tableHead, tableAlias, exTagMap
}

func snakeStr(s string) string {
	data := make([]byte, 0, len(s)*2)
	j := false
	num := len(s)
	for i := 0; i < num; i++ {
		d := s[i]
		// or通过ASCII码进行大小写的转化
		// 65-90（A-Z），97-122（a-z）
		//判断如果字母为大写的A-Z就在前面拼接一个_
		if i > 0 && d >= 'A' && d <= 'Z' && j {
			data = append(data, '_')
		}
		if d != '_' {
			j = true
		}
		data = append(data, d)
	}
	//ToLower把大写字母统一转小写
	return strings.ToLower(string(data[:]))
}
