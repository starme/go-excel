package exload

type csv struct {
	filePath string
	sheet    string
	colCount int
}

func csvReader(filePath, sheet string, colCount int) (rows [][]string, err error) {
	// 读取CSV文件
	//var rowsGet [][]string
	//rowsGet, err = csv.Reader{}(filePath)
	//if err != nil {
	//	return
	//}
	//h := len(rowsGet)
	//// 过滤空格，回车
	//for i := 0; i < h; i++ {
	//	rowData := make([]string, colCount, colCount)
	//	for j := 0; j < colCount; j++ {
	//		v := rowsGet[i][j]
	//		v = strings.ReplaceAll(v, "\r\n", "")
	//		v = strings.ReplaceAll(v, "\n", "")
	//		v = strings.TrimSpace(v)
	//		rowData[j] = v
	//	}
	//	rows = append(rows, rowData)
	//}
	return nil, nil
}
