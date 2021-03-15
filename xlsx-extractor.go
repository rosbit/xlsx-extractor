package xlsx

import (
	"github.com/360EntSecGroup-Skylar/excelize"
	"fmt"
	"os"
	"io"
	"time"
	"strconv"
)

var (
	loc = time.FixedZone("UTF+8", 8*3600)
	xlsDateBase = time.Date(1900, 1, 1, 0, 0, 0, 0, loc)
)

const (
	dateLayout = "2006-01-02"
)

func init() {
	if l := os.Getenv("TZ"); len(l) > 0 {
		if loc2, err := time.LoadLocation(l); err == nil {
			loc = loc2
		}
	}
}

func Rows(book io.Reader, sheet string, titles []string, dateFieldsList ...[]string) (<-chan []string, error) {
	_, c, err := RowsWithTitles(book, sheet, titles, dateFieldsList...)
	return c, err
}

func RowsWithTitles(book io.Reader, sheet string, titles []string, dateFieldsList ...[]string) ([]string, <-chan []string, error) {
	dateFields := map[string]bool{}
	if len(dateFieldsList) > 0 && len(dateFieldsList[0]) > 0 {
		dateFields = parseDateFields(dateFieldsList[0])
	}

	f, err := excelize.OpenReader(book)
	if err != nil {
		return nil, nil, err
	}

	rows, err := f.Rows(sheet)
	if err != nil {
		return nil, nil, err
	}

	var headers []string
	for rows.Next() {
		// 1st row is fields names
		headers = rows.Columns()
		if len(headers) > 0 {
			break
		}
	}
	if len(headers) == 0 {
		return nil, nil, fmt.Errorf("no data")
	}

	var dateCols map[int]bool
	var outHeaders []string
	var headerIdx []int
	colCount := 0
	colCount, outHeaders, headerIdx, dateCols, err = createHeaderTitle(headers, titles, dateFields)
	if err != nil {
		return nil, nil, err
	}

	c := make(chan []string)
	go func() {
		for rows.Next() {
			if cols := adjustCol(rows.Columns(), colCount, dateCols); len(cols) == 0 {
				continue
			} else {
				row := make([]string, len(headerIdx))
				for i, idx := range headerIdx {
					if idx >= 0 {
						row[i] = cols[idx]
					}
				}
				c <- row
			}
		}
		close(c)
	}()

	return outHeaders, c, nil
}

func parseDateFields(dateFieldsList []string) map[string]bool {
	dateFields := map[string]bool{}

	if len(dateFieldsList) > 0 {
		for _, field := range dateFieldsList {
			dateFields[field] = true
		}
	}
	return dateFields
}

// @param headers 导入文件中的标题列表
// @param titles  欲导出的标题列表；如果是nil表示获取全部的列名
// @param dateFields 标题 -> 日期属性
//
// @return colCount    列数
// @return outHeaders  导出的标题
// @return headerIdx   outHeaders对应的列序号 (-1表示该列不存在)
// @return dateColss   列序号 -> 是否是日期属性
func createHeaderTitle(headers []string, titles []string, dateFields map[string]bool) (colCount int, outHeaders []string, headerIdx []int, dateCols map[int]bool, err error) {
	outHeaders, headerIdx, err = indexTitleHeaders(headers, titles)
	if err != nil {
		return
	}

	colCount = len(headers)
	dateCols = indexDateHeaders(headers, dateFields)
	return
}

func indexTitleHeaders(headers, titles []string) ([]string, []int, error) {
	headerIndex := map[string]int{}
	for i, header := range headers {
		headerIndex[header] = i
	}

	if len(titles) == 0 {
		outHeaders := make([]int, len(headers))
		for i, _ := range headers {
			outHeaders[i] = i
		}
		return headers, outHeaders, nil
	}

	outHeaders := make([]int, len(titles))
	for i, title := range titles {
		idx, ok := headerIndex[title]
		if ok {
			outHeaders[i] = idx
			// continue
		} else {
			outHeaders[i] = -1
		}
		// return nil, nil, fmt.Errorf("title %s not found", title)
	}
	return titles, outHeaders, nil
}

func indexDateHeaders(headers []string, dateFields map[string]bool) map[int]bool {
	dateCols := map[int]bool{}
	if len(dateFields) > 0 {
		for i, header := range headers {
			if _, ok := dateFields[header]; ok {
				dateCols[i] = true
			}
		}
	}
	return dateCols
}

func adjustCol(cols []string, colCount int, dateCols map[int]bool) []string {
	aCols := adjustCols(cols, colCount)
	return ajustDate(aCols, dateCols)
}

func adjustCols(cols []string, colCount int) []string {
	var aCols []string
	c := len(cols)
	switch {
	case c == 0:
		return nil
	case c == colCount:
		if isBlankRow(cols, c) {
			return nil
		}
		aCols = cols
	case c > colCount:
		if isBlankRow(cols, colCount) {
			return nil
		}
		aCols = cols[:colCount]
	default:
		if isBlankRow(cols, c) {
			return nil
		}
		aCols = make([]string, colCount)
		for i, cc := range cols {
			aCols[i] = cc
		}
	}

	return aCols
}

func isBlankRow(cols []string, colCount int) bool {
	for i:=0; i<colCount; i++ {
		if len(cols[i]) > 0 {
			return false
		}
	}
	return true
}

func ajustDate(cols []string, dateCols map[int]bool) []string {
	if len(dateCols) == 0 {
		return cols
	}
	for i, _ := range cols {
		if _, ok := dateCols[i]; !ok {
			continue
		}
		col := cols[i]
		if len(col) == 0 {
			continue
		}
		offset, err := strconv.Atoi(col)
		if err != nil {
			fmt.Fprintf(os.Stderr, "failed to convert col #%d value: %s to integer: %v\n", i, col, err)
			continue
		}
		d := xlsDateBase.Add(time.Duration(offset-2)*time.Hour*24)
		cols[i] = d.Format(dateLayout)
	}
	return cols
}

