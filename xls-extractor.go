package xlsx

import (
	"github.com/extrame/xls"
	"fmt"
	"io"
)

func XlsRows(book io.ReadSeeker, charset, sheet string, titles []string, dateFieldsList ...[]string) (<-chan []string, error) {
	_, c, err := XlsRowsWithTitles(book, charset, sheet, titles, dateFieldsList...)
	return c, err
}

func XlsRowsWithTitles(book io.ReadSeeker, charset, sheet string, titles []string, dateFieldsList ...[]string) ([]string, <-chan []string, error) {
	dateFields := map[string]bool{}
	if len(dateFieldsList) > 0 && len(dateFieldsList[0]) > 0 {
		dateFields = parseDateFields(dateFieldsList[0])
	}

	f, err := xls.OpenReader(book, charset)
	if err != nil {
		return nil, nil, err
	}

	sheetNums := f.NumSheets()
	var rows *xls.WorkSheet
	var i int
	for i=0; i<sheetNums; i++ {
		rows = f.GetSheet(i)
		if rows.Name == sheet {
			break
		}
	}
	if rows == nil || i >= sheetNums {
		return nil, nil, fmt.Errorf("sheet %s not found", sheet)
	}

	var headers []string
	colsChan := xlsColumns(rows)
	if colsChan == nil {
		return nil, nil, fmt.Errorf("no data in sheet %s", sheet)
	}
	for cols := range colsChan {
		// 1st row is fields names
		headers = cols
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
		for columns := range colsChan {
			if cols := adjustCol(columns, colCount, dateCols); len(cols) == 0 {
				continue
			} else {
				row := make([]string, len(headerIdx))
				for i, idx := range headerIdx {
					row[i] = cols[idx]
				}
				c <- row
			}
		}
		close(c)
	}()

	return outHeaders, c, nil
}

func xlsColumns(rows *xls.WorkSheet) (<-chan []string) {
	rowsNum := int(rows.MaxRow)
	if rowsNum <= 0 {
		return nil
	}

	cols := make(chan []string)
	go func() {
		for i:=0; i<=rowsNum; i++ {
			row := rows.Row(i)
			colCount := row.LastCol()
			if colCount <= 0 {
				continue
			}
			cs := make([]string, colCount)
			for j:=0; j<colCount; j++ {
				cs[j] = row.Col(j)
			}
			cols <- cs
		}
		close(cols)
	}()

	return cols
}
