package xlsx

import (
	"encoding/csv"
	"fmt"
	"io"
	"strings"
)

const (
	utf8_bom = "\xef\xbb\xbf"
)

func CsvRows(book io.Reader, titles []string) (<-chan []string, error) {
	_, c, err := CsvRowsWithTitles(book, titles)
	return c, err
}

func CsvRowsWithTitles(book io.Reader, titles []string) ([]string, <-chan []string, error) {
	f := csv.NewReader(book)

	var headers []string
	colsChan := csvColumns(f)
	if colsChan == nil {
		return nil, nil, fmt.Errorf("no data in file")
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

	if strings.Index(headers[0], utf8_bom) == 0 {
		headers[0] = headers[0][len(utf8_bom):]
	}

	var outHeaders []string
	var headerIdx []int
	var err error
	colCount := 0
	colCount, outHeaders, headerIdx, err = createCsvHeaderTitle(headers, titles)
	if err != nil {
		return nil, nil, err
	}

	c := make(chan []string)
	go func() {
		for columns := range colsChan {
			if cols := adjustCols(columns, colCount); len(cols) == 0 {
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

func csvColumns(rows *csv.Reader) (<-chan []string) {
	cols := make(chan []string)
	go func() {
		for {
			cs, err := rows.Read()
			if err != nil {
				break
			}
			cols <- cs
		}
		close(cols)
	}()

	return cols
}

func createCsvHeaderTitle(headers []string, titles []string) (colCount int, outHeaders []string, headerIdx []int, err error) {
	outHeaders, headerIdx, err = indexTitleHeaders(headers, titles)
	if err != nil {
		return
	}

	colCount = len(headers)
	return
}

