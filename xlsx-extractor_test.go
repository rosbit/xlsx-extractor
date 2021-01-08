package xlsx

import (
	"testing"
	"os"
	"fmt"
	"strings"
)

func TestXlsx(t *testing.T) {
	testXlxs(t, func(fp *os.File, sheet string, fields, dateFieldsList []string) {
		rows, err := Rows(fp, sheet, fields, dateFieldsList)
		if err != nil {
			fmt.Printf("%v\n", err)
			t.Fatalf("failed to call Rows()\n")
		}

		for row := range rows {
			fmt.Printf("%v\n", row)
		}
	})
}

func TestXlsxWithTitles(t *testing.T) {
	testXlxs(t, func(fp *os.File, sheet string, fields, dateFieldsList []string) {
		titles, rows, err := RowsWithTitles(fp, sheet, fields, dateFieldsList)
		if err != nil {
			fmt.Printf("%v\n", err)
			t.Fatalf("failed to call RowsWithTitles\n")
		}

		fmt.Printf("titles: %v\n", titles)
		for row := range rows {
			fmt.Printf("%v\n", row)
		}
	})
}

func testXlxs(t *testing.T, callback func(fp *os.File, sheet string, fields, dateFieldsList []string)) {
	book, sheet, fields, dateFieldsList, err := getOptions()
	if err != nil {
		fmt.Printf("%v\n", err)
		t.Fatalf("args expected\n")
	}

	fp, err := os.Open(book)
	if err != nil {
		fmt.Printf("%v\n", err)
		t.Fatalf("failed to open book\n")
	}
	defer fp.Close()

	callback(fp, sheet, fields, dateFieldsList)
}

func getOptions() (book, sheet string, fields, dateFieldsList []string, err error) {
	if len(os.Args) < 5 {
		err = fmt.Errorf("Usage: go test -args <book> <sheet> <fields-list>[ <date-fields-list>]")
		return
	}

	var fieldsList string
	book, sheet, fieldsList = os.Args[2], os.Args[3], os.Args[4]

	splitFunc := func(c rune)bool {
		return c==' ' || c == ',' || c == ';'
	}
	fields = strings.FieldsFunc(fieldsList, splitFunc)

	if len(os.Args) >= 6 {
		dateFieldsList = strings.FieldsFunc(os.Args[5], splitFunc)
	}
	return
}

