package xlsx

import (
	"testing"
	"os"
	"fmt"
)

func TestXls(t *testing.T) {
	testXls(t, func(fp *os.File, charset, sheet string, fields []string) {
		rows, err := XlsRows(fp, charset, sheet, fields)
		if err != nil {
			fmt.Printf("%v\n", err)
			t.Fatalf("failed to call Rows()\n")
		}

		for row := range rows {
			fmt.Printf("%v\n", row)
		}
	})
}

func testXls(t *testing.T, callback func(fp *os.File, charst, sheet string, fields []string)) {
	book, sheet, fields := "../auto-delivery/获奖名单-2021-01-29.xls", "获奖名单", []string{"中奖用户", "用户信用", "奖项等级"}

	fp, err := os.Open(book)
	if err != nil {
		fmt.Printf("%v\n", err)
		t.Fatalf("failed to open book\n")
	}
	defer fp.Close()

	callback(fp, "GBK", sheet, fields)
}

