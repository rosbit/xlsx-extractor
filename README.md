# xlsx/csv列提取器

 1. 根据指定标题列表提取列字段
 1. 标题行前可以有空行
 1. 正确转化日期

## 使用方法

```go

import (
	"github.com/rosbit/xlsx-extractor"
	"os"
	"fmt"
)

func main() {
	// --- reading xlsx book sheet ----
	book, err := os.Open("somebook.xlsx")
	if err != nil {
		// error
		return
	}
	defer book.Close()

	rows, err := xlsx.XlsxRows(book, "Sheet1", []string{"title1", "title2", "title3"})
	if err != nil {
		// error
	}
	for row := range rows {
		fmt.Printf("%#v\n", row)
	}

	// --- reading csv ----
	fpCsv, err := os.Open("somename.csv")
	if err != nil {
		// error
	}
	defer fpCsv.Close()
	lines, err := xlsx.CsvRows(fpCsv, []string{"title1", "title2"})
	if err != nil {
		// error
	}
	for line := range lines {
		fmt.Printf("%#v\n", line)
	}
}

```

