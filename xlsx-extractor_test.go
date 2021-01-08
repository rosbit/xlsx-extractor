package xlsx

import (
	"testing"
	"os"
	"fmt"
	"strings"
)

func TestXlsx(t *testing.T) {
	if len(os.Args) < 5 {
		fmt.Printf("Usage: go test -args <book> <sheet> <fields-list>[ <date-fields-list>]\n")
		return
	}

	book, sheet, fieldsList := os.Args[2], os.Args[3], os.Args[4]
	fields := strings.FieldsFunc(fieldsList, func(c rune)bool{
		return c==' ' || c == ',' || c == ';'
	})
	var dateFieldsList []string
	if len(os.Args) >= 6 {
		dateFieldsList = strings.FieldsFunc(os.Args[5], func(c rune)bool{
			return c==' ' || c == ',' || c == ';'
		})
	}

	fp, err := os.Open(book)
	if err != nil {
		fmt.Printf("%v\n", err)
		return
	}
	defer fp.Close()

	c, err := Rows(fp, sheet, fields, dateFieldsList)
	if err != nil {
		fmt.Printf("%v\n", err)
		return
	}

	for row := range c {
		fmt.Printf("%v\n", row)
	}
}
