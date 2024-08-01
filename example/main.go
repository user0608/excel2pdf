package main

import (
	"fmt"

	"github.com/sfperusacdev/excel2pdf"
)

var excelPath = `file.xlsx`

func main() {
	fmt.Println(excel2pdf.ConvertExcelToPdf(excelPath))
}
