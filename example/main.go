package main

import (
	"fmt"

	"github.com/user0608/excel2pdf"
)

var excelPath = `file.xlsx`

func main() {
	fmt.Println(excel2pdf.ConvertExcelToPdf(excelPath))
}
