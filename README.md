# excel2pdf
Converts an Excel file to PDF, compatible with both Windows and Linux. Uses either LibreOffice or Microsoft Excel, with LibreOffice given priority if installed.


```go
var excelPath = `file.xlsx`

func main() {
    pdfFilePath,err := excel2pdf.ConvertExcelToPdf(excelPath)
	fmt.Println(pdfFilePath, err)
}
```