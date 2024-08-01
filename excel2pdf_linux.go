//go:build linux || darwin
// +build linux darwin

package excel2pdf

func convertExcelToPdf(excelFile string) (pdfFile string, err error) {
	return convertExcelToPDFWithLibreOffice(excelFile)
}
