//go:build linux || darwin
// +build linux darwin

package excel2pdf

func convertExcelToPdf(excelFile, pdfFile string) error {
	return convertExcelToPDFWithLibreOffice(excelFile, pdfFile)
}
