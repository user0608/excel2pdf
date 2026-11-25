//go:build linux || darwin
// +build linux darwin

package excel2pdf

import "context"

func convertExcelToPdf(ctx context.Context, excelFile, pdfFile string) error {
	return convertExcelToPDFWithLibreOffice(ctx, excelFile, pdfFile)
}
