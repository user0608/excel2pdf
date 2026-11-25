//go:build linux
// +build linux

package excel2pdf

import "context"

func convertToPDF(ctx context.Context, inputFile, pdfFile string) error {
	return convertWithLibreOffice(ctx, inputFile, pdfFile)
}
