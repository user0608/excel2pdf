//go:build windows
// +build windows

package excel2pdf

import (
	"context"
	"os/exec"
)

func convertExcelToPdf(ctx context.Context, excelFile, pdfFile string) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	engine, err := findOfficeEngine(findLibreOfficeBinPath, isExcelInstalled)
	if err != nil {
		return err
	}
	if engine.useExcel {
		return convertExcelToPDFWithExcel(ctx, excelFile, pdfFile)
	}
	return runLibreOfficeConversion(ctx, engine.libreOfficePath, excelFile, pdfFile, exec.CommandContext)
}
