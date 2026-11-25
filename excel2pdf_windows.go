//go:build windows
// +build windows

package excel2pdf

import (
	"context"
	"os/exec"
)

func convertToPDF(ctx context.Context, inputFile, pdfFile string) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	kind, err := documentKindForPath(inputFile)
	if err != nil {
		return err
	}
	findMicrosoftOffice := isExcelInstalled
	convertMicrosoftOffice := convertExcelToPDFWithExcel
	if kind == wordDocument {
		findMicrosoftOffice = isWordInstalled
		convertMicrosoftOffice = convertDocumentToPDFWithWord
	}
	engine, err := findOfficeEngine(findLibreOfficeBinPath, findMicrosoftOffice)
	if err != nil {
		return err
	}
	if engine.useMicrosoftOffice {
		return convertMicrosoftOffice(ctx, inputFile, pdfFile)
	}
	return runLibreOfficeConversion(ctx, engine.libreOfficePath, inputFile, pdfFile, exec.CommandContext)
}
