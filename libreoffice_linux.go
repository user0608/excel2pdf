package excel2pdf

import (
	"context"
	"os/exec"
)

func findLibreOfficeBinPath() (string, error) {
	if path, configured, err := libreOfficeOverride(); configured {
		return path, err
	}
	return findLibreOfficeOnPath("libreoffice", "libreoffice24.8", "soffice")
}

func convertWithLibreOffice(ctx context.Context, inputPath, pdfFilePath string) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	libreOfficePath, err := findLibreOfficeBinPath()
	if err != nil {
		return err
	}
	return runLibreOfficeConversion(ctx, libreOfficePath, inputPath, pdfFilePath, exec.CommandContext)
}
