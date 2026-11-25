package excel2pdf

import (
	"fmt"
	"log/slog"
	"os"
	"os/exec"
	"path/filepath"
)

func convertExcelToPDFWithLibreOffice(excelFilePath, pdfFilePath string) error {
	pibreOfficePath, err := findLibreOfficeBinPath()
	if err != nil {
		return err
	}
	cmd := exec.Command(
		pibreOfficePath,
		"--headless",
		"--convert-to",
		"pdf", excelFilePath,
		"--outdir", filepath.Dir(pdfFilePath),
	)

	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr
	if err := cmd.Run(); err != nil {
		slog.Error("libreoffice runing", "error", err, "libre_ofice_path", err)
		return fmt.Errorf("failed to convert file: %w", err)
	}
	if cmd.Err != nil {
		slog.Error("libreoffice comand", "error", err, "libre_ofice_path", err)
	}
	return nil
}
