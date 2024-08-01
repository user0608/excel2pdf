package excel2pdf

import (
	"fmt"
	"log/slog"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"time"
)

func convertExcelToPDFWithLibreOffice(excelFilePath string) (pdfFilePath string, err error) {
	pibreOfficePath, err := findLibreOfficeBinPath()
	if err != nil {
		return "", err
	}
	cmd := exec.Command(
		pibreOfficePath,
		"--headless",
		"--convert-to",
		"pdf", excelFilePath,
		"--outdir", os.TempDir(),
	)

	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr
	if err := cmd.Run(); err != nil {
		slog.Error("libreoffice runing", "error", err, "libre_ofice_path", err)
		return "", fmt.Errorf("failed to convert file: %w", err)
	}
	if cmd.Err != nil {
		slog.Error("libreoffice comand", "error", err, "libre_ofice_path", err)
	}
	const pdfSuffix = ".pdf"
	var tmpPdfFilePath = filepath.Join(
		os.TempDir(),
		fmt.Sprintf("%s%s",
			strings.TrimSuffix(
				filepath.Base(excelFilePath),
				filepath.Ext(excelFilePath),
			),
			pdfSuffix,
		),
	)

	pdfFilePath = fmt.Sprintf("%s-%d%s",
		strings.TrimSuffix(tmpPdfFilePath, pdfSuffix),
		time.Now().Unix(),
		pdfSuffix,
	)
	if err := os.Rename(tmpPdfFilePath, pdfFilePath); err != nil {
		slog.Error("renaming pdf file", "error", err, "old_path", tmpPdfFilePath, "new_path", pdfFilePath)
		return tmpPdfFilePath, err
	}
	return pdfFilePath, nil
}
