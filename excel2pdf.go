package excel2pdf

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"sync"
)

var (
	mutex sync.Mutex
)
var ErrExcel2PdfIsProcessing = errors.New("a process is currently running using the resource")

// ConvertExcelToPdf converts a readable, regular Excel file to PDF. The returned
// absolute path is inside a new temporary directory owned by this conversion.
// On success, the caller must remove the PDF and its containing directory when
// no longer needed. On failure, the package removes the temporary directory.
func ConvertExcelToPdf(excelFile string) (pdfFile string, err error) {
	if !mutex.TryLock() {
		return "", ErrExcel2PdfIsProcessing
	}
	defer mutex.Unlock()
	return convertExcelFile(excelFile, convertExcelToPdf)
}

func convertExcelFile(excelFile string, convert func(inputPath, outputPath string) error) (pdfFile string, err error) {
	inputPath, err := absoluteExcelPath(excelFile)
	if err != nil {
		return "", err
	}
	dir, err := os.MkdirTemp("", "excel2pdf-*")
	if err != nil {
		return "", fmt.Errorf("create PDF temporary directory: %w", err)
	}
	keepOutput := false
	defer func(dir string) {
		if !keepOutput {
			if cleanupErr := os.RemoveAll(dir); cleanupErr != nil {
				err = errors.Join(err, fmt.Errorf("remove PDF temporary directory: %w", cleanupErr))
			}
		}
	}(dir)
	dir, err = filepath.Abs(dir)
	if err != nil {
		return "", fmt.Errorf("resolve PDF temporary directory: %w", err)
	}
	name := strings.TrimSuffix(filepath.Base(inputPath), filepath.Ext(inputPath)) + ".pdf"
	outputPath := filepath.Join(dir, name)
	if err := convert(inputPath, outputPath); err != nil {
		return "", fmt.Errorf("convert Excel to PDF: %w", err)
	}
	info, err := os.Lstat(outputPath)
	if err != nil {
		return "", fmt.Errorf("inspect generated PDF: %w", err)
	}
	if !info.Mode().IsRegular() || info.Size() == 0 {
		return "", fmt.Errorf("generated PDF is not a non-empty regular file: %s", outputPath)
	}
	keepOutput = true
	return outputPath, nil
}

func absoluteExcelPath(path string) (string, error) {
	if path == "" {
		return "", errors.New("Excel input path is empty")
	}
	path, err := filepath.Abs(path)
	if err != nil {
		return "", fmt.Errorf("resolve Excel input path: %w", err)
	}
	info, err := os.Stat(path)
	if err != nil {
		return "", fmt.Errorf("inspect Excel input: %w", err)
	}
	if !info.Mode().IsRegular() {
		return "", fmt.Errorf("Excel input is not a regular file: %s", path)
	}
	file, err := os.Open(path)
	if err != nil {
		return "", fmt.Errorf("open Excel input: %w", err)
	}
	if err := file.Close(); err != nil {
		return "", fmt.Errorf("close Excel input: %w", err)
	}
	return path, nil
}
