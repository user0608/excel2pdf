package excel2pdf

import (
	"context"
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
	return ConvertExcelToPdfContext(context.Background(), excelFile)
}

// ConvertExcelToPdfContext is ConvertExcelToPdf with cancellation support. It
// rejects concurrent conversions rather than waiting for the active conversion.
// LibreOffice's launched process is canceled with ctx; on Linux its process group
// is killed as well. Excel checks ctx between COM operations, but a blocked COM
// call cannot be interrupted. No default timeout is imposed. ctx must not be nil.
func ConvertExcelToPdfContext(ctx context.Context, excelFile string) (pdfFile string, err error) {
	if ctx == nil {
		return "", errors.New("conversion context must not be nil")
	}
	if err := ctx.Err(); err != nil {
		return "", err
	}
	if !mutex.TryLock() {
		return "", ErrExcel2PdfIsProcessing
	}
	defer mutex.Unlock()
	return convertExcelFile(ctx, excelFile, convertExcelToPdf)
}

func convertExcelFile(ctx context.Context, excelFile string, convert func(context.Context, string, string) error) (pdfFile string, err error) {
	if err := ctx.Err(); err != nil {
		return "", err
	}
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
	if err := ctx.Err(); err != nil {
		return "", err
	}
	err = convert(ctx, inputPath, outputPath)
	if ctxErr := ctx.Err(); ctxErr != nil {
		err = errors.Join(err, ctxErr)
	}
	if err != nil {
		return "", fmt.Errorf("convert Excel to PDF: %w", err)
	}
	if err := validateGeneratedPDF(outputPath); err != nil {
		return "", err
	}
	keepOutput = true
	return outputPath, nil
}

func validateGeneratedPDF(path string) error {
	info, err := os.Lstat(path)
	if err != nil {
		return fmt.Errorf("inspect generated PDF: %w", err)
	}
	if !info.Mode().IsRegular() || info.Size() == 0 {
		return fmt.Errorf("generated PDF is not a non-empty regular file: %s", path)
	}
	return nil
}

func absoluteExcelPath(path string) (string, error) {
	if path == "" {
		return "", errors.New("input path for Excel is empty")
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
		return "", fmt.Errorf("input for Excel is not a regular file: %s", path)
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
