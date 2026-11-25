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
var ErrConversionInProgress = errors.New("a conversion is currently running")
var ErrUnsupportedFormat = errors.New("unsupported document format")

// ConvertToPDF converts an .xlsx or .docx file to PDF, preferring LibreOffice.
// Windows falls back to Excel or Word when LibreOffice is not installed.
// The caller owns the returned PDF and its exclusive temporary directory and
// must remove them after use. Failed conversions are cleaned up by the package.
// Concurrent conversions are rejected. ctx must not be nil; there is no default
// timeout, and a context cannot interrupt an ongoing Microsoft Office COM call.
func ConvertToPDF(ctx context.Context, inputPath string) (pdfFile string, err error) {
	if ctx == nil {
		return "", errors.New("conversion context must not be nil")
	}
	if err := ctx.Err(); err != nil {
		return "", err
	}
	if !mutex.TryLock() {
		return "", ErrConversionInProgress
	}
	defer mutex.Unlock()
	return convertDocumentFile(ctx, inputPath, convertToPDF)
}

func convertDocumentFile(ctx context.Context, input string, convert func(context.Context, string, string) error) (pdfFile string, err error) {
	if err := ctx.Err(); err != nil {
		return "", err
	}
	inputPath, err := absoluteInputPath(input)
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
		return "", fmt.Errorf("convert document to PDF: %w", err)
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

func absoluteInputPath(path string) (string, error) {
	if path == "" {
		return "", errors.New("input path is empty")
	}
	path, err := filepath.Abs(path)
	if err != nil {
		return "", fmt.Errorf("resolve input path: %w", err)
	}
	if _, err := documentKindForPath(path); err != nil {
		return "", err
	}
	info, err := os.Stat(path)
	if err != nil {
		return "", fmt.Errorf("inspect input: %w", err)
	}
	if !info.Mode().IsRegular() {
		return "", fmt.Errorf("input is not a regular file: %s", path)
	}
	file, err := os.Open(path)
	if err != nil {
		return "", fmt.Errorf("open input: %w", err)
	}
	if err := file.Close(); err != nil {
		return "", fmt.Errorf("close input: %w", err)
	}
	return path, nil
}

type documentKind uint8

const (
	spreadsheetDocument documentKind = iota
	wordDocument
)

func documentKindForPath(path string) (documentKind, error) {
	switch extension := strings.ToLower(filepath.Ext(path)); extension {
	case ".xlsx":
		return spreadsheetDocument, nil
	case ".docx":
		return wordDocument, nil
	default:
		return 0, fmt.Errorf("%w: %q (expected .xlsx or .docx)", ErrUnsupportedFormat, extension)
	}
}
