package excel2pdf

import (
	"errors"
	"fmt"
)

// excelAutomationObject isolates the COM boundary so the workbook lifecycle can
// be tested without Windows or an Office installation.
type excelAutomationObject interface {
	getObject(name string, args ...any) (excelAutomationObject, error)
	call(name string, args ...any) error
	release() error
}

func exportExcelPDF(excel excelAutomationObject, inputPath, outputPath string) (err error) {
	defer func() {
		if recovered := recover(); recovered != nil {
			if panicErr, ok := recovered.(error); ok {
				err = errors.Join(err, fmt.Errorf("excel automation panic: %w", panicErr))
			} else {
				err = errors.Join(err, fmt.Errorf("excel automation panic: %v", recovered))
			}
		}
	}()
	defer func() {
		addExcelCleanupError(&err, "quit Excel", excel.call("Quit"))
	}()

	workbooks, err := excel.getObject("Workbooks")
	if err != nil {
		return fmt.Errorf("get Excel workbooks: %w", err)
	}
	defer func() {
		addExcelCleanupError(&err, "release Excel workbooks", workbooks.release())
	}()

	workbook, err := workbooks.getObject("Open", inputPath)
	if err != nil {
		return fmt.Errorf("open Excel workbook: %w", err)
	}
	defer func() {
		addExcelCleanupError(&err, "release Excel workbook", workbook.release())
	}()
	defer func() {
		addExcelCleanupError(&err, "close Excel workbook", workbook.call("Close", false))
	}()

	if err := workbook.call("ExportAsFixedFormat", 0, outputPath); err != nil {
		return fmt.Errorf("export Excel workbook as PDF: %w", err)
	}
	return nil
}

func addExcelCleanupError(err *error, operation string, cleanupErr error) {
	if cleanupErr != nil {
		*err = errors.Join(*err, fmt.Errorf("%s: %w", operation, cleanupErr))
	}
}
