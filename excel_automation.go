package excel2pdf

import (
	"context"
	"errors"
	"fmt"
)

// officeAutomationObject isolates the COM boundary so the document lifecycle can
// be tested without Windows or an Office installation.
type officeAutomationObject interface {
	getObject(name string, args ...any) (officeAutomationObject, error)
	getBoolProperty(name string) (bool, error)
	setProperty(name string, value any) error
	call(name string, args ...any) error
	release() error
}

func exportExcelPDF(ctx context.Context, excel officeAutomationObject, inputPath, outputPath string) (err error) {
	defer recoverOfficePanic("excel", &err)
	defer func() {
		addOfficeCleanupError(&err, "quit Excel", excel.call("Quit"))
	}()
	return exportWorkbook(ctx, excel, inputPath, outputPath)
}

func recoverOfficePanic(name string, err *error) {
	if recovered := recover(); recovered != nil {
		if cause, ok := recovered.(error); ok {
			*err = errors.Join(*err, fmt.Errorf("%s automation panic: %w", name, cause))
		} else {
			*err = errors.Join(*err, fmt.Errorf("%s automation panic: %v", name, recovered))
		}
	}
}

func exportWorkbook(ctx context.Context, excel officeAutomationObject, inputPath, outputPath string) (err error) {
	defer recoverOfficePanic("excel", &err)
	if err := configureExcelSecurity(ctx, excel); err != nil {
		return err
	}

	workbooks, err := excel.getObject("Workbooks")
	if err != nil {
		return fmt.Errorf("get Excel workbooks: %w", err)
	}
	defer func() {
		addOfficeCleanupError(&err, "release Excel workbooks", workbooks.release())
	}()
	if err := ctx.Err(); err != nil {
		return err
	}

	// UpdateLinks=0 prevents workbook-link updates; ReadOnly=true avoids
	// modifying the source. No password or interactive-opening API is exposed.
	workbook, err := workbooks.getObject("Open", inputPath, 0, true)
	if err != nil {
		return fmt.Errorf("open Excel workbook: %w", err)
	}
	defer func() {
		addOfficeCleanupError(&err, "release Excel workbook", workbook.release())
	}()
	defer func() {
		addOfficeCleanupError(&err, "close Excel workbook", workbook.call("Close", false))
	}()
	if err := ctx.Err(); err != nil {
		return err
	}

	if err := workbook.call("ExportAsFixedFormat", 0, outputPath); err != nil {
		return fmt.Errorf("export Excel workbook as PDF: %w", err)
	}
	return ctx.Err()
}

func configureExcelSecurity(ctx context.Context, excel officeAutomationObject) error {
	for _, property := range []struct {
		name  string
		value any
	}{
		// msoAutomationSecurityForceDisable disables VBA macros for files
		// opened programmatically. Excel 4.0/XLM macros are not covered by it.
		{name: "AutomationSecurity", value: 3},
		{name: "EnableEvents", value: false},
		{name: "AskToUpdateLinks", value: false},
		{name: "DisplayAlerts", value: false},
		{name: "Visible", value: false},
	} {
		if err := ctx.Err(); err != nil {
			return err
		}
		if err := excel.setProperty(property.name, property.value); err != nil {
			return fmt.Errorf("set Excel %s: %w", property.name, err)
		}
	}
	return ctx.Err()
}

func addOfficeCleanupError(err *error, operation string, cleanupErr error) {
	if cleanupErr != nil {
		*err = errors.Join(*err, fmt.Errorf("%s: %w", operation, cleanupErr))
	}
}
