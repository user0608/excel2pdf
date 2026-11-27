package excel2pdf

import (
	"context"
	"fmt"
)

func exportWordPDF(ctx context.Context, word officeAutomationObject, inputPath, outputPath string) (err error) {
	defer recoverOfficePanic("word", &err)
	defer func() {
		addOfficeCleanupError(&err, "quit Word", word.call("Quit", 0)) // wdDoNotSaveChanges
	}()
	return exportDocument(ctx, word, inputPath, outputPath)
}

func exportDocument(ctx context.Context, word officeAutomationObject, inputPath, outputPath string) (err error) {
	defer recoverOfficePanic("word", &err)
	for _, property := range []struct {
		name  string
		value any
	}{
		{name: "AutomationSecurity", value: 3}, // msoAutomationSecurityForceDisable
		{name: "DisplayAlerts", value: 0},      // wdAlertsNone
		{name: "Visible", value: false},
	} {
		if err := ctx.Err(); err != nil {
			return err
		}
		if err := word.setProperty(property.name, property.value); err != nil {
			return fmt.Errorf("set Word %s: %w", property.name, err)
		}
	}
	if err := ctx.Err(); err != nil {
		return err
	}
	options, err := word.getObject("Options")
	if err != nil {
		return fmt.Errorf("get Word options: %w", err)
	}
	defer func() { addOfficeCleanupError(&err, "release Word options", options.release()) }()
	if err := ctx.Err(); err != nil {
		return err
	}
	originalUpdateLinks, err := options.getBoolProperty("UpdateLinksAtOpen")
	if err != nil {
		return fmt.Errorf("read Word link-update setting: %w", err)
	}
	if err := ctx.Err(); err != nil {
		return err
	}
	// Options may persist in the user's Office profile. Restore the original
	// value after the document is closed, even if opening/exporting fails.
	defer func() {
		addOfficeCleanupError(&err, "restore Word link-update setting", options.setProperty("UpdateLinksAtOpen", originalUpdateLinks))
	}()
	if err := options.setProperty("UpdateLinksAtOpen", false); err != nil {
		return fmt.Errorf("disable Word link updates: %w", err)
	}
	if err := ctx.Err(); err != nil {
		return err
	}
	documents, err := word.getObject("Documents")
	if err != nil {
		return fmt.Errorf("get Word documents: %w", err)
	}
	defer func() { addOfficeCleanupError(&err, "release Word documents", documents.release()) }()
	if err := ctx.Err(); err != nil {
		return err
	}
	// FileName, ConfirmConversions=false, ReadOnly=true, AddToRecentFiles=false.
	document, err := documents.getObject("Open", inputPath, false, true, false)
	if err != nil {
		return fmt.Errorf("open Word document: %w", err)
	}
	defer func() { addOfficeCleanupError(&err, "release Word document", document.release()) }()
	defer func() { addOfficeCleanupError(&err, "close Word document", document.call("Close", 0)) }()
	if err := ctx.Err(); err != nil {
		return err
	}
	// OutputFileName, ExportFormat=wdExportFormatPDF, OpenAfterExport=false.
	if err := document.call("ExportAsFixedFormat", outputPath, 17, false); err != nil {
		return fmt.Errorf("export Word document as PDF: %w", err)
	}
	return ctx.Err()
}
