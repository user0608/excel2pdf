//go:build windows
// +build windows

package excel2pdf

import (
	"fmt"
	"log/slog"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func convertExcelToPDFWithExcel(excelFilePath string) (pdfFilePath string, err error) {
	defer func() {
		if r := recover(); r != nil {
			slog.Error("convertExcelToPDF recovery", "recover", r)
		}
	}()
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	err = ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_SPEED_OVER_MEMORY)
	if err != nil {
		return "", fmt.Errorf("failed to initialize OLE: %w", err)
	}
	defer ole.CoUninitialize()

	// Crea una instancia de Excel
	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		return "", fmt.Errorf("failed to create Excel object: %w", err)
	}
	defer unknown.Release()

	// Obtiene la interfaz de Excel
	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return "", fmt.Errorf("failed to query Excel interface: %w", err)
	}
	defer excel.Release()
	workbooks := oleutil.MustGetProperty(excel, "Workbooks").ToIDispatch()
	workbook := oleutil.MustCallMethod(workbooks, "Open", excelFilePath).ToIDispatch()
	defer workbook.Release()
	defer func() {
		if _, err := oleutil.CallMethod(workbook, "Close", false); err != nil {
			slog.Error("failed to close workbook", "error", err)
		}
		if _, err := oleutil.CallMethod(excel, "Quit"); err != nil {

			slog.Error("failed to quit Excel", "error", err)
		}
	}()

	pdfFilePath = fmt.Sprintf("%s-%d.pdf",
		strings.TrimSuffix(
			filepath.Base(excelFilePath),
			filepath.Ext(excelFilePath),
		),
		time.Now().Unix(),
	)
	pdfFilePath = filepath.Join(os.TempDir(), pdfFilePath)
	exportArgs := []interface{}{0, pdfFilePath}

	if _, err := oleutil.CallMethod(workbook, "ExportAsFixedFormat", exportArgs...); err != nil {
		return "", fmt.Errorf("failed to export as PDF: %w", err)
	}
	return pdfFilePath, nil
}
