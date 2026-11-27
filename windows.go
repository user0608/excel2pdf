//go:build windows
// +build windows

package excel2pdf

import (
	"context"
	"errors"
	"fmt"
	"runtime"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func convertExcelToPDFWithExcel(ctx context.Context, excelFilePath, pdfFilePath string) error {
	return convertWithMicrosoftOffice(ctx, excelFilePath, pdfFilePath, "Excel.Application", exportExcelPDF)
}

func convertDocumentToPDFWithWord(ctx context.Context, inputPath, pdfFilePath string) error {
	return convertWithMicrosoftOffice(ctx, inputPath, pdfFilePath, "Word.Application", exportWordPDF)
}

func convertWithMicrosoftOffice(ctx context.Context, inputPath, pdfFilePath, programID string, export func(context.Context, officeAutomationObject, string, string) error) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_SPEED_OVER_MEMORY)
	if !officeCOMInitialized(err) {
		return fmt.Errorf("failed to initialize OLE: %w", err)
	}
	defer ole.CoUninitialize()

	application, err := createOfficeApplication(programID)
	if err != nil {
		return err
	}
	defer application.dispatch.Release()
	return export(ctx, application, inputPath, pdfFilePath)
}

func createOfficeApplication(programID string) (*comOfficeObject, error) {
	unknown, err := oleutil.CreateObject(programID)
	if err != nil {
		return nil, fmt.Errorf("create %s object: %w", programID, err)
	}
	defer unknown.Release()

	application, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, fmt.Errorf("query %s interface: %w", programID, err)
	}
	return &comOfficeObject{dispatch: application}, nil
}

func officeCOMInitialized(err error) bool {
	if err == nil {
		return true
	}
	// go-ole reports S_FALSE as an error, although COM was initialized and
	// CoUninitialize must still be called to balance this initialization.
	var oleErr *ole.OleError
	return errors.As(err, &oleErr) && oleErr.Code() == 1
}

type comOfficeObject struct {
	dispatch *ole.IDispatch
	owner    *ole.VARIANT
}

func (object *comOfficeObject) getObject(name string, args ...any) (officeAutomationObject, error) {
	var value *ole.VARIANT
	var err error
	switch name {
	case "Workbooks", "Documents", "Options":
		value, err = oleutil.GetProperty(object.dispatch, name, args...)
	default:
		value, err = oleutil.CallMethod(object.dispatch, name, args...)
	}
	if err != nil {
		if value != nil {
			addOfficeCleanupError(&err, "clear failed COM result", value.Clear())
		}
		return nil, err
	}
	if value == nil {
		return nil, fmt.Errorf("%s returned no COM value", name)
	}
	dispatch := value.ToIDispatch()
	if dispatch == nil {
		err = fmt.Errorf("%s returned no COM dispatch interface", name)
		addOfficeCleanupError(&err, "clear invalid COM result", value.Clear())
		return nil, err
	}
	return &comOfficeObject{dispatch: dispatch, owner: value}, nil
}

func (object *comOfficeObject) setProperty(name string, value any) error {
	result, err := oleutil.PutProperty(object.dispatch, name, value)
	if result != nil {
		addOfficeCleanupError(&err, "clear COM property result", result.Clear())
	}
	return err
}

func (object *comOfficeObject) getBoolProperty(name string) (value bool, err error) {
	result, err := oleutil.GetProperty(object.dispatch, name)
	if result != nil {
		defer func() { addOfficeCleanupError(&err, "clear COM property result", result.Clear()) }()
	}
	if err != nil {
		return false, err
	}
	if result == nil || result.VT != ole.VT_BOOL {
		return false, fmt.Errorf("%s returned no boolean COM value", name)
	}
	return result.Val&0xffff != 0, nil
}

func (object *comOfficeObject) call(name string, args ...any) error {
	value, err := oleutil.CallMethod(object.dispatch, name, args...)
	if value != nil {
		addOfficeCleanupError(&err, "clear COM call result", value.Clear())
	}
	return err
}

func (object *comOfficeObject) release() error {
	// The VARIANT owns the dispatch reference; clearing it releases the
	// interface. Calling dispatch.Release as well would release it twice.
	return object.owner.Clear()
}
