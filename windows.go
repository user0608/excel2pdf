//go:build windows
// +build windows

package excel2pdf

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func convertExcelToPDFWithExcel(excelFilePath string) (pdfFilePath string, err error) {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	err = ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_SPEED_OVER_MEMORY)
	if !excelCOMInitialized(err) {
		return "", fmt.Errorf("failed to initialize OLE: %w", err)
	}
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		return "", fmt.Errorf("failed to create Excel object: %w", err)
	}
	defer unknown.Release()

	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return "", fmt.Errorf("failed to query Excel interface: %w", err)
	}
	defer excel.Release()

	pdfFilePath = fmt.Sprintf("%s-%d.pdf",
		strings.TrimSuffix(
			filepath.Base(excelFilePath),
			filepath.Ext(excelFilePath),
		),
		time.Now().Unix(),
	)
	pdfFilePath = filepath.Join(os.TempDir(), pdfFilePath)
	if err := exportExcelPDF(&comExcelObject{dispatch: excel}, excelFilePath, pdfFilePath); err != nil {
		return "", err
	}
	return pdfFilePath, nil
}

func excelCOMInitialized(err error) bool {
	if err == nil {
		return true
	}
	// go-ole reports S_FALSE as an error, although COM was initialized and
	// CoUninitialize must still be called to balance this initialization.
	var oleErr *ole.OleError
	return errors.As(err, &oleErr) && oleErr.Code() == 1
}

type comExcelObject struct {
	dispatch *ole.IDispatch
	owner    *ole.VARIANT
}

func (object *comExcelObject) getObject(name string, args ...any) (excelAutomationObject, error) {
	var value *ole.VARIANT
	var err error
	if name == "Workbooks" {
		value, err = oleutil.GetProperty(object.dispatch, name, args...)
	} else {
		value, err = oleutil.CallMethod(object.dispatch, name, args...)
	}
	if err != nil {
		if value != nil {
			addExcelCleanupError(&err, "clear failed COM result", value.Clear())
		}
		return nil, err
	}
	if value == nil {
		return nil, fmt.Errorf("%s returned no COM value", name)
	}
	dispatch := value.ToIDispatch()
	if dispatch == nil {
		err = fmt.Errorf("%s returned no COM dispatch interface", name)
		addExcelCleanupError(&err, "clear invalid COM result", value.Clear())
		return nil, err
	}
	return &comExcelObject{dispatch: dispatch, owner: value}, nil
}

func (object *comExcelObject) call(name string, args ...any) error {
	value, err := oleutil.CallMethod(object.dispatch, name, args...)
	if value != nil {
		addExcelCleanupError(&err, "clear COM call result", value.Clear())
	}
	return err
}

func (object *comExcelObject) release() error {
	// The VARIANT owns the dispatch reference; clearing it releases the
	// interface. Calling dispatch.Release as well would release it twice.
	return object.owner.Clear()
}
