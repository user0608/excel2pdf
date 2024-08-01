package excel2pdf

import (
	"errors"
	"sync"
)

var (
	mutex sync.Mutex
)
var ErrExcel2PdfIsProcessing = errors.New("a process is currently running using the resource")

func ConvertExcelToPdf(excelFile string) (pdfFile string, err error) {
	if !mutex.TryLock() {
		return "", ErrExcel2PdfIsProcessing
	}
	defer mutex.Unlock()
	return convertExcelToPdf(excelFile)
}
