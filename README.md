# excel2pdf
Converts an Excel file to PDF, compatible with both Windows and Linux. Uses either LibreOffice or Microsoft Excel, with LibreOffice given priority if installed.

The input must be a readable, regular file. Relative input paths are resolved to
absolute paths before starting Office. The function returns an error if the
conversion does not produce a non-empty, regular PDF file.

Each conversion writes to its own temporary directory. On failure, the package
removes that directory. On success, the caller owns the PDF and its directory and
must remove them when no longer needed. The returned filename keeps the input's
base name with a `.pdf` extension; uniqueness comes from the containing directory.

```go
package main

import (
    "fmt"
    "log"
    "os"
    "path/filepath"

    "github.com/user0608/excel2pdf"
)

func main() {
    pdfFilePath, err := excel2pdf.ConvertExcelToPdf("file.xlsx")
    if err != nil {
        log.Fatal(err)
    }
    defer func() {
        // Only remove the directory returned by this successful conversion.
        if err := os.RemoveAll(filepath.Dir(pdfFilePath)); err != nil {
            log.Printf("remove PDF temporary directory: %v", err)
        }
    }()

    fmt.Println(pdfFilePath)
    // Read, copy, or send the PDF before this function returns.
}
```
