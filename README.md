# excel2pdf
Converts an Excel file to PDF, compatible with both Windows and Linux. Uses either LibreOffice or Microsoft Excel, with LibreOffice given priority if installed.

The input must be a readable, regular file. Relative input paths are resolved to
absolute paths before starting Office. The function returns an error if the
conversion does not produce a non-empty, regular PDF file.

Each conversion writes to its own temporary directory. On failure, the package
removes that directory. On success, the caller owns the PDF and its directory and
must remove them when no longer needed. The returned filename keeps the input's
base name with a `.pdf` extension; uniqueness comes from the containing directory.

LibreOffice uses a separate temporary user profile for every conversion, rather
than sharing the user's existing profile. The package removes that profile after
the process finishes, including on conversion failures. Process output is not
written to the application's stdout or stderr; errors include captured diagnostics
(up to 64 KiB, with a truncation marker when needed). A successful process exit
without the expected PDF is still treated as a conversion failure.

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

## Cancellation and timeouts

Use `ConvertExcelToPdfContext` to supply a cancellation signal or deadline:

```go
ctx, cancel := context.WithTimeout(context.Background(), 30*time.Second)
defer cancel()

pdfFilePath, err := excel2pdf.ConvertExcelToPdfContext(ctx, "file.xlsx")
// Use errors.Is(err, context.Canceled) or
// errors.Is(err, context.DeadlineExceeded) to distinguish cancellation/timeouts.
// On success, consume and clean up the PDF as in the example above.
```

The snippet requires the `context` and `time` packages. The original
`ConvertExcelToPdf` remains available and imposes **no default timeout**. Neither
function queues concurrent requests: another active conversion in the same Go
process causes `ErrExcel2PdfIsProcessing`. A nil context is rejected with an error.

- **LibreOffice on Linux:** cancellation kills the conversion's own process group,
  including descendants that remain in that group. Detached processes are not
  guaranteed to be terminated. Output-pipe waits after process exit/cancellation
  are capped at two seconds; process termination and filesystem cleanup can add
  time beyond the context's deadline.
- **LibreOffice on Windows:** cancellation kills only the launched process.
  Descendant termination is **not guaranteed**; strict process-tree cancellation
  requires a Windows Job Object or an isolated worker. Descendants may also retain
  file handles and prevent temporary-directory cleanup, which is reported as an
  error. The package never terminates other Office processes by name.
- **Microsoft Excel:** cancellation is checked between COM operations. An ongoing
  or blocked COM call, including cleanup, cannot be interrupted by a context, so
  this API does **not** provide a strict Excel timeout. The package waits for COM
  and attempts resource cleanup before returning. A strict timeout would require
  a separate worker process.

When cancellation is observed, the function returns no PDF path and attempts to
remove its temporary files, even if the backend produced a PDF. Context errors
remain inspectable with `errors.Is`, alongside any backend or cleanup errors.
