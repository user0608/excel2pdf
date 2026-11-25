# excel2pdf
Converts an Excel file to PDF, compatible with both Windows and Linux. Uses either LibreOffice or Microsoft Excel, with LibreOffice given priority if installed.

## Engines and configuration

- **Linux:** requires LibreOffice. Detection searches `PATH` for `libreoffice`,
  then `libreoffice24.8`, then `soffice`, without depending on `which`.
- **Windows:** searches `PATH` for `soffice.com`, `soffice.exe`, `libreoffice`, then
  `libreoffice24.8`. If absent, it checks LibreOffice's UNO installation and
  versioned registry keys for the current user and machine, in both 64-bit and
  32-bit views. Registered directories are resolved to `soffice.com` or
  `soffice.exe`, directly or under `program`; stale entries are skipped.
  Microsoft Excel is selected only if LibreOffice is absent, using Excel's COM
  ProgID and local-server registrations in the same registry scopes and views.
  COM registration is a discovery check, not a guarantee that Excel can start.
- **macOS and other operating systems:** not supported.

Set `LIBREOFFICE_PATH` to override discovery with an executable file path or a
command resolvable through `PATH`. A path containing spaces should be stored
without literal wrapping quotes. The result must be a regular executable file;
a directory, empty/whitespace value, or invalid location is an explicit
configuration error. **Unset** the variable to restore automatic discovery.

Detection/permission errors are returned rather than treated as "not installed".
If LibreOffice is selected but conversion fails, the package returns that error
and does **not** retry using Excel. This avoids silently changing the rendering
engine after a document-specific failure. Absence can be inspected using
`errors.Is(err, excel2pdf.ErrLibreofficeNotInstalled)` on Linux, or
`errors.Is(err, excel2pdf.ErrExcelOrLibreOfficeNotInstalled)` when neither engine
is found on Windows.

## Conversion and temporary files

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
