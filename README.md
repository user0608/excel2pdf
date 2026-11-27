# excel2pdf
Converts `.xlsx` spreadsheets and `.docx` documents to PDF on Linux and Windows.
The single conversion API is `ConvertToPDF(ctx, inputPath)`. LibreOffice has
priority; Windows can use Microsoft Excel for XLSX or Microsoft Word for DOCX.
Extensions are case-insensitive. Other formats return `ErrUnsupportedFormat`.

Requires Go 1.22.5 or newer. Use a currently maintained Go release in production;
the minimum version is tested for API compatibility, not recommended for security.

## Engines and configuration

- **Linux:** requires LibreOffice. Detection searches `PATH` for `libreoffice`,
  then `libreoffice24.8`, then `soffice`, without depending on `which`.
- **Windows:** searches `PATH` for `soffice.com`, `soffice.exe`, `libreoffice`, then
  `libreoffice24.8`. If absent, it checks LibreOffice's UNO installation and
  versioned registry keys for the current user and machine, in both 64-bit and
  32-bit views. Registered directories are resolved to `soffice.com` or
   `soffice.exe`, directly or under `program`; stale entries are skipped.
  If neither PATH nor the registry contains an installation, it checks
  `%ProgramW6432%`, `%ProgramFiles%`, then `%ProgramFiles(x86)%`, under
  `LibreOffice\program`, preferring `soffice.com` to `soffice.exe` in each folder.
  Unset or relative directory variables are ignored; no drive letter is assumed.
  Excel (`Excel.Application`) is selected for `.xlsx`, or Word (`Word.Application`)
  for `.docx`, only if LibreOffice is absent. COM ProgID and local-server
  registrations are checked in the same scopes and views. Having Excel installed
  is not sufficient for DOCX conversion: Word or LibreOffice is required.
  COM registration is a discovery check, not a guarantee that Office can start.
- **macOS and other operating systems:** not supported.

Set `LIBREOFFICE_PATH` to override discovery with an executable file path or a
command resolvable through `PATH`. A path containing spaces should be stored
without literal wrapping quotes. The result must be a regular executable file;
a directory, empty/whitespace value, or invalid location is an explicit
configuration error. **Unset** the variable to restore automatic discovery.

Detection/permission errors are returned rather than treated as "not installed".
If LibreOffice is selected but conversion fails, the package returns that error
and does **not** retry using Microsoft Office. This avoids silently changing the rendering
engine after a document-specific failure. Absence can be inspected using
`errors.Is(err, excel2pdf.ErrLibreofficeNotInstalled)` on Linux, or
`errors.Is(err, excel2pdf.ErrOfficeNotInstalled)` when neither LibreOffice nor the
Microsoft application required for the input format is found on Windows.

To inspect the LibreOffice executable that would be selected:

```go
path, err := excel2pdf.LibreOfficePath()
if err != nil {
    return err
}
fmt.Println("LibreOffice:", path)
```

This uses the same detection as conversion without launching LibreOffice. It
returns an absolute path, does not cache the result, and does not verify the
version. A later conversion detects again, so keep the environment unchanged.
The library does not print discovery logs automatically. Invalid explicit paths
and lookup/permission errors are returned, not treated as absence or a reason to
fall back to Microsoft Office.

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
    "context"
    "fmt"
    "log"
    "os"
    "path/filepath"
    "time"

    "github.com/user0608/excel2pdf"
)

func main() {
    ctx, cancel := context.WithTimeout(context.Background(), 30*time.Second)
    defer cancel()

    pdfFilePath, err := excel2pdf.ConvertToPDF(ctx, "file.docx") // or file.xlsx
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

`ConvertToPDF` always accepts a context. Supply a cancellation signal or deadline:

```go
ctx, cancel := context.WithTimeout(context.Background(), 30*time.Second)
defer cancel()

pdfFilePath, err := excel2pdf.ConvertToPDF(ctx, "file.xlsx")
// Use errors.Is(err, context.Canceled) or
// errors.Is(err, context.DeadlineExceeded) to distinguish cancellation/timeouts.
// On success, consume and clean up the PDF as in the example above.
```

The snippet requires the `context` and `time` packages. There is **no default
timeout**; use `context.Background()` if you deliberately do not want a deadline.
Requests are not queued: another active conversion (of either format) in the same
Go process causes `ErrConversionInProgress`. A nil context is rejected with an
error. The previous Excel-specific conversion functions have been removed.

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
- **Microsoft Excel and Word:** cancellation is checked between COM operations. An ongoing
  or blocked COM call, including cleanup, cannot be interrupted by a context, so
  this API does **not** provide a strict Office timeout. The package waits for COM
  and attempts resource cleanup before returning. A strict timeout would require
  a separate worker process.

When cancellation is observed, the function returns no PDF path and attempts to
remove its temporary files, even if the backend produced a PDF. Context errors
remain inspectable with `errors.Is`, alongside any backend or cleanup errors.

## Document security and deployment

Before accessing Excel workbooks, the package sets all of these application
properties on the new Excel instance:

- `AutomationSecurity = 3` (`msoAutomationSecurityForceDisable`) to disable VBA
  macros in documents opened programmatically.
- `EnableEvents = false`, `AskToUpdateLinks = false`, `DisplayAlerts = false`, and
  `Visible = false` to reduce event execution and interactive prompts.

If any property cannot be applied, the conversion fails **before opening the
document** and attempts to quit Excel. Workbooks are opened with `UpdateLinks = 0`
and `ReadOnly = true`, then closed without saving. Macros and link-dependent
calculations are intentionally not refreshed, so output may differ from an
interactive session that executes document code or updates external values.

For Word, the package applies `AutomationSecurity = 3`, `DisplayAlerts = 0`, and
`Visible = false`, then disables `Options.UpdateLinksAtOpen` before accessing
documents. It opens DOCX with `ConfirmConversions = false`, `ReadOnly = true`, and
`AddToRecentFiles = false`, exports with `wdExportFormatPDF` without opening the
result, and closes/quits using `wdDoNotSaveChanges`. Configuration failures abort
before the document is opened; resources are released on success, failure, and
observed cancellation. The original `UpdateLinksAtOpen` option is restored during
cleanup so the conversion does not intentionally persist that preference change.
Restoration errors are returned. Neither application is reused from the user's session.

These controls have important limits:

- `AutomationSecurity` does **not** disable Excel 4.0/XLM macros. Excel security
  policies, Protected View, password protection, add-ins, or security warnings may
  still block opening or require interaction. The package does not lower Trust
  Center settings or bypass Protected View, and has no password-opening API.
- `UpdateLinks = 0` is not a network sandbox. External data connections, embedded
  objects, parser vulnerabilities, and engine-specific behavior are not fully
  controlled by these properties. A fresh LibreOffice profile is also **not a
  sandbox** or a guarantee that every kind of active content is disabled.
- Microsoft does not recommend or support unattended server-side Office
  automation in typical service scenarios. Excel and Word may require an initialized user
  profile and an interactive desktop, and can still hang on a COM operation.
  Validate the deployment on a dedicated Windows worker; do not assume that a
  web-service or SYSTEM account can safely run Office without interaction.

For documents from third parties, use a dedicated unprivileged account and an
isolated worker/container/VM with no secrets, restricted filesystem access, and
network access denied unless explicitly required. Enforce file-size, CPU, memory,
and execution limits outside this package, and keep Office and LibreOffice
patched. Strict interruption of Excel or Word still requires a separate worker process.

LibreOffice diagnostics redact the known input/output/profile paths and their
file-URI forms, and strip line breaks and terminal control characters. This is
**best-effort redaction**, not a guarantee against sensitive content: other paths,
document-derived text, and underlying inspectable errors may still contain private
data. Do not send raw errors to end users or public logs; translate them to safe
application messages and apply your own logging/redaction policy.

## Tests and CI

Unit tests use controlled subprocesses and COM test doubles; they do not require
an Office installation:

```sh
go test -race -shuffle=on -count=1 -timeout=2m ./...
go vet ./...
golangci-lint run ./...
```

The lint configuration includes integration-tagged sources without running them.
CI pins golangci-lint v2.15.0; an older linter may not understand export data from a
newer Go toolchain. Reachability-based dependency checks can be run with
`govulncheck ./...`; on Linux, `GOOS=windows GOARCH=amd64 govulncheck ./...` also
checks the Windows call graph using the host's installed scanner binary.

### Real LibreOffice integration

The `integration` build tag enables tests that generate small XLSX and DOCX fixtures
without additional Go dependencies, convert each twice through the public API, and
convert same-stem inputs of both formats from two separate processes. They check the PDF header,
unique output directories, profile cleanup, and preservation of the source file.
They do not validate visual fidelity, page layout, or every Office feature.

On Linux, with LibreOffice Calc and Writer installed:

```sh
EXCEL2PDF_REQUIRE_LIBREOFFICE=1 go test -tags=integration -run '^TestIntegrationLibreOffice' -count=1 -timeout=5m -v .
```

Set `LIBREOFFICE_PATH` if the executable is not discovered automatically. Without
`EXCEL2PDF_REQUIRE_LIBREOFFICE=1`, these tests skip when LibreOffice is absent;
invalid explicit configuration still fails. With the variable set, absence fails
the suite instead of producing a misleading green result.

On Windows PowerShell, with LibreOffice installed:

```powershell
$env:LIBREOFFICE_PATH = 'C:\Program Files\LibreOffice\program\soffice.com'
$env:EXCEL2PDF_REQUIRE_LIBREOFFICE = '1'
go test -tags=integration -run '^TestIntegrationLibreOffice' -count=1 -timeout=5m -v .
```

To check interaction with an already-open desktop instance, run this suite while
that instance is open. The tests themselves do not launch a GUI instance.

### Real Excel and Word integration (Windows only)

Excel tests require both the build tag and an explicit environment opt-in. Run
them only on a prepared, isolated Windows machine with Excel installed, a suitable
user session, and no sensitive documents. They select the Excel backend directly
even if LibreOffice is installed, test open/export failures and subsequent
recovery, verify source preservation, and check that the source can be renamed
after conversion. File-handle checks do not prove that every Excel process exited.

```powershell
$env:EXCEL2PDF_TEST_EXCEL = '1'
go test -tags=integration -run '^TestIntegrationExcel' -count=1 -timeout=5m -v .
Remove-Item Env:EXCEL2PDF_TEST_EXCEL
```

The equivalent Word tests require Word installed and their own explicit opt-in:

```powershell
$env:EXCEL2PDF_TEST_WORD = '1'
go test -tags=integration -run '^TestIntegrationWord' -count=1 -timeout=5m -v .
Remove-Item Env:EXCEL2PDF_TEST_WORD
```

Word tests also select their backend directly, exercise opening/export failures
and recovery, and check source preservation and file-handle release.

The context timeout still cannot interrupt a blocked COM call. The test runner's
outer timeout can stop the Go test process but may leave Excel or Word running; inspect
the dedicated test environment afterward. These tests never terminate unrelated
Office instances by name.

### GitHub Actions

`.github/workflows/ci.yml` runs:

- Race-enabled, shuffled unit tests and coverage on Linux and Windows, with Go
  1.22.x and the current stable version; formatting, vet, and module metadata checks.
- Cross-builds and compilation of unit/integration tests for Linux amd64/arm64 and
  Windows 386/amd64/arm64. Cross-compilation does **not** execute those tests.
- Linting and vulnerability reachability checks for Linux and Windows targets.
- Real LibreOffice integration on Linux when manually dispatched with
  `libreoffice_integration` enabled. That job installs Calc, Writer and fonts and
  requires the engine to be present. Excel/Word are not launched by hosted CI.

Coverage is retained as an artifact, not used as an arbitrary pass/fail threshold.
The workflow uses read-only repository permissions and does not publish releases
or deploy anything. Windows Excel, Word and LibreOffice integration remain explicit
local checks on prepared machines.

## Example program

```sh
go run ./example /absolute/path/to/report.xlsx
# Or: go run ./example /absolute/path/to/document.docx
```

The example uses a 30-second context deadline, reports the generated file's size,
and removes its temporary output before exiting. To retain the PDF in an
application, copy it to your destination before cleaning up the returned directory.
