package excel2pdf

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"testing"
	"time"
)

func TestConvertToPDFBusy(t *testing.T) {
	mutex.Lock()
	defer mutex.Unlock()
	path, err := ConvertToPDF(context.Background(), "unused.xlsx")
	if path != "" || !errors.Is(err, ErrConversionInProgress) {
		t.Fatalf("ConvertToPDF while busy = (%q, %v)", path, err)
	}
}

func TestConvertToPDFInvalidInputReleasesLock(t *testing.T) {
	path, err := ConvertToPDF(context.Background(), "")
	if path != "" || err == nil {
		t.Fatalf("ConvertToPDF with empty path = (%q, %v)", path, err)
	}
	if !mutex.TryLock() {
		t.Fatal("conversion failure left the mutex locked")
	}
	mutex.Unlock()
}

func TestConvertToPDFContext(t *testing.T) {
	canceled, cancel := context.WithCancel(context.Background())
	cancel()
	expired, cancelDeadline := context.WithDeadline(context.Background(), time.Now().Add(-time.Second))
	defer cancelDeadline()
	tests := []struct {
		name string
		ctx  context.Context
		want error
	}{
		{name: "nil context"},
		{name: "already canceled", ctx: canceled, want: context.Canceled},
		{name: "expired deadline", ctx: expired, want: context.DeadlineExceeded},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			pdf, err := ConvertToPDF(tt.ctx, "unused.xlsx")
			if pdf != "" || err == nil {
				t.Fatalf("conversion = (%q, %v), want empty path and error", pdf, err)
			}
			if tt.want != nil && !errors.Is(err, tt.want) {
				t.Errorf("error = %v, want %v", err, tt.want)
			}
			if !mutex.TryLock() {
				t.Fatal("rejected context left the mutex locked")
			}
			mutex.Unlock()
		})
	}
	mutex.Lock()
	defer mutex.Unlock()
	if _, err := ConvertToPDF(context.Background(), "unused.xlsx"); !errors.Is(err, ErrConversionInProgress) {
		t.Errorf("concurrent context conversion error = %v", err)
	}
}

func TestConvertExcelFileCancellation(t *testing.T) {
	t.Parallel()
	t.Run("canceled context does not reach backend", func(t *testing.T) {
		ctx, cancel := context.WithCancel(context.Background())
		cancel()
		_, err := convertDocumentFile(ctx, "unused.xlsx", func(context.Context, string, string) error {
			t.Fatal("backend ran with an already canceled context")
			return nil
		})
		if !errors.Is(err, context.Canceled) {
			t.Errorf("error = %v, want context.Canceled", err)
		}
	})
	t.Run("cancellation discards PDF and preserves backend error", func(t *testing.T) {
		ctx, cancel := context.WithCancel(context.Background())
		defer cancel()
		backendErr := errors.New("backend failure")
		input := writeExcelTestInput(t, "input.xlsx")
		var output string
		pdf, err := convertDocumentFile(ctx, input, func(gotCtx context.Context, _, outputPath string) error {
			if gotCtx != ctx {
				t.Fatal("backend received a different context")
			}
			output = outputPath
			if err := os.WriteFile(output, []byte("%PDF-1.7\n"), 0o600); err != nil {
				t.Fatal(err)
			}
			cancel()
			return backendErr
		})
		if pdf != "" || !errors.Is(err, context.Canceled) || !errors.Is(err, backendErr) {
			t.Errorf("canceled conversion = (%q, %v)", pdf, err)
		}
		if _, err := os.Stat(filepath.Dir(output)); !errors.Is(err, os.ErrNotExist) {
			t.Errorf("canceled conversion directory was not removed: %v", err)
		}
	})
}

func TestConvertExcelFile(t *testing.T) {
	t.Parallel()
	backendErr := errors.New("backend failed")
	tests := []struct {
		name    string
		mode    string
		wantErr bool
	}{
		{name: "non-empty PDF", mode: "valid"},
		{name: "missing PDF", mode: "missing", wantErr: true},
		{name: "empty PDF", mode: "empty", wantErr: true},
		{name: "directory instead of PDF", mode: "directory", wantErr: true},
		{name: "partial PDF and backend error", mode: "failed", wantErr: true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			input := writeExcelTestInput(t, "informe anual ñ.xlsx")
			cwd, err := os.Getwd()
			if err != nil {
				t.Fatal(err)
			}
			relativeInput, err := filepath.Rel(cwd, input)
			if err != nil {
				t.Fatal(err)
			}
			var output string
			got, err := convertDocumentFile(context.Background(), relativeInput, func(_ context.Context, inputPath, outputPath string) error {
				output = outputPath
				if inputPath != input || !filepath.IsAbs(outputPath) {
					t.Fatalf("backend paths = (%q, %q), want absolute paths", inputPath, outputPath)
				}
				if filepath.Base(outputPath) != "informe anual ñ.pdf" {
					t.Fatalf("unexpected output name: %s", outputPath)
				}
				if _, err := os.Lstat(outputPath); !errors.Is(err, os.ErrNotExist) {
					t.Fatalf("output already exists before conversion: %v", err)
				}
				switch tt.mode {
				case "missing":
					return nil
				case "directory":
					return os.Mkdir(outputPath, 0o700)
				case "empty":
					return os.WriteFile(outputPath, nil, 0o600)
				default:
					if err := os.WriteFile(outputPath, []byte("%PDF-1.7\n"), 0o600); err != nil {
						return err
					}
					if tt.mode == "failed" {
						return backendErr
					}
					return nil
				}
			})
			if (err != nil) != tt.wantErr {
				t.Fatalf("convertDocumentFile error = %v, want error = %v", err, tt.wantErr)
			}
			if tt.wantErr {
				if got != "" {
					t.Errorf("failure returned PDF path %q", got)
				}
				if _, statErr := os.Stat(filepath.Dir(output)); !errors.Is(statErr, os.ErrNotExist) {
					t.Errorf("failed conversion directory was not removed: %v", statErr)
				}
				if tt.mode == "failed" && !errors.Is(err, backendErr) {
					t.Errorf("backend error was not preserved: %v", err)
				}
				return
			}
			t.Cleanup(func() { removeExcelTestOutput(t, got) })
			if got != output {
				t.Errorf("returned PDF = %q, want %q", got, output)
			}
			if _, err := os.Stat(got); err != nil {
				t.Errorf("successful PDF was removed: %v", err)
			}
		})
	}
}

func TestConvertExcelFileRejectsInvalidInput(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name string
		path string
	}{
		{name: "empty path"},
		{name: "missing file", path: filepath.Join(t.TempDir(), "missing.xlsx")},
		{name: "directory", path: t.TempDir()},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			path, err := convertDocumentFile(context.Background(), tt.path, func(_ context.Context, _, _ string) error {
				t.Fatal("backend must not run for invalid input")
				return nil
			})
			if path != "" || err == nil {
				t.Errorf("invalid input conversion = (%q, %v)", path, err)
			}
			if tt.name == "missing file" && !errors.Is(err, os.ErrNotExist) {
				t.Errorf("missing file error not preserved: %v", err)
			}
		})
	}
}

func TestConvertExcelFileIsolatesSameName(t *testing.T) {
	t.Parallel()
	firstInput := writeExcelTestInput(t, "report.xlsx")
	secondInput := writeExcelTestInput(t, "report.xlsx")
	writePDF := func(_ context.Context, _, output string) error {
		return os.WriteFile(output, []byte("%PDF-1.7\n"), 0o600)
	}
	first, err := convertDocumentFile(context.Background(), firstInput, writePDF)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { removeExcelTestOutput(t, first) })
	second, err := convertDocumentFile(context.Background(), secondInput, writePDF)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { removeExcelTestOutput(t, second) })
	if filepath.Dir(first) == filepath.Dir(second) {
		t.Fatal("conversions share a temporary directory")
	}
	path, err := convertDocumentFile(context.Background(), secondInput, func(_ context.Context, _, _ string) error { return nil })
	if path != "" || err == nil {
		t.Fatalf("missing output reused an earlier PDF: (%q, %v)", path, err)
	}
	for _, output := range []string{first, second} {
		if _, err := os.Stat(output); err != nil {
			t.Errorf("previous conversion was modified: %v", err)
		}
	}
}

func TestConvertExcelFileRejectsOutputSymlink(t *testing.T) {
	t.Parallel()
	input := writeExcelTestInput(t, "report.xlsx")
	target := filepath.Join(t.TempDir(), "existing.pdf")
	if err := os.WriteFile(target, []byte("previous PDF"), 0o600); err != nil {
		t.Fatal(err)
	}
	if err := os.Symlink(target, filepath.Join(filepath.Dir(target), "probe.pdf")); err != nil {
		t.Skipf("symbolic links are unavailable: %v", err)
	}
	var output string
	path, err := convertDocumentFile(context.Background(), input, func(_ context.Context, _, outputPath string) error {
		output = outputPath
		if err := os.Symlink(target, outputPath); err != nil {
			return err
		}
		return nil
	})
	if path != "" || err == nil {
		t.Fatalf("symlink output conversion = (%q, %v)", path, err)
	}
	if _, err := os.Stat(filepath.Dir(output)); !errors.Is(err, os.ErrNotExist) {
		t.Errorf("output directory was not cleaned: %v", err)
	}
	data, err := os.ReadFile(target)
	if err != nil || string(data) != "previous PDF" {
		t.Errorf("symlink target was modified: %q, %v", data, err)
	}
}

func writeExcelTestInput(t *testing.T, name string) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), name)
	if err := os.WriteFile(path, []byte("test input"), 0o600); err != nil {
		t.Fatal(err)
	}
	return path
}

func removeExcelTestOutput(t *testing.T, path string) {
	t.Helper()
	if err := os.RemoveAll(filepath.Dir(path)); err != nil {
		t.Errorf("remove conversion output: %v", err)
	}
}

func TestDocumentKindForPath(t *testing.T) {
	t.Parallel()
	for _, tt := range []struct {
		path string
		want documentKind
	}{
		{"report.xlsx", spreadsheetDocument},
		{"report.XLSX", spreadsheetDocument},
		{"report.docx", wordDocument},
		{"document with spaces ñ.DoCx", wordDocument},
	} {
		t.Run(tt.path, func(t *testing.T) {
			t.Parallel()
			got, err := documentKindForPath(tt.path)
			if got != tt.want || err != nil {
				t.Errorf("document kind = (%v, %v), want %v", got, err, tt.want)
			}
		})
	}
}

func TestConvertToPDFUnsupportedFormat(t *testing.T) {
	for _, path := range []string{"report.xls", "report.doc", "report.pdf", "report.txt", "report.xlsm", "report", "report.docx.exe"} {
		t.Run(path, func(t *testing.T) {
			pdf, err := ConvertToPDF(context.Background(), path)
			if pdf != "" || !errors.Is(err, ErrUnsupportedFormat) {
				t.Errorf("unsupported conversion = (%q, %v)", pdf, err)
			}
			if !mutex.TryLock() {
				t.Fatal("unsupported conversion left the mutex locked")
			}
			mutex.Unlock()
		})
	}
}
