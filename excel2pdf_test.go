package excel2pdf

import (
	"errors"
	"os"
	"path/filepath"
	"testing"
)

func TestConvertExcelToPdfBusy(t *testing.T) {
	mutex.Lock()
	defer mutex.Unlock()
	path, err := ConvertExcelToPdf("unused.xlsx")
	if path != "" || !errors.Is(err, ErrExcel2PdfIsProcessing) {
		t.Fatalf("ConvertExcelToPdf while busy = (%q, %v)", path, err)
	}
}

func TestConvertExcelToPdfInvalidInputReleasesLock(t *testing.T) {
	path, err := ConvertExcelToPdf("")
	if path != "" || err == nil {
		t.Fatalf("ConvertExcelToPdf with empty path = (%q, %v)", path, err)
	}
	if !mutex.TryLock() {
		t.Fatal("conversion failure left the mutex locked")
	}
	mutex.Unlock()
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
			got, err := convertExcelFile(relativeInput, func(inputPath, outputPath string) error {
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
				t.Fatalf("convertExcelFile error = %v, want error = %v", err, tt.wantErr)
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
			path, err := convertExcelFile(tt.path, func(_, _ string) error {
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
	writePDF := func(_, output string) error {
		return os.WriteFile(output, []byte("%PDF-1.7\n"), 0o600)
	}
	first, err := convertExcelFile(firstInput, writePDF)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { removeExcelTestOutput(t, first) })
	second, err := convertExcelFile(secondInput, writePDF)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { removeExcelTestOutput(t, second) })
	if filepath.Dir(first) == filepath.Dir(second) {
		t.Fatal("conversions share a temporary directory")
	}
	path, err := convertExcelFile(secondInput, func(_, _ string) error { return nil })
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
	path, err := convertExcelFile(input, func(_, outputPath string) error {
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
