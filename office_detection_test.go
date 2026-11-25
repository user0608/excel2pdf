package excel2pdf

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"testing"
)

func TestLibreOfficeOverride(t *testing.T) {
	executable := writeOfficeTestExecutable(t, t.TempDir(), "office with spaces ñ")
	tests := []struct {
		name       string
		value      string
		unset      bool
		configured bool
		wantPath   string
		wantErr    bool
	}{
		{name: "unset", unset: true},
		{name: "empty", configured: true, wantErr: true},
		{name: "whitespace", value: "   ", configured: true, wantErr: true},
		{name: "valid executable", value: executable, configured: true, wantPath: executable},
		{name: "missing executable", value: filepath.Join(t.TempDir(), "missing"), configured: true, wantErr: true},
		{name: "directory", value: t.TempDir(), configured: true, wantErr: true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Setenv("LIBREOFFICE_PATH", tt.value)
			if tt.unset {
				unsetOfficeTestEnv(t, "LIBREOFFICE_PATH")
			}
			path, configured, err := libreOfficeOverride()
			if path != tt.wantPath || configured != tt.configured || (err != nil) != tt.wantErr {
				t.Errorf("override = (%q, %v, %v), want (%q, %v, error=%v)", path, configured, err, tt.wantPath, tt.configured, tt.wantErr)
			}
			if tt.wantErr && errors.Is(err, ErrLibreofficeNotInstalled) {
				t.Errorf("invalid explicit configuration was treated as absent: %v", err)
			}
		})
	}
}

func TestFindLibreOfficeOnPath(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("PATH", dir)
	if runtime.GOOS == "windows" {
		t.Setenv("PATHEXT", ".COM;.EXE")
	}
	second := writeOfficeTestExecutable(t, dir, "secondary-office")
	path, err := findLibreOfficeOnPath("primary-office", "secondary-office")
	if path != second || err != nil {
		t.Fatalf("alternative candidate = (%q, %v), want %q", path, err, second)
	}
	first := writeOfficeTestExecutable(t, dir, "primary-office")
	path, err = findLibreOfficeOnPath("primary-office", "secondary-office")
	if path != first || err != nil {
		t.Errorf("preferred candidate = (%q, %v), want %q", path, err, first)
	}
	path, err = findLibreOfficeOnPath("missing-office")
	if path != "" || !errors.Is(err, ErrLibreofficeNotInstalled) {
		t.Errorf("missing candidate = (%q, %v)", path, err)
	}
}

func TestFindOfficeEngine(t *testing.T) {
	t.Parallel()
	lookupErr := errors.New("lookup denied")
	tests := []struct {
		name            string
		libreOfficePath string
		libreOfficeErr  error
		excelInstalled  bool
		excelErr        error
		want            officeEngine
		wantExcelProbe  bool
		wantErr         error
		wantAnyErr      bool
	}{
		{name: "LibreOffice has priority", libreOfficePath: "soffice", excelInstalled: true, want: officeEngine{libreOfficePath: "soffice"}},
		{name: "Excel when LibreOffice absent", libreOfficeErr: ErrLibreofficeNotInstalled, excelInstalled: true, want: officeEngine{useExcel: true}, wantExcelProbe: true},
		{name: "wrapped absence", libreOfficeErr: fmt.Errorf("probe: %w", ErrLibreofficeNotInstalled), excelInstalled: true, want: officeEngine{useExcel: true}, wantExcelProbe: true},
		{name: "neither installed", libreOfficeErr: ErrLibreofficeNotInstalled, wantExcelProbe: true, wantErr: ErrExcelOrLibreOfficeNotInstalled},
		{name: "LibreOffice lookup error is not absence", libreOfficeErr: lookupErr, wantErr: lookupErr},
		{name: "Excel lookup error is not absence", libreOfficeErr: ErrLibreofficeNotInstalled, excelErr: lookupErr, wantExcelProbe: true, wantErr: lookupErr},
		{name: "empty successful detection is rejected", wantAnyErr: true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			excelProbed := false
			engine, err := findOfficeEngine(
				func() (string, error) { return tt.libreOfficePath, tt.libreOfficeErr },
				func() (bool, error) {
					excelProbed = true
					return tt.excelInstalled, tt.excelErr
				},
			)
			if engine != tt.want || excelProbed != tt.wantExcelProbe {
				t.Errorf("engine = %+v, Excel probed = %v; want %+v, %v", engine, excelProbed, tt.want, tt.wantExcelProbe)
			}
			if tt.wantErr != nil {
				if !errors.Is(err, tt.wantErr) {
					t.Errorf("error = %v, want %v", err, tt.wantErr)
				}
			} else if (err != nil) != tt.wantAnyErr {
				t.Errorf("error = %v, want error = %v", err, tt.wantAnyErr)
			}
		})
	}
}

func writeOfficeTestExecutable(t *testing.T, dir, name string) string {
	t.Helper()
	if runtime.GOOS == "windows" {
		name += ".exe"
	}
	path := filepath.Join(dir, name)
	// Detection checks the path, not the binary format; this is never executed.
	if err := os.WriteFile(path, []byte("executable placeholder"), 0o700); err != nil {
		t.Fatal(err)
	}
	return path
}

func unsetOfficeTestEnv(t *testing.T, name string) {
	t.Helper()
	t.Setenv(name, "temporary value")
	if err := os.Unsetenv(name); err != nil {
		t.Fatal(err)
	}
}
