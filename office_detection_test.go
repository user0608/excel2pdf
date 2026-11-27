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

func TestLibreOfficePath(t *testing.T) {
	executable := writeOfficeTestExecutable(t, t.TempDir(), "configured office")
	t.Setenv("LIBREOFFICE_PATH", executable)
	if path, err := LibreOfficePath(); path != executable || err != nil {
		t.Fatalf("LibreOfficePath = (%q, %v), want %q", path, err, executable)
	}
	// Detection must not cache an earlier environment or ignore invalid overrides.
	t.Setenv("LIBREOFFICE_PATH", filepath.Join(t.TempDir(), "missing.exe"))
	if path, err := LibreOfficePath(); path != "" || err == nil || errors.Is(err, ErrLibreofficeNotInstalled) {
		t.Fatalf("invalid override = (%q, %v)", path, err)
	}
}

func TestFindLibreOfficeInStandardDirectories(t *testing.T) {
	t.Parallel()
	for _, tt := range []struct {
		name        string
		variable    string
		executables []string
		want        string
	}{
		{name: "native 64 bit", variable: "ProgramW6432", executables: []string{"soffice.com"}, want: "soffice.com"},
		{name: "program files", variable: "ProgramFiles", executables: []string{"soffice.exe"}, want: "soffice.exe"},
		{name: "32 bit", variable: "ProgramFiles(x86)", executables: []string{"soffice.exe"}, want: "soffice.exe"},
		{name: "prefer console executable", variable: "ProgramFiles", executables: []string{"soffice.exe", "soffice.com"}, want: "soffice.com"},
		{name: "missing executable", variable: "ProgramFiles"},
		{name: "unset roots"},
	} {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			root := t.TempDir()
			dir := filepath.Join(root, "LibreOffice", "program")
			if err := os.MkdirAll(dir, 0o700); err != nil {
				t.Fatal(err)
			}
			for _, name := range tt.executables {
				if err := os.WriteFile(filepath.Join(dir, name), []byte("not executed"), 0o700); err != nil {
					t.Fatal(err)
				}
			}
			getenv := func(name string) string {
				if name == tt.variable {
					return root
				}
				return ""
			}
			path, err := findLibreOfficeInStandardDirectories(getenv)
			if tt.want == "" {
				if path != "" || !errors.Is(err, ErrLibreofficeNotInstalled) {
					t.Fatalf("absent installation = (%q, %v)", path, err)
				}
			} else if want := filepath.Join(dir, tt.want); path != want || err != nil {
				t.Fatalf("standard lookup = (%q, %v), want %q", path, err, want)
			}
		})
	}
}

func TestStandardDirectoriesOrderAndErrors(t *testing.T) {
	t.Parallel()
	first, second := t.TempDir(), t.TempDir()
	for _, root := range []string{first, second} {
		dir := filepath.Join(root, "LibreOffice", "program")
		if err := os.MkdirAll(dir, 0o700); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(filepath.Join(dir, "soffice.exe"), []byte("not executed"), 0o700); err != nil {
			t.Fatal(err)
		}
	}
	getenv := func(name string) string {
		if name == "ProgramW6432" {
			return first
		}
		return second
	}
	if path, err := findLibreOfficeInStandardDirectories(getenv); path != filepath.Join(first, "LibreOffice", "program", "soffice.exe") || err != nil {
		t.Fatalf("directory precedence = (%q, %v)", path, err)
	}
	if path, err := findLibreOfficeInStandardDirectories(func(string) string { return "relative" }); path != "" || !errors.Is(err, ErrLibreofficeNotInstalled) {
		t.Fatalf("relative roots = (%q, %v)", path, err)
	}
	// A directory posing as an executable is a lookup failure, not absence.
	if err := os.Mkdir(filepath.Join(first, "LibreOffice", "program", "soffice.com"), 0o700); err != nil {
		t.Fatal(err)
	}
	if err := os.Remove(filepath.Join(first, "LibreOffice", "program", "soffice.exe")); err != nil {
		t.Fatal(err)
	}
	if path, err := findLibreOfficeInStandardDirectories(getenv); path != filepath.Join(second, "LibreOffice", "program", "soffice.exe") || err != nil {
		t.Fatalf("valid alternative after invalid candidate = (%q, %v)", path, err)
	}
	if path, err := findLibreOfficeInStandardDirectories(func(string) string { return first }); path != "" || err == nil || errors.Is(err, ErrLibreofficeNotInstalled) {
		t.Fatalf("invalid candidate = (%q, %v)", path, err)
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
		{name: "Microsoft Office when LibreOffice absent", libreOfficeErr: ErrLibreofficeNotInstalled, excelInstalled: true, want: officeEngine{useMicrosoftOffice: true}, wantExcelProbe: true},
		{name: "wrapped absence", libreOfficeErr: fmt.Errorf("probe: %w", ErrLibreofficeNotInstalled), excelInstalled: true, want: officeEngine{useMicrosoftOffice: true}, wantExcelProbe: true},
		{name: "neither installed", libreOfficeErr: ErrLibreofficeNotInstalled, wantExcelProbe: true, wantErr: ErrOfficeNotInstalled},
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
