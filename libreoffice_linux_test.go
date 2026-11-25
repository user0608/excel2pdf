package excel2pdf

import (
	"errors"
	"os"
	"testing"
)

func TestFindLibreOfficeBinPathLinux(t *testing.T) {
	tests := []struct {
		name       string
		candidates []string
		want       string
	}{
		{name: "standard name", candidates: []string{"libreoffice"}, want: "libreoffice"},
		{name: "versioned fallback", candidates: []string{"libreoffice24.8"}, want: "libreoffice24.8"},
		{name: "soffice fallback", candidates: []string{"soffice"}, want: "soffice"},
		{name: "standard name preferred", candidates: []string{"libreoffice24.8", "libreoffice", "soffice"}, want: "libreoffice"},
		{name: "no installation"},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			dir := t.TempDir()
			t.Setenv("PATH", dir)
			unsetOfficeTestEnv(t, "LIBREOFFICE_PATH")
			paths := make(map[string]string)
			for _, name := range tt.candidates {
				paths[name] = writeOfficeTestExecutable(t, dir, name)
			}
			path, err := findLibreOfficeBinPath()
			if tt.want == "" {
				if path != "" || !errors.Is(err, ErrLibreofficeNotInstalled) {
					t.Errorf("missing installation = (%q, %v)", path, err)
				}
			} else if path != paths[tt.want] || err != nil {
				t.Errorf("installation = (%q, %v), want %q", path, err, paths[tt.want])
			}
		})
	}
}

func TestFindLibreOfficeBinPathLinuxOverride(t *testing.T) {
	dir := t.TempDir()
	t.Setenv("PATH", dir)
	writeOfficeTestExecutable(t, dir, "libreoffice")
	override := writeOfficeTestExecutable(t, t.TempDir(), "custom-office")
	t.Setenv("LIBREOFFICE_PATH", override)
	if path, err := findLibreOfficeBinPath(); path != override || err != nil {
		t.Errorf("override = (%q, %v), want %q", path, err, override)
	}
	if err := os.Chmod(override, 0o600); err != nil {
		t.Fatal(err)
	}
	if path, err := findLibreOfficeBinPath(); path != "" || err == nil {
		t.Errorf("non-executable override fell back to PATH: (%q, %v)", path, err)
	}
	t.Setenv("LIBREOFFICE_PATH", "")
	if path, err := findLibreOfficeBinPath(); path != "" || err == nil {
		t.Errorf("empty override fell back to PATH: (%q, %v)", path, err)
	}
}
