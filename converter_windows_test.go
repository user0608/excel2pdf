package excel2pdf

import (
	"context"
	"errors"
	"testing"
)

func assertBridgeExited(t *testing.T, s *libreOfficeSession) {
	t.Helper()
	select {
	case <-s.exited:
	default:
		t.Fatal("owned bridge was not reaped")
	}
}

func TestMissingWindowsApplication(t *testing.T) {
	for _, tc := range []struct {
		installed documentKind
		input     string
	}{
		{wordDocument, "input.xlsx"},
		{spreadsheetDocument, "input.docx"},
	} {
		t.Run(tc.input, func(t *testing.T) {
			s := &microsoftSession{
				installed: map[documentKind]bool{tc.installed: true},
				requests:  make(chan func()), done: make(chan error, 1),
			}
			s.done <- nil
			for i := 0; i < 2; i++ {
				if err := s.convert(context.Background(), tc.input, "out.pdf"); !errors.Is(err, ErrOfficeNotInstalled) {
					t.Fatalf("missing application: %v", err)
				}
			}
			if err := s.close(); err != nil {
				t.Fatal(err)
			}
		})
	}
}

func TestNewConverterWindowsInvalidOverride(t *testing.T) {
	t.Setenv("LIBREOFFICE_PATH", "")
	c, err := NewConverter(context.Background())
	if c != nil || err == nil {
		t.Fatalf("Windows persistent constructor = (%v, %v)", c, err)
	}
}

func TestPersistentFactoryPrefersLibreOffice(t *testing.T) {
	dir := t.TempDir()
	binary := writeOfficeTestExecutable(t, dir, "soffice")
	writeOfficeTestExecutable(t, dir, "python")
	t.Setenv("LIBREOFFICE_PATH", binary)
	unsetOfficeTestEnv(t, "EXCEL2PDF_PYTHON")
	start, err := persistentSessionFactory()
	if start == nil || err != nil {
		t.Fatalf("LibreOffice factory = (%v, %v)", start != nil, err)
	}
	// An invalid Python configuration must fail instead of selecting COM.
	t.Setenv("EXCEL2PDF_PYTHON", dir)
	if start, err := persistentSessionFactory(); start != nil || err == nil {
		t.Fatal("invalid LibreOffice Python configuration did not fail")
	}
}
