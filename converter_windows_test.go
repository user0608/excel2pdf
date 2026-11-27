package excel2pdf

import (
	"context"
	"path/filepath"
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

func TestMissingWindowsApplicationReusesLibreOffice(t *testing.T) {
	for _, tc := range []struct {
		installed documentKind
		input     string
	}{
		{wordDocument, "input.xlsx"},
		{spreadsheetDocument, "input.docx"},
	} {
		t.Run(tc.input, func(t *testing.T) {
			fallback := &fakeConversionSession{}
			starts := 0
			s := &microsoftSession{
				installed: map[documentKind]bool{tc.installed: true},
				requests:  make(chan func()), done: make(chan error, 1),
				startFallback: func(context.Context) (conversionSession, error) {
					starts++
					return fallback, nil
				},
			}
			s.done <- nil
			for i := 0; i < 2; i++ {
				if err := s.convert(context.Background(), tc.input, filepath.Join(t.TempDir(), "out.pdf")); err != nil {
					t.Fatal(err)
				}
			}
			if err := s.close(); err != nil {
				t.Fatal(err)
			}
			if starts != 1 || fallback.closes.Load() != 1 {
				t.Fatal("expected a single owned fallback session")
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
