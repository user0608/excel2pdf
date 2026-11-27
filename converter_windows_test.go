package excel2pdf

import (
	"context"
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

func TestNewConverterWindowsInvalidOverride(t *testing.T) {
	t.Setenv("LIBREOFFICE_PATH", "")
	c, err := NewConverter(context.Background())
	if c != nil || err == nil {
		t.Fatalf("Windows persistent constructor = (%v, %v)", c, err)
	}
}
