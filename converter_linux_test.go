package excel2pdf

import (
	"errors"
	"syscall"
	"testing"
)

func assertBridgeExited(t *testing.T, s *libreOfficeSession) {
	t.Helper()
	if err := syscall.Kill(-s.cmd.Process.Pid, 0); !errors.Is(err, syscall.ESRCH) {
		t.Fatalf("owned process group survived: %v", err)
	}
}
