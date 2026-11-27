//go:build integration

package excel2pdf

import (
	"syscall"
	"testing"
)

func suspendBridge(t *testing.T, session *libreOfficeSession) {
	if err := syscall.Kill(-session.cmd.Process.Pid, syscall.SIGSTOP); err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { assertBridgeExited(t, session) })
}
