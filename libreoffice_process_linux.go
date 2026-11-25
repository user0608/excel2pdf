package excel2pdf

import (
	"errors"
	"os"
	"os/exec"
	"syscall"
)

func configureLibreOfficeCommand(cmd *exec.Cmd) {
	// soffice can launch soffice.bin. Give this conversion its own process
	// group so cancellation also reaches descendants that remain in the group.
	cmd.SysProcAttr = &syscall.SysProcAttr{Setpgid: true}
	cmd.Cancel = func() error {
		err := syscall.Kill(-cmd.Process.Pid, syscall.SIGKILL)
		if errors.Is(err, syscall.ESRCH) {
			return os.ErrProcessDone
		}
		return err
	}
}
