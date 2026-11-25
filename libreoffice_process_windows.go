package excel2pdf

import "os/exec"

func configureLibreOfficeCommand(cmd *exec.Cmd) {
	// CommandContext kills only the launched process on Windows. Do not use
	// global process-name termination: other Office instances may be user-owned.
	// Strict descendant termination requires a Job Object or isolated worker.
}
