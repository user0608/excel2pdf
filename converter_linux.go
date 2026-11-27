package excel2pdf

import (
	"context"
	"os"
	"os/exec"
)

func removeSessionProfile(profile string) error { return os.RemoveAll(profile) }

func persistentSessionFactory() (func(context.Context) (conversionSession, error), error) {
	binary, err := LibreOfficePath()
	if err != nil {
		return nil, err
	}
	return libreOfficeSessionFactory(binary, "python3")
}

func startBridgeProcess(cmd *exec.Cmd) (func() error, error) {
	configureLibreOfficeCommand(cmd)
	if err := cmd.Start(); err != nil {
		return nil, err
	}
	return cmd.Cancel, nil
}
