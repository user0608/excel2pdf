//go:build linux || darwin
// +build linux darwin

package excel2pdf

import (
	"errors"
	"log/slog"
	"os/exec"
)

var ErrLibreofficeNotInstalled = errors.New("LibreOffice is not installed")

func findLibreOfficeBinPath() (string, error) {
	cmd := exec.Command("which", "libreoffice")

	if err := cmd.Run(); err != nil {
		slog.Error("running which libreoffice", "error", err)
		return "", ErrLibreofficeNotInstalled
	}
	out, err := cmd.Output()
	if err != nil {
		slog.Error("reading output `which libreoffice`", "error", err)
		return "", ErrLibreofficeNotInstalled
	}
	var libreofficePath = string(out)
	if libreofficePath == "" {
		return libreofficePath, ErrLibreofficeNotInstalled
	}
	return libreofficePath, nil
}
