//go:build linux || darwin
// +build linux darwin

package excel2pdf

import (
	"errors"
	"log/slog"
	"os/exec"
	"strings"
)

var ErrLibreofficeNotInstalled = errors.New("LibreOffice is not installed")

func findLibreOfficeBinPath() (string, error) {
	cmd := exec.Command("which", "libreoffice")
	out, err := cmd.CombinedOutput()
	if err != nil {
		slog.Error("reading output `which libreoffice`", "error", err)
		return "", ErrLibreofficeNotInstalled
	}
	var libreofficePath = strings.TrimSpace(string(out))
	if libreofficePath == "" {
		return libreofficePath, ErrLibreofficeNotInstalled
	}
	return libreofficePath, nil
}
