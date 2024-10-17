//go:build linux || darwin
// +build linux darwin

package excel2pdf

import (
	"errors"
	"log/slog"
	"os"
	"os/exec"
	"strings"
)

var ErrLibreofficeNotInstalled = errors.New("LibreOffice is not installed")

func findLibreOffice() (string, error) {
	cmd := exec.Command("which", "libreoffice")
	out, err := cmd.CombinedOutput()
	if err != nil {
		slog.Error("reading output `which libreoffice`", "error", err)
		return "", ErrLibreofficeNotInstalled
	}
	var libreofficePath = strings.TrimSpace(string(out))
	return libreofficePath, nil
}

func findlibreoffice24_8() (string, error) {
	cmd := exec.Command("which", "libreoffice24.8")
	out, err := cmd.CombinedOutput()
	if err != nil {
		slog.Error("reading output `which libreoffice`", "error", err)
		return "", ErrLibreofficeNotInstalled
	}
	var libreofficePath = strings.TrimSpace(string(out))
	return libreofficePath, nil
}

func findLibreOfficeBinPath() (string, error) {
	value, ok := os.LookupEnv("LIBREOFFICE_PATH")
	if ok {
		return value, nil
	}
	libreofficePath, err := findLibreOffice()
	if err != nil {
		return "", err
	}
	if libreofficePath == "" {
		libreofficePath, err = findlibreoffice24_8()
		if err != nil {
			return "", err
		}
	}
	if libreofficePath == "" {
		return libreofficePath, ErrLibreofficeNotInstalled
	}
	return libreofficePath, nil
}
