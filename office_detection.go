package excel2pdf

import (
	"errors"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
)

var ErrLibreofficeNotInstalled = errors.New("LibreOffice is not installed")
var ErrExcelOrLibreOfficeNotInstalled = errors.New("neither Microsoft Excel nor LibreOffice is installed on this system")

func libreOfficeOverride() (path string, configured bool, err error) {
	value, configured := os.LookupEnv("LIBREOFFICE_PATH")
	if !configured {
		return "", false, nil
	}
	if strings.TrimSpace(value) == "" {
		return "", true, errors.New("LIBREOFFICE_PATH must name an executable, not an empty value")
	}
	path, err = resolveOfficeExecutable(value)
	if err != nil {
		return "", true, fmt.Errorf("invalid LIBREOFFICE_PATH: %w", err)
	}
	return path, true, nil
}

func resolveOfficeExecutable(name string) (string, error) {
	path, err := exec.LookPath(name)
	if err != nil {
		return "", err
	}
	info, err := os.Stat(path)
	if err != nil {
		return "", err
	}
	if !info.Mode().IsRegular() {
		return "", fmt.Errorf("Office executable is not a regular file: %s", path)
	}
	return filepath.Abs(path)
}

func findLibreOfficeOnPath(names ...string) (string, error) {
	for _, name := range names {
		path, err := resolveOfficeExecutable(name)
		if err == nil {
			return path, nil
		}
		if !errors.Is(err, exec.ErrNotFound) && !errors.Is(err, os.ErrNotExist) {
			return "", fmt.Errorf("locate LibreOffice executable %q: %w", name, err)
		}
	}
	return "", ErrLibreofficeNotInstalled
}

type officeEngine struct {
	libreOfficePath string
	useExcel        bool
}

func findOfficeEngine(findLibreOffice func() (string, error), findExcel func() (bool, error)) (officeEngine, error) {
	path, err := findLibreOffice()
	if err == nil {
		if path == "" {
			return officeEngine{}, errors.New("LibreOffice detection returned an empty executable path")
		}
		return officeEngine{libreOfficePath: path}, nil
	}
	if !errors.Is(err, ErrLibreofficeNotInstalled) {
		return officeEngine{}, err
	}
	installed, err := findExcel()
	if err != nil {
		return officeEngine{}, err
	}
	if installed {
		return officeEngine{useExcel: true}, nil
	}
	return officeEngine{}, ErrExcelOrLibreOfficeNotInstalled
}
