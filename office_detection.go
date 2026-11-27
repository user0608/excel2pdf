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
var ErrOfficeNotInstalled = errors.New("neither LibreOffice nor the required Microsoft Office application is installed")

// LibreOfficePath returns the absolute executable path that conversion would use
// with the current environment. It does not start LibreOffice or check its version.
// A later conversion detects the path again; the result is not cached.
// ErrLibreofficeNotInstalled means no installation was found, not a lookup failure.
func LibreOfficePath() (string, error) {
	return findLibreOfficeBinPath()
}

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
		return "", fmt.Errorf("resolved Office executable is not a regular file: %s", path)
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
	return "", fmt.Errorf("%w: none of %s found in PATH", ErrLibreofficeNotInstalled, strings.Join(names, ", "))
}

func findLibreOfficeInStandardDirectories(getenv func(string) string) (string, error) {
	var lookupErrors []error
	seen := make(map[string]bool)
	for _, name := range []string{"ProgramW6432", "ProgramFiles", "ProgramFiles(x86)"} {
		root := getenv(name)
		// Never turn a missing or relative environment value into a CWD lookup.
		if !filepath.IsAbs(root) {
			continue
		}
		dir := filepath.Join(root, "LibreOffice", "program")
		key := strings.ToLower(dir)
		if seen[key] {
			continue
		}
		seen[key] = true
		for _, executable := range []string{"soffice.com", "soffice.exe"} {
			candidate := filepath.Join(dir, executable)
			// Stat the exact filename first: Windows LookPath may hide access
			// errors or try appended extensions such as soffice.com.exe.
			info, err := os.Stat(candidate)
			if errors.Is(err, os.ErrNotExist) {
				continue
			}
			if err == nil && !info.Mode().IsRegular() {
				err = fmt.Errorf("not a regular executable file: %s", candidate)
			}
			if err != nil {
				lookupErrors = append(lookupErrors, fmt.Errorf("inspect LibreOffice candidate %q from %s: %w", candidate, name, err))
				continue
			}
			path, err := resolveOfficeExecutable(candidate)
			if err == nil {
				return path, nil
			}
			if !errors.Is(err, exec.ErrNotFound) && !errors.Is(err, os.ErrNotExist) {
				lookupErrors = append(lookupErrors, fmt.Errorf("inspect LibreOffice candidate %q from %s: %w", candidate, name, err))
			}
		}
	}
	if err := errors.Join(lookupErrors...); err != nil {
		return "", err
	}
	return "", fmt.Errorf("%w: no executable in standard Program Files directories", ErrLibreofficeNotInstalled)
}

type officeEngine struct {
	libreOfficePath    string
	useMicrosoftOffice bool
}

func findOfficeEngine(findLibreOffice func() (string, error), findMicrosoftOffice func() (bool, error)) (officeEngine, error) {
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
	installed, err := findMicrosoftOffice()
	if err != nil {
		return officeEngine{}, err
	}
	if installed {
		return officeEngine{useMicrosoftOffice: true}, nil
	}
	return officeEngine{}, ErrOfficeNotInstalled
}
