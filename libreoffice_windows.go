package excel2pdf

import (
	"errors"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
)

func findLibreOfficeBinPath() (string, error) {
	if path, configured, err := libreOfficeOverride(); configured {
		return path, err
	}
	path, err := findLibreOfficeOnPath("soffice.com", "soffice.exe", "libreoffice", "libreoffice24.8")
	if !errors.Is(err, ErrLibreofficeNotInstalled) {
		return path, err
	}
	return findLibreOfficeInRegistry(readOfficeRegistryString, readOfficeRegistrySubkeys)
}

func findLibreOfficeInRegistry(read officeRegistryStringReader, subkeys officeRegistrySubkeysReader) (string, error) {
	var lookupErrors []error
	for _, location := range officeRegistryLocations() {
		value, err := read(location.root, `SOFTWARE\LibreOffice\UNO\InstallPath`, "", location.view)
		if err != nil {
			lookupErrors = append(lookupErrors, err)
		} else if path, err := resolveLibreOfficeInstallation(value); err != nil {
			lookupErrors = append(lookupErrors, err)
		} else if path != "" {
			return path, nil
		}
		path, err := findLibreOfficeRegistryVersion(location, `SOFTWARE\LibreOffice\LibreOffice`, 0, read, subkeys)
		if err != nil {
			lookupErrors = append(lookupErrors, err)
		} else if path != "" {
			return path, nil
		}
	}
	if err := errors.Join(lookupErrors...); err != nil {
		return "", err
	}
	return "", ErrLibreofficeNotInstalled
}

func findLibreOfficeRegistryVersion(location officeRegistryLocation, keyPath string, depth int, read officeRegistryStringReader, subkeys officeRegistrySubkeysReader) (string, error) {
	if depth >= 16 {
		return "", fmt.Errorf("LibreOffice registry nesting limit reached at %q", keyPath)
	}
	var lookupErrors []error
	value, err := read(location.root, keyPath, "Path", location.view)
	if err != nil {
		lookupErrors = append(lookupErrors, err)
	} else if path, err := resolveLibreOfficeInstallation(value); err != nil {
		lookupErrors = append(lookupErrors, err)
	} else if path != "" {
		return path, nil
	}
	names, err := subkeys(location.root, keyPath, location.view)
	if err != nil {
		lookupErrors = append(lookupErrors, err)
	}
	for _, name := range names {
		path, err := findLibreOfficeRegistryVersion(location, filepath.Join(keyPath, name), depth+1, read, subkeys)
		if err != nil {
			lookupErrors = append(lookupErrors, err)
		} else if path != "" {
			return path, nil
		}
	}
	return "", errors.Join(lookupErrors...)
}

func resolveLibreOfficeInstallation(path string) (string, error) {
	if path == "" {
		return "", nil
	}
	info, err := os.Stat(path)
	if errors.Is(err, os.ErrNotExist) {
		return "", nil // Ignore stale registrations and keep looking.
	}
	if err != nil {
		return "", fmt.Errorf("inspect registered LibreOffice installation: %w", err)
	}
	if !info.IsDir() {
		return resolveOfficeExecutable(path)
	}
	for _, relative := range []string{"soffice.com", "soffice.exe", `program\soffice.com`, `program\soffice.exe`} {
		executable, err := resolveOfficeExecutable(filepath.Join(path, relative))
		if err == nil {
			return executable, nil
		}
		if !errors.Is(err, exec.ErrNotFound) && !errors.Is(err, os.ErrNotExist) {
			return "", fmt.Errorf("resolve registered LibreOffice executable: %w", err)
		}
	}
	return "", nil
}
