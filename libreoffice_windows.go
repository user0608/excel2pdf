//go:build windows
// +build windows

package excel2pdf

import (
	"log/slog"
	"os"
	"path/filepath"
	"strings"

	"golang.org/x/sys/windows/registry"
)

func lookPath(key registry.Key) (string, bool, error) {
	val, _, err := key.GetStringValue("Path")
	if err != nil {
		if err == registry.ErrNotExist {
			return "", false, nil
		}
		slog.Error("error reading key", "error", err)
		return "", false, err
	}
	return val, true, nil
}

func readLibreofficePath(names ...string) (string, error) {
	const prefix = `SOFTWARE\LibreOffice\LibreOffice`
	var keyPath = filepath.Join(append([]string{prefix}, names...)...)
	key, err := registry.OpenKey(registry.LOCAL_MACHINE, keyPath, registry.READ)
	if err != nil {
		if err == registry.ErrNotExist {
			return "", nil
		}
		slog.Error(`opening registry.LOCAL_MACHINE`, "erorr", err, "key_path", keyPath)
		return "", err
	}
	path, ok, err := lookPath(key)
	if err != nil {
		return "", err
	}
	if ok {
		return path, nil
	}
	subkeys, err := key.ReadSubKeyNames(-1)
	if err != nil {
		slog.Error("reading  registry.LOCAL_MACHINE sub keys", "error", err, "key_path", keyPath)
		return "", err
	}
	if len(subkeys) == 0 {
		return "", nil
	}
	for _, name := range subkeys {
		path, err := readLibreofficePath(strings.TrimPrefix(keyPath, prefix), name)
		if err != nil {
			return "", err
		}
		if path != "" {
			return path, nil
		}
	}
	return "", nil
}

func findLibreOfficeBinPath() (string, error) {
	value, ok := os.LookupEnv("LIBREOFFICE_PATH")
	if ok {
		return value, nil
	}
	return readLibreofficePath()
}
