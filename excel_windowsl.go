//go:build windows
// +build windows

package excel2pdf

import (
	"log/slog"
	"path/filepath"
	"strings"

	"golang.org/x/sys/windows/registry"
)

func existExcel(names ...string) (bool, error) {
	const prefix = `SOFTWARE\Microsoft\Office`
	var keyPath = filepath.Join(append([]string{prefix}, names...)...)
	key, err := registry.OpenKey(registry.LOCAL_MACHINE, keyPath, registry.READ)
	if err != nil {
		slog.Error(`opening registry.LOCAL_MACHINE`, "erorr", err, "key_path", keyPath)
		return false, err
	}
	subkeys, err := key.ReadSubKeyNames(-1)
	if err != nil {
		slog.Error("reading  registry.LOCAL_MACHINE sub keys", "error", err, "key_path", keyPath)
		return false, err
	}
	for _, name := range subkeys {
		switch name {
		case "Excel":
			return true, nil
		case "ClickToRun", "Common", "Access",
			"ClickToRunStore", "Outlook", "PowerPoint",
			"Project", "SDXHelper", "Visio", "Word":
			continue
		}

		ok, err := existExcel(strings.TrimPrefix(keyPath, prefix), name)
		if err != nil {
			return false, err
		}
		if ok {
			return ok, nil
		}
	}
	return false, nil
}

func isExcelInstalled() (bool, error) { return existExcel() }
