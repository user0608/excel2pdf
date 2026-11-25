package excel2pdf

import (
	"errors"
	"fmt"
	"os"

	"golang.org/x/sys/windows/registry"
)

type officeRegistryLocation struct {
	root registry.Key
	view uint32
}

func officeRegistryLocations() []officeRegistryLocation {
	return []officeRegistryLocation{
		{root: registry.CURRENT_USER, view: registry.WOW64_64KEY},
		{root: registry.CURRENT_USER, view: registry.WOW64_32KEY},
		{root: registry.LOCAL_MACHINE, view: registry.WOW64_64KEY},
		{root: registry.LOCAL_MACHINE, view: registry.WOW64_32KEY},
	}
}

type officeRegistryStringReader func(registry.Key, string, string, uint32) (string, error)
type officeRegistrySubkeysReader func(registry.Key, string, uint32) ([]string, error)

func readOfficeRegistryString(root registry.Key, path, name string, view uint32) (value string, err error) {
	key, err := registry.OpenKey(root, path, registry.QUERY_VALUE|view)
	if errors.Is(err, os.ErrNotExist) {
		return "", nil
	}
	if err != nil {
		return "", fmt.Errorf("open registry key %q: %w", path, err)
	}
	defer func() {
		if closeErr := key.Close(); closeErr != nil {
			err = errors.Join(err, fmt.Errorf("close registry key %q: %w", path, closeErr))
		}
	}()
	value, valueType, err := key.GetStringValue(name)
	if errors.Is(err, os.ErrNotExist) {
		return "", nil
	}
	if err != nil {
		return "", fmt.Errorf("read registry value %q in %q: %w", name, path, err)
	}
	if valueType == registry.EXPAND_SZ {
		value, err = registry.ExpandString(value)
		if err != nil {
			return "", fmt.Errorf("expand registry value in %q: %w", path, err)
		}
	}
	return value, nil
}

func readOfficeRegistrySubkeys(root registry.Key, path string, view uint32) (names []string, err error) {
	key, err := registry.OpenKey(root, path, registry.ENUMERATE_SUB_KEYS|view)
	if errors.Is(err, os.ErrNotExist) {
		return nil, nil
	}
	if err != nil {
		return nil, fmt.Errorf("open registry key %q: %w", path, err)
	}
	defer func() {
		if closeErr := key.Close(); closeErr != nil {
			err = errors.Join(err, fmt.Errorf("close registry key %q: %w", path, closeErr))
		}
	}()
	names, err = key.ReadSubKeyNames(-1)
	if err != nil {
		return nil, fmt.Errorf("enumerate registry key %q: %w", path, err)
	}
	return names, nil
}
