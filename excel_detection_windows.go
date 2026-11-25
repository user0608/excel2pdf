package excel2pdf

import (
	"errors"
	"fmt"

	"github.com/go-ole/go-ole"
	"golang.org/x/sys/windows/registry"
)

func isExcelInstalled() (bool, error) {
	return findExcelRegistration(readOfficeRegistryString)
}

func findExcelRegistration(read officeRegistryStringReader) (bool, error) {
	var lookupErrors []error
	for _, location := range officeRegistryLocations() {
		classID, err := read(location.root, `SOFTWARE\Classes\Excel.Application\CLSID`, "", location.view)
		if err != nil {
			lookupErrors = append(lookupErrors, err)
			continue
		}
		if classID == "" {
			continue
		}
		guid := ole.NewGUID(classID)
		if guid == nil {
			lookupErrors = append(lookupErrors, fmt.Errorf("invalid Excel COM class ID %q", classID))
			continue
		}
		// COM combines per-user and machine class registrations. A per-user
		// ProgID can refer to a server registered for the entire machine.
		for _, root := range []registry.Key{registry.CURRENT_USER, registry.LOCAL_MACHINE} {
			server, err := read(root, `SOFTWARE\Classes\CLSID\`+guid.String()+`\LocalServer32`, "", location.view)
			if err != nil {
				lookupErrors = append(lookupErrors, err)
				continue
			}
			if server != "" {
				return true, nil
			}
		}
	}
	return false, errors.Join(lookupErrors...)
}
