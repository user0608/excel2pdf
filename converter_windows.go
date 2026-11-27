package excel2pdf

import (
	"context"
	"errors"
	"os"
	"path/filepath"
)

func persistentSessionFactory() (func(context.Context) (conversionSession, error), error) {
	if _, explicit := os.LookupEnv("LIBREOFFICE_PATH"); !explicit {
		start, err := microsoftSessionFactory()
		if !errors.Is(err, ErrOfficeNotInstalled) {
			return start, err
		}
	}
	return windowsLibreOfficeFactory()
}

func windowsLibreOfficeFactory() (func(context.Context) (conversionSession, error), error) {
	binary, err := LibreOfficePath()
	if errors.Is(err, ErrLibreofficeNotInstalled) {
		return nil, ErrOfficeNotInstalled
	}
	if err != nil {
		return nil, err
	}
	return libreOfficeSessionFactory(binary, filepath.Join(filepath.Dir(binary), "python.exe"))
}

func startWindowsLibreOffice(ctx context.Context) (conversionSession, error) {
	start, err := windowsLibreOfficeFactory()
	if err != nil {
		return nil, err
	}
	return start(ctx)
}
