package excel2pdf

import (
	"context"
	"errors"
	"path/filepath"
)

func persistentSessionFactory() (func(context.Context) (conversionSession, error), error) {
	binary, err := LibreOfficePath()
	if err == nil {
		return libreOfficeSessionFactory(binary, filepath.Join(filepath.Dir(binary), "python.exe"))
	}
	if !errors.Is(err, ErrLibreofficeNotInstalled) {
		return nil, err
	}
	return microsoftSessionFactory()
}
