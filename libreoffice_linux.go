package excel2pdf

func findLibreOfficeBinPath() (string, error) {
	if path, configured, err := libreOfficeOverride(); configured {
		return path, err
	}
	return findLibreOfficeOnPath("libreoffice", "libreoffice24.8", "soffice")
}
