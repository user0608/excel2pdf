//go:build linux || darwin
// +build linux darwin

package excel2pdf

func isExcelInstalled() (bool, error) { return false, nil }
