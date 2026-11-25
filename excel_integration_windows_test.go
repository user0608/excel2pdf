//go:build integration && windows

package excel2pdf

import (
	"bytes"
	"context"
	"os"
	"path/filepath"
	"testing"
	"time"
)

func TestIntegrationExcelConversion(t *testing.T) {
	requireIntegrationExcel(t)
	input := writeIntegrationWorkbook(t, "report with spaces ñ.xlsx")
	original, err := os.ReadFile(input)
	if err != nil {
		t.Fatal(err)
	}
	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()
	// Call the Excel backend explicitly, even when LibreOffice is installed.
	pdf, err := convertExcelFile(ctx, input, convertExcelToPDFWithExcel)
	if err != nil {
		t.Fatal(err)
	}
	assertIntegrationPDF(t, pdf)
	assertIntegrationExcelInputReleased(t, input, original)
}

func TestIntegrationExcelFailureRecovery(t *testing.T) {
	requireIntegrationExcel(t)
	t.Run("open failure", func(t *testing.T) {
		input := filepath.Join(t.TempDir(), "missing.xlsx")
		ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
		defer cancel()
		// Bypass public input validation deliberately to reach Workbooks.Open.
		output := filepath.Join(t.TempDir(), "output.pdf")
		if err := convertExcelToPDFWithExcel(ctx, input, output); err == nil {
			t.Fatal("opening a missing workbook unexpectedly succeeded")
		}
	})
	t.Run("export failure", func(t *testing.T) {
		input := writeIntegrationWorkbook(t, "valid.xlsx")
		original, err := os.ReadFile(input)
		if err != nil {
			t.Fatal(err)
		}
		ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
		defer cancel()
		// The parent directory is deliberately absent, so Excel cannot export.
		output := filepath.Join(t.TempDir(), "missing-directory", "output.pdf")
		if err := convertExcelToPDFWithExcel(ctx, input, output); err == nil {
			t.Fatal("export to a missing directory unexpectedly succeeded")
		}
		assertIntegrationExcelInputReleased(t, input, original)
	})
	// A new conversion must work after both failure paths.
	input := writeIntegrationWorkbook(t, "recovery.xlsx")
	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()
	pdf, err := convertExcelFile(ctx, input, convertExcelToPDFWithExcel)
	if err != nil {
		t.Fatalf("conversion after failure paths: %v", err)
	}
	assertIntegrationPDF(t, pdf)
}

func requireIntegrationExcel(t *testing.T) {
	t.Helper()
	if os.Getenv("EXCEL2PDF_TEST_EXCEL") != "1" {
		t.Skip("set EXCEL2PDF_TEST_EXCEL=1 on a prepared Windows machine to launch Excel")
	}
	installed, err := isExcelInstalled()
	if err != nil {
		t.Fatal(err)
	}
	if !installed {
		t.Fatal("Excel integration was requested, but no COM registration was found")
	}
}

func assertIntegrationExcelInputReleased(t *testing.T, input string, original []byte) {
	t.Helper()
	after, err := os.ReadFile(input)
	if err != nil || !bytes.Equal(after, original) {
		t.Fatalf("Excel modified the source workbook: %v", err)
	}
	// Windows won't allow this operation if Excel retains a conflicting file
	// handle. This tests workbook release, not the absence of every Excel PID.
	renamed := input + ".released"
	if err := os.Rename(input, renamed); err != nil {
		t.Fatalf("workbook remains locked after conversion: %v", err)
	}
	if err := os.Rename(renamed, input); err != nil {
		t.Fatal(err)
	}
}
