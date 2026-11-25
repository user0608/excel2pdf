//go:build integration && windows

package excel2pdf

import (
	"context"
	"os"
	"path/filepath"
	"testing"
	"time"
)

func TestIntegrationWordConversion(t *testing.T) {
	requireIntegrationWord(t)
	input := writeIntegrationWordDocument(t, "document with spaces ñ.docx")
	original, err := os.ReadFile(input)
	if err != nil {
		t.Fatal(err)
	}
	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()
	pdf, err := convertDocumentFile(ctx, input, convertDocumentToPDFWithWord)
	if err != nil {
		t.Fatal(err)
	}
	assertIntegrationPDF(t, pdf)
	assertIntegrationOfficeInputReleased(t, input, original)
}

func TestIntegrationWordFailureRecovery(t *testing.T) {
	requireIntegrationWord(t)
	t.Run("open failure", func(t *testing.T) {
		ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
		defer cancel()
		if err := convertDocumentToPDFWithWord(ctx, filepath.Join(t.TempDir(), "missing.docx"), filepath.Join(t.TempDir(), "output.pdf")); err == nil {
			t.Fatal("opening a missing Word document unexpectedly succeeded")
		}
	})
	t.Run("export failure", func(t *testing.T) {
		input := writeIntegrationWordDocument(t, "valid.docx")
		original, err := os.ReadFile(input)
		if err != nil {
			t.Fatal(err)
		}
		ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
		defer cancel()
		output := filepath.Join(t.TempDir(), "missing-directory", "output.pdf")
		if err := convertDocumentToPDFWithWord(ctx, input, output); err == nil {
			t.Fatal("Word export to a missing directory unexpectedly succeeded")
		}
		assertIntegrationOfficeInputReleased(t, input, original)
	})
	input := writeIntegrationWordDocument(t, "recovery.docx")
	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()
	pdf, err := convertDocumentFile(ctx, input, convertDocumentToPDFWithWord)
	if err != nil {
		t.Fatalf("Word conversion after failures: %v", err)
	}
	assertIntegrationPDF(t, pdf)
}

func requireIntegrationWord(t *testing.T) {
	t.Helper()
	if os.Getenv("EXCEL2PDF_TEST_WORD") != "1" {
		t.Skip("set EXCEL2PDF_TEST_WORD=1 on a prepared Windows machine to launch Word")
	}
	installed, err := isWordInstalled()
	if err != nil {
		t.Fatal(err)
	}
	if !installed {
		t.Fatal("Word integration was requested, but no COM registration was found")
	}
}
