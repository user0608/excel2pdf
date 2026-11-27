//go:build integration

package excel2pdf

import (
	"bufio"
	"context"
	"os"
	"path/filepath"
	"strconv"
	"testing"
	"time"
)

func suspendBridge(t *testing.T, session *libreOfficeSession) {
	// Simulate a lost reply from the real session. Keep the writer open until
	// cleanup so only cancellation can unblock the protocol reader.
	t.Helper()
	reader, writer, err := os.Pipe()
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { writer.Close() })
	original := session.reader
	t.Cleanup(func() { original.Close() })
	session.reader = reader
	session.output = bufio.NewScanner(reader)
}

func TestIntegrationPersistentMicrosoftOffice(t *testing.T) {
	t.Run("Word", func(t *testing.T) {
		requireIntegrationWord(t)
		testPersistentApplication(t, writeIntegrationWordDocument(t, "texto ñ.docx"), wordDocument)
	})
	t.Run("Excel", func(t *testing.T) {
		requireIntegrationExcel(t)
		testPersistentApplication(t, writeIntegrationWorkbook(t, "hoja ñ.xlsx"), spreadsheetDocument)
	})
}

func testPersistentApplication(t *testing.T, input string, kind documentKind) {
	t.Helper()
	start, err := microsoftSessionFactory()
	if err != nil {
		t.Fatal(err)
	}
	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()
	instances := make([]*Converter, 2)
	for i := range instances {
		c, err := newConverter(ctx, start)
		if err != nil {
			t.Fatal(err)
		}
		t.Cleanup(func() {
			if err := c.Close(); err != nil {
				t.Error(err)
			}
		})
		instances[i] = c
	}
	convert := func(t *testing.T, c *Converter) {
		t.Helper()
		pdf, err := c.ConvertToPDF(ctx, input)
		if err != nil {
			t.Fatal(err)
		}
		assertIntegrationPDF(t, pdf)
	}
	t.Run("independent instances", func(t *testing.T) {
		for i, c := range instances {
			t.Run(strconv.Itoa(i), func(t *testing.T) {
				t.Parallel()
				convert(t, c)
			})
		}
	})
	first := instances[0].session.(*microsoftSession)
	app := first.apps[kind]
	convert(t, instances[0])
	if instances[0].session != first || app == nil || first.apps[kind] != app {
		t.Fatal("successful conversion did not reuse the Office application")
	}
	// Deny Office its source after Go's input checks, so failure occurs in
	// the real Documents.Open / Workbooks.Open rather than input validation.
	instances[0].session = missingOfficeInput{first}
	if pdf, err := instances[0].ConvertToPDF(ctx, input); pdf != "" || err == nil {
		t.Fatalf("missing Office source = (%q, %v)", pdf, err)
	}
	if instances[0].session != nil {
		t.Fatal("failed COM session retained")
	}
	convert(t, instances[0])
	if instances[0].session == first {
		t.Fatal("failed COM session reused")
	}
	if err := instances[0].Close(); err != nil {
		t.Fatal(err)
	}
	convert(t, instances[1])
}

type missingOfficeInput struct{ conversionSession }

func (s missingOfficeInput) convert(ctx context.Context, input, output string) error {
	return s.conversionSession.convert(ctx, filepath.Join(filepath.Dir(input), "missing"+filepath.Ext(input)), output)
}
