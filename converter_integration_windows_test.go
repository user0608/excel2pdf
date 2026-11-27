//go:build integration

package excel2pdf

import (
	"bufio"
	"context"
	"os"
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
	requireIntegrationExcel(t)
	requireIntegrationWord(t)
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
	inputs := []string{writeIntegrationWorkbook(t, "hoja ñ.xlsx"), writeIntegrationWordDocument(t, "texto ñ.docx")}
	convert := func(t *testing.T, c *Converter) {
		t.Helper()
		for _, input := range inputs {
			pdf, err := c.ConvertToPDF(ctx, input)
			if err != nil {
				t.Fatal(err)
			}
			assertIntegrationPDF(t, pdf)
		}
	}
	t.Run("independent instances", func(t *testing.T) {
		for i, c := range instances {
			t.Run(strconv.Itoa(i), func(t *testing.T) {
				t.Parallel()
				convert(t, c)
			})
		}
	})
	convert(t, instances[0])
	if err := instances[0].Close(); err != nil {
		t.Fatal(err)
	}
	convert(t, instances[1])
}
