//go:build integration

package excel2pdf

import (
	"bytes"
	"context"
	"errors"
	"os"
	"path/filepath"
	"sync"
	"testing"
	"time"
)

func requirePersistentIntegration(t *testing.T) {
	t.Helper()
	if os.Getenv("EXCEL2PDF_TEST_PERSISTENT") != "1" {
		t.Skip("set EXCEL2PDF_TEST_PERSISTENT=1 with Python 3/PyUNO to test persistent sessions")
	}
	t.Setenv("LIBREOFFICE_PATH", requireIntegrationLibreOffice(t))
}

func TestIntegrationPersistentConversion(t *testing.T) {
	requirePersistentIntegration(t)
	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()
	started := time.Now()
	c, err := NewConverter(ctx)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() {
		if err := c.Close(); err != nil {
			t.Error(err)
		}
	})
	t.Logf("persistent startup: %v", time.Since(started))
	session := c.session.(*libreOfficeSession)
	profile := session.profile
	for _, input := range []string{
		writeIntegrationWordDocument(t, "document with spaces ñ.docx"),
		writeIntegrationWorkbook(t, "report with spaces ñ.xlsx"),
	} {
		original, err := os.ReadFile(input)
		if err != nil {
			t.Fatal(err)
		}
		for i := 0; i < 3; i++ {
			started = time.Now()
			pdf, err := c.ConvertToPDF(ctx, input)
			if err != nil {
				t.Fatal(err)
			}
			t.Logf("persistent %s #%d: %v", filepath.Ext(input), i+1, time.Since(started))
			assertIntegrationPDF(t, pdf)
			if c.session != session {
				t.Fatal("successful conversion replaced session")
			}
		}
		after, err := os.ReadFile(input)
		if err != nil || !bytes.Equal(original, after) {
			t.Fatalf("source changed: %v", err)
		}
		// Baseline includes startup and shutdown, unlike the warm timings above.
		started = time.Now()
		pdf, err := ConvertToPDF(ctx, input)
		if err != nil {
			t.Fatal(err)
		}
		t.Logf("one-shot %s: %v", filepath.Ext(input), time.Since(started))
		assertIntegrationPDF(t, pdf)
	}
	if err := c.Close(); err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(profile); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("session profile survived Close: %v", err)
	}
	assertBridgeExited(t, session)
}

func TestIntegrationPersistentIndependentInstances(t *testing.T) {
	requirePersistentIntegration(t)
	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()
	var converters []*Converter
	for i := 0; i < 2; i++ {
		c, err := NewConverter(ctx)
		if err != nil {
			t.Fatal(err)
		}
		t.Cleanup(func() {
			if err := c.Close(); err != nil {
				t.Error(err)
			}
		})
		converters = append(converters, c)
	}
	if converters[0].session.(*libreOfficeSession).profile == converters[1].session.(*libreOfficeSession).profile {
		t.Fatal("instances share profile")
	}
	input := writeIntegrationWordDocument(t, "same-name.docx")
	var group sync.WaitGroup
	for _, c := range converters {
		group.Add(1)
		go func() {
			defer group.Done()
			pdf, err := c.ConvertToPDF(ctx, input)
			if err != nil {
				t.Error(err)
				return
			}
			assertIntegrationPDF(t, pdf)
		}()
	}
	group.Wait()
	if err := converters[0].Close(); err != nil {
		t.Fatal(err)
	}
	// Closing one instance must not terminate another's LibreOffice.
	pdf, err := converters[1].ConvertToPDF(ctx, input)
	if err != nil {
		t.Fatal(err)
	}
	assertIntegrationPDF(t, pdf)
}

func TestIntegrationPersistentFailureRecovery(t *testing.T) {
	requirePersistentIntegration(t)
	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()
	c, err := NewConverter(ctx)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() {
		if err := c.Close(); err != nil {
			t.Error(err)
		}
	})
	first := c.session.(*libreOfficeSession)
	input := writeIntegrationWordDocument(t, "valid.docx")
	// Exercise actual process death, not an input-validation error.
	if err := first.cmd.Process.Kill(); err != nil {
		t.Fatal(err)
	}
	<-first.exited
	if pdf, err := c.ConvertToPDF(ctx, input); pdf != "" || err == nil {
		t.Fatalf("dead session conversion = (%q, %v)", pdf, err)
	} else {
		t.Logf("dead session: %v", err)
	}
	if c.session != nil {
		t.Fatal("failed session retained")
	}
	if _, err := os.Stat(first.profile); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("failed session profile survived: %v", err)
	}
	pdf, err := c.ConvertToPDF(ctx, input)
	if err != nil {
		t.Fatal(err)
	}
	assertIntegrationPDF(t, pdf)
	if c.session == first {
		t.Fatal("failed session was not replaced")
	}
	checkHungBridge(t, c, ctx, input)
}

func checkHungBridge(t *testing.T, c *Converter, ctx context.Context, input string) {
	hung := c.session.(*libreOfficeSession)
	suspendBridge(t, hung)
	operation, stopOperation := context.WithTimeout(ctx, 50*time.Millisecond)
	defer stopOperation()
	if pdf, err := c.ConvertToPDF(operation, input); pdf != "" || !errors.Is(err, context.DeadlineExceeded) {
		t.Fatalf("hung session timeout = (%q, %v)", pdf, err)
	}
	if _, err := os.Stat(hung.profile); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("canceled session profile survived: %v", err)
	}
	pdf, err := c.ConvertToPDF(ctx, input)
	if err != nil {
		t.Fatal(err)
	}
	assertIntegrationPDF(t, pdf)
}
