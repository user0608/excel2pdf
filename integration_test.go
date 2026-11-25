//go:build integration && (linux || windows)

package excel2pdf

import (
	"archive/zip"
	"bytes"
	"context"
	"errors"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"testing"
	"time"
)

func TestIntegrationLibreOfficeConversion(t *testing.T) {
	path := requireIntegrationLibreOffice(t)
	t.Setenv("LIBREOFFICE_PATH", path)
	t.Run("xlsx", func(t *testing.T) {
		assertIntegrationPublicConversion(t, writeIntegrationWorkbook(t, "report with spaces ñ.xlsx"))
	})
	t.Run("docx", func(t *testing.T) {
		assertIntegrationPublicConversion(t, writeIntegrationWordDocument(t, "document with spaces ñ.DOCX"))
	})
}

func assertIntegrationPublicConversion(t *testing.T, input string) {
	t.Helper()
	original, err := os.ReadFile(input)
	if err != nil {
		t.Fatal(err)
	}
	var previous string
	for attempt := 0; attempt < 2; attempt++ {
		ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
		pdf, err := ConvertToPDF(ctx, input)
		cancel()
		if err != nil {
			t.Fatal(err)
		}
		assertIntegrationPDF(t, pdf)
		if previous != "" && filepath.Dir(pdf) == filepath.Dir(previous) {
			t.Fatal("successive conversions reused an output directory")
		}
		previous = pdf
	}
	after, err := os.ReadFile(input)
	if err != nil || !bytes.Equal(after, original) {
		t.Fatalf("conversion changed the input workbook: %v", err)
	}
}

func TestIntegrationLibreOfficeSeparateProcesses(t *testing.T) {
	path := requireIntegrationLibreOffice(t)
	ctx, cancel := context.WithTimeout(context.Background(), 2*time.Minute)
	defer cancel()
	var commands []*exec.Cmd
	var outputs []*bytes.Buffer
	t.Cleanup(func() {
		cancel()
		for _, cmd := range commands {
			if cmd.ProcessState == nil {
				if err := cmd.Wait(); err != nil {
					t.Logf("wait for integration child during cleanup: %v", err)
				}
			}
		}
	})
	for child := 0; child < 2; child++ {
		input := writeIntegrationWorkbook(t, "same-name.xlsx")
		if child == 1 {
			input = writeIntegrationWordDocument(t, "same-name.docx")
		}
		cmd := exec.CommandContext(ctx, os.Args[0], "-test.run=^TestIntegrationLibreOfficeChildProcess$", "-test.timeout=90s")
		cmd.Env = append(os.Environ(),
			"EXCEL2PDF_LIBREOFFICE_INTEGRATION_CHILD=1",
			"EXCEL2PDF_INTEGRATION_INPUT="+input,
			"LIBREOFFICE_PATH="+path,
		)
		output := &bytes.Buffer{}
		cmd.Stdout = output
		cmd.Stderr = output
		if err := cmd.Start(); err != nil {
			t.Fatal(err)
		}
		commands = append(commands, cmd)
		outputs = append(outputs, output)
	}
	for index, cmd := range commands {
		if err := cmd.Wait(); err != nil {
			t.Errorf("integration child %d: %v\n%s", index, err, outputs[index].String())
		}
	}
}

func TestIntegrationLibreOfficeChildProcess(t *testing.T) {
	if os.Getenv("EXCEL2PDF_LIBREOFFICE_INTEGRATION_CHILD") != "1" {
		t.Skip("only runs in an integration child process")
	}
	ctx, cancel := context.WithTimeout(context.Background(), time.Minute)
	defer cancel()
	pdf, err := ConvertToPDF(ctx, os.Getenv("EXCEL2PDF_INTEGRATION_INPUT"))
	if err != nil {
		t.Fatal(err)
	}
	assertIntegrationPDF(t, pdf)
}

func requireIntegrationLibreOffice(t *testing.T) string {
	t.Helper()
	path, err := findLibreOfficeBinPath()
	if errors.Is(err, ErrLibreofficeNotInstalled) && os.Getenv("EXCEL2PDF_REQUIRE_LIBREOFFICE") != "1" {
		t.Skip("LibreOffice is not installed; set EXCEL2PDF_REQUIRE_LIBREOFFICE=1 to require it")
	}
	if err != nil {
		t.Fatal(err)
	}
	return path
}

func writeIntegrationWorkbook(t *testing.T, name string) string {
	t.Helper()
	files := []integrationXMLFile{
		{"[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`},
		{"_rels/.rels", `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`},
		{"xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Report" sheetId="1" r:id="rId1"/></sheets></workbook>`},
		{"xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`},
		{"xl/worksheets/sheet1.xml", `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>excel2pdf integration</t></is></c></row><row r="2"><c r="A2"><v>123.45</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait"/></worksheet>`},
	}
	return writeIntegrationOfficePackage(t, name, files)
}

type integrationXMLFile struct{ name, xml string }

func writeIntegrationWordDocument(t *testing.T, name string) string {
	t.Helper()
	return writeIntegrationOfficePackage(t, name, []integrationXMLFile{
		{"[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`},
		{"_rels/.rels", `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`},
		{"word/document.xml", `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>DOCX to PDF integration — documento de prueba</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`},
	})
}

func writeIntegrationOfficePackage(t *testing.T, name string, files []integrationXMLFile) string {
	t.Helper()
	var buffer bytes.Buffer
	archive := zip.NewWriter(&buffer)
	for _, file := range files {
		entry, err := archive.Create(file.name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := io.WriteString(entry, file.xml); err != nil {
			t.Fatal(err)
		}
	}
	if err := archive.Close(); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), name)
	if err := os.WriteFile(path, buffer.Bytes(), 0o600); err != nil {
		t.Fatal(err)
	}
	return path
}

func assertIntegrationPDF(t *testing.T, path string) {
	t.Helper()
	t.Cleanup(func() { removeExcelTestOutput(t, path) })
	file, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	header := make([]byte, 5)
	_, readErr := io.ReadFull(file, header)
	closeErr := file.Close()
	if err := errors.Join(readErr, closeErr); err != nil {
		t.Fatal(err)
	}
	if string(header) != "%PDF-" {
		t.Fatalf("output lacks a PDF header: %q", header)
	}
	entries, err := os.ReadDir(filepath.Dir(path))
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 1 || entries[0].Name() != filepath.Base(path) {
		t.Fatalf("auxiliary files survived conversion: %v", entries)
	}
}
