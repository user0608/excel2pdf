package excel2pdf

import (
	"context"
	"errors"
	"fmt"
	"net/url"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"time"
)

func convertExcelToPDFWithLibreOffice(ctx context.Context, excelFilePath, pdfFilePath string) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	libreOfficePath, err := findLibreOfficeBinPath()
	if err != nil {
		return err
	}
	return runLibreOfficeConversion(ctx, libreOfficePath, excelFilePath, pdfFilePath, exec.CommandContext)
}

func runLibreOfficeConversion(ctx context.Context, binaryPath, inputPath, outputPath string, command func(context.Context, string, ...string) *exec.Cmd) (err error) {
	if err := ctx.Err(); err != nil {
		return err
	}
	profileDir, err := os.MkdirTemp(filepath.Dir(outputPath), ".libreoffice-profile-*")
	if err != nil {
		return fmt.Errorf("create LibreOffice temporary profile: %w", err)
	}
	defer func() {
		if cleanupErr := os.RemoveAll(profileDir); cleanupErr != nil {
			err = errors.Join(err, fmt.Errorf("remove LibreOffice temporary profile: %w", cleanupErr))
		}
	}()
	profileURI, err := libreOfficeProfileURI(profileDir)
	if err != nil {
		return fmt.Errorf("resolve LibreOffice temporary profile: %w", err)
	}
	cmd := command(
		ctx,
		binaryPath,
		"-env:UserInstallation="+profileURI,
		"--headless",
		"--convert-to",
		"pdf", inputPath,
		"--outdir", filepath.Dir(outputPath),
	)
	configureLibreOfficeCommand(cmd)
	// Bound the wait for inherited output pipes after process exit/cancellation.
	cmd.WaitDelay = 2 * time.Second

	var output libreOfficeOutput
	cmd.Stdout = &output
	cmd.Stderr = &output
	err = cmd.Run()
	if ctxErr := ctx.Err(); ctxErr != nil {
		err = errors.Join(err, ctxErr)
	}
	if err == nil {
		err = validateGeneratedPDF(outputPath)
	}
	if err != nil {
		if diagnostic := strings.TrimSpace(output.String()); diagnostic != "" {
			return fmt.Errorf("run LibreOffice: %w (output: %s)", err, diagnostic)
		}
		return fmt.Errorf("run LibreOffice: %w", err)
	}
	return nil
}

func libreOfficeProfileURI(path string) (string, error) {
	absPath, err := filepath.Abs(path)
	if err != nil {
		return "", err
	}
	path = filepath.ToSlash(absPath)
	uri := url.URL{Scheme: "file"}
	if strings.HasPrefix(path, "//") {
		// Windows UNC paths encode the server as the URI authority.
		uri.Host, path, _ = strings.Cut(strings.TrimPrefix(path, "//"), "/")
		path = "/" + path
	} else if !strings.HasPrefix(path, "/") {
		// Windows drive paths must be file:///C:/..., not file://C:/....
		path = "/" + path
	}
	uri.Path = path
	return uri.String(), nil
}

// Bound retained diagnostics while continuing to drain the child's output.
const libreOfficeOutputLimit = 64 * 1024

type libreOfficeOutput struct {
	data      []byte
	truncated bool
}

func (output *libreOfficeOutput) Write(data []byte) (int, error) {
	length := len(data)
	remaining := libreOfficeOutputLimit - len(output.data)
	if len(data) > remaining {
		data = data[:remaining]
		output.truncated = true
	}
	output.data = append(output.data, data...)
	return length, nil
}

func (output *libreOfficeOutput) String() string {
	text := string(output.data)
	if output.truncated {
		text += "\n[LibreOffice output truncated]"
	}
	return text
}
