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
	"unicode"
)

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
		if diagnostic := sanitizeLibreOfficeDiagnostic(output.String(), inputPath, outputPath, profileDir); diagnostic != "" {
			return fmt.Errorf("run LibreOffice: %w (output: %s)", err, diagnostic)
		}
		return fmt.Errorf("run LibreOffice: %w", err)
	}
	return nil
}

func sanitizeLibreOfficeDiagnostic(text, inputPath, outputPath, profileDir string) string {
	var outputDir string
	if outputPath != "" {
		outputDir = filepath.Dir(outputPath)
	}
	for _, target := range []struct {
		path  string
		label string
	}{
		{inputPath, "[input]"},
		{outputPath, "[output]"},
		{profileDir, "[profile]"},
		{outputDir, "[temporary-directory]"},
	} {
		if target.path == "" {
			continue
		}
		if uri, err := libreOfficeProfileURI(target.path); err == nil {
			text = strings.ReplaceAll(text, uri, target.label)
		}
		text = strings.ReplaceAll(text, filepath.ToSlash(target.path), target.label)
		text = strings.ReplaceAll(text, target.path, target.label)
	}
	// Keep backend output on one line and remove terminal control characters.
	text = strings.Map(func(character rune) rune {
		switch character {
		case '\r', '\n', '\t':
			return ' '
		}
		if unicode.IsControl(character) {
			return -1
		}
		return character
	}, text)
	text = strings.TrimSpace(text)
	if len(text) > libreOfficeOutputLimit {
		text = text[:libreOfficeOutputLimit] + " [LibreOffice output truncated]"
	}
	return text
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
