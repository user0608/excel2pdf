package excel2pdf

import (
	"bufio"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"os/exec"
	"strings"
	"testing"
	"time"
)

func sessionTestCommand(mode string, profile *string) func(context.Context, string, ...string) *exec.Cmd {
	return func(ctx context.Context, _ string, args ...string) *exec.Cmd {
		*profile = args[len(args)-1]
		cmd := exec.CommandContext(ctx, os.Args[0], "-test.run=^TestSessionProcessHelper$")
		cmd.Env = append(os.Environ(), "EXCEL2PDF_SESSION_HELPER="+mode, "EXCEL2PDF_SESSION_PROFILE="+*profile)
		return cmd
	}
}

func TestLibreOfficeSessionProtocol(t *testing.T) {
	t.Parallel()
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	var profile string
	session, err := startLibreOfficeBridge(ctx, "unused", "unused", sessionTestCommand("success", &profile))
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() {
		if err := session.close(); err != nil {
			t.Error(err)
		}
	})
	s := session.(*libreOfficeSession)
	for i := 0; i < 2; i++ {
		pdf, err := convertDocumentFile(ctx, writeExcelTestInput(t, "name ñ.docx"), s.convert)
		if err != nil {
			t.Fatal(err)
		}
		t.Cleanup(func() { removeExcelTestOutput(t, pdf) })
	}
	if err := session.close(); err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(profile); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("profile survived close: %v", err)
	}
}

func TestLibreOfficeSessionStartupFailureAndCancellation(t *testing.T) {
	t.Parallel()
	for _, mode := range []string{"startup-error", "startup-hang", "malformed"} {
		t.Run(mode, func(t *testing.T) {
			t.Parallel()
			ctx, cancel := context.WithTimeout(context.Background(), 200*time.Millisecond)
			defer cancel()
			var profile string
			session, err := startLibreOfficeBridge(ctx, "unused", "unused", sessionTestCommand(mode, &profile))
			if session != nil || err == nil {
				t.Fatalf("startup failure = (%v, %v)", session, err)
			}
			if mode == "startup-hang" && !errors.Is(err, context.DeadlineExceeded) {
				t.Fatalf("startup timeout lost: %v", err)
			}
			if strings.Contains(err.Error(), profile) {
				t.Fatalf("profile leaked in error: %v", err)
			}
			if _, err := os.Stat(profile); !errors.Is(err, os.ErrNotExist) {
				t.Fatalf("profile survived startup failure: %v", err)
			}
		})
	}
}

func TestLibreOfficeSessionActiveCancellation(t *testing.T) {
	t.Parallel()
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	var profile string
	session, err := startLibreOfficeBridge(ctx, "unused", "unused", sessionTestCommand("conversion-hang", &profile))
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() {
		if err := session.close(); err != nil {
			t.Error(err)
		}
	})
	s := session.(*libreOfficeSession)
	operation, cancelOperation := context.WithTimeout(ctx, 20*time.Millisecond)
	defer cancelOperation()
	if _, err := convertDocumentFile(operation, writeExcelTestInput(t, "input.xlsx"), s.convert); !errors.Is(err, context.DeadlineExceeded) {
		t.Fatalf("active cancellation = %v", err)
	}
	assertBridgeExited(t, s)
}

func TestLibreOfficeSessionDiagnostics(t *testing.T) {
	t.Parallel()
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	var profile string
	session, err := startLibreOfficeBridge(ctx, "unused", "unused", sessionTestCommand("conversion-error", &profile))
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() {
		if err := session.close(); err != nil {
			t.Error(err)
		}
	})
	input := writeExcelTestInput(t, "private-name.docx")
	_, err = convertDocumentFile(ctx, input, session.convert)
	if err == nil || strings.Contains(err.Error(), input) || strings.Contains(err.Error(), profile) || strings.ContainsAny(err.Error(), "\n\r\x1b") {
		t.Fatalf("unsafe diagnostic: %v", err)
	}
}

func TestSessionProcessHelper(t *testing.T) {
	mode := os.Getenv("EXCEL2PDF_SESSION_HELPER")
	if mode == "" {
		return
	}
	write := func(value any) {
		if err := json.NewEncoder(os.Stdout).Encode(value); err != nil {
			t.Fatal(err)
		}
	}
	hang := func() {
		for {
			time.Sleep(time.Hour)
		}
	}
	switch mode {
	case "startup-error":
		write(bridgeResponse{Error: "missing PyUNO " + os.Getenv("EXCEL2PDF_SESSION_PROFILE")})
		os.Exit(0)
	case "startup-hang":
		hang()
	case "malformed":
		if _, err := fmt.Fprintln(os.Stdout, "invalid JSON"); err != nil {
			t.Fatal(err)
		}
		os.Exit(0)
	}
	write(bridgeResponse{OK: true})
	reader := bufio.NewScanner(os.Stdin)
	for reader.Scan() {
		var request map[string]string
		if err := json.Unmarshal(reader.Bytes(), &request); err != nil {
			t.Fatal(err)
		}
		if mode == "conversion-hang" {
			hang()
		}
		if mode == "conversion-error" {
			write(bridgeResponse{Error: request["input"] + "\n\r\x1b " + request["output"]})
			continue
		}
		if err := os.WriteFile(request["output"], []byte("%PDF-1.7\n"), 0o600); err != nil {
			t.Fatal(err)
		}
		write(bridgeResponse{OK: true})
	}
	if err := reader.Err(); err != nil {
		t.Fatal(err)
	}
	// Avoid the Go test runner's PASS text on our protocol stream.
	os.Exit(0)
}
