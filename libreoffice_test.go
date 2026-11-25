package excel2pdf

import (
	"context"
	"errors"
	"fmt"
	"net/url"
	"os"
	"os/exec"
	"path/filepath"
	"reflect"
	"runtime"
	"strings"
	"testing"
	"time"
)

func TestRunLibreOfficeConversion(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name    string
		mode    string
		wantErr bool
	}{
		{name: "successful process", mode: "success"},
		{name: "non-zero exit", mode: "fail", wantErr: true},
		{name: "zero exit without PDF", mode: "missing", wantErr: true},
		{name: "excessive diagnostics", mode: "large", wantErr: true},
		{name: "cannot start executable", mode: "start-failure", wantErr: true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			input := writeExcelTestInput(t, "report with spaces ñ.xlsx")
			var profilePath string
			command := libreOfficeHelperCommand(t, tt.mode, &profilePath)
			pdf, err := convertDocumentFile(context.Background(), input, func(ctx context.Context, inputPath, outputPath string) error {
				return runLibreOfficeConversion(ctx, "fake-libreoffice", inputPath, outputPath, command)
			})
			if (err != nil) != tt.wantErr {
				t.Fatalf("conversion error = %v, want error = %v", err, tt.wantErr)
			}
			if profilePath == "" {
				t.Fatal("command did not receive an isolated profile")
			}
			if _, err := os.Stat(profilePath); !errors.Is(err, os.ErrNotExist) {
				t.Errorf("temporary profile was not removed: %v", err)
			}
			if !tt.wantErr {
				t.Cleanup(func() { removeExcelTestOutput(t, pdf) })
				entries, err := os.ReadDir(filepath.Dir(pdf))
				if err != nil {
					t.Fatal(err)
				}
				if len(entries) != 1 || entries[0].Name() != filepath.Base(pdf) {
					t.Errorf("successful conversion retained auxiliary files: %v", entries)
				}
				return
			}
			if pdf != "" {
				t.Errorf("failed conversion returned path %q", pdf)
			}
			if _, err := os.Stat(filepath.Dir(profilePath)); !errors.Is(err, os.ErrNotExist) {
				t.Errorf("failed conversion retained output directory: %v", err)
			}
			if tt.mode == "fail" || tt.mode == "missing" {
				for _, diagnostic := range []string{"helper stdout", "helper stderr"} {
					if !strings.Contains(err.Error(), diagnostic) {
						t.Errorf("error omitted %q: %v", diagnostic, err)
					}
				}
			}
			if tt.mode == "fail" {
				if strings.Contains(err.Error(), input) || strings.Contains(err.Error(), profilePath) {
					t.Errorf("process diagnostics exposed known paths: %v", err)
				}
				if !strings.Contains(err.Error(), "[input]") || !strings.Contains(err.Error(), "[profile]") {
					t.Errorf("process diagnostics were not sanitized: %v", err)
				}
			}
			if tt.mode == "fail" || tt.mode == "large" {
				var exitErr *exec.ExitError
				if !errors.As(err, &exitErr) || exitErr.ExitCode() != 7 {
					t.Errorf("process exit error was not preserved: %v", err)
				}
			}
			if tt.mode == "large" {
				if !strings.Contains(err.Error(), "output truncated") || len(err.Error()) > libreOfficeOutputLimit+1024 {
					t.Errorf("diagnostics were not bounded: length %d", len(err.Error()))
				}
			}
		})
	}
}

func TestRunLibreOfficeConversionCancellation(t *testing.T) {
	t.Parallel()
	for _, mode := range []string{"cancel", "deadline"} {
		t.Run(mode, func(t *testing.T) {
			t.Parallel()
			limit, stop := context.WithTimeout(context.Background(), 10*time.Second)
			defer stop()
			ctx, cancel := context.WithCancel(limit)
			wantErr := context.Canceled
			if mode == "deadline" {
				cancel()
				ctx, cancel = context.WithTimeout(limit, 3*time.Second)
				wantErr = context.DeadlineExceeded
			}
			defer cancel()
			input := writeExcelTestInput(t, "cancel.xlsx")
			type processInfo struct {
				profile string
				cmd     *exec.Cmd
			}
			type conversionResult struct {
				pdf string
				err error
			}
			started := make(chan processInfo, 1)
			results := make(chan conversionResult, 1)
			done := make(chan struct{})
			go func() {
				defer close(done)
				pdf, err := convertDocumentFile(ctx, input, func(ctx context.Context, input, output string) error {
					var profile string
					helper := libreOfficeHelperCommand(t, "hang", &profile)
					command := func(ctx context.Context, binary string, args ...string) *exec.Cmd {
						cmd := helper(ctx, binary, args...)
						started <- processInfo{profile: profile, cmd: cmd}
						return cmd
					}
					return runLibreOfficeConversion(ctx, "fake-libreoffice", input, output, command)
				})
				results <- conversionResult{pdf: pdf, err: err}
			}()
			t.Cleanup(func() {
				cancel()
				select {
				case <-done:
				case <-time.After(5 * time.Second):
					t.Error("conversion goroutine did not stop after cancellation")
				}
			})
			var process processInfo
			select {
			case process = <-started:
			case <-limit.Done():
				t.Fatal("command was not constructed")
			}
			if err := waitForLibreOfficeTestFile(ctx, filepath.Join(process.profile, "profile-used")); err != nil {
				t.Fatalf("helper did not start before cancellation: %v", err)
			}
			if mode == "cancel" {
				cancel()
			}
			var result conversionResult
			select {
			case result = <-results:
			case <-limit.Done():
				t.Fatal("conversion did not stop")
			}
			if result.pdf != "" || !errors.Is(result.err, wantErr) {
				t.Errorf("conversion = (%q, %v), want %v", result.pdf, result.err, wantErr)
			}
			if process.cmd.ProcessState == nil {
				t.Error("canceled process was not waited for")
			}
			if process.cmd.WaitDelay <= 0 {
				t.Error("output pipe wait is unbounded")
			}
			if _, err := os.Stat(filepath.Dir(process.profile)); !errors.Is(err, os.ErrNotExist) {
				t.Errorf("canceled conversion directory was not removed: %v", err)
			}
		})
	}
}

func waitForLibreOfficeTestFile(ctx context.Context, path string) error {
	ticker := time.NewTicker(10 * time.Millisecond)
	defer ticker.Stop()
	for {
		if _, err := os.Stat(path); err == nil {
			return nil
		} else if !errors.Is(err, os.ErrNotExist) {
			return err
		}
		select {
		case <-ctx.Done():
			return ctx.Err()
		case <-ticker.C:
		}
	}
}

func TestLibreOfficeProfileURI(t *testing.T) {
	t.Parallel()
	path := filepath.Join(t.TempDir(), "profile space ñ # %")
	uri, err := libreOfficeProfileURI(path)
	if err != nil {
		t.Fatal(err)
	}
	parsed, err := url.Parse(uri)
	if err != nil {
		t.Fatal(err)
	}
	if parsed.Scheme != "file" || parsed.RawQuery != "" || parsed.Fragment != "" {
		t.Fatalf("invalid profile URI: %q", uri)
	}
	if got := libreOfficeTestURIPath(parsed); got != path {
		t.Errorf("URI round-trip = %q, want %q", got, path)
	}
	for _, escaped := range []string{"%20", "%C3%B1", "%23", "%25"} {
		if !strings.Contains(uri, escaped) {
			t.Errorf("profile URI does not escape %s: %q", escaped, uri)
		}
	}
}

func TestLibreOfficeOutput(t *testing.T) {
	t.Parallel()
	var output libreOfficeOutput
	for _, data := range []string{"first", strings.Repeat("x", libreOfficeOutputLimit), "last"} {
		n, err := output.Write([]byte(data))
		if n != len(data) || err != nil {
			t.Fatalf("Write = (%d, %v), want (%d, nil)", n, err, len(data))
		}
	}
	if len(output.data) != libreOfficeOutputLimit || !output.truncated {
		t.Errorf("unexpected retained diagnostics: length %d, truncated %v", len(output.data), output.truncated)
	}
	if !strings.HasPrefix(output.String(), "first") || !strings.HasSuffix(output.String(), "[LibreOffice output truncated]") {
		t.Error("diagnostic prefix or truncation marker missing")
	}
}

func TestSanitizeLibreOfficeDiagnostic(t *testing.T) {
	t.Parallel()
	dir := filepath.Join(t.TempDir(), "private data ñ # %")
	input := filepath.Join(dir, "private.xlsx")
	output := filepath.Join(dir, "conversion", "private.pdf")
	profile := filepath.Join(dir, "conversion", ".profile")
	profileURI, err := libreOfficeProfileURI(profile)
	if err != nil {
		t.Fatal(err)
	}
	inputURI, err := libreOfficeProfileURI(input)
	if err != nil {
		t.Fatal(err)
	}
	raw := "convert " + input + " and " + inputURI + "\noutput " + output + "\rprofile " + profileURI + "\t" + profile + " directory " + filepath.Dir(output) + "\x1b\x00"
	diagnostic := sanitizeLibreOfficeDiagnostic(raw, input, output, profile)
	for _, sensitive := range []string{input, output, profile, profileURI, inputURI, filepath.Dir(output), "\n", "\r", "\t", "\x1b", "\x00"} {
		if strings.Contains(diagnostic, sensitive) {
			t.Errorf("diagnostic retained sensitive path or control %q: %q", sensitive, diagnostic)
		}
	}
	for _, useful := range []string{"convert", "[input]", "[output]", "[profile]", "[temporary-directory]"} {
		if !strings.Contains(diagnostic, useful) {
			t.Errorf("diagnostic omitted %q: %q", useful, diagnostic)
		}
	}
	large := sanitizeLibreOfficeDiagnostic(strings.Repeat("x", libreOfficeOutputLimit*2), input, output, profile)
	if len(large) > libreOfficeOutputLimit+64 || !strings.Contains(large, "output truncated") {
		t.Errorf("sanitized diagnostics are unbounded: %d bytes", len(large))
	}
}

func libreOfficeHelperCommand(t *testing.T, mode string, profilePath *string) func(context.Context, string, ...string) *exec.Cmd {
	t.Helper()
	return func(ctx context.Context, binary string, args ...string) *exec.Cmd {
		t.Helper()
		if binary != "fake-libreoffice" || len(args) != 7 {
			t.Fatalf("unexpected command: %q %q", binary, args)
		}
		if !strings.HasPrefix(args[0], "-env:UserInstallation=") || !reflect.DeepEqual(args[1:4], []string{"--headless", "--convert-to", "pdf"}) || args[5] != "--outdir" {
			t.Fatalf("unexpected LibreOffice arguments: %q", args)
		}
		uri, err := url.Parse(strings.TrimPrefix(args[0], "-env:UserInstallation="))
		if err != nil || uri.Scheme != "file" {
			t.Fatalf("invalid profile URI: %v, %v", uri, err)
		}
		*profilePath = libreOfficeTestURIPath(uri)
		if filepath.Dir(*profilePath) != args[6] {
			t.Fatalf("profile is outside this conversion's output directory: %q", *profilePath)
		}
		if mode == "start-failure" {
			return exec.CommandContext(ctx, filepath.Join(t.TempDir(), "nonexistent-executable"))
		}
		childArgs := append([]string{"-test.run=^TestLibreOfficeHelperProcess$", "--"}, args...)
		cmd := exec.CommandContext(ctx, os.Args[0], childArgs...)
		cmd.Env = append(os.Environ(), "EXCEL2PDF_LIBREOFFICE_HELPER="+mode)
		return cmd
	}
}

func libreOfficeTestURIPath(uri *url.URL) string {
	path := uri.Path
	if uri.Host != "" {
		path = "//" + uri.Host + path
	} else if runtime.GOOS == "windows" {
		path = strings.TrimPrefix(path, "/")
	}
	return filepath.FromSlash(path)
}

// TestLibreOfficeHelperProcess only performs work in the explicitly selected
// child test process. It simulates the CLI contract, not LibreOffice rendering.
func TestLibreOfficeHelperProcess(t *testing.T) {
	mode := os.Getenv("EXCEL2PDF_LIBREOFFICE_HELPER")
	if mode == "" {
		return
	}
	separator := -1
	for i, arg := range os.Args {
		if arg == "--" {
			separator = i
			break
		}
	}
	if separator < 0 || len(os.Args[separator+1:]) != 7 {
		t.Fatal("missing LibreOffice arguments")
	}
	args := os.Args[separator+1:]
	uri, err := url.Parse(strings.TrimPrefix(args[0], "-env:UserInstallation="))
	if err != nil {
		t.Fatal(err)
	}
	profile := libreOfficeTestURIPath(uri)
	if err := os.WriteFile(filepath.Join(profile, "profile-used"), []byte("isolated"), 0o600); err != nil {
		t.Fatal(err)
	}
	fmt.Fprintln(os.Stdout, "helper stdout")
	fmt.Fprintln(os.Stderr, "helper stderr")
	fmt.Fprintln(os.Stdout, "input", args[4], "profile", uri.String())
	if mode == "hang" {
		for {
			time.Sleep(time.Hour)
		}
	}
	if mode == "large" {
		fmt.Fprint(os.Stdout, strings.Repeat("x", libreOfficeOutputLimit*2))
		os.Exit(7)
	}
	if mode == "missing" {
		return
	}
	input := args[4]
	name := strings.TrimSuffix(filepath.Base(input), filepath.Ext(input)) + ".pdf"
	if err := os.WriteFile(filepath.Join(args[6], name), []byte("%PDF-1.7\n"), 0o600); err != nil {
		t.Fatal(err)
	}
	if mode == "fail" {
		os.Exit(7)
	}
}
