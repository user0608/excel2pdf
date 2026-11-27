package excel2pdf

import (
	"bufio"
	"context"
	_ "embed"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"os/exec"
	"time"
)

//go:embed libreoffice_bridge.py
var libreOfficeBridge string

func libreOfficeSessionFactory(binary, defaultPython string) (func(context.Context) (conversionSession, error), error) {
	python := os.Getenv("EXCEL2PDF_PYTHON")
	if python == "" {
		python = defaultPython
	}
	python, err := resolveOfficeExecutable(python)
	if err != nil {
		return nil, fmt.Errorf("locate Python 3 for persistent LibreOffice (EXCEL2PDF_PYTHON): %w", err)
	}
	return func(ctx context.Context) (conversionSession, error) {
		return startLibreOfficeSession(ctx, binary, python)
	}, nil
}

type libreOfficeSession struct {
	cmd      *exec.Cmd
	stop     context.CancelFunc
	input    io.WriteCloser
	output   *bufio.Scanner
	reader   *os.File
	exited   chan struct{}
	waitErr  error // published by exited
	profile  string
	logs     libreOfficeOutput // read only after exited
	closed   bool
	killed   bool
	stopTree func() error
}

type bridgeResponse struct {
	OK    bool   `json:"ok"`
	Error string `json:"error"`
}

func startLibreOfficeSession(ctx context.Context, binary, python string) (_ conversionSession, err error) {
	return startLibreOfficeBridge(ctx, binary, python, exec.CommandContext)
}

func startLibreOfficeBridge(ctx context.Context, binary, python string, command func(context.Context, string, ...string) *exec.Cmd) (_ conversionSession, err error) {
	startup, cancelStartup := context.WithTimeout(ctx, 30*time.Second)
	defer cancelStartup()
	profile, err := os.MkdirTemp("", "excel2pdf-session-*")
	if err != nil {
		return nil, fmt.Errorf("create persistent LibreOffice profile: %w", err)
	}
	lifetime, stop := context.WithCancel(context.WithoutCancel(ctx))
	s := &libreOfficeSession{profile: profile, stop: stop, exited: make(chan struct{})}
	// Isolated Python excludes CWD, user-site packages and PYTHONPATH imports.
	s.cmd = command(lifetime, python, "-I", "-u", "-c", libreOfficeBridge, binary, profile)
	s.cmd.WaitDelay = 2 * time.Second
	s.cmd.Stderr = &s.logs
	defer func() {
		if err != nil {
			err = errors.Join(err, s.close())
		}
	}()
	s.input, err = s.cmd.StdinPipe()
	if err != nil {
		return nil, fmt.Errorf("open UNO bridge input: %w", err)
	}
	stdout, writer, err := os.Pipe()
	if err != nil {
		return nil, fmt.Errorf("open UNO bridge output: %w", err)
	}
	defer writer.Close()
	s.reader = stdout
	s.cmd.Stdout = writer
	s.output = bufio.NewScanner(stdout)
	s.output.Buffer(make([]byte, 4096), libreOfficeOutputLimit)
	if s.stopTree, err = startBridgeProcess(s.cmd); err != nil {
		return nil, fmt.Errorf("start UNO bridge: %w", err)
	}
	go func() {
		s.waitErr = s.cmd.Wait()
		close(s.exited)
	}()
	if err = writer.Close(); err != nil {
		return nil, fmt.Errorf("close parent UNO output writer: %w", err)
	}
	if err = s.exchange(startup, nil); err != nil {
		killErr := s.kill()
		diagnostic := sanitizeLibreOfficeDiagnostic(err.Error()+" "+s.logs.String(), "", "", profile)
		return nil, errors.Join(startup.Err(), killErr, fmt.Errorf("initialize persistent LibreOffice (requires Python 3 and compatible PyUNO): %s", diagnostic))
	}
	return s, nil
}

// exchange has exclusive access to the protocol. On cancellation it kills and
// reaps the owned process group before joining the I/O goroutine.
func (s *libreOfficeSession) exchange(ctx context.Context, request any) error {
	result := make(chan error, 1)
	go func() {
		if request != nil {
			if err := json.NewEncoder(s.input).Encode(request); err != nil {
				result <- fmt.Errorf("write UNO request: %w", err)
				return
			}
		}
		if !s.output.Scan() {
			err := s.output.Err()
			if err == nil {
				err = io.ErrUnexpectedEOF
			}
			result <- fmt.Errorf("read UNO response: %w", err)
			return
		}
		var response bridgeResponse
		if err := json.Unmarshal(s.output.Bytes(), &response); err != nil {
			result <- fmt.Errorf("decode UNO response: %w", err)
			return
		}
		if !response.OK {
			result <- fmt.Errorf("UNO bridge: %s", response.Error)
			return
		}
		result <- nil
	}()
	select {
	case err := <-result:
		return err
	case <-s.exited:
		// A launcher can exit while its interpreter still holds both pipes.
		killErr := s.kill()
		return errors.Join(errors.New("UNO bridge exited"), killErr, <-result)
	case <-ctx.Done():
		killErr := s.kill()
		<-result
		return errors.Join(ctx.Err(), killErr)
	}
}

func (s *libreOfficeSession) convert(ctx context.Context, input, output string) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	err := s.exchange(ctx, map[string]string{"input": input, "output": output})
	if err != nil {
		// Stop before reading stderr, whose writer belongs to os/exec.
		killErr := s.kill()
		diagnostic := sanitizeLibreOfficeDiagnostic(err.Error()+" "+s.logs.String(), input, output, s.profile)
		return errors.Join(ctx.Err(), killErr, fmt.Errorf("persistent LibreOffice: %s", diagnostic))
	}
	return nil
}

func (s *libreOfficeSession) kill() error {
	if s.stopTree != nil && !s.killed {
		s.killed = true
		// Also reach descendants if the bridge itself has already exited.
		err := s.stopTree()
		if errors.Is(err, os.ErrProcessDone) {
			err = nil
		}
		s.stop()
		<-s.exited
		// Unblock a reader even if an unexpected child inherited stdout.
		if s.reader != nil {
			err = errors.Join(err, s.reader.Close())
			s.reader = nil
		}
		return err
	}
	return nil
}

func (s *libreOfficeSession) close() error {
	if s.closed {
		return nil
	}
	s.closed = true
	var err error
	if s.input != nil {
		// EOF asks the bridge to terminate its desktop and wait for LibreOffice.
		if closeErr := s.input.Close(); closeErr != nil && !errors.Is(closeErr, os.ErrClosed) {
			err = errors.Join(err, fmt.Errorf("close UNO bridge input: %w", closeErr))
		}
	}
	if s.stopTree != nil {
		timer := time.NewTimer(5 * time.Second)
		defer timer.Stop()
		select {
		case <-s.exited:
			if !s.killed && s.waitErr != nil {
				err = errors.Join(err, fmt.Errorf("UNO bridge exited: %w", s.waitErr))
			}
		case <-timer.C:
			err = errors.Join(err, errors.New("persistent LibreOffice did not stop within 5 seconds"))
		}
		err = errors.Join(err, s.kill())
	}
	if s.reader != nil {
		err = errors.Join(err, s.reader.Close())
		s.reader = nil
	}
	s.stop()
	if cleanupErr := removeSessionProfile(s.profile); cleanupErr != nil {
		err = errors.Join(err, fmt.Errorf("remove persistent LibreOffice profile: %w", cleanupErr))
	}
	return err
}
