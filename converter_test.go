package excel2pdf

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"sync"
	"sync/atomic"
	"testing"
	"time"
)

type fakeConversionSession struct {
	convertFunc func(context.Context, string, string) error
	closes      atomic.Int32
	closeErr    error
}

func (s *fakeConversionSession) convert(ctx context.Context, input, output string) error {
	if s.convertFunc != nil {
		return s.convertFunc(ctx, input, output)
	}
	return os.WriteFile(output, []byte("%PDF-1.7\n"), 0o600)
}

func (s *fakeConversionSession) close() error {
	s.closes.Add(1)
	return s.closeErr
}

func testConverter(t *testing.T, start func(context.Context) (conversionSession, error)) *Converter {
	t.Helper()
	c, err := newConverter(context.Background(), start)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() {
		if err := c.Close(); err != nil {
			t.Errorf("close converter: %v", err)
		}
	})
	return c
}

func TestNewConverterInvalidContext(t *testing.T) {
	t.Parallel()
	if c, err := NewConverter(nil); c != nil || err == nil { //nolint:staticcheck // Verify intentional rejection of nil contexts.
		t.Fatalf("nil context = (%v, %v)", c, err)
	}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	if c, err := NewConverter(ctx); c != nil || !errors.Is(err, context.Canceled) {
		t.Fatalf("canceled constructor = (%v, %v)", c, err)
	}
}

func TestConverterReusesSession(t *testing.T) {
	t.Parallel()
	s := &fakeConversionSession{}
	var starts atomic.Int32
	c := testConverter(t, func(context.Context) (conversionSession, error) {
		starts.Add(1)
		return s, nil
	})
	var previous string
	for _, name := range []string{"report.xlsx", "document.docx"} {
		pdf, err := c.ConvertToPDF(context.Background(), writeExcelTestInput(t, name))
		if err != nil {
			t.Fatal(err)
		}
		t.Cleanup(func() { removeExcelTestOutput(t, pdf) })
		if previous == filepath.Dir(pdf) {
			t.Fatal("reused PDF output directory")
		}
		previous = filepath.Dir(pdf)
	}
	if starts.Load() != 1 || s.closes.Load() != 0 {
		t.Fatal("session was restarted between successful conversions")
	}
	if err := c.Close(); err != nil {
		t.Fatal(err)
	}
	if s.closes.Load() != 1 {
		t.Fatal("session was not closed exactly once")
	}
}

func TestConverterConcurrentCallsAndWaitingCancellation(t *testing.T) {
	t.Parallel()
	input := writeExcelTestInput(t, "input.docx")
	entered, release := make(chan struct{}), make(chan struct{})
	var calls atomic.Int32
	var active atomic.Int32
	s := &fakeConversionSession{convertFunc: func(ctx context.Context, _, output string) error {
		if active.Add(1) != 1 {
			t.Error("overlapping conversions in one session")
		}
		defer active.Add(-1)
		if calls.Add(1) == 1 {
			close(entered)
			select {
			case <-release:
			case <-ctx.Done():
				return ctx.Err()
			}
		}
		return os.WriteFile(output, []byte("%PDF-1.7\n"), 0o600)
	}}
	c := testConverter(t, func(context.Context) (conversionSession, error) { return s, nil })
	result := make(chan error, 1)
	go func() {
		pdf, err := c.ConvertToPDF(context.Background(), input)
		if pdf != "" {
			err = errors.Join(err, os.RemoveAll(filepath.Dir(pdf)))
		}
		result <- err
	}()
	<-entered
	ctx, cancel := context.WithTimeout(context.Background(), 20*time.Millisecond)
	defer cancel()
	if _, err := c.ConvertToPDF(ctx, input); !errors.Is(err, context.DeadlineExceeded) {
		t.Fatalf("waiting cancellation = %v", err)
	}
	if calls.Load() != 1 || s.closes.Load() != 0 {
		t.Fatal("waiting cancellation affected active session")
	}
	close(release)
	if err := <-result; err != nil {
		t.Fatal(err)
	}
	var group sync.WaitGroup
	for i := 0; i < 5; i++ {
		group.Add(1)
		go func() {
			defer group.Done()
			pdf, err := c.ConvertToPDF(context.Background(), input)
			if err != nil {
				t.Error(err)
				return
			}
			if err := os.RemoveAll(filepath.Dir(pdf)); err != nil {
				t.Error(err)
			}
		}()
	}
	group.Wait()
	if calls.Load() != 6 {
		t.Fatalf("completed calls = %d, want 6", calls.Load())
	}
}

func TestConverterFailureAndRestart(t *testing.T) {
	t.Parallel()
	failure := errors.New("export failed")
	first := &fakeConversionSession{convertFunc: func(context.Context, string, string) error { return failure }}
	second := &fakeConversionSession{}
	var starts int
	c := testConverter(t, func(context.Context) (conversionSession, error) {
		starts++
		if starts == 1 {
			return first, nil
		}
		return second, nil
	})
	input := writeExcelTestInput(t, "input.xlsx")
	if pdf, err := c.ConvertToPDF(context.Background(), input); pdf != "" || !errors.Is(err, failure) {
		t.Fatalf("failed conversion = (%q, %v)", pdf, err)
	}
	if first.closes.Load() != 1 {
		t.Fatal("failed session not closed")
	}
	pdf, err := c.ConvertToPDF(context.Background(), input)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { removeExcelTestOutput(t, pdf) })
	if starts != 2 {
		t.Fatalf("session starts = %d", starts)
	}
}

func TestConverterCloseCancelsActiveAndWaitingCalls(t *testing.T) {
	t.Parallel()
	entered := make(chan struct{})
	s := &fakeConversionSession{convertFunc: func(ctx context.Context, _, _ string) error {
		close(entered)
		<-ctx.Done()
		return ctx.Err()
	}}
	c := testConverter(t, func(context.Context) (conversionSession, error) { return s, nil })
	input := writeExcelTestInput(t, "input.docx")
	active, waiting := make(chan error, 1), make(chan error, 1)
	go func() { _, err := c.ConvertToPDF(context.Background(), input); active <- err }()
	<-entered
	go func() { _, err := c.ConvertToPDF(context.Background(), input); waiting <- err }()
	var group sync.WaitGroup
	for i := 0; i < 3; i++ {
		group.Add(1)
		go func() {
			defer group.Done()
			if err := c.Close(); err != nil {
				t.Error(err)
			}
		}()
	}
	group.Wait()
	for _, result := range []chan error{active, waiting} {
		if err := <-result; !errors.Is(err, ErrConverterClosed) {
			t.Fatalf("closed conversion = %v", err)
		}
	}
	if s.closes.Load() != 1 {
		t.Fatalf("session closes = %d", s.closes.Load())
	}
}

func TestConverterInvalidInputAndZeroValue(t *testing.T) {
	t.Parallel()
	s := &fakeConversionSession{}
	c := testConverter(t, func(context.Context) (conversionSession, error) { return s, nil })
	for _, input := range []string{"", "unsupported.txt"} {
		if pdf, err := c.ConvertToPDF(context.Background(), input); pdf != "" || err == nil {
			t.Fatalf("invalid input = (%q, %v)", pdf, err)
		}
	}
	if _, err := c.ConvertToPDF(nil, "unused.xlsx"); err == nil { //nolint:staticcheck // Verify intentional rejection of nil contexts.
		t.Fatal("nil conversion context accepted")
	}
	if s.closes.Load() != 0 {
		t.Fatal("invalid input destroyed session")
	}
	var zero Converter
	if _, err := zero.ConvertToPDF(context.Background(), "unused.xlsx"); !errors.Is(err, ErrConverterClosed) {
		t.Fatalf("zero value = %v", err)
	}
	if err := zero.Close(); err != nil {
		t.Fatal(err)
	}
}

func TestConverterConstructorContextAndCloseError(t *testing.T) {
	t.Parallel()
	closeErr := errors.New("cleanup failed")
	s := &fakeConversionSession{closeErr: closeErr}
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	c, err := newConverter(ctx, func(context.Context) (conversionSession, error) { return s, nil })
	if err != nil {
		t.Fatal(err)
	}
	cancel() // Constructor context must not become the session lifetime.
	pdf, err := c.ConvertToPDF(context.Background(), writeExcelTestInput(t, "input.docx"))
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { removeExcelTestOutput(t, pdf) })
	for i := 0; i < 2; i++ {
		if err := c.Close(); !errors.Is(err, closeErr) {
			t.Fatalf("close error = %v", err)
		}
	}
	if _, err := os.Stat(pdf); err != nil {
		t.Fatalf("Close removed successful output: %v", err)
	}
	if s.closes.Load() != 1 {
		t.Fatalf("session closes = %d", s.closes.Load())
	}
}
