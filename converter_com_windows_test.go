package excel2pdf

import (
	"context"
	"errors"
	"os"
	"path/filepath"
	"sync"
	"testing"
	"time"
)

// Substitute only the Office call; dispatch and shutdown use the real apartment.
type blockedOfficeCall struct {
	*microsoftSession
	entered chan string
	release chan struct{}
}

func (s *blockedOfficeCall) convert(ctx context.Context, _, output string) error {
	return s.invoke(ctx, func() error {
		s.entered <- output
		<-s.release
		return os.WriteFile(output, []byte("%PDF-1.7\n"), 0600)
	})
}

func TestPersistentOfficeCancellation(t *testing.T) {
	worker := &microsoftSession{requests: make(chan func()), done: make(chan error, 1)}
	ready := make(chan error, 1)
	go worker.run(ready)
	if err := <-ready; err != nil {
		t.Fatal(err)
	}
	blocked := &blockedOfficeCall{worker, make(chan string, 1), make(chan struct{})}
	starts := 0
	c := testConverter(t, func(context.Context) (conversionSession, error) {
		starts++
		if starts == 1 {
			return blocked, nil
		}
		return &fakeConversionSession{}, nil
	})
	release := sync.OnceFunc(func() { close(blocked.release) })
	defer release()
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	input := writeExcelTestInput(t, "office.docx")
	result := make(chan error, 1)
	go func() {
		_, err := c.ConvertToPDF(ctx, input)
		result <- err
	}()
	var output string
	select {
	case output = <-blocked.entered:
	case <-time.After(5 * time.Second):
		t.Fatal("Office call never started")
	}
	wait, stop := context.WithTimeout(context.Background(), 25*time.Millisecond)
	defer stop()
	if _, err := c.ConvertToPDF(wait, input); !errors.Is(err, context.DeadlineExceeded) {
		t.Fatalf("waiting cancellation = %v", err)
	}
	cancel()
	select {
	case err := <-result:
		t.Fatalf("returned before the Office call finished: %v", err)
	case <-time.After(25 * time.Millisecond):
	}
	release()
	select {
	case err := <-result:
		if !errors.Is(err, context.Canceled) {
			t.Fatalf("active cancellation = %v", err)
		}
	case <-time.After(5 * time.Second):
		t.Fatal("worker did not finish after release")
	}
	if c.session != nil || starts != 1 {
		t.Fatal("canceled session was not discarded")
	}
	if _, err := os.Stat(filepath.Dir(output)); !errors.Is(err, os.ErrNotExist) {
		t.Fatalf("canceled output directory survived: %v", err)
	}
	pdf, err := c.ConvertToPDF(context.Background(), input)
	if err != nil {
		t.Fatal(err)
	}
	removeExcelTestOutput(t, pdf)
	if starts != 2 {
		t.Fatal("expected a fresh session after cancellation")
	}
}
