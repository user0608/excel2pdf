package excel2pdf

import (
	"context"
	"errors"
	"sync"
)

var ErrConverterClosed = errors.New("converter is closed")
var ErrPersistentConversionUnsupported = errors.New("persistent conversion is not supported on this platform")

type conversionSession interface {
	convert(context.Context, string, string) error
	close() error
}

// Converter owns a reusable Office session. Concurrent calls wait for exclusive
// access, respecting their contexts; strict FIFO ordering is not guaranteed.
// A Converter must not be copied. Its zero value is closed.
type Converter struct {
	turn      chan struct{}
	closed    chan struct{}
	stop      context.CancelFunc
	lifetime  context.Context
	closeOnce sync.Once
	closeErr  error
	start     func(context.Context) (conversionSession, error)
	session   conversionSession // accessed only while holding turn
}

// NewConverter starts a reusable LibreOffice session on Linux. It requires
// Python 3 with compatible PyUNO bindings. Windows persistence is not yet
// supported; use the package-level ConvertToPDF there.
// ctx controls startup only. Call Close to release the session after use.
func NewConverter(ctx context.Context) (*Converter, error) {
	if ctx == nil {
		return nil, errors.New("converter context must not be nil")
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	start, err := persistentSessionFactory()
	if err != nil {
		return nil, err
	}
	return newConverter(ctx, start)
}

func newConverter(ctx context.Context, start func(context.Context) (conversionSession, error)) (*Converter, error) {
	session, err := start(ctx)
	if err != nil {
		return nil, err
	}
	if err := ctx.Err(); err != nil {
		return nil, errors.Join(err, session.close())
	}
	lifetime, stop := context.WithCancel(context.WithoutCancel(ctx))
	return &Converter{
		turn: make(chan struct{}, 1), closed: make(chan struct{}),
		lifetime: lifetime, stop: stop, start: start, session: session,
	}, nil
}

// ConvertToPDF converts an XLSX or DOCX, reusing this converter's session.
// It waits for the active conversion rather than returning ErrConversionInProgress.
// The caller owns the returned PDF and its exclusive directory, as with the
// package-level function. An active cancellation stops the owned session; the
// next call starts a new one. Waiting cancellation does not affect other calls.
func (c *Converter) ConvertToPDF(ctx context.Context, inputPath string) (string, error) {
	if ctx == nil {
		return "", errors.New("conversion context must not be nil")
	}
	if err := ctx.Err(); err != nil {
		return "", err
	}
	if c == nil || c.turn == nil {
		return "", ErrConverterClosed
	}
	select {
	case c.turn <- struct{}{}:
		defer func() { <-c.turn }()
	case <-ctx.Done():
		return "", ctx.Err()
	case <-c.closed:
		return "", ErrConverterClosed
	}
	select {
	case <-c.closed:
		return "", ErrConverterClosed
	default:
	}
	operation, cancel := context.WithCancel(ctx)
	stop := context.AfterFunc(c.lifetime, cancel)
	defer stop()
	defer cancel()
	pdf, err := convertDocumentFile(operation, inputPath, func(ctx context.Context, input, output string) error {
		if c.session == nil {
			var err error
			c.session, err = c.start(ctx)
			if err != nil {
				return err
			}
		}
		if err := c.session.convert(ctx, input, output); err != nil {
			cleanupErr := c.session.close()
			c.session = nil
			return errors.Join(err, cleanupErr)
		}
		return nil
	})
	if err != nil && c.lifetime.Err() != nil {
		err = errors.Join(err, ErrConverterClosed)
	}
	return pdf, err
}

// Close rejects new calls, cancels the active conversion, waits for cleanup,
// and releases the owned session. It is safe to call concurrently or repeatedly.
// It does not remove PDFs returned by successful conversions.
func (c *Converter) Close() error {
	if c == nil || c.turn == nil {
		return nil
	}
	c.closeOnce.Do(func() {
		close(c.closed)
		c.stop()
		c.turn <- struct{}{}
		defer func() { <-c.turn }()
		if c.session != nil {
			c.closeErr = c.session.close()
			c.session = nil
		}
	})
	return c.closeErr
}
