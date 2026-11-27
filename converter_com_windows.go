package excel2pdf

import (
	"context"
	"errors"
	"runtime"

	"github.com/go-ole/go-ole"
)

type microsoftSession struct {
	requests  chan func()
	done      chan error
	installed map[documentKind]bool
	apps      map[documentKind]*comOfficeObject // only accessed on the apartment thread
}

func microsoftSessionFactory() (func(context.Context) (conversionSession, error), error) {
	excel, err := isExcelInstalled()
	if err != nil {
		return nil, err
	}
	word, err := isWordInstalled()
	if err != nil {
		return nil, err
	}
	if !excel && !word {
		return nil, ErrOfficeNotInstalled
	}
	return func(ctx context.Context) (conversionSession, error) {
		s := &microsoftSession{
			requests: make(chan func()), done: make(chan error, 1),
			installed: map[documentKind]bool{spreadsheetDocument: excel, wordDocument: word},
			apps:      make(map[documentKind]*comOfficeObject),
		}
		ready := make(chan error, 1)
		go s.run(ready)
		if err := <-ready; err != nil {
			return nil, err
		}
		if err := ctx.Err(); err != nil {
			return nil, errors.Join(err, s.close())
		}
		return s, nil
	}, nil
}

func (s *microsoftSession) run(ready chan<- error) {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_SPEED_OVER_MEMORY)
	if !officeCOMInitialized(err) {
		ready <- err
		return
	}
	ready <- nil
	for request := range s.requests {
		request()
	}
	var cleanup error
	for kind, app := range s.apps {
		cleanup = errors.Join(cleanup, quitOfficeApplication(kind, app))
	}
	ole.CoUninitialize()
	s.done <- cleanup
}

func quitOfficeApplication(kind documentKind, app *comOfficeObject) (err error) {
	defer recoverOfficePanic("office", &err)
	defer app.dispatch.Release()
	if kind == wordDocument {
		return app.call("Quit", 0)
	}
	return app.call("Quit")
}

func (s *microsoftSession) convert(ctx context.Context, input, output string) error {
	return s.invoke(ctx, func() error { return s.export(ctx, input, output) })
}

func (s *microsoftSession) invoke(ctx context.Context, call func() error) error {
	result := make(chan error, 1)
	select {
	case <-ctx.Done():
		return ctx.Err()
	case s.requests <- func() { result <- call() }:
		// Join even on cancellation: COM must release the document before
		// the caller can clean up or start another conversion.
		return <-result
	}
}

func (s *microsoftSession) export(ctx context.Context, input, output string) (err error) {
	defer recoverOfficePanic("office", &err)
	if err := ctx.Err(); err != nil {
		return err
	}
	kind, err := documentKindForPath(input)
	if err != nil {
		return err
	}
	if !s.installed[kind] {
		return ErrOfficeNotInstalled
	}
	program, export := "Excel.Application", exportWorkbook
	if kind == wordDocument {
		program, export = "Word.Application", exportDocument
	}
	app := s.apps[kind]
	if app == nil {
		app, err = createOfficeApplication(program)
		if err != nil {
			return err
		}
		s.apps[kind] = app
	}
	return export(ctx, app, input, output)
}

func (s *microsoftSession) close() error {
	close(s.requests)
	return <-s.done
}
