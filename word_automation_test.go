package excel2pdf

import (
	"context"
	"errors"
	"reflect"
	"testing"
)

func wordTestEvents() []string {
	return []string{
		"word.set.AutomationSecurity", "word.set.DisplayAlerts", "word.set.Visible",
		"word.Options", "options.get.UpdateLinksAtOpen", "options.set.UpdateLinksAtOpen", "word.Documents",
		"documents.Open", "document.ExportAsFixedFormat", "document.Close",
		"document.release", "documents.release", "options.restore.UpdateLinksAtOpen", "options.release", "word.Quit",
	}
}

func TestExportWordPDF(t *testing.T) {
	t.Parallel()
	state := &fakeExcelState{}
	word := &fakeWordObject{fakeExcelObject{name: "word", state: state}}
	if err := exportWordPDF(context.Background(), word, "input.docx", "output.pdf"); err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(state.events, wordTestEvents()) {
		t.Errorf("Word lifecycle = %v", state.events)
	}
	for event, want := range map[string][]any{
		"word.set.AutomationSecurity":       {3},
		"word.set.DisplayAlerts":            {0},
		"word.set.Visible":                  {false},
		"options.set.UpdateLinksAtOpen":     {false},
		"options.restore.UpdateLinksAtOpen": {true},
		"documents.Open":                    {"input.docx", false, true, false},
		"document.ExportAsFixedFormat":      {"output.pdf", 17, false},
		"document.Close":                    {0},
		"word.Quit":                         {0},
	} {
		if !reflect.DeepEqual(state.args[event], want) {
			t.Errorf("%s arguments = %v, want %v", event, state.args[event], want)
		}
	}
}

func TestExportWordPDFFailures(t *testing.T) {
	t.Parallel()
	operationErr := errors.New("Word operation failed")
	for index, event := range wordTestEvents() {
		for _, mode := range []string{"error", "panic"} {
			t.Run(event+"/"+mode, func(t *testing.T) {
				t.Parallel()
				state := &fakeExcelState{}
				if mode == "error" {
					state.failures = map[string]error{event: operationErr}
				} else {
					state.panicAt, state.panicValue = event, operationErr
				}
				err := exportWordPDF(context.Background(), &fakeWordObject{fakeExcelObject{name: "word", state: state}}, "input.docx", "output.pdf")
				if !errors.Is(err, operationErr) {
					t.Errorf("error = %v, want preserved operation error", err)
				}
				want := wordTestEvents()[:index+1]
				switch {
				case index < 4:
					want = append(want, "word.Quit")
				case index == 4:
					want = append(want, "options.release", "word.Quit")
				case index < 7:
					want = append(want, "options.restore.UpdateLinksAtOpen", "options.release", "word.Quit")
				case index == 7:
					want = append(want, "documents.release", "options.restore.UpdateLinksAtOpen", "options.release", "word.Quit")
				default:
					want = wordTestEvents()
				}
				if !reflect.DeepEqual(state.events, want) {
					t.Errorf("events = %v, want %v", state.events, want)
				}
			})
		}
	}
}

func TestExportWordPDFCancellation(t *testing.T) {
	t.Parallel()
	for _, tt := range []struct {
		name       string
		cancelAt   string
		wantEvents []string
	}{
		{name: "before security", wantEvents: []string{"word.Quit"}},
		{name: "during security", cancelAt: "word.set.AutomationSecurity", wantEvents: []string{"word.set.AutomationSecurity", "word.Quit"}},
		{name: "before open", cancelAt: "word.Documents", wantEvents: append(wordTestEvents()[:7], "documents.release", "options.restore.UpdateLinksAtOpen", "options.release", "word.Quit")},
		{name: "after open", cancelAt: "documents.Open", wantEvents: append(wordTestEvents()[:8], wordTestEvents()[9:]...)},
		{name: "during export", cancelAt: "document.ExportAsFixedFormat", wantEvents: wordTestEvents()},
	} {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			ctx, cancel := context.WithCancel(context.Background())
			defer cancel()
			if tt.cancelAt == "" {
				cancel()
			}
			state := &fakeExcelState{after: map[string]func(){tt.cancelAt: cancel}}
			err := exportWordPDF(ctx, &fakeWordObject{fakeExcelObject{name: "word", state: state}}, "input.docx", "output.pdf")
			if !errors.Is(err, context.Canceled) {
				t.Errorf("error = %v, want context.Canceled", err)
			}
			if !reflect.DeepEqual(state.events, tt.wantEvents) {
				t.Errorf("events = %v, want %v", state.events, tt.wantEvents)
			}
		})
	}
}

func TestExportWordPDFPreservesCleanupErrors(t *testing.T) {
	t.Parallel()
	exportErr := errors.New("export failed")
	closeErr := errors.New("close failed")
	quitErr := errors.New("quit failed")
	state := &fakeExcelState{failures: map[string]error{
		"document.ExportAsFixedFormat": exportErr, "document.Close": closeErr, "word.Quit": quitErr,
	}}
	err := exportWordPDF(context.Background(), &fakeWordObject{fakeExcelObject{name: "word", state: state}}, "input.docx", "output.pdf")
	for _, want := range []error{exportErr, closeErr, quitErr} {
		if !errors.Is(err, want) {
			t.Errorf("lost error %v: %v", want, err)
		}
	}
}

type fakeWordObject struct{ fakeExcelObject }

func (object *fakeWordObject) setProperty(name string, value any) error {
	if object.name == "options" && name == "UpdateLinksAtOpen" && value == true {
		return object.record("restore."+name, value)
	}
	return object.fakeExcelObject.setProperty(name, value)
}

func (object *fakeWordObject) getObject(name string, args ...any) (officeAutomationObject, error) {
	if err := object.record(name, args...); err != nil {
		return nil, err
	}
	childName := map[string]string{"Options": "options", "Documents": "documents", "Open": "document"}[name]
	return &fakeWordObject{fakeExcelObject{name: childName, state: object.state}}, nil
}
