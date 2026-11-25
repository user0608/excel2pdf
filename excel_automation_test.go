package excel2pdf

import (
	"context"
	"errors"
	"reflect"
	"strings"
	"testing"
)

func TestExportExcelPDF(t *testing.T) {
	t.Parallel()
	operationErr := errors.New("operation failed")
	cleanupErr := errors.New("cleanup failed")
	allEvents := []string{
		"excel.Workbooks", "workbooks.Open", "workbook.ExportAsFixedFormat",
		"workbook.Close", "workbook.release", "workbooks.release", "excel.Quit",
	}
	tests := []struct {
		name       string
		failures   map[string]error
		panicAt    string
		panicValue any
		wantEvents []string
		wantErrors []error
		wantPanic  bool
	}{
		{name: "success", wantEvents: allEvents},
		{
			name:       "workbooks failure still quits Excel",
			failures:   map[string]error{"excel.Workbooks": operationErr},
			wantEvents: []string{"excel.Workbooks", "excel.Quit"},
			wantErrors: []error{operationErr},
		},
		{
			name:       "open failure releases workbooks and quits Excel",
			failures:   map[string]error{"workbooks.Open": operationErr},
			wantEvents: []string{"excel.Workbooks", "workbooks.Open", "workbooks.release", "excel.Quit"},
			wantErrors: []error{operationErr},
		},
		{
			name:       "export failure cleans up all resources",
			failures:   map[string]error{"workbook.ExportAsFixedFormat": operationErr},
			wantEvents: allEvents, wantErrors: []error{operationErr},
		},
		{
			name:       "close failure does not skip release or quit",
			failures:   map[string]error{"workbook.Close": cleanupErr},
			wantEvents: allEvents, wantErrors: []error{cleanupErr},
		},
		{
			name:       "workbook release failure does not skip remaining cleanup",
			failures:   map[string]error{"workbook.release": cleanupErr},
			wantEvents: allEvents, wantErrors: []error{cleanupErr},
		},
		{
			name:       "workbooks release failure still quits Excel",
			failures:   map[string]error{"workbooks.release": cleanupErr},
			wantEvents: allEvents, wantErrors: []error{cleanupErr},
		},
		{
			name:       "quit failure is returned",
			failures:   map[string]error{"excel.Quit": cleanupErr},
			wantEvents: allEvents, wantErrors: []error{cleanupErr},
		},
		{
			name:       "export and cleanup errors remain inspectable",
			failures:   map[string]error{"workbook.ExportAsFixedFormat": operationErr, "workbook.Close": cleanupErr},
			wantEvents: allEvents, wantErrors: []error{operationErr, cleanupErr},
		},
		{
			name:    "open panic becomes an error and quits Excel",
			panicAt: "workbooks.Open", panicValue: operationErr,
			wantEvents: []string{"excel.Workbooks", "workbooks.Open", "workbooks.release", "excel.Quit"},
			wantErrors: []error{operationErr}, wantPanic: true,
		},
		{
			name:    "export string panic becomes an error and closes workbook",
			panicAt: "workbook.ExportAsFixedFormat", panicValue: "unexpected failure",
			wantEvents: allEvents, wantPanic: true,
		},
		{
			name:     "cleanup panic preserves export error and continues cleanup",
			failures: map[string]error{"workbook.ExportAsFixedFormat": operationErr},
			panicAt:  "workbook.Close", panicValue: cleanupErr,
			wantEvents: allEvents, wantErrors: []error{operationErr, cleanupErr}, wantPanic: true,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			state := &fakeExcelState{failures: tt.failures, panicAt: tt.panicAt, panicValue: tt.panicValue}
			err := exportExcelPDF(context.Background(), &fakeExcelObject{name: "excel", state: state}, "input.xlsx", "output.pdf")
			wantEvents := append(excelSecurityTestEvents(), tt.wantEvents...)
			if !reflect.DeepEqual(state.events, wantEvents) {
				t.Errorf("events = %v, want %v", state.events, wantEvents)
			}
			if len(tt.wantErrors) == 0 && !tt.wantPanic && err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			for _, want := range tt.wantErrors {
				if !errors.Is(err, want) {
					t.Errorf("error = %v, want errors.Is(_, %v)", err, want)
				}
			}
			if tt.wantPanic && (err == nil || !strings.Contains(err.Error(), "excel automation panic")) {
				t.Errorf("panic did not produce an explicit error: %v", err)
			}
			if got := state.args["workbooks.Open"]; got != nil && !reflect.DeepEqual(got, []any{"input.xlsx", 0, true}) {
				t.Errorf("Open args = %v", got)
			}
			if got := state.args["workbook.ExportAsFixedFormat"]; got != nil && !reflect.DeepEqual(got, []any{0, "output.pdf"}) {
				t.Errorf("ExportAsFixedFormat args = %v", got)
			}
			if got := state.args["workbook.Close"]; got != nil && !reflect.DeepEqual(got, []any{false}) {
				t.Errorf("Close args = %v, want saving disabled", got)
			}
		})
	}
}

func TestExportExcelPDFCancellation(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name       string
		cancelAt   string
		wantEvents []string
	}{
		{name: "before workbooks", wantEvents: []string{"excel.Quit"}},
		{
			name: "before open", cancelAt: "excel.Workbooks",
			wantEvents: []string{"excel.Workbooks", "workbooks.release", "excel.Quit"},
		},
		{
			name: "after open", cancelAt: "workbooks.Open",
			wantEvents: []string{"excel.Workbooks", "workbooks.Open", "workbook.Close", "workbook.release", "workbooks.release", "excel.Quit"},
		},
		{
			name: "during export", cancelAt: "workbook.ExportAsFixedFormat",
			wantEvents: []string{"excel.Workbooks", "workbooks.Open", "workbook.ExportAsFixedFormat", "workbook.Close", "workbook.release", "workbooks.release", "excel.Quit"},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			ctx, cancel := context.WithCancel(context.Background())
			defer cancel()
			if tt.cancelAt == "" {
				cancel()
			}
			state := &fakeExcelState{after: map[string]func(){tt.cancelAt: cancel}}
			err := exportExcelPDF(ctx, &fakeExcelObject{name: "excel", state: state}, "input.xlsx", "output.pdf")
			if !errors.Is(err, context.Canceled) {
				t.Errorf("error = %v, want context.Canceled", err)
			}
			wantEvents := tt.wantEvents
			if tt.cancelAt != "" {
				wantEvents = append(excelSecurityTestEvents(), wantEvents...)
			}
			if !reflect.DeepEqual(state.events, wantEvents) {
				t.Errorf("events = %v, want %v", state.events, wantEvents)
			}
		})
	}
}

func TestExportExcelPDFSecurityPolicy(t *testing.T) {
	t.Parallel()
	state := &fakeExcelState{}
	if err := exportExcelPDF(context.Background(), &fakeExcelObject{name: "excel", state: state}, "input.xlsx", "output.pdf"); err != nil {
		t.Fatal(err)
	}
	for name, want := range map[string]any{
		"AutomationSecurity": 3,
		"EnableEvents":       false,
		"AskToUpdateLinks":   false,
		"DisplayAlerts":      false,
		"Visible":            false,
	} {
		if got := state.args["excel.set."+name]; !reflect.DeepEqual(got, []any{want}) {
			t.Errorf("Excel %s = %v, want %v", name, got, want)
		}
	}
	if len(state.events) < 6 {
		t.Fatalf("incomplete security/open sequence: %v", state.events)
	}
	if !reflect.DeepEqual(state.events[:5], excelSecurityTestEvents()) || state.events[5] != "excel.Workbooks" {
		t.Errorf("security policy was not applied before accessing workbooks: %v", state.events)
	}
	if got := state.args["workbooks.Open"]; !reflect.DeepEqual(got, []any{"input.xlsx", 0, true}) {
		t.Errorf("Open did not disable updates and enable read-only mode: %v", got)
	}
}

func TestExportExcelPDFFailsClosedOnSecurityConfiguration(t *testing.T) {
	t.Parallel()
	policyErr := errors.New("policy unavailable")
	for index, event := range excelSecurityTestEvents() {
		for _, mode := range []string{"error", "cancel", "panic"} {
			t.Run(event+"/"+mode, func(t *testing.T) {
				t.Parallel()
				ctx, cancel := context.WithCancel(context.Background())
				defer cancel()
				state := &fakeExcelState{}
				wantErr := policyErr
				switch mode {
				case "error":
					state.failures = map[string]error{event: policyErr}
				case "cancel":
					state.after = map[string]func(){event: cancel}
					wantErr = context.Canceled
				case "panic":
					state.panicAt, state.panicValue = event, policyErr
				}
				err := exportExcelPDF(ctx, &fakeExcelObject{name: "excel", state: state}, "input.xlsx", "output.pdf")
				if !errors.Is(err, wantErr) {
					t.Errorf("error = %v, want %v", err, wantErr)
				}
				wantEvents := append(excelSecurityTestEvents()[:index+1], "excel.Quit")
				if !reflect.DeepEqual(state.events, wantEvents) {
					t.Errorf("workbook accessed after failed security configuration: events %v, want %v", state.events, wantEvents)
				}
			})
		}
	}
}

type fakeExcelState struct {
	events     []string
	args       map[string][]any
	failures   map[string]error
	panicAt    string
	panicValue any
	after      map[string]func()
}

type fakeExcelObject struct {
	name  string
	state *fakeExcelState
}

func (object *fakeExcelObject) record(name string, args ...any) error {
	event := object.name + "." + name
	object.state.events = append(object.state.events, event)
	if object.state.args == nil {
		object.state.args = make(map[string][]any)
	}
	object.state.args[event] = args
	if after := object.state.after[event]; after != nil {
		after()
	}
	if event == object.state.panicAt {
		panic(object.state.panicValue)
	}
	return object.state.failures[event]
}

func (object *fakeExcelObject) getObject(name string, args ...any) (officeAutomationObject, error) {
	if err := object.record(name, args...); err != nil {
		return nil, err
	}
	childName := "workbooks"
	if name == "Open" {
		childName = "workbook"
	}
	return &fakeExcelObject{name: childName, state: object.state}, nil
}

func (object *fakeExcelObject) call(name string, args ...any) error {
	return object.record(name, args...)
}

func (object *fakeExcelObject) setProperty(name string, value any) error {
	return object.record("set."+name, value)
}

func (object *fakeExcelObject) getBoolProperty(name string) (bool, error) {
	return true, object.record("get." + name)
}

func excelSecurityTestEvents() []string {
	return []string{
		"excel.set.AutomationSecurity", "excel.set.EnableEvents",
		"excel.set.AskToUpdateLinks", "excel.set.DisplayAlerts", "excel.set.Visible",
	}
}

func (object *fakeExcelObject) release() error {
	return object.record("release")
}
