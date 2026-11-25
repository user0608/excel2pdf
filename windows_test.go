//go:build windows

package excel2pdf

import (
	"errors"
	"fmt"
	"testing"

	"github.com/go-ole/go-ole"
)

func TestExcelCOMInitialized(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name string
		err  error
		want bool
	}{
		{name: "initialized", want: true},
		{name: "already initialized in same apartment", err: ole.NewError(1), want: true},
		{name: "wrapped S_FALSE", err: fmt.Errorf("initialize: %w", ole.NewError(1)), want: true},
		{name: "incompatible apartment", err: ole.NewError(0x80010106)},
		{name: "unrelated error", err: errors.New("failed")},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			t.Parallel()
			if got := excelCOMInitialized(tt.err); got != tt.want {
				t.Errorf("excelCOMInitialized(%v) = %v, want %v", tt.err, got, tt.want)
			}
		})
	}
}
