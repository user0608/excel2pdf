package excel2pdf

import (
	"context"
	"errors"
	"testing"
)

func TestNewConverterWindowsUnsupported(t *testing.T) {
	t.Parallel()
	c, err := NewConverter(context.Background())
	if c != nil || !errors.Is(err, ErrPersistentConversionUnsupported) {
		t.Fatalf("Windows persistent constructor = (%v, %v)", c, err)
	}
}
