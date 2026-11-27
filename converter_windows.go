package excel2pdf

import "context"

func persistentSessionFactory() (func(context.Context) (conversionSession, error), error) {
	return nil, ErrPersistentConversionUnsupported
}
