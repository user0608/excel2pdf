package main

import (
	"context"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"time"

	"github.com/user0608/excel2pdf"
)

func main() {
	if err := run(); err != nil {
		log.Fatal(err)
	}
}

func run() error {
	input := "file.xlsx"
	if len(os.Args) > 1 {
		input = os.Args[1]
	}
	ctx, cancel := context.WithTimeout(context.Background(), 30*time.Second)
	defer cancel()
	pdf, err := excel2pdf.ConvertExcelToPdfContext(ctx, input)
	if err != nil {
		return err
	}
	defer func() {
		if err := os.RemoveAll(filepath.Dir(pdf)); err != nil {
			log.Printf("remove PDF temporary directory: %v", err)
		}
	}()
	// Consume the PDF before cleaning up its directory. An application can copy
	// or send it here; this example simply reports its size.
	info, err := os.Stat(pdf)
	if err != nil {
		return err
	}
	fmt.Printf("Generated %s (%d bytes); temporary output will now be removed\n", filepath.Base(pdf), info.Size())
	return nil
}
