package excel2pdf

import (
	"bufio"
	"context"
	"fmt"
	"os"
	"os/exec"
	"strconv"
	"strings"
	"syscall"
	"testing"
	"time"
)

func TestConfigureLibreOfficeCommandOwnProcessGroup(t *testing.T) {
	t.Parallel()
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()
	cmd := exec.CommandContext(ctx, os.Args[0], "-test.run=^TestLibreOfficeProcessGroupHelper$")
	cmd.Env = append(os.Environ(), "EXCEL2PDF_GROUP_HELPER=1")
	configureLibreOfficeCommand(cmd)
	cmd.WaitDelay = time.Second
	stdout, err := cmd.StdoutPipe()
	if err != nil {
		t.Fatal(err)
	}
	if err := cmd.Start(); err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() {
		cancel()
		if cmd.ProcessState == nil {
			if err := cmd.Wait(); err == nil {
				t.Error("hanging helper exited without cancellation error")
			}
		}
	})
	line, err := bufio.NewReader(stdout).ReadString('\n')
	if err != nil {
		t.Fatalf("read helper process group: %v", err)
	}
	group, err := strconv.Atoi(strings.TrimSpace(line))
	if err != nil {
		t.Fatal(err)
	}
	if group != cmd.Process.Pid || group == syscall.Getpgrp() {
		t.Fatalf("conversion group = %d, process = %d, caller group = %d", group, cmd.Process.Pid, syscall.Getpgrp())
	}
	cancel()
	if err := cmd.Wait(); err == nil {
		t.Fatal("canceled process group exited successfully")
	}
}

func TestLibreOfficeProcessGroupHelper(t *testing.T) {
	if os.Getenv("EXCEL2PDF_GROUP_HELPER") != "1" {
		return
	}
	if _, err := fmt.Fprintln(os.Stdout, syscall.Getpgrp()); err != nil {
		t.Fatal(err)
	}
	for {
		time.Sleep(time.Hour)
	}
}
