//go:build integration

package excel2pdf

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"testing"
	"time"
	"unsafe"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"golang.org/x/sys/windows"
)

// Exercise explicit Office launch instead of COM activation: the process is
// born inside the worker's job, even if acquiring the automation object hangs.
type ownershipMessage struct {
	PID           uint32
	Thread        uint32
	ExportStarted bool
}

const ownershipHelper = "EXCEL2PDF_OWNERSHIP_TEST_HELPER"

func TestOfficeOwnershipHelper(t *testing.T) {
	if os.Getenv(ownershipHelper) != "1" {
		t.Skip("child process only")
	}
	if err := runOwnershipHelper(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
	os.Exit(0)
}

func runOwnershipHelper() error {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	if err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED); !officeCOMInitialized(err) {
		return err
	}
	defer ole.CoUninitialize()
	bin := os.Getenv("EXCEL2PDF_OWNERSHIP_BIN")
	input := os.Getenv("EXCEL2PDF_OWNERSHIP_INPUT")
	kind, err := documentKindForPath(input)
	if err != nil {
		return err
	}
	args := []string{"/w", "/q"}
	class := "_WwG"
	export := exportDocument
	if kind == spreadsheetDocument {
		// A blank trusted fixture creates the EXCEL7 document window without
		// opening the document under test before automation security is set.
		args = []string{"/x", os.Getenv("EXCEL2PDF_OWNERSHIP_BLANK")}
		class, export = "EXCEL7", exportWorkbook
	}
	process := exec.Command(bin, args...)
	if err := process.Start(); err != nil {
		return err
	}
	defer process.Process.Release()
	pid := uint32(process.Process.Pid)
	enc := json.NewEncoder(os.Stdout)
	if err := enc.Encode(ownershipMessage{PID: pid}); err != nil {
		return err
	}
	var hwnd windows.HWND
	deadline := time.Now().Add(20 * time.Second)
	for hwnd == 0 && time.Now().Before(deadline) {
		hwnd = ownedOfficeWindow(pid, class)
		if hwnd == 0 {
			time.Sleep(50 * time.Millisecond)
		}
	}
	if hwnd == 0 {
		return fmt.Errorf("no %s window for launched PID %d", class, pid)
	}
	var window *ole.IDispatch
	getObject := windows.NewLazySystemDLL("oleacc.dll").NewProc("AccessibleObjectFromWindow")
	hr, _, _ := getObject.Call(uintptr(hwnd), 0xfffffff0, uintptr(unsafe.Pointer(ole.IID_IDispatch)), uintptr(unsafe.Pointer(&window)))
	if int32(hr) < 0 || window == nil {
		return fmt.Errorf("AccessibleObjectFromWindow: %#x", hr)
	}
	defer window.Release()
	app, err := oleutil.GetProperty(window, "Application")
	if err != nil {
		return err
	}
	defer app.Clear()
	dispatch := app.ToIDispatch()
	if dispatch == nil {
		return errors.New("no Application dispatch")
	}
	var owner uint32
	thread, err := windows.GetWindowThreadProcessId(hwnd, &owner)
	if err != nil || owner != pid {
		return fmt.Errorf("Office window changed owner: %d, %v", owner, err)
	}
	if err := enc.Encode(ownershipMessage{PID: pid, Thread: thread}); err != nil {
		return err
	}
	dec := json.NewDecoder(os.Stdin)
	for {
		var output string
		if err := dec.Decode(&output); err != nil {
			return err
		}
		if err := enc.Encode(ownershipMessage{ExportStarted: true}); err != nil {
			return err
		}
		if err := export(context.Background(), &comOfficeObject{dispatch: dispatch}, input, output); err != nil {
			return err
		}
		if err := enc.Encode(ownershipMessage{PID: pid}); err != nil {
			return err
		}
	}
}

// The callbacks are created once; syscall callback slots are process-lifetime.
type windowSearch struct {
	pid   uint32
	class string
	found windows.HWND
}

var inspectOfficeChild = windows.NewCallback(func(hwnd windows.HWND, search *windowSearch) uintptr {
	var pid uint32
	_, err := windows.GetWindowThreadProcessId(hwnd, &pid)
	if err == nil && pid == search.pid {
		var name [128]uint16
		n, err := windows.GetClassName(hwnd, &name[0], int32(len(name)))
		if err == nil && windows.UTF16ToString(name[:n]) == search.class {
			search.found = hwnd
			return 0
		}
	}
	return 1
})

var inspectOfficeTop = windows.NewCallback(func(hwnd windows.HWND, search *windowSearch) uintptr {
	var pid uint32
	_, err := windows.GetWindowThreadProcessId(hwnd, &pid)
	if err == nil && pid == search.pid {
		windows.EnumChildWindows(hwnd, inspectOfficeChild, unsafe.Pointer(search))
	}
	if search.found != 0 {
		return 0
	}
	return 1
})

func ownedOfficeWindow(pid uint32, class string) windows.HWND {
	search := windowSearch{pid: pid, class: class}
	_ = windows.EnumWindows(inspectOfficeTop, unsafe.Pointer(&search))
	return search.found
}

type ownershipProbe struct {
	input     io.WriteCloser
	responses chan ownershipMessage
	stop      func()
	process   windows.Handle
	thread    windows.Handle
}

func startOwnershipProbe(t *testing.T, bin, input string, stopAtLaunch bool) *ownershipProbe {
	t.Helper()
	var blank string
	if filepath.Ext(input) == ".xlsx" {
		blank = writeIntegrationWorkbook(t, "bootstrap.xlsx")
	}
	cmd := exec.Command(os.Args[0], "-test.run=^TestOfficeOwnershipHelper$")
	cmd.Env = append(os.Environ(), ownershipHelper+"=1", "EXCEL2PDF_OWNERSHIP_BIN="+bin, "EXCEL2PDF_OWNERSHIP_INPUT="+input, "EXCEL2PDF_OWNERSHIP_BLANK="+blank)
	cmd.Stderr = os.Stderr
	in, err := cmd.StdinPipe()
	if err != nil {
		t.Fatal(err)
	}
	out, err := cmd.StdoutPipe()
	if err != nil {
		t.Fatal(err)
	}
	terminate, err := startBridgeProcess(cmd)
	if err != nil {
		t.Fatal(err)
	}
	p := &ownershipProbe{input: in, responses: make(chan ownershipMessage, 4)}
	go func() {
		defer close(p.responses)
		dec := json.NewDecoder(out)
		for {
			var msg ownershipMessage
			if dec.Decode(&msg) != nil {
				return
			}
			p.responses <- msg
		}
	}()
	stopped := false
	p.stop = func() {
		if stopped {
			return
		}
		stopped = true
		if err := terminate(); err != nil {
			t.Error(err)
		}
		_ = in.Close()
		_ = cmd.Wait() // an externally terminated helper exits non-zero
	}
	t.Cleanup(p.stop)
	msg := p.receive(t)
	p.process, err = windows.OpenProcess(windows.SYNCHRONIZE, false, msg.PID)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { windows.CloseHandle(p.process) })
	if !stopAtLaunch {
		ready := p.receive(t)
		if ready.PID != msg.PID || ready.Thread == 0 {
			t.Fatalf("invalid ownership reply: %+v", ready)
		}
		p.thread, err = windows.OpenThread(windows.THREAD_SUSPEND_RESUME, false, ready.Thread)
		if err != nil {
			t.Fatal(err)
		}
		t.Cleanup(func() { windows.CloseHandle(p.thread) })
	}
	return p
}

func (p *ownershipProbe) receive(t *testing.T) ownershipMessage {
	t.Helper()
	select {
	case msg, ok := <-p.responses:
		if !ok {
			t.Fatal("Office helper exited before replying")
		}
		return msg
	case <-time.After(30 * time.Second):
		t.Fatal("Office helper timed out")
	}
	return ownershipMessage{}
}

func (p *ownershipProbe) request(t *testing.T) string {
	t.Helper()
	path := filepath.Join(t.TempDir(), "result.pdf")
	if err := json.NewEncoder(p.input).Encode(path); err != nil {
		t.Fatal(err)
	}
	if msg := p.receive(t); !msg.ExportStarted {
		t.Fatal("worker did not acknowledge the export")
	}
	return path
}

func (p *ownershipProbe) convert(t *testing.T) {
	t.Helper()
	path := p.request(t)
	p.receive(t)
	assertIntegrationPDF(t, path)
}

func (p *ownershipProbe) stopAndCheck(t *testing.T) {
	t.Helper()
	start := time.Now()
	p.stop()
	state, err := windows.WaitForSingleObject(p.process, 5000)
	if err != nil || state != windows.WAIT_OBJECT_0 {
		t.Fatalf("Office survived job termination: %d, %v", state, err)
	}
	elapsed := time.Since(start)
	t.Logf("terminated owned worker and Office in %v", elapsed)
	if elapsed > 5*time.Second {
		t.Fatalf("termination exceeded test budget: %v", elapsed)
	}
}

func TestIntegrationOfficeProcessOwnership(t *testing.T) {
	if os.Getenv("EXCEL2PDF_TEST_OFFICE_OWNERSHIP") != "1" {
		t.Skip("set EXCEL2PDF_TEST_OFFICE_OWNERSHIP=1 to exercise forced termination")
	}
	for _, app := range []string{"WINWORD.EXE", "EXCEL.EXE"} {
		t.Run(app, func(t *testing.T) {
			var input string
			if app == "WINWORD.EXE" {
				requireIntegrationWord(t)
				input = writeIntegrationWordDocument(t, "ownership.docx")
			} else {
				requireIntegrationExcel(t)
				input = writeIntegrationWorkbook(t, "ownership.xlsx")
			}
			var bin string
			for _, loc := range officeRegistryLocations() {
				var err error
				bin, err = readOfficeRegistryString(loc.root, `SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\`+app, "", loc.view)
				if err != nil {
					t.Fatal(err)
				}
				if bin != "" {
					break
				}
			}
			if !filepath.IsAbs(bin) {
				t.Fatalf("no registered absolute path for %s", app)
			}
			startOwnershipProbe(t, bin, input, true).stopAndCheck(t)
			survivor := startOwnershipProbe(t, bin, input, false)
			survivor.convert(t)
			hung := startOwnershipProbe(t, bin, input, false)
			hung.convert(t)
			suspend := windows.NewLazySystemDLL("kernel32.dll").NewProc("SuspendThread")
			r, _, err := suspend.Call(uintptr(hung.thread))
			if uint32(r) == 0xffffffff {
				t.Fatal(err)
			}
			output := hung.request(t)
			select {
			case <-hung.responses:
				t.Fatal("suspended Office unexpectedly replied or exited")
			case <-time.After(300 * time.Millisecond):
			}
			hung.stopAndCheck(t)
			if _, err := os.Stat(output); !errors.Is(err, os.ErrNotExist) {
				t.Fatalf("blocked request produced output: %v", err)
			}
			survivor.convert(t)
			fresh := startOwnershipProbe(t, bin, input, false)
			fresh.convert(t)
			fresh.stopAndCheck(t)
			survivor.stopAndCheck(t)
		})
	}
}
