package excel2pdf

import (
	"errors"
	"fmt"
	"os"
	"os/exec"
	"syscall"
	"time"
	"unsafe"

	"golang.org/x/sys/windows"
)

func removeSessionProfile(profile string) error {
	// Windows can briefly retain mappings of the extension cache even after
	// the job has no active processes. Retry only transient sharing failures.
	deadline := time.Now().Add(3 * time.Second)
	for {
		err := os.RemoveAll(profile)
		if !errors.Is(err, windows.ERROR_SHARING_VIOLATION) || time.Now().After(deadline) {
			return err
		}
		time.Sleep(50 * time.Millisecond)
	}
}

// Start suspended: even a bundled Python launcher must not spawn its real
// interpreter before becoming a job member. Descendants inherit membership.
func startBridgeProcess(cmd *exec.Cmd) (func() error, error) {
	job, err := windows.CreateJobObject(nil, nil)
	if err != nil {
		return nil, err
	}
	var limits windows.JOBOBJECT_EXTENDED_LIMIT_INFORMATION
	limits.BasicLimitInformation.LimitFlags = windows.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
	_, err = windows.SetInformationJobObject(job, windows.JobObjectExtendedLimitInformation,
		uintptr(unsafe.Pointer(&limits)), uint32(unsafe.Sizeof(limits)))
	if err != nil {
		return nil, errors.Join(err, windows.CloseHandle(job))
	}
	cmd.SysProcAttr = &syscall.SysProcAttr{CreationFlags: windows.CREATE_SUSPENDED, HideWindow: true}
	if err := cmd.Start(); err != nil {
		return nil, errors.Join(err, windows.CloseHandle(job))
	}
	if err := assignAndResume(job, uint32(cmd.Process.Pid)); err != nil {
		killErr := cmd.Process.Kill()
		_ = cmd.Wait()
		return nil, errors.Join(err, killErr, windows.CloseHandle(job))
	}
	return func() error {
		err := windows.TerminateJobObject(job, 1)
		if err == nil {
			err = waitForEmptyJob(job)
		}
		return errors.Join(err, windows.CloseHandle(job))
	}, nil
}

func waitForEmptyJob(job windows.Handle) error {
	// JOBOBJECT_BASIC_ACCOUNTING_INFORMATION (winnt.h).
	var accounting struct {
		Times                                 [4]int64
		PageFaults, Total, Active, Terminated uint32
	}
	deadline := time.Now().Add(5 * time.Second)
	for {
		err := windows.QueryInformationJobObject(job, windows.JobObjectBasicAccountingInformation,
			uintptr(unsafe.Pointer(&accounting)), uint32(unsafe.Sizeof(accounting)), nil)
		if err != nil || accounting.Active == 0 {
			return err
		}
		if time.Now().After(deadline) {
			return fmt.Errorf("%d bridge job processes still terminating", accounting.Active)
		}
		time.Sleep(10 * time.Millisecond)
	}
}

func assignAndResume(job windows.Handle, pid uint32) error {
	process, err := windows.OpenProcess(windows.PROCESS_SET_QUOTA|windows.PROCESS_TERMINATE, false, pid)
	if err != nil {
		return err
	}
	defer windows.CloseHandle(process)
	if err := windows.AssignProcessToJobObject(job, process); err != nil {
		return err
	}
	snapshot, err := windows.CreateToolhelp32Snapshot(windows.TH32CS_SNAPTHREAD, 0)
	if err != nil {
		return err
	}
	defer windows.CloseHandle(snapshot)
	var entry windows.ThreadEntry32
	entry.Size = uint32(unsafe.Sizeof(entry))
	for err = windows.Thread32First(snapshot, &entry); err == nil; err = windows.Thread32Next(snapshot, &entry) {
		if entry.OwnerProcessID == pid {
			thread, err := windows.OpenThread(windows.THREAD_SUSPEND_RESUME, false, entry.ThreadID)
			if err != nil {
				return err
			}
			defer windows.CloseHandle(thread)
			_, err = windows.ResumeThread(thread)
			return err
		}
	}
	return fmt.Errorf("find suspended bridge thread: %w", err)
}
