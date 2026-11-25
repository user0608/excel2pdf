package excel2pdf

import (
	"errors"
	"os"
	"path/filepath"
	"testing"

	"golang.org/x/sys/windows/registry"
)

type officeRegistryTestKey struct {
	root registry.Key
	view uint32
	path string
	name string
}

type officeRegistryTestData struct {
	values map[officeRegistryTestKey]string
	names  map[officeRegistryTestKey][]string
	errors map[officeRegistryTestKey]error
}

func (data officeRegistryTestData) read(root registry.Key, path, name string, view uint32) (string, error) {
	key := officeRegistryTestKey{root: root, view: view, path: path, name: name}
	return data.values[key], data.errors[key]
}

func (data officeRegistryTestData) subkeys(root registry.Key, path string, view uint32) ([]string, error) {
	key := officeRegistryTestKey{root: root, view: view, path: path}
	return data.names[key], data.errors[key]
}

func TestFindLibreOfficeInRegistryLocations(t *testing.T) {
	t.Parallel()
	for i, location := range officeRegistryLocations() {
		t.Run([]string{"user64", "user32", "machine64", "machine32"}[i], func(t *testing.T) {
			t.Parallel()
			dir := t.TempDir()
			executable := writeOfficeTestExecutable(t, dir, "soffice")
			data := officeRegistryTestData{values: map[officeRegistryTestKey]string{
				{root: location.root, view: location.view, path: `SOFTWARE\LibreOffice\UNO\InstallPath`}: dir,
			}}
			path, err := findLibreOfficeInRegistry(data.read, data.subkeys)
			if path != executable || err != nil {
				t.Errorf("registry lookup = (%q, %v), want %q", path, err, executable)
			}
		})
	}
}

func TestFindLibreOfficeInRegistryLegacyAndStalePaths(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	program := filepath.Join(dir, "program")
	if err := os.Mkdir(program, 0o700); err != nil {
		t.Fatal(err)
	}
	executable := writeOfficeTestExecutable(t, program, "soffice")
	root := registry.CURRENT_USER
	view := uint32(registry.WOW64_32KEY)
	versionRoot := `SOFTWARE\LibreOffice\LibreOffice`
	data := officeRegistryTestData{
		values: map[officeRegistryTestKey]string{
			{root: root, view: view, path: `SOFTWARE\LibreOffice\UNO\InstallPath`}: filepath.Join(dir, "stale"),
			{root: root, view: view, path: versionRoot + `\24.8`, name: "Path"}:    dir,
		},
		names: map[officeRegistryTestKey][]string{
			{root: root, view: view, path: versionRoot}: {"24.8"},
		},
	}
	path, err := findLibreOfficeInRegistry(data.read, data.subkeys)
	if path != executable || err != nil {
		t.Errorf("legacy lookup = (%q, %v), want %q", path, err, executable)
	}
}

func TestFindLibreOfficeInRegistryErrors(t *testing.T) {
	t.Parallel()
	denied := errors.New("registry access denied")
	for _, lookupErr := range []error{denied, registry.ErrUnexpectedType} {
		t.Run(lookupErr.Error(), func(t *testing.T) {
			t.Parallel()
			data := officeRegistryTestData{errors: map[officeRegistryTestKey]error{
				{root: registry.CURRENT_USER, view: registry.WOW64_64KEY, path: `SOFTWARE\LibreOffice\UNO\InstallPath`}: lookupErr,
			}}
			path, err := findLibreOfficeInRegistry(data.read, data.subkeys)
			if path != "" || !errors.Is(err, lookupErr) || errors.Is(err, ErrLibreofficeNotInstalled) {
				t.Errorf("registry error was treated as absent: (%q, %v)", path, err)
			}
		})
	}
	data := officeRegistryTestData{}
	if path, err := findLibreOfficeInRegistry(data.read, data.subkeys); path != "" || !errors.Is(err, ErrLibreofficeNotInstalled) {
		t.Errorf("empty registry = (%q, %v)", path, err)
	}
}

func TestFindLibreOfficeInRegistryContinuesAfterQueryError(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	executable := writeOfficeTestExecutable(t, dir, "soffice")
	data := officeRegistryTestData{
		values: map[officeRegistryTestKey]string{
			{root: registry.LOCAL_MACHINE, view: registry.WOW64_32KEY, path: `SOFTWARE\LibreOffice\UNO\InstallPath`}: dir,
		},
		errors: map[officeRegistryTestKey]error{
			{root: registry.CURRENT_USER, view: registry.WOW64_64KEY, path: `SOFTWARE\LibreOffice\UNO\InstallPath`}: errors.New("access denied"),
		},
	}
	if path, err := findLibreOfficeInRegistry(data.read, data.subkeys); path != executable || err != nil {
		t.Errorf("valid alternative registration = (%q, %v), want %q", path, err, executable)
	}
}

func TestFindLibreOfficeRegistryVersionDepthLimit(t *testing.T) {
	t.Parallel()
	data := officeRegistryTestData{}
	subkeys := func(registry.Key, string, uint32) ([]string, error) { return []string{"nested"}, nil }
	if _, err := findLibreOfficeRegistryVersion(officeRegistryLocation{root: registry.CURRENT_USER}, "root", 0, data.read, subkeys); err == nil {
		t.Error("unbounded registry recursion")
	}
}

func TestResolveLibreOfficeInstallation(t *testing.T) {
	t.Parallel()
	dir := t.TempDir()
	executable := writeOfficeTestExecutable(t, dir, "soffice")
	for _, path := range []string{executable, dir} {
		got, err := resolveLibreOfficeInstallation(path)
		if got != executable || err != nil {
			t.Errorf("resolve(%q) = (%q, %v), want %q", path, got, err, executable)
		}
	}
	com := filepath.Join(dir, "soffice.com")
	if err := os.WriteFile(com, []byte("executable placeholder"), 0o600); err != nil {
		t.Fatal(err)
	}
	if got, err := resolveLibreOfficeInstallation(dir); got != com || err != nil {
		t.Errorf("console executable preference = (%q, %v), want %q", got, err, com)
	}
	for _, path := range []string{"", t.TempDir(), filepath.Join(dir, "missing")} {
		if got, err := resolveLibreOfficeInstallation(path); got != "" || err != nil {
			t.Errorf("stale or empty installation = (%q, %v)", got, err)
		}
	}
}

func TestFindExcelRegistration(t *testing.T) {
	t.Parallel()
	const classID = "{00024500-0000-0000-C000-000000000046}"
	const progIDPath = `SOFTWARE\Classes\Excel.Application\CLSID`
	const serverPath = `SOFTWARE\Classes\CLSID\` + classID + `\LocalServer32`
	for i, location := range officeRegistryLocations() {
		t.Run([]string{"user64", "user32", "machine64", "machine32"}[i], func(t *testing.T) {
			t.Parallel()
			data := officeRegistryTestData{values: map[officeRegistryTestKey]string{
				{root: location.root, view: location.view, path: progIDPath}: classID,
				{root: location.root, view: location.view, path: serverPath}: `"C:\Office\EXCEL.EXE" /automation`,
			}}
			if installed, err := findExcelRegistration(data.read); !installed || err != nil {
				t.Errorf("Excel registration = (%v, %v)", installed, err)
			}
		})
	}
	t.Run("per-user ProgID and machine server", func(t *testing.T) {
		t.Parallel()
		data := officeRegistryTestData{values: map[officeRegistryTestKey]string{
			{root: registry.CURRENT_USER, view: registry.WOW64_64KEY, path: progIDPath}:  classID,
			{root: registry.LOCAL_MACHINE, view: registry.WOW64_64KEY, path: serverPath}: "EXCEL.EXE",
		}}
		if installed, err := findExcelRegistration(data.read); !installed || err != nil {
			t.Errorf("merged registration = (%v, %v)", installed, err)
		}
	})
	t.Run("no registration", func(t *testing.T) {
		t.Parallel()
		data := officeRegistryTestData{}
		if installed, err := findExcelRegistration(data.read); installed || err != nil {
			t.Errorf("missing registration = (%v, %v)", installed, err)
		}
	})
	t.Run("incomplete registration", func(t *testing.T) {
		t.Parallel()
		data := officeRegistryTestData{values: map[officeRegistryTestKey]string{
			{root: registry.CURRENT_USER, view: registry.WOW64_64KEY, path: progIDPath}: classID,
		}}
		if installed, err := findExcelRegistration(data.read); installed || err != nil {
			t.Errorf("incomplete registration = (%v, %v)", installed, err)
		}
	})
	t.Run("query errors are returned", func(t *testing.T) {
		t.Parallel()
		denied := errors.New("registry access denied")
		data := officeRegistryTestData{errors: map[officeRegistryTestKey]error{
			{root: registry.CURRENT_USER, view: registry.WOW64_64KEY, path: progIDPath}: denied,
		}}
		if installed, err := findExcelRegistration(data.read); installed || !errors.Is(err, denied) {
			t.Errorf("query failure = (%v, %v)", installed, err)
		}
	})
	t.Run("invalid class ID is returned as an error", func(t *testing.T) {
		t.Parallel()
		data := officeRegistryTestData{values: map[officeRegistryTestKey]string{
			{root: registry.CURRENT_USER, view: registry.WOW64_64KEY, path: progIDPath}: "not-a-guid",
		}}
		if installed, err := findExcelRegistration(data.read); installed || err == nil {
			t.Errorf("invalid COM registration = (%v, %v)", installed, err)
		}
	})
}

func TestFindWordRegistration(t *testing.T) {
	t.Parallel()
	const classID = "{000209FF-0000-0000-C000-000000000046}"
	for i, location := range officeRegistryLocations() {
		t.Run([]string{"user64", "user32", "machine64", "machine32"}[i], func(t *testing.T) {
			t.Parallel()
			data := officeRegistryTestData{values: map[officeRegistryTestKey]string{
				{root: location.root, view: location.view, path: `SOFTWARE\Classes\Word.Application\CLSID`}:              classID,
				{root: location.root, view: location.view, path: `SOFTWARE\Classes\CLSID\` + classID + `\LocalServer32`}: "WINWORD.EXE",
			}}
			if installed, err := findOfficeRegistration("Word.Application", data.read); !installed || err != nil {
				t.Errorf("Word registration = (%v, %v)", installed, err)
			}
			if installed, err := findExcelRegistration(data.read); installed || err != nil {
				t.Errorf("Word registration incorrectly detected as Excel: (%v, %v)", installed, err)
			}
		})
	}
	data := officeRegistryTestData{values: map[officeRegistryTestKey]string{
		{root: registry.CURRENT_USER, view: registry.WOW64_64KEY, path: `SOFTWARE\Classes\Excel.Application\CLSID`}:                                    "{00024500-0000-0000-C000-000000000046}",
		{root: registry.CURRENT_USER, view: registry.WOW64_64KEY, path: `SOFTWARE\Classes\CLSID\{00024500-0000-0000-C000-000000000046}\LocalServer32`}: "EXCEL.EXE",
	}}
	if installed, err := findOfficeRegistration("Word.Application", data.read); installed || err != nil {
		t.Errorf("Excel registration incorrectly detected as Word: (%v, %v)", installed, err)
	}
}
