"""Private newline-JSON bridge. Only the owning Go process writes requests."""
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import uuid


def reply(ok, error=""):
    print(json.dumps({"ok": ok, "error": error}), flush=True)


def main():
    binary, profile = sys.argv[1:]
    # Official LibreOffice tarball/DEB installations may ship their own uno.py.
    program = os.path.dirname(os.path.realpath(binary))
    if os.path.isfile(os.path.join(program, "uno.py")):
        sys.path.insert(0, program)
        fundamental = os.path.join(program, "fundamentalrc")
        if os.path.isfile(fundamental):
            os.environ["URE_BOOTSTRAP"] = Path(fundamental).as_uri()
    import uno
    import unohelper
    from com.sun.star.task import XInteractionHandler
    from com.sun.star.document.MacroExecMode import NEVER_EXECUTE
    from com.sun.star.document.UpdateDocMode import NO_UPDATE

    class RejectInteraction(unohelper.Base, XInteractionHandler):
        def handle(self, request):
            abort_type = uno.getTypeByName("com.sun.star.task.XInteractionAbort")
            for continuation in request.getContinuations():
                abort = continuation.queryInterface(abort_type)
                if abort is not None:
                    abort.select()
                    return

    def prop(name, value):
        p = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        p.Name, p.Value = name, value
        return p

    pipe = "excel2pdf_" + uuid.uuid4().hex
    process = subprocess.Popen(
        [binary, "-env:UserInstallation=" + uno.systemPathToFileUrl(profile),
         "--headless", "--nologo", "--nodefault", "--nofirststartwizard",
         "--accept=pipe,name=" + pipe + ";urp;StarOffice.ComponentContext"],
        stdin=subprocess.DEVNULL, stdout=sys.stderr, stderr=sys.stderr,
    )
    desktop = None
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local)
        deadline = time.monotonic() + 25
        while True:
            try:
                remote = resolver.resolve(
                    "uno:pipe,name=" + pipe + ";urp;StarOffice.ComponentContext")
                break
            except Exception:
                if process.poll() is not None or time.monotonic() >= deadline:
                    raise RuntimeError("LibreOffice UNO listener did not become ready")
                time.sleep(0.05)
        desktop = remote.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", remote)
        reply(True)
        for line in sys.stdin:
            request = json.loads(line)
            document = None
            try:
                document = desktop.loadComponentFromURL(
                    uno.systemPathToFileUrl(request["input"]), "_blank", 0,
                    (prop("Hidden", True), prop("ReadOnly", True),
                     prop("MacroExecutionMode", uno.Any("short", NEVER_EXECUTE)),
                     prop("UpdateDocMode", uno.Any("short", NO_UPDATE)),
                     prop("InteractionHandler", RejectInteraction()),
                     prop("AddToRecentDocs", False)))
                if document is None:
                    raise RuntimeError("LibreOffice could not open document")
                if document.supportsService("com.sun.star.text.TextDocument"):
                    export_filter = "writer_pdf_Export"
                elif document.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
                    export_filter = "calc_pdf_Export"
                else:
                    raise RuntimeError("document is not a Writer or Calc document")
                document.storeToURL(uno.systemPathToFileUrl(request["output"]),
                                    (prop("FilterName", export_filter),
                                     prop("Overwrite", True)))
                document.close(True)
                document = None
                reply(True)
            except Exception as error:
                # No error text from a document is interpreted as a command.
                if document is not None:
                    document.dispose()
                reply(False, str(error)[:8192])
    finally:
        if desktop is not None:
            desktop.terminate()
        try:
            process.wait(timeout=3)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait()


try:
    main()
except Exception as error:
    reply(False, str(error)[:8192])
    sys.exit(1)
