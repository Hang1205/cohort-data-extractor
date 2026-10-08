"""Shared desktop entry-point support. Diagnostic reports exclude inputs and note text."""
from datetime import datetime,timezone
import json
from pathlib import Path
import platform
import sys
import tempfile
import traceback
import tkinter as tk
from tkinter import messagebox


def home():
    return Path(sys.executable).resolve().parent if getattr(sys,'frozen',False) else Path(__file__).resolve().parent


def safe_diagnostic(error):
    report={'time_utc':datetime.now(timezone.utc).isoformat(),'error_type':type(error).__name__,
            'python_version':platform.python_version(),'frozen':bool(getattr(sys,'frozen',False)),
            'frames':[{'module':Path(frame.filename).name,'function':frame.name,'line':frame.lineno} for frame in traceback.extract_tb(error.__traceback__)]}
    # Never include str(error), repr(error), local variables, note text or selected IDs.
    for parent in (home()/'Diagnostics',Path(tempfile.gettempdir())/'Source_Tools_Diagnostics'):
        try:
            parent.mkdir(parents=True,exist_ok=True)
            file=parent/('startup_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')+'.json')
            file.write_text(json.dumps(report,indent=2),encoding='utf-8')
            return file
        except OSError:
            continue
    return None


def desktop_main(factory):
    root=None
    try:
        root=tk.Tk()
        def handle_callback(kind,error,tb):
            error=error.with_traceback(tb)
            log=safe_diagnostic(error)
            messagebox.showerror('Operation stopped','An unexpected desktop error occurred. No id data was included in the diagnostic.'+
                                 ('\nDiagnostic: '+str(log) if log else '\nThe diagnostic could not be saved.'),parent=root)
        root.report_callback_exception=handle_callback
        factory(root)
        root.mainloop()
        return 0
    except Exception as error:
        log=safe_diagnostic(error)
        try:
            messagebox.showerror('Cannot start the desktop app','Extract the ENTIRE desktop ZIP to a local folder and keep _internal/ beside the EXE.'+
                                 ('\nDiagnostic: '+str(log) if log else ''))
        except Exception:
            pass
        return 1
    finally:
        if root is not None:
            try:root.destroy()
            except tk.TclError:pass


def run_self_test(test):
    # This mode processes only built-in fictional fixtures. Its test failures may
    # include module names to diagnose packaging; normal-use diagnostics never do.
    result={'app_version':'1.0.0','status':'failed','frozen':bool(getattr(sys,'frozen',False))}
    code=1
    try:
        result.update(test());result['status']='passed';code=0
    except Exception as error:
        result.update(error_type=type(error).__name__,test_error=str(error),traceback=traceback.format_exc())
    if '--report' in sys.argv:
        file=Path(sys.argv[sys.argv.index('--report')+1])
    else:
        file=home()/'Desktop_Self_Test.json'
    file.parent.mkdir(parents=True,exist_ok=True)
    file.write_text(json.dumps(result,indent=2),encoding='utf-8')
    return code
