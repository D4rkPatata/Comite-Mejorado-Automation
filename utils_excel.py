import gc
import os
import time
from contextlib import contextmanager

import pythoncom
import pywintypes
import win32api
import win32event
import win32con
import win32process
import win32com.client as win32


# Excel answers these while it is busy (e.g. still finishing PowerPoint /
# ChartData work from a previous macro). The call was not executed, so it is
# safe to retry.
RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846


def _retry_busy(func, *args, retries=60, delay=2):
    for attempt in range(retries):
        try:
            return func(*args)
        except pywintypes.com_error as e:
            if e.hresult not in (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER) or attempt == retries - 1:
                raise
            if attempt == 0:
                print("  Excel is busy, retrying...")
            time.sleep(delay)


def _wait_process_exit(pid, timeout=15):
    # Excel keeps running a few seconds after Quit(); a new DispatchEx in that
    # window attaches to the dying instance. Wait for it, kill it if it hangs.
    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, pid)
    except Exception:
        return
    try:
        if win32event.WaitForSingleObject(handle, timeout * 1000) != win32event.WAIT_OBJECT_0:
            win32api.TerminateProcess(handle, 1)
    finally:
        win32api.CloseHandle(handle)


@contextmanager
def excel_app(visible=False):
    """
    Opens a separate, hidden Excel instance (does not touch the user's open
    Excel) and always closes it, even on error.
    """
    pythoncom.CoInitialize()
    app = win32.DispatchEx("Excel.Application")
    _, pid = win32process.GetWindowThreadProcessId(app.Hwnd)
    app.Visible = visible
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    try:
        yield app
    finally:
        try:
            app.Quit()
        except Exception:
            pass
        del app
        gc.collect()
        _wait_process_exit(pid)


def _refresh_open_workbook(app, wb):
    # Background refresh makes RefreshAll return before the data arrives;
    # turn it off so each connection (SQL, Power Query, ...) finishes in order.
    for conn in wb.Connections:
        for kind in ("OLEDBConnection", "ODBCConnection"):
            try:
                getattr(conn, kind).BackgroundQuery = False
            except Exception:
                pass

    wb.RefreshAll()
    app.CalculateUntilAsyncQueriesDone()

    # Pivots fed by Power Query tables may have refreshed before their query.
    for sheet in wb.Worksheets:
        for pivot in sheet.PivotTables():
            pivot.PivotCache().Refresh()
    app.CalculateFull()


def refresh_workbook(file_path, save=True):
    """
    Refreshes all data in an .xlsx / .xlsm: connections, Power Query queries
    and pivot tables, then saves it.

    Args:
        file_path (str): Path to the workbook.
        save (bool): Save the workbook after refreshing.
    """
    file_path = os.path.abspath(file_path)
    with excel_app() as app:
        wb = app.Workbooks.Open(file_path, UpdateLinks=0)
        try:
            _refresh_open_workbook(app, wb)
            if save:
                wb.Save()
        finally:
            wb.Close(SaveChanges=False)
    print(f"Refreshed: {file_path}")


def refresh_workbooks(file_paths, save=True):
    """
    Refreshes several workbooks, one after another.

    Args:
        file_paths (list): Paths to .xlsx / .xlsm files.
        save (bool): Save each workbook after refreshing.

    Returns:
        dict: path -> None if it worked, or the error message.
    """
    results = {}
    for path in file_paths:
        try:
            refresh_workbook(path, save)
            results[path] = None
        except Exception as e:
            print(f"Error refreshing {path}: {e}")
            results[path] = str(e)
    return results


def run_macros(file_path, macros, save=True):
    """
    Runs one or more VBA macros of an .xlsm, in order, in a single session.

    Args:
        file_path (str): Path to the .xlsm.
        macros (str | list): Macro name, e.g. "Module1.Actualizar", or a list
            of names / (name, arg1, arg2, ...) tuples for macros with arguments.
        save (bool): Save the workbook after running the macros.

    Returns:
        list: The value returned by each macro (None for Subs).
    """
    file_path = os.path.abspath(file_path)
    if isinstance(macros, (str, tuple)):
        macros = [macros]

    results = []
    with excel_app() as app:
        wb = app.Workbooks.Open(file_path, UpdateLinks=0)
        # Resolve everything now: while Excel is busy after a macro, even an
        # attribute lookup fails (pywin32 raises it as AttributeError).
        wb_name, wb_save, wb_close = wb.Name, wb.Save, wb.Close
        try:
            for macro in macros:
                name, *args = (macro,) if isinstance(macro, str) else macro
                print(f"Running macro: {name}")
                results.append(_retry_busy(app.Run, f"'{wb_name}'!{name}", *args))
            if save:
                _retry_busy(wb_save)
        finally:
            _retry_busy(wb_close, False)
    return results


def write_cells(file_path, sheet, values, save=True):
    """
    Writes values into cells of a workbook (keeps macros, formats, queries).

    Args:
        file_path (str): Path to the workbook.
        sheet (str): Sheet name.
        values (dict): Cell -> value, e.g. {"O2": 41}.
        save (bool): Save the workbook after writing.
    """
    file_path = os.path.abspath(file_path)
    with excel_app() as app:
        wb = app.Workbooks.Open(file_path, UpdateLinks=0)
        try:
            ws = wb.Worksheets(sheet)
            for cell, value in values.items():
                ws.Range(cell).Value = value
            if save:
                wb.Save()
        finally:
            wb.Close(SaveChanges=False)
    print(f"Written {values} in {sheet}: {file_path}")
