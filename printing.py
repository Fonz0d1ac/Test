"""
Windows printing helpers for the ticket/label feature.

All win32 imports are LAZY so this module imports cleanly off-Windows (dev/test
box). On the production board PC (Windows + pywin32) these drive the real queues.

- Zebra box label  → print_zpl_raw()  (RAW passthrough of the ZPL text)
- Canon BPT ticket → primarily printed from the browser (window.print on the
  approved HTML), so it needs no server-side path; print_pdf_to() is provided for
  a future silent path once a PDF is produced on the PC.
"""


def available():
    """True if this process can talk to Windows printers (pywin32 present)."""
    try:
        import win32print  # noqa: F401
        return True
    except Exception:
        return False


def list_printers():
    """Names of installed printers (local + network). [] off-Windows."""
    try:
        import win32print
    except Exception:
        return []
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    try:
        return [p[2] for p in win32print.EnumPrinters(flags, None, 1)]
    except Exception:
        return []


def print_zpl_raw(zpl, printer_name, doc="Box label"):
    """Send raw ZPL bytes to a Windows print queue (RAW datatype = passthrough).
    Raises on failure so the caller can report it."""
    import win32print
    data = zpl.encode("utf-8") if isinstance(zpl, str) else zpl
    h = win32print.OpenPrinter(printer_name)
    try:
        win32print.StartDocPrinter(h, 1, (doc, None, "RAW"))
        win32print.StartPagePrinter(h)
        win32print.WritePrinter(h, data)
        win32print.EndPagePrinter(h)
        win32print.EndDocPrinter(h)
    finally:
        win32print.ClosePrinter(h)


def print_pdf_to(printer_name, pdf_path):
    """Best-effort silent PDF print to a named printer via the registered
    handler's 'printto' verb (Adobe/Foxit/SumatraPDF support it; Edge is spotty).
    Not used by the default flow (BPT prints from the browser) — here for later."""
    import win32api
    win32api.ShellExecute(0, "printto", pdf_path, f'"{printer_name}"', ".", 0)
