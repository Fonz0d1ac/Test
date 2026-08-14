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
    Legacy fallback — the reliable path is print_html_silent() below."""
    import win32api
    win32api.ShellExecute(0, "printto", pdf_path, f'"{printer_name}"', ".", 0)


# ── Silent BPT path: HTML → PDF (headless Chromium) → named printer (Sumatra) ──
# Windows can render a PDF but cannot send one to a *named* printer without UI,
# so the paper side needs two steps. Headless Edge/Chrome is already on every
# Windows PC; SumatraPDF is a single portable .exe (no installer, no admin) that
# takes `-print-to "<queue>"`. Either piece missing → we say so and the caller
# falls back to the browser-tab print dialog, so printing never hard-fails.

import os
import pathlib
import shutil
import subprocess
import tempfile

_BROWSER_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]

_SUMATRA_CANDIDATES = [
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "tools", "SumatraPDF.exe"),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "SumatraPDF.exe"),
    os.path.expandvars(r"%LOCALAPPDATA%\SumatraPDF\SumatraPDF.exe"),
    r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
    r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
]


def find_browser(explicit=""):
    """Path to a headless-capable Chromium (Edge first — always present on Win10+)."""
    for p in ([explicit] if explicit else []) + _BROWSER_CANDIDATES:
        if p and os.path.exists(p):
            return p
    for exe in ("msedge", "chrome"):
        p = shutil.which(exe)
        if p:
            return p
    return ""


def find_sumatra(explicit=""):
    """Path to SumatraPDF.exe (the PDF → named-printer step)."""
    for p in ([explicit] if explicit else []) + _SUMATRA_CANDIDATES:
        if p and os.path.exists(p):
            return p
    return shutil.which("SumatraPDF") or ""


def silent_ready(sumatra_path="", browser_path=""):
    """(ok, detail) — whether the fully silent BPT path can run on this PC."""
    b, s = find_browser(browser_path), find_sumatra(sumatra_path)
    if not b:
        return False, "no headless browser found (Edge/Chrome)"
    if not s:
        return False, "SumatraPDF.exe not found — put it in a 'tools' folder next to app.py"
    return True, f"{os.path.basename(b)} + {os.path.basename(s)}"


# Dedicated, PERSISTENT browser profile. It must not be the operator's own Edge
# profile — headless would attach to their running instance and silently produce
# nothing — but making a fresh one per print is expensive on Windows, where every
# newly created file gets antivirus-scanned. Created once, reused forever.
_PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "data", "_bpt_browser_profile")

# Startup work we don't need for a one-shot local render.
_LEAN_FLAGS = [
    "--headless=new", "--disable-gpu", "--no-first-run", "--no-sandbox",
    "--no-pdf-header-footer",
    "--disable-extensions", "--disable-background-networking", "--disable-sync",
    "--disable-default-apps", "--no-default-browser-check",
    "--disable-component-update", "--metrics-recording-only", "--mute-audio",
    # The page is entirely local (data: URI barcodes), so it settles immediately;
    # the budget is just a ceiling so a slow PC can't emit a half-drawn page.
    # --run-all-compositor-stages-before-draw was dropped: it's for screenshots
    # and only added latency here.
    "--virtual-time-budget=2000",
]


def html_to_pdf(html_text, pdf_path, browser_path="", timeout=90):
    """Render HTML to PDF with headless Chromium. Returns the pdf path."""
    browser = find_browser(browser_path)
    if not browser:
        raise RuntimeError("no headless browser found (Edge/Chrome)")
    tmpdir = tempfile.mkdtemp(prefix="bpt_")
    html_path = os.path.join(tmpdir, "ticket.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_text)
    os.makedirs(_PROFILE_DIR, exist_ok=True)
    cmd = ([browser] + _LEAN_FLAGS + [f"--user-data-dir={_PROFILE_DIR}",
           f"--print-to-pdf={pdf_path}", pathlib.Path(html_path).as_uri()])
    r = subprocess.run(cmd, capture_output=True, timeout=timeout)
    if not os.path.exists(pdf_path) or os.path.getsize(pdf_path) == 0:
        err = (r.stderr or b"").decode("utf-8", "replace").strip()[:300]
        raise RuntimeError(f"headless render produced no PDF ({err or 'no output'})")
    return pdf_path


def prewarm(browser_path=""):
    """Render a trivial page to PDF and throw it away.

    The first headless launch after boot is dramatically slower than later ones —
    measured 12.2s cold vs 0.65s warm — because the browser binary isn't in the OS
    file cache and the profile doesn't exist yet. Paying that once at startup, and
    again when the operator opens the Issue modal, keeps it off the print click.
    Silent and best-effort: failure here must never affect printing."""
    try:
        tmp = os.path.join(tempfile.mkdtemp(prefix="bptwarm_"), "w.pdf")
        html_to_pdf("<html><body>warm</body></html>", tmp, browser_path, timeout=120)
        shutil.rmtree(os.path.dirname(tmp), ignore_errors=True)
        return True
    except Exception:
        return False


def print_pdf_silent(pdf_path, printer_name, sumatra_path="", timeout=120,
                     settings="simplex,noscale"):
    """Send a PDF to a named Windows queue with no dialog, via SumatraPDF.

    `simplex` forces SINGLE-SIDED regardless of the queue's duplex default — a BPT
    is one ticket per box and must never share a sheet with the next box's ticket.
    `noscale` keeps the A5 layout at true size instead of shrink-to-fit."""
    sumatra = find_sumatra(sumatra_path)
    if not sumatra:
        raise RuntimeError("SumatraPDF.exe not found")
    cmd = [sumatra, "-print-to", printer_name]
    if settings:
        cmd += ["-print-settings", settings]
    cmd += ["-silent", "-exit-when-done", pdf_path]
    r = subprocess.run(cmd, capture_output=True, timeout=timeout)
    if r.returncode != 0:
        err = (r.stderr or b"").decode("utf-8", "replace").strip()[:300]
        raise RuntimeError(f"SumatraPDF exit {r.returncode} ({err or 'no message'})")
    return True


def print_html_silent(html_text, printer_name, sumatra_path="", browser_path="",
                      print_settings="simplex,noscale"):
    """Full silent paper path: HTML → PDF → named printer. Returns (ok, detail);
    never raises, so the caller can fall back to the browser dialog."""
    if not printer_name:
        return False, "no Canon printer set in print settings"
    tmp_pdf = os.path.join(tempfile.mkdtemp(prefix="bpt_pdf_"), "ticket.pdf")
    import time as _t
    t0 = _t.time()
    try:
        html_to_pdf(html_text, tmp_pdf, browser_path)
        t1 = _t.time()
        print_pdf_silent(tmp_pdf, printer_name, sumatra_path, settings=print_settings)
        t2 = _t.time()
        # Split so a slow BPT can be blamed on the right half — browser render vs
        # handing the PDF to the Canon queue.
        print(f'[bpt print] render {t1 - t0:.1f}s · spool {t2 - t1:.1f}s → {printer_name}')
        return True, f"printed to {printer_name} (render {t1 - t0:.1f}s, spool {t2 - t1:.1f}s)"
    except Exception as e:
        return False, str(e)
