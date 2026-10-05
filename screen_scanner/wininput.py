"""Minimal Windows mouse/keyboard input via SendInput (no dependencies).

Importing is safe on any OS; calling the input functions off Windows raises.
"""
import ctypes
import sys
import time
from ctypes import wintypes

INPUT_MOUSE, INPUT_KEYBOARD = 0, 1
MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP, MOUSEEVENTF_WHEEL = 0x0002, 0x0004, 0x0800
KEYEVENTF_KEYUP = 0x0002
WHEEL_DELTA = 120
VK_BACK, VK_END = 0x08, 0x23

user32 = None
if sys.platform == 'win32':
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    # Real pixel coordinates even with Windows display scaling (125%, 150%...)
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        user32.SetProcessDPIAware()


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [('dx', wintypes.LONG), ('dy', wintypes.LONG),
                ('mouseData', wintypes.LONG),  # DWORD in the API; signed here so -120 works
                ('dwFlags', wintypes.DWORD), ('time', wintypes.DWORD),
                ('dwExtraInfo', ctypes.c_size_t)]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [('wVk', wintypes.WORD), ('wScan', wintypes.WORD),
                ('dwFlags', wintypes.DWORD), ('time', wintypes.DWORD),
                ('dwExtraInfo', ctypes.c_size_t)]


class INPUT(ctypes.Structure):
    class _U(ctypes.Union):
        _fields_ = [('mi', MOUSEINPUT), ('ki', KEYBDINPUT)]
    _anonymous_ = ('u',)
    _fields_ = [('type', wintypes.DWORD), ('u', _U)]


def _send(*inputs: INPUT) -> None:
    if user32 is None:
        raise RuntimeError('Input simulation only works on Windows.')
    arr = (INPUT * len(inputs))(*inputs)
    if user32.SendInput(len(inputs), arr, ctypes.sizeof(INPUT)) != len(inputs):
        raise RuntimeError('SendInput was blocked. If the game runs as administrator, '
                           'run this script from an administrator terminal too.')


def _mouse(flags: int, data: int = 0) -> INPUT:
    inp = INPUT(type=INPUT_MOUSE)
    inp.mi = MOUSEINPUT(0, 0, data, flags, 0, 0)
    return inp


def _key(vk: int, up: bool = False) -> INPUT:
    inp = INPUT(type=INPUT_KEYBOARD)
    inp.ki = KEYBDINPUT(vk, user32.MapVirtualKeyW(vk, 0), KEYEVENTF_KEYUP if up else 0, 0, 0)
    return inp


def cursor_pos() -> tuple:
    pt = wintypes.POINT()
    user32.GetCursorPos(ctypes.byref(pt))
    return pt.x, pt.y


def move(x: int, y: int) -> None:
    user32.SetCursorPos(int(x), int(y))


def click(x: int, y: int) -> None:
    move(x, y)
    time.sleep(0.05)
    _send(_mouse(MOUSEEVENTF_LEFTDOWN))
    time.sleep(0.05)
    _send(_mouse(MOUSEEVENTF_LEFTUP))


def wheel(notches: int) -> None:
    """Positive = up (away from you), negative = down."""
    _send(_mouse(MOUSEEVENTF_WHEEL, notches * WHEEL_DELTA))


def press(vk: int, times: int = 1) -> None:
    for _ in range(times):
        _send(_key(vk))
        time.sleep(0.02)
        _send(_key(vk, up=True))
        time.sleep(0.02)


def type_digits(text: str) -> None:
    for ch in text:
        press(ord(ch))  # VK codes for 0-9 are their ASCII codes


def type_keys(vks: list, delay: float = 0.0) -> None:
    """Press and release each key in order. delay = pause between events;
    0 sends the whole sequence in a single call (fastest)."""
    events = [e for vk in vks for e in (_key(vk), _key(vk, up=True))]
    if not delay:
        _send(*events)
        return
    for e in events:
        _send(e)
        time.sleep(delay)


def client_rect(title: str) -> tuple:
    """Screen position and size of a window's drawable area: (left, top, width, height)."""
    if user32 is None:
        raise RuntimeError('Window lookup only works on Windows.')
    hwnd = user32.FindWindowW(None, title)
    if not hwnd:
        raise RuntimeError(f'No window titled "{title}" - is the game open?')
    if user32.IsIconic(hwnd):
        raise RuntimeError(f'"{title}" is minimized - restore it first.')
    rect = wintypes.RECT()
    user32.GetClientRect(hwnd, ctypes.byref(rect))
    origin = wintypes.POINT(0, 0)
    user32.ClientToScreen(hwnd, ctypes.byref(origin))
    return origin.x, origin.y, rect.right, rect.bottom
