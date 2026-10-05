"""Send mouse-wheel-down notches to whatever window is under the cursor (Windows).

  python zoom_test.py                  5 notches down after a 3 s countdown
  python zoom_test.py --count 3 --up   3 notches up instead
  python zoom_test.py --at 960 540     move the cursor there first
  python zoom_test.py --shots          also save a screenshot after each notch

During the countdown, click the game window and park the cursor on open
ground. --shots saves zoom_0.png (before) .. zoom_N.png into zoom_shots/ so
you can check whether every notch zooms by the same amount.
"""
import argparse
import ctypes
import sys
import time
from ctypes import wintypes
from pathlib import Path

if sys.platform != 'win32':
    sys.exit('zoom_test.py only runs on Windows.')

INPUT_MOUSE = 0
MOUSEEVENTF_WHEEL = 0x0800
WHEEL_DELTA = 120

user32 = ctypes.WinDLL('user32', use_last_error=True)


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [('dx', wintypes.LONG), ('dy', wintypes.LONG),
                ('mouseData', wintypes.LONG),  # DWORD in the API; signed here so -120 works
                ('dwFlags', wintypes.DWORD), ('time', wintypes.DWORD),
                ('dwExtraInfo', ctypes.c_size_t)]


class INPUT(ctypes.Structure):
    class _U(ctypes.Union):
        _fields_ = [('mi', MOUSEINPUT)]  # largest member, so the size is right
    _anonymous_ = ('u',)
    _fields_ = [('type', wintypes.DWORD), ('u', _U)]


def wheel(notches: int) -> bool:
    """One wheel event; positive = up (away from you), negative = down."""
    inp = INPUT(type=INPUT_MOUSE)
    inp.mi = MOUSEINPUT(0, 0, notches * WHEEL_DELTA, MOUSEEVENTF_WHEEL, 0, 0)
    return user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT)) == 1


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--count', type=int, default=5, help='number of notches (default 5)')
    p.add_argument('--up', action='store_true', help='scroll up instead of down')
    p.add_argument('--interval', type=float, default=0.3, help='seconds between notches (default 0.3)')
    p.add_argument('--delay', type=float, default=3, help='countdown before starting (default 3)')
    p.add_argument('--at', type=int, nargs=2, metavar=('X', 'Y'), help='move the cursor here first')
    p.add_argument('--shots', action='store_true', help='save a screenshot before and after each notch')
    args = p.parse_args()

    # Use real pixel coordinates even with Windows display scaling (125%, 150%...)
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        user32.SetProcessDPIAware()

    sct = None
    if args.shots:
        import mss
        import mss.tools
        sct = (getattr(mss, 'MSS', None) or mss.mss)()
        out_dir = Path(__file__).parent / 'zoom_shots'
        out_dir.mkdir(exist_ok=True)

    def shot(i: int) -> None:
        if sct:
            img = sct.grab(sct.monitors[1])
            mss.tools.to_png(img.rgb, img.size, output=str(out_dir / f'zoom_{i}.png'))

    for s in range(int(args.delay), 0, -1):
        print(f'Starting in {s}... (focus the game, cursor on open ground)')
        time.sleep(1)
    time.sleep(args.delay - int(args.delay))

    if args.at:
        user32.SetCursorPos(*args.at)
        time.sleep(0.1)
    pt = wintypes.POINT()
    user32.GetCursorPos(ctypes.byref(pt))
    direction = 'up' if args.up else 'down'
    print(f'Scrolling {direction} {args.count}x at cursor ({pt.x}, {pt.y})')

    shot(0)
    for i in range(1, args.count + 1):
        if not wheel(1 if args.up else -1):
            print('SendInput was blocked. If the game runs as administrator, '
                  'run this script from an administrator terminal too.')
            return 1
        time.sleep(args.interval)
        shot(i)
        print(f'  notch {i}')
    if sct:
        print(f'Screenshots saved in {out_dir}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
