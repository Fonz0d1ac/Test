"""Lightweight on-screen object scanner.

Watches the screen for reference images ("targets") and reports when each
one appears or disappears. Uses plain OpenCV template matching, so there is
no model to train and nothing heavier than mss + OpenCV to install.

  python screen_scan.py snip NAME       crop a new target from the screen
  python screen_scan.py                 watch the screen continuously
  python screen_scan.py --once          scan once, print results, exit
"""
import argparse
import csv
import sys
import time
from datetime import datetime
from pathlib import Path

import cv2
import mss
import numpy as np

HERE = Path(__file__).parent.resolve()
TARGET_DIR = HERE / 'targets'
LOG_FILE = HERE / 'scan_log.csv'
# mss 10 renamed the entry point; support both
ScreenCapture = getattr(mss, 'MSS', None) or mss.mss


def grab(sct, monitor: int) -> np.ndarray:
    """Screenshot one monitor (0 = all monitors combined) as a BGR image."""
    shot = sct.grab(sct.monitors[monitor])
    return cv2.cvtColor(np.asarray(shot), cv2.COLOR_BGRA2BGR)


def load_targets() -> dict:
    targets = {}
    for f in sorted(TARGET_DIR.glob('*.png')):
        img = cv2.imread(str(f), cv2.IMREAD_GRAYSCALE)
        if img is not None:
            targets[f.stem] = img
    return targets


def match(screen_gray: np.ndarray, target: np.ndarray):
    """Return (score, (x, y)) of the best match; score is 0..1."""
    if target.shape[0] > screen_gray.shape[0] or target.shape[1] > screen_gray.shape[1]:
        return 0.0, (0, 0)
    res = cv2.matchTemplate(screen_gray, target, cv2.TM_CCOEFF_NORMED)
    _, score, _, loc = cv2.minMaxLoc(res)
    return float(score), loc


def beep() -> None:
    try:
        import winsound
        winsound.Beep(1000, 150)
    except Exception:
        print('\a', end='', flush=True)


def log_event(name: str, event: str, score: float, loc) -> None:
    new = not LOG_FILE.exists()
    with open(LOG_FILE, 'a', newline='') as f:
        w = csv.writer(f)
        if new:
            w.writerow(['time', 'target', 'event', 'score', 'x', 'y'])
        w.writerow([datetime.now().isoformat(timespec='seconds'), name, event,
                    f'{score:.3f}', loc[0], loc[1]])


def cmd_snip(args) -> None:
    """Take a screenshot after a short delay and let the user crop a target."""
    TARGET_DIR.mkdir(exist_ok=True)
    print(f'Capturing in {args.delay}s - bring the object on screen...')
    time.sleep(args.delay)
    with ScreenCapture() as sct:
        screen = grab(sct, args.monitor)
    title = 'Drag a box around the object, then ENTER (ESC to cancel)'
    x, y, w, h = cv2.selectROI(title, screen, showCrosshair=False)
    cv2.destroyAllWindows()
    if w == 0 or h == 0:
        print('Cancelled.')
        return
    out = TARGET_DIR / f'{args.name}.png'
    cv2.imwrite(str(out), screen[y:y + h, x:x + w])
    print(f'Saved {out} ({w}x{h})')


def cmd_scan(args) -> int:
    targets = load_targets()
    if not targets:
        print(f'No targets in {TARGET_DIR}. Create one with: python screen_scan.py snip NAME')
        return 2

    if not args.once:
        print(f'Watching monitor {args.monitor} for: {", ".join(targets)} '
              f'(threshold {args.threshold}, every {args.interval}s). Ctrl+C to stop.')
    visible = {name: False for name in targets}
    with ScreenCapture() as sct:
        while True:
            screen = cv2.cvtColor(grab(sct, args.monitor), cv2.COLOR_BGR2GRAY)
            for name, tmpl in targets.items():
                score, loc = match(screen, tmpl)
                found = score >= args.threshold
                if args.once:
                    status = 'FOUND' if found else 'not found'
                    print(f'{name:20s} {status:10s} score={score:.2f} at {loc}')
                elif found != visible[name]:
                    stamp = datetime.now().strftime('%H:%M:%S')
                    event = 'appeared' if found else 'gone'
                    print(f'[{stamp}] {name} {event} (score {score:.2f} at {loc})')
                    if found and not args.quiet:
                        beep()
                    if args.log:
                        log_event(name, event, score, loc)
                visible[name] = found
            if args.once:
                return 0 if all(visible.values()) else 1
            time.sleep(args.interval)


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--monitor', type=int, default=1, help='1 = primary, 2 = second..., 0 = all (default 1)')
    sub = p.add_subparsers(dest='cmd')

    s = sub.add_parser('snip', help='crop a new target image from the screen')
    s.add_argument('name', help='target name (saved as targets/NAME.png)')
    s.add_argument('--delay', type=float, default=3, help='seconds before capture (default 3)')
    s.add_argument('--monitor', type=int, default=argparse.SUPPRESS, help='same as above')

    p.add_argument('--threshold', type=float, default=0.85, help='match score 0..1 (default 0.85)')
    p.add_argument('--interval', type=float, default=1.0, help='seconds between scans (default 1)')
    p.add_argument('--once', action='store_true', help='scan once and exit (exit code 0 if all found)')
    p.add_argument('--quiet', action='store_true', help='no beep')
    p.add_argument('--log', action='store_true', help=f'append events to {LOG_FILE.name}')

    args = p.parse_args()
    if args.cmd == 'snip':
        cmd_snip(args)
        return 0
    try:
        return cmd_scan(args)
    except KeyboardInterrupt:
        return 0


if __name__ == '__main__':
    sys.exit(main())
