"""Scan a region of the game map for a target sprite and record its map coordinates.

Steps, in order (game on screen, UI hidden with the blue button, at the
fly-home zoom level):

  python node_scan.py setup                    record where the search button / X / Y / Go are
  python screen_scan.py snip cropland          crop the target sprite (once)
  python node_scan.py goto 312 930             test that a coordinate jump works
  python node_scan.py calibrate 312 930        measure pixels per map tile (needs open map there)
  python node_scan.py check 312 930            detect on the current screen, save check.png
  python node_scan.py scan 200 800 400 1000    scan X 200..400, Y 800..1000, save nodes.json
  python node_scan.py scan ... --loop 10       re-scan every 10 minutes

Emergency stop: push the mouse into the top-left corner of the screen.
"""
import argparse
import json
import sys
import time
from datetime import datetime
from pathlib import Path

import cv2
import numpy as np

import wininput as wi
from screen_scan import ScreenCapture, grab

HERE = Path(__file__).parent.resolve()
CONFIG_FILE = HERE / 'scan_config.json'
NODES_FILE = HERE / 'nodes.json'

DEFAULTS = {
    'target': 'targets/cropland.png',
    'threshold': 0.85,
    'load_wait': 1.2,    # seconds after a jump before taking the screenshot
    'step': 16,          # map tiles between scan points
    'monitor': 1,
    'center': None,      # screen pixel of the map centre; None = middle of the monitor
    'tile_offset': [0, 0],  # added to every result, to correct a constant error
    'points': {},        # search_button, x_field, y_field, go_button (absolute screen px)
    'vx': None,          # screen px moved per +1 map X
    'vy': None,          # screen px moved per +1 map Y
}
POINTS = [
    ('search_button', 'the magnifying glass on the coordinate bar (top left)'),
    ('x_field', 'the X box in the search popup (click the magnifier yourself to open it)'),
    ('y_field', 'the Y box in the search popup'),
    ('go_button', 'the blue search button in the popup'),
]
DEDUPE_TILES = 2


class Stop(Exception):
    pass


# ---------- config ----------

def load_config() -> dict:
    cfg = dict(DEFAULTS)
    if CONFIG_FILE.exists():
        cfg.update(json.loads(CONFIG_FILE.read_text()))
    return cfg


def save_config(cfg: dict) -> None:
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2))


def need(cfg: dict, points: bool = False, calibration: bool = False) -> None:
    missing = [p for p, _ in POINTS if points and p not in cfg['points']]
    missing += [k for k in ('vx', 'vy') if calibration and not cfg[k]]
    if missing:
        sys.exit(f'Missing in {CONFIG_FILE.name}: {", ".join(missing)}. '
                 f'Run "setup" (click points) and "calibrate" (vx/vy) first.')


# ---------- detection (pure functions, no input) ----------

def whiteness(img: np.ndarray) -> np.ndarray:
    """Darkest channel per pixel: white sprites stay bright, coloured ground
    (green grass, pink or teal territory) goes dark, so matches ignore the
    background colour."""
    return img.min(axis=2)


def find_all(screen: np.ndarray, target: np.ndarray, threshold: float) -> list:
    """Every match of target (both BGR) above threshold: [(score, cx, cy)]."""
    r = cv2.matchTemplate(whiteness(screen), whiteness(target), cv2.TM_CCOEFF_NORMED)
    th, tw = target.shape[:2]
    peaks = (r >= threshold) & (r == cv2.dilate(r, np.ones((th, tw), np.uint8)))
    return [(float(r[y, x]), x + tw // 2, y + th // 2) for y, x in zip(*np.nonzero(peaks))]


def screen_center(cfg: dict, screen: np.ndarray) -> tuple:
    if cfg['center']:
        return tuple(cfg['center'])
    h, w = screen.shape[:2]
    return w / 2, h / 2


def pixel_to_tile(cfg: dict, center_px: tuple, px: tuple) -> tuple:
    """Map-tile offset (dX, dY) of a screen pixel relative to the screen centre."""
    m = np.array([cfg['vx'], cfg['vy']], dtype=float).T  # columns: px per +1 X, px per +1 Y
    d = np.linalg.solve(m, np.subtract(px, center_px))
    return float(d[0]), float(d[1])


def detect(cfg: dict, screen: np.ndarray, target: np.ndarray, cx: int, cy: int,
           step: int = None) -> list:
    """Targets on screen as map coords [(x, y, score, px, py)]. With step, keep only
    those inside this scan point's own step x step cell (the most accurate part of
    the screen; neighbouring scan points cover the rest)."""
    center = screen_center(cfg, screen)
    ox, oy = cfg['tile_offset']
    found = []
    for score, px, py in find_all(screen, target, cfg['threshold']):
        dx, dy = pixel_to_tile(cfg, center, (px, py))
        if step and not (-step / 2 <= dx < step / 2 and -step / 2 <= dy < step / 2):
            continue
        found.append((cx + round(dx) + ox, cy + round(dy) + oy, score, px, py))
    return found


# ---------- game actions ----------

def check_stop() -> None:
    x, y = wi.cursor_pos()
    if x <= 5 and y <= 5:
        raise Stop()


def goto(cfg: dict, x: int, y: int) -> None:
    """Jump the map centre to (x, y) through the coordinate search popup."""
    p = cfg['points']
    check_stop()
    wi.click(*p['search_button'])
    time.sleep(0.5)
    for field, value in (('x_field', x), ('y_field', y)):
        check_stop()
        wi.click(*p[field])
        time.sleep(0.15)
        wi.press(wi.VK_END)
        wi.press(wi.VK_BACK, 6)
        wi.type_digits(str(value))
        time.sleep(0.1)
    check_stop()
    wi.click(*p['go_button'])
    time.sleep(cfg['load_wait'])


def countdown(seconds: int, msg: str) -> None:
    for s in range(seconds, 0, -1):
        print(f'  {msg} {s}...', end='\r', flush=True)
        time.sleep(1)
    print(' ' * 70, end='\r')


# ---------- node list ----------

def load_nodes() -> list:
    return json.loads(NODES_FILE.read_text()) if NODES_FILE.exists() else []


def merge(nodes: list, x: int, y: int, score: float, now: str) -> bool:
    """Add or refresh a node; True if it is new."""
    for n in nodes:
        if abs(n['x'] - x) <= DEDUPE_TILES and abs(n['y'] - y) <= DEDUPE_TILES:
            n['last_seen'] = now
            n['score'] = round(score, 3)
            return False
    nodes.append({'x': x, 'y': y, 'score': round(score, 3), 'first_seen': now, 'last_seen': now})
    return True


# ---------- commands ----------

def cmd_setup(cfg: dict, args) -> None:
    print('For each item, put the mouse over it and keep it still until the countdown ends.')
    for key, desc in POINTS:
        print(f'\n{key}: hover over {desc}')
        countdown(6, 'recording in')
        cfg['points'][key] = list(wi.cursor_pos())
        print(f'  {key} = {cfg["points"][key]}')
    save_config(cfg)
    print(f'\nSaved to {CONFIG_FILE.name}. Close the popup, then try: python node_scan.py goto X Y')


def cmd_goto(cfg: dict, args) -> None:
    need(cfg, points=True)
    countdown(3, 'focus the game, starting in')
    goto(cfg, args.x, args.y)
    print(f'Jumped to X:{args.x} Y:{args.y} - check the coordinate bar matches.')


def cmd_calibrate(cfg: dict, args) -> None:
    """Jump d tiles along X and along Y and measure how far the ground moves."""
    need(cfg, points=True)
    d = args.d
    countdown(3, 'focus the game, starting in')
    shots = []
    with ScreenCapture() as sct:
        for x, y in ((args.x, args.y), (args.x + d, args.y), (args.x, args.y + d)):
            goto(cfg, x, y)
            time.sleep(0.5)
            shots.append(cv2.cvtColor(grab(sct, cfg['monitor']), cv2.COLOR_BGR2GRAY).astype(np.float32))
    goto(cfg, args.x, args.y)

    h, w = shots[0].shape
    crop = (slice(h // 5, h * 4 // 5), slice(w // 5, w * 4 // 5))  # ignore UI at the edges
    win = cv2.createHanningWindow((w * 4 // 5 - w // 5, h * 4 // 5 - h // 5), cv2.CV_32F)
    vec = []
    for i, axis in ((1, 'X'), (2, 'Y')):
        (sx, sy), conf = cv2.phaseCorrelate(shots[0][crop], shots[i][crop], win)
        # camera +d tiles => ground moves the opposite way on screen
        vec.append([-sx / d, -sy / d])
        print(f'+1 {axis} = ({-sx / d:+.1f}, {-sy / d:+.1f}) px   (match confidence {conf:.2f})')
        if conf < 0.05:  # perspective keeps this low (~0.1) even when the result is right
            print('  Low confidence: pick a spot with more ground detail, or check the jumps worked.')
    cfg['vx'], cfg['vy'] = vec

    step = cfg['step']
    cx, cy = w / 2, h / 2
    corners = [np.add((cx, cy), np.multiply(cfg['vx'], sx * step / 2) + np.multiply(cfg['vy'], sy * step / 2))
               for sx in (-1, 1) for sy in (-1, 1)]
    if any(not (0 <= px < w and 0 <= py < h) for px, py in corners):
        print(f'Warning: a {step}x{step} tile cell does not fit on screen; lower "step" in the config.')
    save_config(cfg)
    print(f'Saved to {CONFIG_FILE.name}.')


def cmd_check(cfg: dict, args) -> None:
    need(cfg, calibration=True)
    target = cv2.imread(str(HERE / cfg['target']))
    if target is None:
        sys.exit(f'No target image at {cfg["target"]}. Create it with: python screen_scan.py snip cropland')
    countdown(3, 'focus the game, capturing in')
    with ScreenCapture() as sct:
        screen = grab(sct, cfg['monitor'])
    found = detect(cfg, screen, target, args.x, args.y)
    th, tw = target.shape[:2]
    for x, y, score, px, py in found:
        print(f'X:{x} Y:{y}  score {score:.2f}  at pixel ({px}, {py})')
        cv2.rectangle(screen, (px - tw // 2, py - th // 2), (px + tw // 2, py + th // 2), (0, 0, 255), 2)
        cv2.putText(screen, f'{x},{y}', (px - tw // 2, py - th // 2 - 6),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
    cv2.imwrite(str(HERE / 'check.png'), screen)
    print(f'{len(found)} found. Annotated screenshot: check.png - click a few in game and '
          f'compare with the popup coordinates.')


def scan_pass(cfg: dict, args, target: np.ndarray, sct) -> None:
    step = cfg['step']
    xs = list(range(args.x1, args.x2 + 1, step))
    ys = list(range(args.y1, args.y2 + 1, step))
    nodes = load_nodes()
    start = datetime.now().isoformat(timespec='seconds')
    print(f'\n[{start}] Scanning {len(xs) * len(ys)} points...')
    for row, y in enumerate(ys):
        for x in (xs if row % 2 == 0 else xs[::-1]):  # snake order
            goto(cfg, x, y)
            now = datetime.now().isoformat(timespec='seconds')
            for nx, ny, score, _, _ in detect(cfg, grab(sct, cfg['monitor']), target, x, y, step):
                if merge(nodes, nx, ny, score, now):
                    print(f'  NEW  X:{nx} Y:{ny}  (score {score:.2f})')
            NODES_FILE.write_text(json.dumps(nodes, indent=2))

    # drop nodes inside the scanned area that were not seen this pass
    half = step / 2
    inside = lambda n: args.x1 - half <= n['x'] < args.x2 + half and args.y1 - half <= n['y'] < args.y2 + half
    gone = [n for n in nodes if inside(n) and n['last_seen'] < start]
    for n in gone:
        print(f'  GONE X:{n["x"]} Y:{n["y"]}')
    nodes = [n for n in nodes if n not in gone]
    NODES_FILE.write_text(json.dumps(nodes, indent=2))
    print(f'Pass done: {len(nodes)} nodes in {NODES_FILE.name}')


def cmd_scan(cfg: dict, args) -> None:
    need(cfg, points=True, calibration=True)
    target = cv2.imread(str(HERE / cfg['target']))
    if target is None:
        sys.exit(f'No target image at {cfg["target"]}. Create it with: python screen_scan.py snip cropland')
    args.x1, args.x2 = sorted((args.x1, args.x2))
    args.y1, args.y2 = sorted((args.y1, args.y2))
    countdown(3, 'focus the game, starting in')
    with ScreenCapture() as sct:
        while True:
            t0 = time.monotonic()
            scan_pass(cfg, args, target, sct)
            if not args.loop:
                return
            wait = args.loop * 60 - (time.monotonic() - t0)
            print(f'Next pass in {max(0, wait) / 60:.1f} min (move the mouse to the top-left corner to stop)')
            end = time.monotonic() + wait
            while time.monotonic() < end:
                check_stop()
                time.sleep(0.5)


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest='cmd', required=True)
    sub.add_parser('setup', help='record the click positions of the coordinate search')
    for name, helptext in (('goto', 'jump to map coordinates'),
                           ('check', 'detect on the current screen (map centre at X Y)'),
                           ('calibrate', 'measure screen pixels per map tile, starting at X Y')):
        s = sub.add_parser(name, help=helptext)
        s.add_argument('x', type=int)
        s.add_argument('y', type=int)
        if name == 'calibrate':
            s.add_argument('--d', type=int, default=4, help='tiles to move per axis (default 4)')
    s = sub.add_parser('scan', help='scan the rectangle X1,Y1 - X2,Y2')
    for a in ('x1', 'y1', 'x2', 'y2'):
        s.add_argument(a, type=int)
    s.add_argument('--loop', type=float, metavar='MIN', help='repeat a full pass every MIN minutes')
    args = p.parse_args()

    cfg = load_config()
    try:
        {'setup': cmd_setup, 'goto': cmd_goto, 'calibrate': cmd_calibrate,
         'check': cmd_check, 'scan': cmd_scan}[args.cmd](cfg, args)
    except (Stop, KeyboardInterrupt):
        print('\nStopped.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
