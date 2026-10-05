"""Scan a region of the game map for a target sprite and record its map coordinates.

Steps, in order (game window visible and not covered, UI hidden with the blue
button, at the fly-home zoom level). Keep the game window the same size
afterwards; moving it is fine. If you resize it, redo snip and calibrate.

  python node_scan.py setup                    record where the search button / X / Y / Go are
  python screen_scan.py snip cropland          crop the target sprite (once)
  python node_scan.py goto 312 930             test that a coordinate jump works
  python node_scan.py calibrate 312 930        measure pixels per map tile (needs open map there)
  python node_scan.py check 312 930            detect on the current screen, save check.png
  python node_scan.py scan 200 800 400 1000    scan X 200..400, Y 800..1000, save nodes.json
  python node_scan.py scan ... --loop 10       re-scan every 10 minutes
  python node_scan.py scan ... --target gem    look for targets/gem.png instead (own nodes_gem.json)
  python node_scan.py set-offset gem -1 0      shift every gem result by X -1, Y 0

Emergency stop: push the mouse into the top-left corner of the screen.
"""
import argparse
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

import cv2
import numpy as np

import wininput as wi
from screen_scan import ScreenCapture

if not hasattr(wi, 'type_keys'):
    sys.exit(f'wininput.py in {Path(__file__).parent} is older than node_scan.py. '
             f'Replace the whole screen_scanner folder with the latest download, then copy '
             f'scan_config.json, targets/ and nodes_*.json back in.')

HERE = Path(__file__).parent.resolve()
CONFIG_FILE = HERE / 'scan_config.json'
NODES_DIR = HERE  # results go in nodes_<target>.json here

CONFIG_VERSION = 4
DEFAULTS = {
    'version': CONFIG_VERSION,
    'window_title': 'Rise of Kingdoms',
    'target': 'targets/cropland.png',
    'threshold': 0.9,    # croplands score ~0.95+, food nodes (no shield) up to ~0.85
    'scales': [0.75, 1.35],  # sprite size range searched, relative to the target image
    'load_wait': 1.2,    # seconds after a jump before taking the screenshot
    'popup_wait': 3.0,   # longest to wait for the search popup to open
    'click_wait': 0.15,  # pause after each click and after typing
    'key_delay': 0.01,   # pause between key presses; 0 types a field instantly
    'step': 16,          # map tiles between scan points
    'tile_offset': [0, 0],  # added to every result, to correct a constant error
    'target_offsets': {},   # per target, e.g. {"gem": [-1, 0]}; replaces tile_offset
    'window_points': {},  # search_button, x_field, y_field, go_button (px inside the game window)
    'vx': None,          # screen px moved per +1 map X, at the middle of the window
    'vy': None,          # screen px moved per +1 map Y, at the middle of the window
    'perspective': None,  # how fast things grow per px further down the window (tilted camera)
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
        saved = json.loads(CONFIG_FILE.read_text())
        version = saved.get('version', 1)
        if version < 2:
            # v1 measured from the whole monitor: drop its screen positions
            for key in ('points', 'monitor', 'center'):
                saved.pop(key, None)
            print('Config updated: click positions are now relative to the game window. '
                  'Run "setup" again.')
        if version < 3:
            # old defaults; v3 adds perspective, so calibrate again
            for key in ('threshold', 'tile_offset', 'version'):
                saved.pop(key, None)
            print('Config updated: run "calibrate" again to measure the camera tilt.')
        if version < 4 and saved.get('click_wait') == 0.3:
            saved.pop('click_wait')  # old default; typing is faster now
        cfg.update(saved)
        if version < CONFIG_VERSION:
            cfg['version'] = CONFIG_VERSION
            save_config(cfg)
    return cfg


def save_config(cfg: dict) -> None:
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2))


def need(cfg: dict, points: bool = False, calibration: bool = False) -> None:
    missing = [p for p, _ in POINTS if points and p not in cfg['window_points']]
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


def find_all(screen: np.ndarray, target: np.ndarray, threshold: float, scales=(1, 1)) -> list:
    """Every match of target (both BGR) above threshold, trying sizes from
    scales[0] to scales[1] (the tilted camera draws sprites lower on the screen
    bigger). Returns [(score, cx, cy, scale)], best match per spot."""
    screen_w = whiteness(screen)
    lo, hi = scales
    candidates = []
    for s in np.arange(lo, hi + 1e-9, 0.05) if hi > lo else [lo]:
        t = cv2.resize(target, None, fx=s, fy=s,
                       interpolation=cv2.INTER_AREA if s < 1 else cv2.INTER_LINEAR)
        th, tw = t.shape[:2]
        if th > screen.shape[0] or tw > screen.shape[1]:
            continue
        r = cv2.matchTemplate(screen_w, whiteness(t), cv2.TM_CCOEFF_NORMED)
        peaks = (r >= threshold) & (r == cv2.dilate(r, np.ones((th, tw), np.uint8)))
        candidates += [(float(r[y, x]), x + tw // 2, y + th // 2, float(s))
                       for y, x in zip(*np.nonzero(peaks))]
    # one result per spot: keep the best-scoring size
    radius = max(target.shape[:2]) * lo * 0.6
    kept = []
    for c in sorted(candidates, reverse=True):
        if all(abs(c[1] - k[1]) > radius or abs(c[2] - k[2]) > radius for k in kept):
            kept.append(c)
    return kept


def screen_center(screen: np.ndarray) -> tuple:
    """The map coordinate in the coordinate bar is the middle of the game view."""
    h, w = screen.shape[:2]
    return w / 2, h / 2


def pixel_to_tile(cfg: dict, center_px: tuple, px: tuple) -> tuple:
    """Map-tile offset (dX, dY) of a screen pixel relative to the screen centre.

    The camera is tilted, so the ground is drawn bigger lower on the screen:
    k(y) = 1 + c * (y - centre) is the size there relative to the middle. Across
    the screen, tiles scale by k; down the screen, by k squared. Undo that, then
    use the vx/vy measured at the middle."""
    c = cfg.get('perspective') or 0.0
    dx, dy = np.subtract(px, center_px)
    k = 1 + c * dy
    if c:
        dx, dy = dx / k, (1 - 1 / k) / c
    m = np.array([cfg['vx'], cfg['vy']], dtype=float).T  # columns: px per +1 X, px per +1 Y
    d = np.linalg.solve(m, (dx, dy))
    return float(d[0]), float(d[1])


def untilt(img: np.ndarray, c: float) -> np.ndarray:
    """Redraw a screenshot as if the camera looked straight down, keeping the
    scale it has at the middle (inverse of the k(y) model in pixel_to_tile)."""
    if not c:
        return img
    h, w = img.shape[:2]
    ry, rx = np.mgrid[0:h, 0:w].astype(np.float32)
    ry -= h / 2
    k = 1 / (1 - c * ry)
    map_y = h / 2 + (k - 1) / c
    map_x = w / 2 + k * (rx - w / 2)
    return cv2.remap(img, map_x.astype(np.float32), map_y.astype(np.float32), cv2.INTER_LINEAR)


def fit_calibration(shots: list, d: int) -> dict:
    """From three screenshots (start, +d X, +d Y) find the camera tilt and the
    px per tile at the middle of the window.

    Each candidate tilt is undone on all three shots and the X and Y jumps are
    measured with phase correlation: with the right tilt the ground moves the
    same amount everywhere and both matches are sharp (high confidence), with
    a wrong one they smear. Keep the tilt with the best combined confidence."""
    h, w = shots[0].shape
    crop = (slice(h // 5, h * 4 // 5), slice(w // 5, w * 4 // 5))  # ignore UI at the edges
    win = cv2.createHanningWindow((w * 4 // 5 - w // 5, h * 4 // 5 - h // 5), cv2.CV_32F)

    def measure(ratio):
        # ratio = size at 3/4 down the window / size at 1/4 down
        c = 4 * (ratio - 1) / (h * (ratio + 1))
        flat = [untilt(s, c) for s in shots]
        (xs, ys), cx_ = cv2.phaseCorrelate(flat[0][crop], flat[1][crop], win)
        (xs2, ys2), cy_ = cv2.phaseCorrelate(flat[0][crop], flat[2][crop], win)
        # A d-tile jump moves the ground well over 10 px per tile; a near-zero
        # shift means it matched something that stayed put (resampling
        # patterns on plain ground, leftover UI), so that tilt is not a fit.
        moved = np.hypot(xs, ys) > 10 * d and np.hypot(xs2, ys2) > 10 * d
        # camera +d tiles => ground moves the opposite way on screen
        return {'ratio': ratio, 'c': c, 'vx': [-xs / d, -ys / d], 'vy': [-xs2 / d, -ys2 / d],
                'conf_x': cx_, 'conf_y': cy_, 'score': cx_ + cy_ if moved else -1}

    lo, hi = 1.0, 1.8
    best = max((measure(r) for r in np.arange(lo, hi + 1e-9, 0.04)), key=lambda m: m['score'])
    r0 = best['ratio']
    best = max((measure(r) for r in np.arange(max(lo, r0 - 0.04), min(hi, r0 + 0.04) + 1e-9, 0.01)),
               key=lambda m: m['score'])
    if best['score'] < 0:
        raise RuntimeError('Calibration failed: the ground did not seem to move between jumps. '
                           'Check the jumps land (try "goto"), the UI is hidden, and pick a spot '
                           'with more ground detail.')
    best['at_limit'] = best['ratio'] >= hi - 0.01  # 1.0 (no tilt) is a valid answer
    return best


def detect(cfg: dict, screen: np.ndarray, target: np.ndarray, cx: int, cy: int,
           step: int = None) -> list:
    """Targets on screen as map coords [(x, y, score, px, py)]. With step, keep only
    those inside this scan point's own step x step cell (the most accurate part of
    the screen; neighbouring scan points cover the rest)."""
    center = screen_center(screen)
    ox, oy = cfg['tile_offset']
    found = []
    for score, px, py, scale in find_all(screen, target, cfg['threshold'], cfg['scales']):
        dx, dy = pixel_to_tile(cfg, center, (px, py))
        # 1 tile of overlap so a node on a cell border is never lost to rounding;
        # merge() drops the duplicate
        if step and not (abs(dx) <= step / 2 + 1 and abs(dy) <= step / 2 + 1):
            continue
        fx, fy = cx + dx + ox, cy + dy + oy
        found.append((round(fx), round(fy), score, px, py, scale, fx, fy))
    return found


# ---------- game window ----------

def grab_game(sct, cfg: dict) -> np.ndarray:
    """Screenshot of just the game window's drawable area, as BGR."""
    left, top, width, height = wi.client_rect(cfg['window_title'])
    shot = sct.grab({'left': left, 'top': top, 'width': width, 'height': height})
    return cv2.cvtColor(np.asarray(shot), cv2.COLOR_BGRA2BGR)


def click_point(cfg: dict, name: str) -> None:
    left, top, _, _ = wi.client_rect(cfg['window_title'])
    x, y = cfg['window_points'][name]
    wi.click(left + x, top + y)


_patch_sct = None


def patch_at(cfg: dict, name: str) -> np.ndarray:
    """Small screenshot around one of the recorded points."""
    global _patch_sct
    if _patch_sct is None:
        _patch_sct = ScreenCapture()
    left, top, _, _ = wi.client_rect(cfg['window_title'])
    x, y = cfg['window_points'][name]
    shot = _patch_sct.grab({'left': left + x - 30, 'top': top + y - 12, 'width': 60, 'height': 24})
    return np.asarray(shot)[:, :, :3].astype(np.int16)


def wait_for_change(cfg: dict, name: str, before: np.ndarray, timeout: float) -> bool:
    """Poll until the area around a point looks different (e.g. a popup opened there)."""
    end = time.monotonic() + timeout
    while time.monotonic() < end:
        time.sleep(0.1)
        if np.abs(patch_at(cfg, name) - before).mean() > 20:
            return True
    return False


# ---------- game actions ----------

def check_stop() -> None:
    x, y = wi.cursor_pos()
    if x <= 5 and y <= 5:
        raise Stop()


def goto(cfg: dict, x: int, y: int) -> None:
    """Jump the map centre to (x, y) through the coordinate search popup."""
    wait = cfg['click_wait']
    for attempt in range(2):
        check_stop()
        before = patch_at(cfg, 'x_field')
        click_point(cfg, 'search_button')
        # wait until the popup's X box actually appears instead of guessing a delay
        if wait_for_change(cfg, 'x_field', before, cfg['popup_wait']):
            break
        if attempt == 0:
            print('  search popup did not open, clicking the magnifier again...')
    else:
        raise RuntimeError('The search popup did not open. Check the game window is in front, '
                           'the UI is hidden, and re-run "setup" if the window layout changed.')
    time.sleep(wait)  # let the popup finish its opening animation
    for field, value in (('x_field', x), ('y_field', y)):
        check_stop()
        click_point(cfg, field)
        time.sleep(wait)
        # coordinates are at most 4 digits: End + 4 backspaces clears the box
        wi.type_keys([wi.VK_END] + [wi.VK_BACK] * 4 + [ord(ch) for ch in str(value)],
                     cfg['key_delay'])
        time.sleep(wait)
    check_stop()
    click_point(cfg, 'go_button')
    time.sleep(cfg['load_wait'])


def countdown(seconds: int, msg: str) -> None:
    for s in range(seconds, 0, -1):
        print(f'  {msg} {s}...', end='\r', flush=True)
        time.sleep(1)
    print(' ' * 70, end='\r')


# ---------- node list ----------

def nodes_file(target_path: Path) -> Path:
    """Each target keeps its own results, so scanning for one never marks
    another's nodes as gone."""
    return NODES_DIR / f'nodes_{target_path.stem}.json'


def load_nodes(path: Path) -> list:
    return json.loads(path.read_text()) if path.exists() else []


def load_target(cfg: dict, args) -> tuple:
    """Target image from --target NAME (targets/NAME.png, or a .png path) or the
    config, plus --threshold if given. Returns (image, path)."""
    name = getattr(args, 'target', None) or cfg['target']
    path = Path(name) if name.lower().endswith('.png') else Path('targets') / f'{name}.png'
    if not path.is_absolute():
        path = HERE / path
    img = cv2.imread(str(path))
    if img is None:
        sys.exit(f'No target image at {path}. Create it with: python screen_scan.py snip {path.stem}')
    if getattr(args, 'threshold', None):
        cfg['threshold'] = args.threshold
    if path.stem in cfg['target_offsets']:
        cfg['tile_offset'] = cfg['target_offsets'][path.stem]
    return img, path


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
        left, top, _, _ = wi.client_rect(cfg['window_title'])
        x, y = wi.cursor_pos()
        cfg['window_points'][key] = [x - left, y - top]
        print(f'  {key} = {cfg["window_points"][key]} (inside the game window)')
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
            shots.append(cv2.cvtColor(grab_game(sct, cfg), cv2.COLOR_BGR2GRAY).astype(np.float32))
    goto(cfg, args.x, args.y)

    h, w = shots[0].shape
    print('Measuring...')
    fit = fit_calibration(shots, d)
    cfg['perspective'], cfg['vx'], cfg['vy'] = fit['c'], fit['vx'], fit['vy']
    print(f'Camera tilt: ground near the bottom is drawn {fit["ratio"]:.2f}x the size of near the top.')
    for axis, v, conf in (('X', fit['vx'], fit['conf_x']), ('Y', fit['vy'], fit['conf_y'])):
        print(f'+1 {axis} = ({v[0]:+.1f}, {v[1]:+.1f}) px   (match confidence {conf:.2f})')
    if min(fit['conf_x'], fit['conf_y']) < 0.1:
        print('  Low confidence: make sure the UI is hidden and nothing covers the game, and pick\n'
              '  open ground with some detail (not water, not a big building), then run it again.')
    if fit['at_limit']:
        print('  The tilt hit the edge of the search range; the result is probably wrong.')

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
    target, target_path = load_target(cfg, args)
    if not cfg.get('perspective'):
        print('Note: no camera-tilt measurement yet; run "calibrate" again for accurate coordinates.')
    print(f'Looking for {target_path.stem} (threshold {cfg["threshold"]})')
    countdown(3, 'focus the game, capturing in')
    with ScreenCapture() as sct:
        screen = grab_game(sct, cfg)
    img = screen.copy()  # draw here so the boxes never affect detection
    th, tw = target.shape[:2]

    def box(px, py, scale, colour, thick):
        hw, hh = int(tw * scale / 2), int(th * scale / 2)
        cv2.rectangle(img, (px - hw, py - hh), (px + hw, py + hh), colour, thick)
        return hw, hh

    # near misses in yellow help pick the threshold
    for score, px, py, scale in find_all(screen, target, min(0.7, cfg['threshold']), cfg['scales']):
        if score < cfg['threshold']:
            print(f'  (rejected: score {score:.2f} at pixel ({px}, {py}))')
            hw, hh = box(px, py, scale, (0, 255, 255), 1)
            cv2.putText(img, f'{score:.2f}', (px - hw, py + hh + 16),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 255), 1)
    found = detect(cfg, screen, target, args.x, args.y)
    for x, y, score, px, py, scale, fx, fy in found:
        print(f'X:{x} Y:{y}  (exact {fx:.1f}, {fy:.1f})  score {score:.2f}  '
              f'at pixel ({px}, {py}), size {scale:.2f}x')
        hw, hh = box(px, py, scale, (0, 0, 255), 2)
        cv2.putText(img, f'{x},{y} ({score:.2f})', (px - hw, py - hh - 6),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
    out = HERE / 'check.png'
    if not cv2.imwrite(str(out), img):
        sys.exit(f'{len(found)} found, but could not save {out}')
    print(f'{len(found)} found. Annotated screenshot: {out}\n'
          f'Click a few in game and compare with the popup coordinates.')
    if sys.platform == 'win32':
        os.startfile(out)  # open it in the default image viewer


def scan_pass(cfg: dict, args, target: np.ndarray, sct, results: Path) -> None:
    step = cfg['step']
    xs = list(range(args.x1, args.x2 + 1, step))
    ys = list(range(args.y1, args.y2 + 1, step))
    # each point covers +-step/2; add one at the far edge if the last one stops short
    if args.x2 - xs[-1] > step / 2:
        xs.append(args.x2)
    if args.y2 - ys[-1] > step / 2:
        ys.append(args.y2)
    nodes = load_nodes(results)
    start = datetime.now().isoformat(timespec='seconds')
    print(f'\n[{start}] Scanning {len(xs) * len(ys)} points...')
    for row, y in enumerate(ys):
        for x in (xs if row % 2 == 0 else xs[::-1]):  # snake order
            goto(cfg, x, y)
            now = datetime.now().isoformat(timespec='seconds')
            for nx, ny, score, *_ in detect(cfg, grab_game(sct, cfg), target, x, y, step):
                if merge(nodes, nx, ny, score, now):
                    print(f'  NEW  X:{nx} Y:{ny}  (score {score:.2f})')
            results.write_text(json.dumps(nodes, indent=2))

    # drop nodes inside the scanned area that were not seen this pass
    half = step / 2
    inside = lambda n: args.x1 - half <= n['x'] < args.x2 + half and args.y1 - half <= n['y'] < args.y2 + half
    gone = [n for n in nodes if inside(n) and n['last_seen'] < start]
    for n in gone:
        print(f'  GONE X:{n["x"]} Y:{n["y"]}')
    nodes = [n for n in nodes if n not in gone]
    results.write_text(json.dumps(nodes, indent=2))
    print(f'Pass done: {len(nodes)} nodes in {results.name}')


def cmd_offset(cfg: dict, args) -> None:
    """Save a correction for one target: added to every coordinate found for it."""
    cfg['target_offsets'][args.name] = [args.dx, args.dy]
    save_config(cfg)
    print(f'{args.name}: results will be shifted by X {args.dx:+g}, Y {args.dy:+g}.')


def cmd_scan(cfg: dict, args) -> None:
    need(cfg, points=True, calibration=True)
    target, target_path = load_target(cfg, args)
    results = nodes_file(target_path)
    legacy = NODES_DIR / 'nodes.json'  # before per-target files, everything went here
    if legacy.exists() and not results.exists() and target_path.stem == 'cropland':
        legacy.rename(results)
        print(f'Moved nodes.json to {results.name}.')
    print(f'Looking for {target_path.stem}; results in {results.name}')
    args.x1, args.x2 = sorted((args.x1, args.x2))
    args.y1, args.y2 = sorted((args.y1, args.y2))
    countdown(3, 'focus the game, starting in')
    with ScreenCapture() as sct:
        while True:
            t0 = time.monotonic()
            scan_pass(cfg, args, target, sct, results)
            if not args.loop:
                return
            wait = args.loop * 60 - (time.monotonic() - t0)
            print(f'Next pass in {max(0, wait) / 60:.1f} min (move the mouse to the top-left corner to stop)')
            end = time.monotonic() + wait
            while time.monotonic() < end:
                check_stop()
                time.sleep(0.5)


def target_options(p) -> None:
    p.add_argument('--target', metavar='NAME',
                   help='look for targets/NAME.png instead of the config\'s target')
    p.add_argument('--threshold', type=float, help='match score for this run (default from config)')


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
        if name == 'check':
            target_options(s)
    s = sub.add_parser('scan', help='scan the rectangle X1,Y1 - X2,Y2')
    for a in ('x1', 'y1', 'x2', 'y2'):
        s.add_argument(a, type=int)
    target_options(s)
    s.add_argument('--loop', type=float, metavar='MIN', help='repeat a full pass every MIN minutes')
    s = sub.add_parser('set-offset', help='correct a target\'s results by DX DY tiles')
    s.add_argument('name', help='target name, e.g. gem')
    s.add_argument('dx', type=float)
    s.add_argument('dy', type=float)
    args = p.parse_args()

    cfg = load_config()
    try:
        {'setup': cmd_setup, 'goto': cmd_goto, 'calibrate': cmd_calibrate,
         'check': cmd_check, 'scan': cmd_scan, 'set-offset': cmd_offset}[args.cmd](cfg, args)
    except (Stop, KeyboardInterrupt):
        print('\nStopped.')
    except RuntimeError as e:
        sys.exit(str(e))
    return 0


if __name__ == '__main__':
    sys.exit(main())
