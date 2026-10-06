"""Offline regression tests for node_scan.py: no game, no Windows, no screen.

Simulates the game's tilted camera (ground drawn bigger lower in the window,
sprites as billboards scaled with it), fakes the window/input layer, and checks
calibration, detection, coordinate maths, scanning, the popup handling in
goto, per-target results and config migration.

  python tests/test_offline.py        (or: python -m pytest tests)
"""
import json
import sys
import tempfile
import time
import types
from pathlib import Path

import cv2
import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import node_scan as ns  # noqa: E402

W, H = 1280, 720
CX, CY = W / 2, H / 2
VXX, VYY = 40.0, -28.0          # true px per +1 X / +1 Y at the window middle
GP = 12                          # world texture px per tile
X0, NT = 900, 240                # simulated world covers tiles X0..X0+NT
rng = np.random.default_rng(7)
TEX = cv2.GaussianBlur((rng.random((NT * GP, NT * GP)) * 255).astype(np.uint8), (0, 0), 1.5)
TEX = cv2.merge([TEX // 3 + 40, TEX // 2 + 110, TEX // 3 + 60])  # greenish ground


def make_sprite() -> np.ndarray:
    """A white icon on grass, standing in for the cropland/gem sprite."""
    s = np.zeros((34, 28, 3), np.uint8)
    s[:] = (60, 150, 90)
    cv2.ellipse(s, (18, 9), (7, 6), 0, 0, 360, (235, 235, 235), -1)                 # corn head
    cv2.fillPoly(s, [np.array([[13, 16], [25, 16], [25, 24], [19, 31], [13, 24]])],
                 (240, 240, 240))                                                    # shield
    cv2.fillPoly(s, [np.array([[4, 12], [8, 10], [10, 28], [5, 30]])], (225, 225, 225))  # stalk
    cv2.line(s, (14, 20), (24, 20), (150, 150, 150), 1)
    return s


SPRITE = make_sprite()


def tilt_c(ratio: float) -> float:
    """Perspective constant for 'size at 3/4 down / size at 1/4 down' = ratio."""
    return 4 * (ratio - 1) / (H * (ratio + 1))


class World:
    """Renders the game view centred on map tile (X, Y)."""

    def __init__(self, c: float, nodes: list, ui: bool = False):
        self.c, self.nodes, self.ui = c, list(nodes), ui
        yy, xx = np.mgrid[0:H, 0:W].astype(np.float32)
        k = 1 + c * (yy - CY)
        self.dX = ((xx - CX) / k) / VXX
        self.dY = (((1 - 1 / k) / c) if c else (yy - CY)) / VYY
        self.cam = [1000, 1000]

    def render(self) -> np.ndarray:
        X, Y = self.cam
        mx = ((X + self.dX) - X0) * GP
        my = (X0 + NT - (Y + self.dY)) * GP
        img = cv2.remap(TEX, mx.astype(np.float32), my.astype(np.float32), cv2.INTER_LINEAR)
        for tx, ty in self.nodes:  # billboards: forward-project the centre, scale by k
            ry = (ty - Y) * VYY
            if 1 - self.c * ry <= 0.2:
                continue  # at or beyond the horizon
            k = 1 / (1 - self.c * ry)
            py = CY + ((k - 1) / self.c if self.c else ry)
            px = CX + k * (tx - X) * VXX
            if not (-50 < px < W + 50 and -50 < py < H + 50):
                continue
            s = cv2.resize(SPRITE, None, fx=k, fy=k, interpolation=cv2.INTER_LINEAR)
            h, w = s.shape[:2]
            x0, y0 = int(round(px - w / 2)), int(round(py - h / 2))
            if 0 <= x0 and x0 + w < W and 0 <= y0 and y0 + h < H:
                img[y0:y0 + h, x0:x0 + w] = s
        if self.ui:  # static overlays that do not move with the map
            img[0:40, 0:330] = 230
            img[H - 90:H, 0:450] = 40
            img[H - 90:H, W - 240:W] = 180
        return img


REAL_TIME = ns.time
ORIGINAL = {name: getattr(ns, name) for name in
            ('goto', 'grab_game', 'ScreenCapture', 'countdown', 'time', 'detect', 'click_point',
             'patch_at', 'HERE', 'NODES_DIR', 'CONFIG_FILE')}
ORIGINAL_WI = {name: getattr(ns.wi, name) for name in ('cursor_pos', 'type_keys')}


def reset() -> None:
    """Undo every patch, so each test starts from the real module."""
    for name, value in ORIGINAL.items():
        setattr(ns, name, value)
    for name, value in ORIGINAL_WI.items():
        setattr(ns.wi, name, value)


def setup_function(fn) -> None:  # pytest hook; main() calls reset() itself
    reset()


class FakeSct:
    def __enter__(self):
        return self

    def __exit__(self, *a):
        pass


def patch_game(world: World, tmp: Path) -> None:
    """Route node_scan's window, input and file access to the simulation."""
    ns.goto = lambda cfg, x, y: world.cam.__setitem__(slice(None), [x, y])
    ns.grab_game = lambda sct, cfg: world.render()
    ns.ScreenCapture = FakeSct
    ns.countdown = lambda *a: None
    ns.wi.cursor_pos = lambda: (500, 500)
    ns.time = types.SimpleNamespace(sleep=lambda s: None, monotonic=REAL_TIME.monotonic)
    ns.HERE = ns.NODES_DIR = tmp
    ns.CONFIG_FILE = tmp / 'scan_config.json'
    (tmp / 'targets').mkdir(exist_ok=True)
    cv2.imwrite(str(tmp / 'targets' / 'cropland.png'), SPRITE)


def base_cfg() -> dict:
    return dict(ns.DEFAULTS, window_points={k: [0, 0] for k, _ in ns.POINTS})


# ---------------------------------------------------------------- tests

def test_pixel_to_tile_inverts_the_tilt_model():
    c = tilt_c(1.3)
    cfg = dict(ns.DEFAULTS, vx=[VXX, 0], vy=[0, VYY], perspective=c)
    for tx, ty in ((7, -5), (-9, 6), (0, 0), (3.5, 8)):
        ry = ty * VYY
        k = 1 / (1 - c * ry)
        px, py = CX + k * tx * VXX, CY + (k - 1) / c
        dx, dy = ns.pixel_to_tile(cfg, (CX, CY), (px, py))
        assert abs(dx - tx) < 1e-6 and abs(dy - ty) < 1e-6, (tx, ty, dx, dy)


def test_calibration_recovers_tilt_and_tile_size():
    for ratio, ui in ((1.0, False), (1.3, False), (1.3, True), (1.5, True)):
        world = World(tilt_c(ratio), [], ui=ui)
        shots = []
        for dx, dy in ((0, 0), (4, 0), (0, 4)):
            world.cam = [1030 + dx, 1030 + dy]
            shots.append(cv2.cvtColor(world.render(), cv2.COLOR_BGR2GRAY).astype(np.float32))
        fit = ns.fit_calibration(shots, 4)
        assert abs(fit['ratio'] - ratio) <= 0.03, (ratio, fit['ratio'])
        assert abs(fit['vx'][0] - VXX) < 0.5 and abs(fit['vx'][1]) < 0.5, fit['vx']
        assert abs(fit['vy'][1] - VYY) < 0.5 and abs(fit['vy'][0]) < 0.5, fit['vy']


def test_calibration_rejects_static_screens():
    world = World(tilt_c(1.3), [])
    still = cv2.cvtColor(world.render(), cv2.COLOR_BGR2GRAY).astype(np.float32)
    try:
        ns.fit_calibration([still, still.copy(), still.copy()], 4)
    except RuntimeError:
        return
    raise AssertionError('a screen that never moved was accepted as a calibration')


def test_scan_finds_every_node_at_its_exact_tile():
    nodes = [(1004, 1060), (1010, 1010), (1025, 1047), (1040, 1036), (1050, 1020),
             (1061, 1003), (1066, 1061), (1032, 1032)]  # 1040/1032 sit on cell borders
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        world = World(tilt_c(1.3), nodes)
        patch_game(world, tmp)
        cfg = base_cfg()
        ns.cmd_calibrate(cfg, types.SimpleNamespace(x=1030, y=1030, d=4))
        ns.cmd_scan(cfg, types.SimpleNamespace(x1=1000, y1=1000, x2=1064, y2=1064, loop=None,
                                               target=None, threshold=None))
        found = sorted((n['x'], n['y']) for n in json.loads((tmp / 'nodes_cropland.json').read_text()))
        assert found == sorted(nodes), (found, sorted(nodes))

        # a node that disappears is reported GONE on the next pass, others stay
        world.nodes.remove((1010, 1010))
        ns.cmd_scan(cfg, types.SimpleNamespace(x1=1000, y1=1000, x2=1064, y2=1064, loop=None,
                                               target=None, threshold=None))
        left = sorted((n['x'], n['y']) for n in json.loads((tmp / 'nodes_cropland.json').read_text()))
        assert left == sorted(set(nodes) - {(1010, 1010)}), left


def test_scan_points_cover_the_far_edge():
    pts = []
    saved = ns.goto, ns.grab_game, ns.detect
    ns.goto = lambda cfg, x, y: pts.append((x, y))
    ns.grab_game = lambda *a: None
    ns.detect = lambda *a, **k: []
    try:
        with tempfile.TemporaryDirectory() as d:
            res = Path(d) / 'n.json'
            for (x1, y1, x2, y2), xs, ys in (((302, 920, 334, 952), [302, 318, 334], [920, 936, 952]),
                                             ((300, 900, 330, 931), [300, 316, 330], [900, 916, 931])):
                pts.clear()
                ns.scan_pass(dict(ns.DEFAULTS), types.SimpleNamespace(x1=x1, y1=y1, x2=x2, y2=y2),
                             None, None, res)
                assert sorted({p[0] for p in pts}) == xs and sorted({p[1] for p in pts}) == ys
    finally:
        ns.goto, ns.grab_game, ns.detect = saved


def test_goto_waits_for_popup_types_and_gives_up():
    log, state = [], {}

    def click(cfg, name):
        log.append(name)
        if name == 'search_button' and not state.get('never'):
            state['open_at'] = time.monotonic() + 0.3

    ns.click_point = click
    ns.patch_at = lambda cfg, n: np.full((24, 60, 3), 200 if state.get('open_at') and
                                          time.monotonic() > state['open_at'] else 50, np.int16)
    ns.wi.type_keys = lambda vks, delay: log.append(vks)
    ns.wi.cursor_pos = lambda: (500, 500)
    cfg = dict(ns.DEFAULTS, load_wait=0, click_wait=0)
    ns.goto(cfg, 1234, 56)
    clicks = [e for e in log if isinstance(e, str)]
    typed = [e for e in log if isinstance(e, list)]
    assert clicks == ['search_button', 'x_field', 'y_field', 'go_button'], clicks
    clear = [ns.wi.VK_END] + [ns.wi.VK_BACK] * 4
    assert typed == [clear + [ord(c) for c in '1234'], clear + [ord(c) for c in '56']], typed

    # popup never opens: one retry, then stop without typing anything
    log.clear(); state.clear(); state['never'] = True
    cfg['popup_wait'] = 0.2
    try:
        ns.goto(cfg, 1, 2)
        raise AssertionError('goto typed without the popup open')
    except RuntimeError:
        pass
    assert log == ['search_button', 'search_button'], log

    # mouse in the top-left corner: emergency stop before any click
    log.clear()
    ns.wi.cursor_pos = lambda: (0, 0)
    try:
        ns.goto(cfg, 1, 2)
        raise AssertionError('goto ignored the emergency stop')
    except ns.Stop:
        pass
    assert log == [], log


def test_targets_keep_separate_results_and_legacy_file_moves():
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        world = World(0.0, [])
        patch_game(world, tmp)
        cv2.imwrite(str(tmp / 'targets' / 'gem.png'), SPRITE)
        (tmp / 'nodes.json').write_text(json.dumps([{'x': 318, 'y': 936, 'score': 0.96,
                                                     'first_seen': '2026-10-05T10:00:00',
                                                     'last_seen': '2026-10-05T10:00:00'}]))
        cfg = dict(base_cfg(), vx=[VXX, 0], vy=[0, VYY])
        saved = ns.detect
        ns.detect = lambda cfg, screen, target, x, y, step=None: (
            [(320, 930, 0.95, 0, 0, 1.0, 320.0, 930.0)] if (x, y) == (318, 936) else [])
        try:
            ns.cmd_scan(dict(cfg), types.SimpleNamespace(x1=302, y1=920, x2=334, y2=952, loop=None,
                                                         target='gem', threshold=None))
        finally:
            ns.detect = saved
        assert (tmp / 'nodes.json').exists(), 'a gem scan touched the cropland results'
        gems = json.loads((tmp / 'nodes_gem.json').read_text())
        assert [(n['x'], n['y']) for n in gems] == [(320, 930)]
        ns.load_target(cfg, types.SimpleNamespace(target=None, threshold=None))
        results = ns.nodes_file(tmp / 'targets' / 'cropland.png')
        assert results.name == 'nodes_cropland.json'


def test_per_target_offset_and_threshold_override():
    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        patch_game(World(0.0, []), tmp)
        cv2.imwrite(str(tmp / 'targets' / 'gem.png'), SPRITE)
        cfg = base_cfg()
        ns.cmd_offset(cfg, types.SimpleNamespace(name='gem', dx=-1, dy=0.5))
        cfg = ns.load_config()
        c1 = dict(cfg)
        ns.load_target(c1, types.SimpleNamespace(target='gem', threshold=0.93))
        assert c1['tile_offset'] == [-1, 0.5] and c1['threshold'] == 0.93
        c2 = dict(cfg)
        ns.load_target(c2, types.SimpleNamespace(target=None, threshold=None))
        assert c2['tile_offset'] == [0, 0] and c2['threshold'] == ns.DEFAULTS['threshold']


def test_config_migrates_once_from_v1():
    with tempfile.TemporaryDirectory() as d:
        ns.CONFIG_FILE = Path(d) / 'scan_config.json'
        ns.CONFIG_FILE.write_text(json.dumps({'points': {'a': [1, 2]}, 'monitor': 1, 'threshold': 0.85,
                                              'click_wait': 0.3, 'vx': [60.5, 0], 'vy': [2, -41.5]}))
        cfg = ns.load_config()
        saved = json.loads(ns.CONFIG_FILE.read_text())
        assert 'points' not in saved and saved['version'] == ns.CONFIG_VERSION
        assert cfg['threshold'] == ns.DEFAULTS['threshold'] and cfg['click_wait'] == ns.DEFAULTS['click_wait']
        assert cfg['vx'] == [60.5, 0]  # measurements are kept


def main() -> int:
    tests = [(n, f) for n, f in globals().items() if n.startswith('test_') and callable(f)]
    failed = 0
    for name, fn in tests:
        reset()
        t = time.time()
        try:
            fn()
            print(f'PASS  {name}  ({time.time() - t:.1f}s)')
        except Exception as e:  # noqa: BLE001 - report every failure
            failed += 1
            print(f'FAIL  {name}: {type(e).__name__}: {e}')
    print(f'\n{len(tests) - failed}/{len(tests)} passed')
    return 1 if failed else 0


if __name__ == '__main__':
    sys.exit(main())
