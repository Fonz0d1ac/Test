# Screen Scanner

A small script that watches your screen for images you choose and tells you when each one appears or disappears. It uses OpenCV template matching: no model, no training, no GUI app. The only dependencies are `mss`, `opencv-python` and `numpy`.

## Setup

```
pip install -r requirements.txt
```

## Use

1. **Add a target.** Get the object on screen, then run:
   ```
   python screen_scan.py snip error_dialog
   ```
   You get 3 seconds to switch windows. A screenshot then opens: drag a box around the object and press Enter. The crop is saved as `targets/error_dialog.png`. You can also drop any PNG cropped from a screenshot into `targets/`.

2. **Watch the screen:**
   ```
   python screen_scan.py
   ```
   ```
   [14:02:11] error_dialog appeared (score 0.97 at (812, 404))
   [14:02:30] error_dialog gone (score 0.31 at (120, 88))
   ```
   It beeps when a target appears. Press Ctrl+C to stop.

3. **One-off check** (handy in scripts; the exit code is 0 only if every target was found):
   ```
   python screen_scan.py --once
   ```

## Options

| Option | Default | Meaning |
|---|---|---|
| `--threshold 0.85` | 0.85 | How close a match must be (0–1). Lower it if targets are missed; raise it for false hits. |
| `--interval 1` | 1 s | Time between scans. |
| `--monitor 1` | 1 | 1 = primary, 2 = second screen, 0 = all screens. |
| `--log` | off | Append appear/gone events to `scan_log.csv`. |
| `--quiet` | off | No beep. |

## Limits

- Matching is pixel-exact in size. Take targets at the same resolution and Windows display scaling (100%, 125%…) you'll scan at. A target that is resized, recolored or on a different theme won't match.
- It finds the best match for each target, not every copy of it on screen.
- Your targets and the log stay local: both are git-ignored.

## Map node scanner (`node_scan.py`)

Scans a rectangle of the game map for a target sprite by jumping through the coordinate search popup, and records each sprite's map coordinates in `nodes.json`. Run each step with the game window visible and not covered, the UI hidden (blue button, top left) and the camera at the fly-home zoom level. It finds the window by its title (`window_title`, default `Rise of Kingdoms`) and works only inside it, so you can move the window. If you resize it, redo steps 2 and 4. The script clicks and types, so leave the mouse alone while it runs. **To stop it, push the mouse into the top-left corner of the screen.**

1. `python node_scan.py setup`: hover over each point it asks for (the magnifier on the coordinate bar, the X box, the Y box, the blue search button) until its countdown ends. Open the popup yourself before the X box step.
2. `python screen_scan.py snip cropland`: crop the target sprite tightly. It must be taken at the same zoom and window size you scan at.
3. `python node_scan.py goto 312 930`: check that the jump lands where it should.
4. `python node_scan.py calibrate 312 930`: jumps +4 X and +4 Y from there and measures how far the ground moves, including how much bigger the ground is drawn near the bottom of the window than the top (the camera is tilted). Pick open ground with some detail.
5. `python node_scan.py check 312 930`: with the map centred on X:312 Y:930, finds the sprites in the game window and opens `check.png`: red boxes are accepted, with their computed coordinates, and yellow boxes are near misses with their scores. Click a couple in game and compare with the popup. If they are all off by the same amount, set `tile_offset` in `scan_config.json`.
6. `python node_scan.py scan 200 800 400 1000`: scans X 200–400, Y 800–1000 in 16-tile steps and prints NEW/GONE nodes. Add `--loop 10` to re-scan every 10 minutes.

Tuning lives in `scan_config.json`: `threshold` (match score, default 0.9; croplands score about 0.95, food nodes up to about 0.85), `scales` (sprite sizes searched relative to the target image, default 0.75–1.35, because sprites lower in the window are drawn bigger), `load_wait` (seconds after each jump, default 1.2), `click_wait` (pause after each click and after typing, default 0.3), `popup_wait` (longest to wait for the search popup, default 3) and `step` (tiles between scan points, default 16). Only sprites inside each scan point's own 16×16 cell (plus 1 tile of overlap) are recorded, which keeps the coordinate maths in the accurate middle of the screen. Expect coordinates to be within a tile of the popup's.
