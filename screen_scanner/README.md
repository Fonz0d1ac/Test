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
