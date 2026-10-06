# Screen Scanner: Session Handoff

This is the complete hand-off for the `screen_scanner/` project. Read it before changing any code.

The repository's **root** `HANDOFF.md` belongs to an unrelated project (labels, printing, dashboards). Do not touch it or the files beside it.

- **Branch:** `claude/ecstatic-feynman-libbqk`. Develop, commit and push there.
- **Last code commit before this document:** `f662322`.

---

## 1. Where things stand (TL;DR)

- **The map node scanner works end to end on the user's PC.** It jumps the game camera across a rectangle of the *Rise of Kingdoms* map using the in-game coordinate search. On each screen it finds a target sprite (croplands, now gem nodes), converts each sprite's screen position to map coordinates, and keeps a per-target node list.
- **Accuracy:** the user reports gem coordinates are "mostly good, sometimes off by 1", which is acceptable.
- **The camera model is solid.** Camera tilt, multi-size matching and calibration were validated on the user's machine. The latest calibration had confidence 0.71 for X and 0.58 for Y.
- **Next agreed work, not started:**
  1. Scan speed-ups and a timing breakdown.
  2. Per-area freshness tracking and a cadenced `sweep`.
  3. Checks for whether a node is free (occupied, targeted).
  4. March automation: watch-only first, then dry run, then one live march, then all 5.
- **Steps 3 and 4 are blocked on screenshots and answers from the user** (see §11).
- **Offline tests:** `python tests/test_offline.py` runs 9 tests, all passing at hand-off. Run them after every change.

---

## 2. The user and how they work

- **Platform:** Windows PC with Python. They play the *Rise of Kingdoms* **PC client in a window**.
  - Window title: `Rise of Kingdoms`.
  - Client area is about 1871 × 1052, on a larger monitor that looks like 2560 × 1440.
- **How they get code:** they **download the branch as a ZIP**; they are not using git. The folder is
  `C:\Users\fonz0\Desktop\Test-claude-ecstatic-feynman-libbqk\Test-claude-ecstatic-feynman-libbqk\screen_scanner`.
  - They often run commands from `C:\Windows\System32` with the script's full path. Everything resolves paths from the script's own folder (`HERE`), so this works.
  - **After every update, tell them:** download the ZIP again, replace the whole `screen_scanner` folder, then copy back `scan_config.json`, `targets/` and `nodes_*.json`.
  - A partial copy once left an old `wininput.py` in place. `node_scan.py` now checks for this and says so.
- **Their testing loop:** they run the commands, then paste the output and screenshots back. Code that controls the mouse, keyboard or screen **cannot be run in the cloud session**. Validate the logic offline (see §10) and be clear about what was not run.
- **Communication style:** they like plain explanations, short command lists and expected outputs. They ask "think it through before coding" for big features. When they do, answer with a plan and no code.
- **History:** the session started with them asking about an industrial PLC + camera inspection app (`main.py`, `plc_communication.py`, uploaded, not in the repo). The "trigger → detect → act" architecture was reused here, but none of that code is.

---

## 3. Boundaries already agreed: keep them

These came up explicitly in this session. Hold the same lines; do not drift.

1. **No anti-detection features.**
   - The user asked for "simulated mouse dragging with some added randomization" because coordinate typing felt "sus". This was **declined**: adding randomization or humanization so the bot slips past the game's anti-bot checks is anti-cheat evasion.
   - Do not add random delays, jitter, human-like mouse paths, scheduling to "look human", or anything whose purpose is avoiding detection.
   - Plain, deterministic dragging for *speed* is fine; see §9.
2. **No bypassing input filtering.**
   - When the zoom test first did nothing, the user was told: if the game deliberately ignores injected input (`LLMHF_INJECTED`), we will **not** work around it (kernel or driver input, Interception, and so on).
   - Benign causes were fixed instead, for example running as administrator. The game accepted `SendInput`, so this has not been an issue since.
3. **Be honest about ToS risk, once, without lecturing.**
   - The plan moved from "scanner on an alt account, user sends marches by hand" to "auto-send marches". The user was told once that auto-sending is full botting and raises the account's ban risk, and that it is their call.
   - Do not repeat this every turn. Do state it once if scope grows again.
   - Background running was discussed (§9). Recommending an Android emulator with ADB is a *convenience* choice, not an evasion technique. Keep it framed that way.
4. **The user cares about their account.** Prefer safe failure: skip a node, stop, or ask. Never risky actions taken on unknown screen state.

---

## 4. Files

```
screen_scanner/
  node_scan.py      main tool: setup / goto / calibrate / check / scan / set-offset
  wininput.py       Windows input + window lookup via ctypes (SendInput, FindWindow, ...)
  screen_scan.py    first, generic tool: watch screen for templates; also `snip` (crop a target);
                    exports ScreenCapture (mss) used by node_scan
  zoom_test.py      standalone mouse-wheel sequence sender (default d4,u1,d1, 0.4 s) – early test
  tests/test_offline.py   offline regression suite (simulated tilted map, fakes for input)
  README.md         user-facing instructions
  HANDOFF.md        this file
  requirements.txt  mss, numpy, opencv-python   (tests also run under pytest if installed)
  .gitignore        scan_log.csv, targets/*.png, zoom_shots/, scan_config.json, nodes.json,
                    check.png, nodes_*.json   (all user-local data stays out of git)
  targets/.gitkeep  user's template PNGs live here locally (cropland.png, gem.png)
```

The user's local, untracked files are `scan_config.json`, `targets/*.png`, `nodes_cropland.json`, `nodes_gem.json` and `check.png`.

---

## 5. Architecture

### 5.1 Layers in `node_scan.py`

| Layer | Functions | Notes |
|---|---|---|
| Config | `load_config`, `save_config`, `need`, `DEFAULTS`, `CONFIG_VERSION=4` | JSON next to the script. Migrations v1→v4 run once and are saved. |
| Detection (pure) | `whiteness`, `find_all`, `screen_center`, `pixel_to_tile`, `untilt`, `fit_calibration`, `detect` | No input or screen access, so fully testable offline. |
| Game window | `grab_game`, `click_point`, `patch_at`, `wait_for_change` | Everything is relative to the game window's client area. |
| Game actions | `check_stop`, `goto`, `countdown` | `goto` is the only thing that drives the search popup. |
| Node list | `nodes_file`, `load_nodes`, `load_target`, `merge` | One file per target: `nodes_<target>.json`. |
| Commands | `cmd_setup`, `cmd_goto`, `cmd_calibrate`, `cmd_check`, `scan_pass`, `cmd_scan`, `cmd_offset` | CLI in `main()`. `Stop`/Ctrl+C prints "Stopped."; `RuntimeError` prints its message and exits. |

### 5.2 Data flow of a scan

```
cmd_scan → load_target (image, per-target offset, --threshold) → results = nodes_<target>.json
  for each scan point (snake order, step 16, far-edge point added):
     goto(x,y): click magnifier → wait until X box appears → click X, type → click Y, type → click Go → sleep load_wait
     grab_game() → detect(): find_all (multi-scale whiteness match) → pixel_to_tile (tilt-corrected)
                   → keep if inside this point's cell (±step/2 + 1 tile) → round(+offset)
     merge() into node list (dedupe ±2 tiles), write file after every point
  after the pass: nodes inside the scanned rectangle not seen this pass → GONE (removed)
```

### 5.3 Config keys (`scan_config.json`, v4)

| Key | Default | Meaning |
|---|---|---|
| `version` | 4 | Config schema version (migrations in `load_config`). |
| `window_title` | `Rise of Kingdoms` | Exact title for `FindWindowW`. |
| `target` | `targets/cropland.png` | Default target when `--target` is not given. |
| `threshold` | 0.9 | Match score cut-off. Croplands score ~0.95–1.0 and food-node lookalikes ≤ 0.85. |
| `scales` | [0.75, 1.35] | Sprite sizes tried relative to the target PNG, in steps of 0.05. |
| `load_wait` | 1.2 | Seconds after pressing search before the screenshot. |
| `popup_wait` | 3.0 | Longest wait for the search popup to appear. |
| `click_wait` | 0.15 | Pause after each click and after typing. |
| `key_delay` | 0.01 | Pause between key events; 0 sends a whole box in one `SendInput`. |
| `step` | 16 | Tiles between scan points (square cells). |
| `tile_offset` | [0, 0] | Added to every result (fractional allowed) before rounding. |
| `target_offsets` | {} | Per target, e.g. `{"gem": [-1, 0]}`; replaces `tile_offset` for that target. Set with `set-offset`. |
| `window_points` | {} | `search_button`, `x_field`, `y_field`, `go_button`, in px inside the client area. Recorded by `setup`. |
| `vx`, `vy` | null | Screen px per +1 map X / +1 Y **at the window middle**. Set by `calibrate`. |
| `perspective` | null | Tilt constant `c` (per px). Set by `calibrate`. |

### 5.4 Node record (`nodes_<target>.json`)

```json
{"x": 384, "y": 54, "score": 0.99, "first_seen": "2026-10-05T16:20:01", "last_seen": "2026-10-05T16:31:12"}
```

---

## 6. Core logic, and why it is that way

1. **Coordinate jumps, not dragging.** The coordinate bar shows the map centre. Typing X/Y into the search popup puts the camera exactly on a known tile, so there is no drift. The user found the feature: magnifier on the top-left coordinate bar → popup `#: [kingdom] X: [ ] Y: [ ] [search]`.

2. **Whiteness matching.** `whiteness(img) = min(B, G, R)`. Map icons are white. Grass, pink territory and teal territory all have one low channel, so the background turns dark whatever its colour. This made matching background-independent. It also works with the game's day/night tint: a night cropland scored 0.98.

3. **Multi-scale matching.** The camera is tilted, so icons lower on screen are drawn bigger. One cropland near the bottom measured 1.22× the size of one near the top. `find_all` tries 0.75–1.35×, keeps local maxima above the threshold, then keeps the best-scoring size per spot (radius 0.6 × template).

4. **Tilt model.** Ground size relative to the window middle is `k(y) = 1 + c·(y − centre)`.
   - Horizontal tile size scales with `k`; vertical tile size with `k²` (ground-plane foreshortening).
   - `pixel_to_tile` undoes it with `dx' = dx/k` and `dy' = (1 − 1/k)/c`, then solves `[vx vy]·d = (dx', dy')`.
   - `untilt()` is the matching inverse warp, used only in calibration.

5. **Calibration (`fit_calibration`).**
   - It takes three shots: the start, start +4 X, and start +4 Y.
   - For each candidate tilt ratio from 1.0 to 1.8 (size at ¾ height ÷ size at ¼ height; steps of 0.04, then refined in steps of 0.01), it untilts all three shots and phase-correlates the middle 60% × 60% crop.
   - It keeps the ratio with the best combined confidence.
   - The **"moved" guard** rejects fits whose shift is under 10 px per tile. Without it, plain ground's resampling pattern matched itself at zero shift.
   - The camera moving +d tiles means the ground moves the opposite way, hence the minus signs.
   - The first version, a simple two-band ratio, gave a wrong 1.52 on the user's PC with Y confidence 0.04. The search version is robust; the simulated tilt is recovered to within 0.03.

6. **Cells.** Each scan point keeps only sprites within ±step/2 tiles **plus 1 tile of overlap**. That keeps the coordinate maths in the accurate middle of the screen. The overlap exists because, without it, a node exactly on a cell border was lost to rounding by both neighbours (found in simulation). `merge` drops the duplicate (±2 tiles).

7. **Far edge.** If the rectangle's width isn't a multiple of the step, an extra point is added at `x2`/`y2`. Previously `scan 300 900 330 931` never looked at X 325–330.

8. **GONE logic.** After a pass, nodes inside the rectangle (expanded by step/2) whose `last_seen` predates the pass start are removed. Nodes outside the rectangle are never touched. Separate files per target mean a gem scan never "GONEs" croplands.

9. **Popup handling in `goto`.**
   - Before clicking the magnifier, it grabs a 60 × 24 patch around the X box.
   - After clicking, it polls every 0.1 s until the patch changes (mean absolute difference > 20), up to `popup_wait`.
   - It retries the magnifier once. If the popup still hasn't appeared, it raises `RuntimeError` and types nothing.
   - The earlier fixed 0.5 s wait was too short on the user's PC.

10. **Typing.** Each box is cleared and filled with End + 4×Backspace + digits as one `type_keys` call (coordinates are at most 4 digits). Jump input time fell from about 3 s to about 1.3 s.

11. **Window-relative everything.**
    - `client_rect(title)` gives the drawable area's screen position and size.
    - Screenshots cover only that area, and clicks are window-relative. The window can be moved; a **resize requires re-snipping and recalibrating**.
    - The map centre is the middle of the client area. Measuring from the monitor centre had put results about a tile off.
    - The process is DPI-aware (`SetProcessDpiAwareness(2)`), so pixels are physical.

12. **Offsets.**
    - The popup's coordinate for a node may differ from where its icon is drawn.
    - `check` prints unrounded "exact" values. `set-offset NAME DX DY` stores a per-target correction.
    - Two gems once showed **opposite** errors. That was a **zoom mismatch**, not an offset (see §7), so offsets are only for constant errors.

---

## 7. Measured facts about the user's setup

- **The game accepts `SendInput`** for wheel, clicks and keys.
- **UI hide:** the blue square button at the top left hides the UI. The coordinate bar stays visible.
- **Fly-home:** the castle button labelled `Space` flies home and resets the zoom.
- **Map loading** takes about 1 s after a jump. Icons pop in late; a screenshot taken before loading finishes showed no icons.
- **Zoom drifts between sessions.** Calibrations measured:
  - 60.5 / −41.5 px per tile (first; monitor-based, no tilt);
  - 55.2 / −39.0 (tilt 1.31);
  - **67.4 / −47.7, tilt 1.31, confidence 0.71 / 0.58, at 380,59 (current).**

  The Y/X ratio stays about 0.70 and the tilt is stable, so only the zoom changed. **A changed zoom silently ruins coordinates away from the screen centre.** The two gems that were off by +1.6 and −1.4 were explained exactly by a 1.2–1.3× zoom difference. A zoom check before scanning is planned (§8).
- **View size:** at the current zoom, a screen covers about 28 × 22 tiles. The 16-tile step is conservative.
- **Scores:** croplands (corn + shield) score 0.95–1.0; food nodes (corn, no shield) up to 0.85; other white icons ≤ 0.78. Gem nodes detect fine (0.96–0.99).
- **Reference point:** an alliance cropland at X:318 Y:936 (popup) was used to verify early calibrations.
- **The mouse cursor can cover sprites** in screenshots (user report).

---

## 8. Safety guards

### 8.1 In the code now

| Guard | Where | What it does |
|---|---|---|
| Emergency stop | `check_stop()` before every click and typing step in `goto`, and every 0.5 s while `scan --loop` waits | Mouse in the screen's top-left corner (≤ 5, 5) raises `Stop` → "Stopped.". Ctrl+C also works. |
| Countdown | `countdown()` before setup (6 s per point), goto, calibrate, check and scan (3 s) | Gives the user time to focus the game before any input. |
| Window lookup | `wi.client_rect()` | Clear error if the window is missing or minimized. Captures and clicks stay inside the client area. |
| Popup verification | `goto` + `wait_for_change` | Never types unless the X box visibly appeared; one retry, then a `RuntimeError` with fix hints. |
| Field clearing | `goto` | End + 4×Backspace before digits, so leftovers can't corrupt coordinates (coordinates are ≤ 4 digits). |
| Input blocked | `wi._send` | `SendInput` returning short raises `RuntimeError` with an "administrator terminal" hint (UIPI). |
| Missing setup or calibration | `need()` | Exits with which keys are missing and which command creates them. |
| Missing target | `load_target` | Exits with the exact `snip` command to run. |
| Stale module | top of `node_scan.py` | Exits if `wininput.py` is older (no `type_keys`) and says how to update. |
| Config migration | `load_config` | v1 → drop monitor-based points; v2/3 → drop old threshold and offset, ask to recalibrate; v4 → old `click_wait` 0.3 to 0.15. Saved once. |
| Calibration sanity | `fit_calibration`, `cmd_calibrate` | Rejects no-movement fits (`RuntimeError`); warns on confidence < 0.1, a tilt at the search limit, and a 16×16 cell not fitting on screen. |
| Detection margin | threshold 0.9 | Above the food-node lookalike (≤ 0.85). `check` shows near misses (≥ 0.7) in yellow with scores so the threshold can be tuned per target. |
| Clean detection | `cmd_check` | Boxes are drawn on a copy, so annotation never affects detection. The save result is checked; `check.png` opens automatically. |
| Coordinate integrity | `detect` / `merge` / `scan_pass` | Cell filter + 1-tile overlap, ±2 dedupe, GONE limited to the scanned rectangle, separate files per target, far-edge coverage. |
| Partial progress | `scan_pass` | Node file written after every scan point, so a stop or crash keeps results. |

### 8.2 Known gaps (not guarded yet)

1. **No check that a jump landed.** Nothing reads the coordinate bar. If the typing was garbled, the scan records nodes at the wrong coordinates.
   - Fix idea: read the coordinate bar after each jump (fixed font, fixed place, digit templates) and compare.
   - Alternative: a ground-shift check against the previous screen.
2. **Zoom drift is undetected.** Planned: at scan start, make one test jump of +4 X and phase-correlate. If the px per tile differs from `vx` by more than 5%, stop with a "press fly-home / recalibrate" message.
3. **A popup already open at start** may be closed by the magnifier click. The patch change is then mistaken for the popup opening, and the typing goes nowhere. Possible fix: also confirm the X box looks open (template of the empty box).
4. **Unexpected game screens** (event popups, disconnect, "logged in elsewhere", city view) are not detected. Scans continue blindly.
5. **The user touching the mouse or keyboard** mid-scan is not detected, only the corner stop.
6. **`key_delay` 0** could drop keys in the game. If used, it is the user's choice; the README says to revert to 0.01.
7. **`zoom_test.py`** sends wheel events to whatever is under the cursor, with no window check. It is a test tool only.
8. **No log file** of actions taken. Only console output and the node files exist.

### 8.3 Required before any march automation

These are non-negotiable for the march loop (§9):
- **Watch-only mode first:** read and print march states, click nothing.
- **`--dry-run`:** decide and print the intended sends, perform none.
- **One live march** before five.
- **Verify the node is free immediately before sending** (occupied and targeted checks on a fresh screen). If unsure, skip the node.
- **Never act on an unknown screen state.** Every action is preceded by a check that the expected UI is visible (template anchors). Anything unexpected → stop and report; don't guess.
- **Claimed set:** two marches never target the same node.
- **Per-march cap:** a counter for the user's "~50 nodes then return".
- **An action log** (timestamped CSV/JSON) of every send, stop and skip.
- **Keep the corner stop and Ctrl+C working inside every loop and wait.**
- **No randomization or humanization** (see §3).

---

## 9. Agreed plan (in order)

### Step 1: scan speed-ups (no user input needed; next thing to build)

A jump is about 2.5–3 s now; the target is about 1.3–1.7 s.
- **Tab instead of clicking the Y box.** The user confirmed Tab moves X → Y. Drop `y_field` from setup, or keep it as a fallback.
- **Park the cursor outside the game window** before each screenshot. It blocks sprites. Do *not* park it in the top-left corner: that is the stop signal.
- **Adaptive load wait:** screenshot every 0.2 s after Go and continue once the detections (or the frame) are stable for two frames. Minimum about 0.4 s; maximum `load_wait`.
- **Pipeline:** run detection on the previous screenshot in a thread while the next jump's clicks and typing happen.
- **Predicted size per row:** use the tilt so `find_all` doesn't try 13 scales everywhere. This needs the row where the target was snipped; record it at snip time, or keep a narrow scan.
- **Non-square step:** about 20 X × 16 Y at the current zoom. The cell filter and edge logic must then take separate X and Y steps.
- **Timing breakdown** printed per jump: popup, input, load, detect. Measure before optimizing further.
- **Zoom check at scan start** (gap 2 in §8.2).
- **Clipboard paste was rejected:** it is no faster than batched typing.
- **Dragging is deferred** to the sweep only, if it's still too slow after the above. It needs coordinate-bar reading after each drag, plus a coordinate jump at row starts to cancel drift. Drags must start on empty ground. Estimated gain is 20–30%.

### Step 2: freshness + `sweep`

- **Nodes regenerate randomly**, so a list goes stale in minutes. The user wants a **full sweep about every 10 minutes**, plus a **local scan around each march** when it finishes.
- **Track `last_scanned` per cell** (a grid of step-sized cells), not per pass.
  - **GONE:** the cell was scanned after the node's `last_seen`.
  - **Stale:** the cell is older than about 10 minutes. Stale nodes stay candidates but must be checked on arrival.
- **Sweep scans the stalest cells first.** It is **interruptible**: a finished march pre-empts it, and the sweep resumes where it left off.
- **Print the expected sweep time** (jumps × measured seconds per jump) and warn if it exceeds the cadence. Rough sizes:

  | Area | Jumps | Time |
  |---|---|---|
  | 200 × 200 tiles | ~130 | ~3–4 min |
  | 300 × 300 tiles | ~290 | ~7–8 min |
  | 400 × 400 tiles | ~500 | too long for 10 min |

### Step 3: is the node free? (needs screenshots)

- **Occupied:** another icon on top of the node while someone gathers. Template-match a second image in a small area above each detected node.
- **Targeted:** a dotted march line from any direction, in several colours. Lines animate, so take two frames about 0.3 s apart and diff them in a **ring around the node** (excluding the node itself). If the changed pixels form a line pointing at the node, it's targeted.
  - Colour-agnostic.
  - A march passing by can cause a false "targeted". That's the safe direction: skip the node.
- **When:** snapshot during scans; **always re-check immediately before sending.**

### Step 4: march automation (needs screenshots + clicks)

**The user's spec:**
- 5 marches, each hard-coded at **50k troops**.
- A march chains about **50 nodes**, then returns home.
- When a node is finished, the game auto-sends the march home. The tool then selects that march, presses **S** to stop it in the field, picks the **nearest free node** and sends it there. Repeat.

**Reading march states:** from the **march list UI** (fixed position; status icons for gathering, marching, returning, idle), not from the map. The tool tracks each march's position itself, as the node it was sent to. No coordinate reading is needed for that.

**UI conflict:** scanning has used hidden UI, but the march list is UI. **Decision: keep the UI visible and exclude UI rectangles from detection.** Cells are mid-screen anyway. Avoid toggling the UI.

**Zones:**
- The user defines one or more rectangles that marches can reach, splitting areas at **blockades**.
- Each march works within one zone.
- The nearest node is chosen by straight-line distance within the zone, skipping claimed nodes.

**On a march finishing:**
1. Local 3×3 scan around it (about 15 s; includes the free checks).
2. Pick the nearest free node in the zone.
3. If none, widen to 5×5, or fall back to the sweep list and verify on arrival.
4. Send, claim the node, and increment the march's counter.

**Later optimization:** read each march's remaining-time timer (digit templates) and pre-scan about 30 s before it finishes.

**Rollout:** watch-only → dry run → one live march → five (see §8.3).

### Step 5: possible later

- **Running in the background** (the user asked).
  - Today the tool takes over the real mouse, keyboard and visible screen.
  - Options discussed: a spare PC; a second Windows session over RDP; posting messages to the window (many games ignore them); a VM (needs GPU passthrough, no).
  - **Recommended later:** the **Android emulator + ADB** version of the game. Taps and typing via `adb shell input` and screenshots via `adb exec-out screencap` mean the user's mouse is untouched and the window can be covered. It would run whichever account is logged in inside the emulator.
  - **Design implication now:** keep input and capture behind a small backend interface (today: `wininput` + mss), so an ADB backend can be swapped in by config.
- **Notifications** (Discord webhook, Telegram or ntfy) with coordinates. This came from the earlier "scanner on alt, send by hand" plan.

---

## 10. Testing

- **Offline suite:** `python tests/test_offline.py`, or `python -m pytest -q tests`. 9 tests, about 30 s, all passing at hand-off.
  - It renders a **simulated tilted map** (perspective ground via `cv2.remap`, billboards scaled by `k`, optional static "UI" blocks) and **fakes** `goto`, `grab_game`, `ScreenCapture`, input and time.
  - It covers: the tilt model inverse; calibration recovery (ratios 1.0, 1.3, 1.5; with and without UI; within 0.03 ratio and 0.5 px); rejecting static screens; a full scan finding every node at exact tiles, including cell-border nodes; GONE on the next pass; far-edge points; `goto` waiting for the popup, typing the exact key sequence, retrying once then aborting, and honouring the corner stop; per-target results and legacy file move; offsets and threshold overrides; config migration.
  - `reset()` restores every patched global between tests. Follow that pattern: earlier, leaked patches made a test exercise a fake.
- **Real-screenshot checks** used the user's uploads during this session. They are **not in the repo** (private game screenshots). They confirmed:
  - croplands 0.95–1.0 vs food ≤ 0.85, on day and night tints;
  - correct rejection of an unloaded screen;
  - cropland at (318, 936) computed correctly.
- **Never claim Windows behaviour was verified** unless the user ran it. Say which parts ran where.

---

## 11. Open questions for the user (blocking steps 3–4)

1. **Screenshots** at the scanning zoom with the UI visible:
   - a gem node someone is gathering on;
   - a gem with a march line coming in (several directions and colours if possible);
   - the march list showing gathering, marching, returning and idle.
2. **The exact click flow** to send a field march to a node: what to click, which menus appear, and which buttons to press.
   - Does **S** stop the march in place?
   - Can a stopped march be sent straight to gather?
3. **Is the "~50 nodes then return" a load or capacity limit or a count?** Is 50k troops per march just a fixed setting?
4. **Farming zones:** rectangles, plus where the blockades are.
5. **Do they want background running soon?** If so, build the swappable input and capture backend now.

---

## 12. Commit history (screen_scanner)

```
f662322 Explain a stale wininput.py instead of crashing
eb7692a Type coordinates faster and support per-target offsets
c10199e Pick the target per run and keep separate results per target
b69fc8d Cover the far edge of a scan area that is not a multiple of the step
9b69fb4 Fit the camera tilt by searching instead of a two-band ratio
4b64dc3 Wait for the search popup to open before typing
ac859b9 Correct for the tilted camera and match sprites at several sizes
82ef0d8 Scan inside the game window instead of the whole monitor
e41557a Verify check.png was saved and open it on Windows
220a670 Print the full path of check.png
956b750 Lower the calibrate low-confidence warning to 0.05
41b9b3f Add a map node scanner driven by the coordinate search
9cffa76 Make the zoom test run a wheel sequence (default down 4, up 1, down 1)
81e2254 Add a mouse-wheel zoom test script
9a5de72 Add a lightweight on-screen object scanner
```
Plus the commit that adds this file and `tests/test_offline.py`.

## 13. Conventions

- **Python style:** match the existing style. Short docstrings that explain *why*, small pure functions for anything testable, and `RuntimeError` for user-actionable failures (main prints them cleanly).
- **New config keys:** bump `CONFIG_VERSION` when changing defaults that users have saved (`save_config` writes the full config). Add a migration in `load_config` that runs once and re-saves.
- **Commits:** descriptive messages, then push to `claude/ecstatic-feynman-libbqk`. **Do not open a PR** unless the user asks.
- **After each change, give the user:** the ZIP re-download instruction (§2), the exact commands to run, and the expected output.
