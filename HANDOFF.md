# Production Control System — Full Handoff

**Repo:** `Fonz0d1ac/Test` · **Branch:** `claude/production-control-board-157ift`
**Build:** `2026.07.30.11` (`BUILD_VERSION` / `BUILD_NOTE` at the top of `app.py`, shown in the
setup window and the first console line)
**Site:** NPV (Northstar Precision Vietnam) — Packaging / Assembly / Raw Part

> **PUSH WORKS.** It was 403 for two whole sessions and the old handoff told you to hand files
> over in chat instead — **that is no longer true and must not be repeated.** Everything is on
> GitHub: 18 files, full history. Commit and push normally. If a push ever 403s again, it is a
> repo-write permission on the Claude GitHub App, not a code problem.

---

## 0. Start here

### Resuming
The repo is cloned for you at session start. Read this file, then `git log --oneline` — every
build has a commit message explaining *why*, not just what. You do **not** need the user to
upload anything.

### Context discipline — three files must never be read whole
| File | Size | Why |
|---|---|---|
| `label_print.py` | 372 KB | **95% is embedded hex logo data.** Only ~17 KB is code. `grep`/`offset` only. |
| `shared_aheadbehind.html` | 233 KB | Chart.js is inlined so the page needs no CDN. |
| `app.py` | 218 KB / 4421 lines | Read the section you need (`grep -n "^# ─"` prints the section map). |

### The three load-bearing concepts (still sacred)
1. **`AREA` is computed once at startup from the Excel plan filename** (Packaging→PK / AssyD→AD /
   RawPart→RP). Everything derives from it. One codebase serves all three areas; **never hardcode
   an area.**
2. **`remaining_qty` is the single source of truth** for ETC / priority / report hours. Changed
   ONLY through `rebuild_final_pdo_cache()` / `inject_splits_into_cache()`.
3. **`_worker_cache` holds today's FULL roster; availability is filtered LIVE** per board GET via
   `_worker_available_now(w, now)` = `(start − 15min) ≤ now ≤ end`.

**Read §6 (safety guards) before touching anything.** They exist because each one has already
cost someone real production state.

---

## 1. Architecture

### Apps & data flow
```
        Excel plan (per area)   SQL Server (LOCALNPV / TECH_DATA)   Attendance Excel   MainDatabase.xlsx
               │                          │                              │                    │
               ▼                          ▼                              ▼                    ▼
 ┌──────────────────────── app.py (Control Board, 1 per area, Windows) ─────────────────────────┐
 │ plan→PDO list · SQL FG/WIP/License-Plate poll (60s) · attendance→live roster · state→data/*   │
 │ serves board.html :5050 · writes {AREA}_status.{json,js} every 5s to the share                │
 │ printing: BPT pick ticket (Canon, silent) + Zebra box label (raw ZPL) + queue-on-station      │
 └───────────────────────────────────────────────────────────────────────────────────────────────┘
                          │ shared folder: \\npvshare\...\dashboard_status\                 │
        {PK,AD,RP}_status.{json,js}  +  aheadbehind_status.js (written by combined_dashboard)│
              ┌───────────────────────────┼───────────────────────────────┬──────────────────┘
              ▼                            ▼                               ▼
 combined_dashboard.py :8080     shared_dashboard.html /            stations_only.html (file://,
 (rotator: /station + /aheadbehind, shared_aheadbehind.html /        PK+RP wall board, fits 1080p)
  reads *_status.json server-side, shared_rotator.html (file://,
  publishes aheadbehind_status.js) read *_status.js client-side)
```

### Deployment topology
- **3 board PCs** (one per area, Windows), each `python app.py` pointed at that area's plan Excel;
  each writes its own `{AREA}_status.{json,js}`. Board viewed locally at `http://127.0.0.1:5050/`.
  Printing (Zebra + Canon) runs here — pywin32 and MainDatabase must be reachable. **Run from
  LOCAL disk, not the share.**
- **1 big-screen PC** running `combined_dashboard.py` (needs the shared folder + SQL + attendance);
  it also publishes `aheadbehind_status.js` so the shared viewers can show Ahead/Behind.
- **Anyone** opens `\\npvshare\...\dashboard_status\shared_rotator.html` — no server, no firewall
  rule, no CDN.
- Share: `\\npvshare\Data\06_Operation\02.Production\02.Component\06. Planning\dashboard_status`

### Repo contents (all pushed)
```
app.py                        Control board backend (Flask + Tkinter) + print endpoints  4421 lines
templates/board.html          Control board UI (capacity window, at-risk alert, print modals)
templates/dashboard.html      Big-screen station status
templates/rotator.html        Big-screen rotator (fades /station ↔ /aheadbehind)
templates/aheadbehind.html    Big-screen Ahead/Behind chart
combined_dashboard.py         Big-screen app + aheadbehind_status.js writer
shared_dashboard.html         Firewall-free station viewer (file://, reads *_status.js)
shared_aheadbehind.html       Firewall-free Ahead/Behind (Chart.js INLINED)
shared_rotator.html           Firewall-free rotator
stations_only.html            PK+RP wall board, fits 1920×1080 without scrolling
bpt.py                        BPT pick-ticket engine + MainDatabase reader + A5 HTML render
boxlabel.py                   Zebra box-label engine (ModeQty formulas, DUMMY LP, stack/GW math)
label_print.py                ZPL renderer (raster → compressed ^GF) — 95% embedded logo data
printing.py                   win32 + silent-print helpers (lazy imports)
requirements-print.txt        openpyxl · pillow · python-barcode · pywin32(win)
PRINTING_README.md            Install + test steps for printing
HANDOFF.md                    This document
.gitignore                    data/, startup_config.json, worker_reset_*, print_settings_*, …
```
`data/`, `startup_config.json`, `worker_reset_{AREA}.json`, `print_settings_{AREA}.json`,
`lp_suppress_{AREA}.json`, `__pycache__` are **git-ignored runtime state** — which is what makes
`git pull` safe on a board PC (§6.10).

---

## 2. Core logic & pipelines — READ before changing anything

### 2a. PDO cache pipeline (`remaining_qty` is sacred)
`_raw_pdo_cache` (Excel only) → `rebuild_final_pdo_cache()` → `_pdo_cache` (served).
- `read_pdos()` reads the plan Excel (filtered to this area by column G) and attaches TimeStudy
  (SQL). **No date-window restriction — ALL orders load.** Each PDO carries: `id`(PO), `part`,
  `raw_part`, `desc`, `qty`, `dest`(F), `pack_type`(G), `pii_po`(Q), `vendor`(S), `ship_date`,
  `receive_date`, `days_left`, `finished_hrs`/`std_hours`, `remaining_qty`, `completed_qty`, splits.
- `rebuild_final_pdo_cache()`: `completed = FG_completed + WIP`, `remaining = qty − completed`,
  **drop if remaining ≤ 0** (finished orders leave the active list → Finished report only).
- `inject_splits_into_cache()`: adds synthetic split children (`PDO-xxx-A/B…`), reconciled against
  the authoritative PO-level FG total (`_fg_completed_cache[parent]`) so a fully-complete PO drives
  all its splits to 0 and flushes.
- `build_board_snapshot()` is **READ-ONLY** — never prunes, never saves. Pruning happens ONLY in
  the 60s poll on a freshly-rebuilt cache.

### 2b. WIP — PO-level, sticky across station moves
`Nhaplecuoingay_All` holds end-of-shift partial-box progress. `get_wip_deduction_latest(po)`
returns the MOST RECENT WIP entry across whatever station last logged it; applied at the **PO
level** so WIP survives X→Y reassignment.

### 2c. Attendance → `read_attendance_shift_info(area, date, include_lent_away=False)`
Reads the area's attendance `.xlsm` (monthly sheets like `T7-2026`), 3 rows/worker, **FULL
(non-read-only) mode** so red-corner comments are readable, from a temp copy (never saved). Sets
`_last_attendance_ok` (True only when TODAY's file+sheet+column read cleanly) — this guards the
prune (§6.2).

### 2d. Capacity model (net of breaks) — 6.73 / 8.47 / 9.08 / 6.90
`CAPACITY_HOURS_BY_TIER` (`CA`: {0:6.73, 2:8.47, 3:9.08}; `HC`: {0:6.90, 2:8.47, 3:9.08}) for full
shifts; `net_hours_for_windows(...)` for partial/noted windows; `compute_capacity_by_day(area,
from, to)` per-day for the report chart.

### 2e. Attendance time-window notes (partial-day) + shift inference
A red-corner comment on the STATUS cell spells out hours/area (e.g. `RP 6:00-11:00, PK 11:00-14:10`).
`parse_area_note(text, area)` keeps only this area's segments (times parsed `hh + mm/60`). A
half-day off-code + note → keep the worker, infer the shift from the work window.
**Fixed:** the Ca1/Ca2 boundary compared the note start against `CA1_END` (the rounded literal
`14.167`), but a `14:10` window start parses to `14.16667` < `14.167` → mis-tagged **Ca 1**. Now
compared against `_hm(14,10)`, computed the same way as the note time, so 14:10 → **Ca 2**.

### 2f. Worker roster & live availability
`refresh_cache()` (throttled ~5 min) sets `_worker_cache = get_available_workers(AREA)` = today's
FULL roster (each with `start`/`end`). **`_worker_available_now(w, now)` = `(start − 15min) ≤ now
≤ end`** is the single availability check, applied live in `build_board_snapshot()` and
`prune_expired_workers()`. `_active_worker_id_set()` = `start ≤ now ≤ end` (drives ETC).
`NEXT_SHIFT_OPEN_EARLY_HRS = 15/60`.

### 2g. ETC
`calc_etc()` uses `remaining_qty`, `time_study`, ACTIVE worker count. Runs only 06:00–22:20 with
≥1 active worker; outside production it FREEZES (`off_hours`); no active worker → PAUSED. Required
hours are per SINGLE operator; ETC scales by actual operators. Shift-boundary resets in the 60s poll.

### 2h. License Plate auto-queue — SELF-HEALING
`auto_queue_from_license_plate()` places a printed PO on its station iff it is (1) active/unfinished
(`remaining > 0`), (2) not already on the board, (3) not manually removed
(`lp_suppress_{AREA}.json`), (4) on a configured station. It **does NOT skip on existing FG/WIP** —
those log independently of the board, and the old behaviour permanently stranded any printed order
that picked up WIP before it was ever placed. A stale License-Plate SQL query returns `None` → the
cache is KEPT and a loud warning printed.

**Keep this.** The print action also places the PDO (`_issue_place_pdo`, §4h), which makes
auto-queue look redundant — it is not. See §7 "Why LP auto-queue stays" and §5.6.

### 2i. Splits
Synthetic children `PDO-xxx-A/B…` in `splits_{AREA}.json`. Per-split completion via
`get_split_live_completed(parent, station, cutoff)` (cutoff bound as a real `datetime`, not the raw
ISO string — that caused SQL error 241). `inject_splits_into_cache` reconciles against the PO-level
FG total. `parse_split_id()` verifies children against the splits file rather than pattern-matching
a prefix, which is why `SO-123-A` worked with no change (§2k).

### 2j. Dev / test mode
`DEV_MODE` (top of `app.py`, from `BOARD_DEV_MODE`; setup-window checkbox). When ON:
`write_dashboard_status()` no-ops (no big-screen/shared feed) and **`sql_write(cursor, sql, params,
label)`** logs `[dev-mode] SQL write SKIPPED …` instead of executing. **All board-originated
INSERT/UPDATE/DELETE must go through `sql_write()`.** Surfaced at startup, in `/api/board.dev_mode`,
and as a banner in `board.html`.
⚠ Printing is **not** gated by dev mode (it is not a SQL write, and the LP is currently a dummy).
When the real mint lands this becomes a trap — see §6.17.

### 2k. Work-order prefixes + the OVERDUE scan grace (build `.11`)
**`WORK_ORDER_PREFIXES = ('PDO', 'SO-')`** replaces the hardcoded `startswith('PDO')` in
`read_pdos()`.
- `PDO` stays **hyphen-less**, exactly as before, so no order that used to load can start being
  dropped.
- `SO-` carries the hyphen **deliberately** — bare `SO` starts plenty of ordinary words, and a
  stray cell like `SOMETHING` must not become a work order.
- Matching is **case-insensitive**: a lowercase `so-1234` vanishing silently is far worse than a
  stray match, which the area / qty>0 / ship-date checks drop anyway.

**`OVERDUE_GRACE_MIN = 20`** — the warehouse pulls and scans FG in ~20 min after the operator
actually finishes the box, so without a grace every on-time order flashed OVERDUE for 20 minutes.
Past ETC the countdown **holds at `00:00`**; only after the grace does it read **OVERDUE** and turn
the station card **red**.
- Published in the board snapshot as **`overdue_grace_min`** so `board.html`, `stations_only.html`,
  `dashboard.html` and `shared_dashboard.html` all read ONE number. Change it in `app.py`, not in
  each HTML file; the HTML falls back to 20 only for an old snapshot.
- Only the **in-progress ETC** colours a station — a long `queue_clear` must not (`isDuration`
  guard in `startETC`).
- A **paused** timer is frozen, so it can never be overdue. Every view checks this (§6.19).
- On the big screen and shared viewer, **`tick()` must also toggle the card class** — data refreshes
  only every 60s, so otherwise a station stays green for up to a minute past the grace. The OVERDUE
  pill counts `.card.overdue` in `tick()` for the same reason.
- `aheadbehind.html`, `rotator.html`, `shared_aheadbehind.html`, `shared_rotator.html` needed **no
  change** — no per-order ETC (verified: zero countdown references), and the rotators are iframe
  switchers.
- ⚠ Flask caches Jinja templates when `debug=False`. **Restart `combined_dashboard.py` after
  editing a template** or you will test the old one.

---

## 3. Excel plan column map (`COL`, 0-based) and per-area state files
`A`=pdo, `B`=part, `C`=raw_part, `D`=desc, `E`=qty, `F`=dest, `G`=pack_type (area filter +
production type), `H`=receive_date, `I`=ship_date, `J`=ship_mode, `K`=week (sample flag), `L`=note,
`M`=docs, `N`=vendor_name, `O`=received, `P`=progress, `Q`=pii_po, `R`=finished_hrs, `S`=vendor,
`T`=station_t, `U`=boxes, `W`=station_w. SAMPLE = the word "sample" in column K.

Per-area JSON in `data/` (git-ignored): `assignments_{AREA}.json` (**PDO placements**),
`workers_{AREA}.json` (**operator staffing**), `timers_{AREA}.json`, `splits_{AREA}.json`,
`priority_{AREA}.json`, `lp_suppress_{AREA}.json`, `notes_{AREA}.json`, `ng_{AREA}.json`,
`stations_{AREA}.json`, `worker_reset_{AREA}.json`, `print_settings_{AREA}.json`. Plus
`startup_config.json` (plan + attendance paths).

`{AREA}_status.{json,js}` (written every 5s): the full snapshot `{area, area_label, stations[],
workers[], queue[](pool), board{sid:{inprogress, queue, workers, timer{…}}}, pdo_status, fg_status,
shift_windows, capacity_hours, in_production, dev_mode, overdue_grace_min, ts, __written}`. The
`.js` is a JSONP wrapper (`window.NPV_STATUS["PK"]={…}`) for `file://` pages.

---

## 4. Printing subsystem — replaces the Excel `PrintLabel` macro

### 4a. MainDatabase.xlsx — real column map (validated against the live file)
`\\npvshare\Data\06_Operation\02.Production\02.Component\01. General\MainDatabase\MainDatabase.xlsx`,
sheet `MainDatabase`. **Headers on row 2, data from row 3; the sheet is several stacked sub-tables**
(each row-1 section header is its own independent list — rows are NOT aligned across sub-tables).

| Cols | Content | Keyed by |
|---|---|---|
| B, C, D | Time study: B=Main FG, C=TS, D=OP Q'ty setup (TargetOp) | B |
| G, H | Drw Revision: G=Part, H=Rev | G (blank for most FGs — sparse) |
| M, N | WI location: M=Part no, N=Vị trí | M |
| P…X | Qty/Box & Weight: P=Part(FG), Q=Qty/box, R=Weight/pcs, S=Max stack/1box, T=Box weight, U=Pallet weight, V=WeightPerBox, W=Box/pallet, X=Destination(bucket) | composite **Part & PartDes** |
| Y…AC | Component BOM: Y=FinishedGood, Z=Material code, AA=desc, AB=Usage, AC=Lot FG Qty | Y |
| AF…AK | **Packaging BOM: AF=Finished Good, AG=Material code, AH=Description, AI=Usage, AJ=Lot FG Qty, AK=Destination(bucket)** | AF + AK |
| BI, BJ | Type (big/small) | BI |
| BL, BM | WH location | BL |
| AV, AW | **Destination lookup: AV=Destination NAME, AW=Cont CODE** → CheckDes | AV |

**Market buckets are exactly `{PVN, Poland, Other}`.** `PartDes = F` if `F in ("PVN","Poland")` else
`"Other"` (macro-exact; `bpt.market_bucket`). **`CheckDes = XLOOKUP(F, AV, AW)`** — the plan-F
Destination NAME (e.g. `TRICAP`) resolves to the short code (e.g. `TRI`) that prints + is barcoded.
Validated: `1029887-329`/`TRICAP` → WI K20, qty/box 48, dest TRI, 7 materials.

### 4b. `bpt.py` — BPT packaging pick ticket (macro `ModePackPick` → Canon)
- `MainDB(path)` loads the file ONCE and builds indexes: `pack_by_fg`, `qtybox`/`qtybox_any`,
  `wt_by_key[(fg,bucket)]`, `wi_by_fg`, `ts_by_fg`, `targetop_by_fg`, `rev_by_part`, `loc_by_part`,
  `type_by_part`, `des_by_raw`.
- `build_bpt(...)`: per box → materials with **pick qty = RoundUp(box_qty / AJ) * AI**; last box
  gets the remainder `(Final)`; pallet-level materials get qty 1 on the box that opens each
  `box/pallet` group, else 0; `WC` prefix → `wooden_crate` routing flag.
- `render_html(ticket)` → **A5 portrait, one page per box**: title, the "LẤY TEM & TIÊU CHUẨN CV"
  banner as a full-width band under the header (first ticket only), three real Code128 barcodes
  (ticket code, destination, PDO), meta grid incl. `Box n/N · Pallet p/P · Standard`, and the 8-col
  material table. `.ticket:last-child` does **not** force a page break (that caused a trailing blank
  sheet on every job).

### 4c. `boxlabel.py` + `label_print.py` — Zebra box label (macro `ModeQty` → ZT421)
`build_box_labels(...)` ports the macro formulas verbatim: `QtyLblPrint = RoundUp(POQty/QtyPerBoxSTD)`;
per-box qty (last = remainder); stack limit (B9) / pallet GW (B12) with the partial-vs-full
last-pallet branches and the ModeStandard-off branch. Weights from MainDatabase (R/S/T/U/V/W).
Returns one label dict per box plus an assembled `sql_row` in the macro's exact INSERT column order.
**License Plate is still a DUMMY** and revision is blank — see §5.

### 4d. Packing standard = the single source of the box split (build `.8`)
**`bpt.resolve_pack_std(mdb, fg, bucket)` is the ONE resolver** for qty/box + box/pallet, used by
BOTH `build_bpt` and `boxlabel.build_box_labels` — they must never disagree, or a PDO prints N
tickets and a different number of labels.

**STRICT on the market bucket.** The old code did `qtybox.get((fg,bucket)) or qtybox_any.get(fg)`,
where `qtybox_any` is *first row seen* for that FG. On a bucket miss it silently returned **another
market's** standard. Reproduced: a 100/box PVN row used for a 50/box `Other` order turned a 100 pc
PDO into **one** ticket carrying the pallet, while `boxlabel` (no fallback) refused outright. That
was the reported "not splitting by boxes" bug.
- The one fallback kept is unambiguous: an FG with **exactly one** packing row in the file uses it
  and reports `source='only-row (…)'`. Two or more candidates and no bucket match = hard miss,
  listing the buckets that exist.
- Pallet (`PL`/`SPL`) and crate (`WC`) prefixes live in **print settings**, not code.
- **No `PL`/`SPL` material for the FG ⇒ no pallet tag at all**, instead of stamping every box
  "no pallet".

### 4e. Silent BPT printing (builds `.8`–`.10`)
Zebra was always silent (raw ZPL via pywin32). The **paper** side needs two steps because Windows
cannot send a PDF to a *named* printer without UI: **headless Edge/Chrome → PDF → SumatraPDF
`-print-to "<queue>" -print-settings "simplex,noscale"`**. `simplex` forces single-sided regardless
of the queue default — one ticket per box must never share a sheet.
- SumatraPDF is a single portable exe (auto-detected: `tools\` next to `app.py`, `%LOCALAPPDATA%`,
  Program Files, PATH; or `sumatra_path` in settings).
- `html_to_pdf` uses a **persistent dedicated profile** (`printing._PROFILE_DIR`, under `data/`).
  It must NOT be the operator's own Edge profile — headless would attach to their running instance
  and silently emit nothing — but a fresh `mkdtemp` per print cost real time (**12.19 s cold vs
  0.65 s warm**, and on Windows every new file is antivirus-scanned).
- **`printing.prewarm()`** runs at startup (`_print_warm_startup`) and again when the Issue modal
  opens (`/api/prewarm_print`), so the cold launch happens while the operator types the station.
- Missing browser or Sumatra → `print_html_silent` returns `(False, reason)` and the client falls
  back to the browser tab. `/api/printers` reports `silent_ready` / `silent_detail`.

### 4f. Print performance — measured, not guessed
Real board-PC timing that drove this:
`maindb 1.38s · zebra 0.48s · zebra_render 0.21s · bpt 7.97s · total 9.91s · 4 labels`

**Two theories died on that data. Trust the instrumentation, not extrapolation.**
- Zebra spooling is NOT slow (0.48 s for four labels). An earlier 2.25 s/label reading was a
  first-print artefact. **Batching ZPL into one job was planned and dropped — it buys nothing.**
- BPT is ~8 s and **roughly fixed regardless of box count**, so it dominates every issue.

Fixes applied:
- **`label_print._gfa_hex`** walked all 2.6 M pixels in Python per label (~0.19 s each). Now
  `img.tobytes()` + a 256-entry `translate` table. **Byte-for-byte identical**, verified at rotate
  0/90/270. ⚠ Watch the row padding: PIL leaves each row's spare bits at 0, and inverting turns them
  BLACK (a 1px stripe down every row) — they must be masked off.
- **ZPL ASCII compression** (`_compress_rows`): repeat counts (`G`–`Y` = 1–19, `g`–`z` = 20–400),
  `,` = fill row white, `:` = repeat previous row. **7029 KB → 835 KB** for 11 labels, lossless —
  verified with an independent decoder over 200 random-noise cases, all-white/all-black/>419-run
  edges, and the real label at every rotation. `^GFA` byte counts stay UNCOMPRESSED.
- **Logo + font caches.** `_bold_font` was re-opening the TTF per text element. ⚠ A "faster" rewrite
  of `_decode_gfa` produced wrong pixels and was reverted — the cache alone removes the cost, so
  **leave the per-pixel loop be.**
- **MainDatabase warm loop** (`_maindb_warm_loop`, refreshes just inside the 10-min TTL) and a
  re-warm after a print-settings save (which invalidates the cache — that is what the 1.38 s
  `maindb` reading was).
- **The BPT spool moved off the operator's wait** (`_start_bpt_job` → `/api/print_job/<id>`).
  Response time **9.91 s → ~0.05 s**. Silent-path readiness is still resolved synchronously so the
  client knows whether to use the tab fallback.

Net in-sandbox: 11 labels 2.88 s → 0.46 s to generate, 8.4× less data on the wire.

### 4g. `app.py` integration + `printing.py`
- **MainDatabase** parsed once, cached (`get_maindb`, 10-min TTL, path from settings, reload on
  change), kept warm in the background.
- Endpoints: **`/api/printers`** · **`/api/print_settings`** GET/POST · **`/api/issue_ticket`** POST
  · **`/api/print_job/<id>`** · **`/api/prewarm_print`** POST · **`/bpt/<token>`** GET.
- **`/api/issue_ticket`** resolves the PDO from `_pdo_cache`, opens MainDatabase, resolves the
  packing standard once, builds the BPT + box labels, prints the Zebra labels (RAW ZPL), starts the
  background Canon job, and **`_issue_place_pdo(pdo_id, station)`** places the PDO on the station
  (in-progress if free, else queue). Returns `{ok, lps, boxes[], std, bpt_token, bpt_pages,
  bpt_job, placed, timings, warnings}`.
- **UI:** a **🖨 button on each PDO card** → Issue modal; a top-bar **🖨 Print** → settings modal
  (Zebra + Canon printer, silent toggle + live readiness line, SumatraPDF path, pallet prefixes,
  MainDatabase path, email recipients).
- **Print-result dialog (`ov-printres`) — BLOCKING.** Every issue ends here: resolved packing
  standard (market, qty/box, box/pallet → box + pallet counts), BPT page count and where it went,
  a per-box table (box, qty, License Plate, ✓/✕ with the error), and warnings. `.ov` overlays in
  `board.html` already ignore backdrop/Esc, so **OK is the only exit**; `forceLoad()` runs on OK,
  not before, so the board cannot shift under the operator while they read.
  ⚠ **`bpt_printed` is False whenever a job is running** — the tab fallback must also check
  `!d.bpt_job` or the ticket prints TWICE (that bug was introduced and caught here).
  A background failure arriving after OK re-opens as the persistent banner (§6.20).
- **`printing.py`**: all win32 imports are **lazy** so it imports off-Windows.

---

## 5. NEXT PHASE — real License Plate mint + SQL insert

Everything below is analysed and agreed in principle but **not written**. The dummy LP and the
assembled `sql_row` are the seams.

### 5.1 LP format (verified against production data)
```
LicensePlate = Left(vendor,6) + yymmdd + serial(6)      e.g. V-0023 260812 200042
serial = area_base + MAX(last 5 digits today, same area) + 1
```
| Area | digit | base | first serial of day |
|---|---|---|---|
| AD Assembly | `0` | 0 | `000001` |
| PK Packaging | `2` | 200000 | `200001` |
| RP Raw Part | `5` | 500000 | `500001` |

The **first digit of the serial is the area code** — the convention exists precisely because all
three areas share one `[License Plate]` table. The counter is **per (area, day), shared across
vendors**. The SQL anchors from the right (`len-11`, `len-5`, `right(...,5)`), so a vendor code
shorter than 6 chars still parses.

⚠ **`boxlabel.make_dummy_lp` hardcodes `200000 + i`** — Packaging. It must become area-driven or
AD and RP will mint into PK's range.

### 5.2 The three area macros are otherwise identical
Byte-identical stack/GW formulas, `UsedLP` accumulation, reuse query, INSERT columns, and the
small-box logic incl. the `1024884-329` hardcode. **The only functional difference is the serial
digit.** One cosmetic difference: RP's small-box label sets `A1 = station` while AD/PK set
`A1 = station & "-ThungNho"` — all three still write the `-ThungNho` suffix to SQL. Looks like an
RP oversight; do not copy it.

### 5.3 Reuse-before-mint — why it exists
At month-end, planning splits an in-progress PDO (`PDO-0001` 75/100 → `PDO-0001` closed + `PDO-0002`
continues). Whatever partial quantity was logged against the OLD PDO must keep the **same LP**, or
the efficiency calculation — which joins on LP — loses that WIP. `PII_PO` is the family ID that
survives the split (assigned upstream in the MRP/D365 import, not in any macro here).

```sql
SELECT TOP 1 [License Plate] FROM [Nhaplecuoingay_All] a
WHERE PII_PO = ? AND PO <> ?
  AND NOT EXISTS (SELECT 1 FROM [FG_Database_All] b WHERE a.[License Plate] = b.[License Plate])
  <UsedLP exclusions>
GROUP BY [License Plate]
```
The macro runs this **per label**, accumulating `UsedLP`, so an 11-box PDO can reuse up to 11
orphaned plates.

### 5.4 Two macro bugs NOT to port
1. **Reuse has no `ORDER BY`** — confirmed live: **113 `PII_PO` groups with 2+ simultaneously
   eligible plates**, so SQL Server's pick is arbitrary. Fix: `ORDER BY MAX([Production date]) DESC`
   (the `GROUP BY` forces the aggregate).
2. **Print-then-insert.** If printing throws, the serial is consumed but never recorded, so the next
   run re-mints the same plate onto a physical label. **Flip to mint+insert committed, then print** —
   worst case becomes an orphan row and a gap, not a duplicate plate.

### 5.5 Why NOT the `area_counters` design from `License_Plate_Logic_Handoff.md`
That doc proposes MySQL with new tables (`license_plates`, `license_plate_events`, `area_counters`).
**It cannot be adopted while the Excel macros are live.** Our counter would say "next is 200043"
while the macro reads `MAX(right(lp,5))+1` off the shared table and mints 200043 too — collisions
in both directions. `MAX+1` also self-heals against unknown writers, which matters because
`Gen_other_info` (a legacy "print N extra copies" tool) mints into AD's `'0'` pool with
`VendorCode="V-0000"`. It is the right **phase-2** target, once the macros retire — and it would
need rewriting for SQL Server anyway (`ON DUPLICATE KEY UPDATE`/`LAST_INSERT_ID` are MySQL-only).

### 5.6 Proposed design
**New `lp.py`**, two queries instead of the macro's 2N:
1. **Reuse pass** — `SELECT TOP (n) …` with the ordering fixed; taking `TOP (n)` in one shot makes
   the within-run exclusion structural and replaces the `UsedLP` string entirely.
2. **Mint pass** — `MAX(...)+1` once, hand out consecutive serials for the remaining boxes.
3. Assign reused plates to the earliest boxes, then minted — matching the macro's order.

`boxlabel.build_box_labels(..., plates=None)` takes the list; the dummy stays as the dev/offline
fallback.

**Concurrency — the obvious fix is a trap.** `SELECT MAX(right(lp,5)) WHERE SUBSTRING(...)` is
**non-sargable** and table-scans; wrapping it in `XLOCK, HOLDLOCK` risks lock escalation to a
table-level exclusive lock on a production table — worse than the bug. Use instead:
- **`sp_getapplock`** (e.g. `LP_MINT_PK_260814`) — a cheap named mutex, no table locks. Fully
  serialises app-vs-app.
- **Read-back verification with retry** for app-vs-macro. Nothing can lock out the macro safely,
  and the macro never locked either. Residual risk is small: areas use disjoint serial ranges, so a
  collision needs the app *and* the macro minting in the same area in the same instant.

**Out of scope for the first pass:** small-box `-ThungNho` (incl. the `1024884-329` hardcode), the
component BOM ticket (`ModeBOMPick` → Canon02, `CB_PO_BOM_Ticket_Database`), the BPT audit INSERT
into `CB_Production_BOM_Print`, the wooden-crate dual-printer routing, and the new-part Outlook-COM
email.

### 5.7 Open questions blocking the build
1. **What are `tg_user` / `tg_dev`** in the macro — Windows username and machine name?
2. **Does `svcsqllocal` have DDL rights?** Not needed now; decides whether phase 2 needs a DBA.
3. **Is `Gen_other_info` still in use?** If yes it keeps consuming AD's pool — an argument for
   staying on `MAX+1` indefinitely.
4. Sign-off on the deviations in §5.4 (insert-then-print, and the `ORDER BY` fix, which makes us
   pick a *different* plate than the macro in those 113 groups).
5. A **single-box test PDO** and a safe window for the first live mint.

---

## 6. SAFETY GUARDS — every guard that protects operator state or production data

1. **Snapshot is READ-ONLY.** `build_board_snapshot()` never prunes or saves. A stale id renders as
   an empty bin; the next 60s poll cleans it authoritatively. (Fixed the old "orders bounce to
   queue" bug.)
2. **Prune wipe-guards.** `prune_stale_assignments` no-ops unless `_last_pdo_status == 'ok'`;
   `prune_expired_workers` no-ops unless `_last_attendance_ok`. A failed Excel read looks identical
   to "empty" — the guard skips and retries rather than wiping the board.
3. **Pruning ONLY in the 60s poll**, on a freshly-rebuilt cache — never on a board GET or dashboard
   write.
4. **Daily worker-staffing reset** (`reset_daily_worker_staffing`) clears carried-over **operator
   staffing only** (never PDO assignments), once per calendar day, persisted in
   `worker_reset_{AREA}.json`.
5. **One-operator-one-station** (server-side, `api_workers_assign`): staffing a worker strips them
   from every other station — the authoritative anti-double-booking, independent of how the client
   dragged.
6. **`remaining_qty` single source of truth** — mutated ONLY via `rebuild_final_pdo_cache` /
   `inject_splits_into_cache`. Splits reconcile against the PO-level FG total.
7. **Worker availability = one check** (`_worker_available_now`) used by the pool, each station's
   display, and the prune, so they can never disagree.
8. **LP auto-queue keyed on `remaining_qty`, not FG/WIP presence.** A stale License-Plate SQL query
   returns `None` → **keep the cache, warn loudly** (never wipe).
9. **Dev mode gates ALL SQL writes** (`sql_write`) + the big-screen/shared feed — one flag, surfaced
   at startup, in `/api/board.dev_mode`, and as a board banner.
10. **Data persistence across code updates.** Operator state lives in `data/*.json`, which is
    **git-ignored** — so `git pull` on a board PC cannot touch it. Never delete/overwrite `data/`,
    `startup_config.json`, `worker_reset_{AREA}.json`, `print_settings_{AREA}.json`. (Operators —
    not PDOs — auto-clear once at the start of a new production day by design, guard #4.)
11. **Worker-drag: newest drop wins** (`_wdropSeq`) — a slow earlier `/api/board` response cannot
    clobber the board after a newer drag.
12. **Printing today cannot write to production:** the LP is a DUMMY and **no SQL INSERT / email is
    performed**, so a mis-print cannot touch `[License Plate]`, `CB_Production_BOM_Print`, or send
    mail.
13. **MainDatabase opened read-only + cached**; a bad path fails the print call cleanly
    (`"Cannot open MainDatabase"`) without crashing the board. `printing.py` win32 imports are lazy,
    so a non-Windows / no-pywin32 host degrades to a "not printed" warning.
14. **Hardcoded secrets to remove (do NOT propagate):** the plan-sheet password `"12345"`, the SQL
    creds `svcsqllocal / KetnoilocalApp` (in `app.py` and `combined_dashboard.py` — move to
    env/secrets), and the six D365 plaintext passwords in the separate `Download_PO_received.py`.
15. **Packing standard is strict per market** (§4d). Never reintroduce a silent cross-market
    fallback — it collapses the box split invisibly. Refuse and name the buckets that exist.
16. **The two print engines must share one resolver.** `bpt` and `boxlabel` both call
    `resolve_pack_std`; if they ever diverge, a PDO prints N tickets and a different number of
    labels.
17. **When the real mint lands, dev mode MUST use dummy plates.** The mint is a `SELECT`, so
    `sql_write()` will *not* gate it — a dev run would consume real serials while skipping the
    INSERT, corrupting the live sequence. This needs an explicit `if DEV_MODE` branch.
18. **BPT double-print guard.** `bpt_printed` is False while a background job runs; the browser-tab
    fallback must also check `!d.bpt_job`, or the ticket prints twice.
19. **A paused timer can never be overdue.** Off-hours / no-operator freezes the estimate; without
    this check every station goes red overnight.
20. **A background print failure must never vanish.** If the Canon job fails after the operator has
    clicked OK, it re-opens as the persistent banner rather than a toast.
21. **Silent-path readiness is resolved synchronously**, before the background job starts, so the
    client always knows whether to open the tab fallback.
22. **Work-order matching is case-insensitive and `SO-` is hyphen-anchored** — the first prevents
    silently losing orders, the second prevents ordinary words becoming work orders.

---

## 7. Key decisions & rationale

- **AREA from the plan filename** — one codebase, three areas; never hardcode.
- **`remaining_qty` single source of truth**; **snapshot read-only; pruning only in the 60s poll**.
- **Prune must never wipe on a failed load** — destroying operator state is far worse than skipping
  a cycle.
- **Worker staffing is per-shift, reset once per day**; **availability filtered live** via one check.
- **Splits reconcile against the PO-level FG total** (authoritative "is it done").
- **Dev mode gates the feed + all SQL writes** — one loud flag.
- **Priority red = due within 1 day** — reserves red for genuinely imminent orders.
- **Planned capacity = per-day editor** (operators × Ca net-hour tiers), separate window, per area.
- **At-risk alert = banner + station-stripe pulse**, gated past 14:00; the big-screen banner covers
  all at-risk orders (in-progress, queued, unassigned).
- **Shared viewers read `.js` via `<script>` + hidden-iframe refresh** — the only CORS-free way to
  read local data from a `file://` page; **Chart.js inlined** and fonts system so no CDN is needed.
- **Printing = two paths, two printers.** Zebra = raw ZPL (crisp thermal, native barcodes, rotation
  handled once). BPT = the approved A5 HTML → PDF → named queue.
- **The print action queues the PDO on the station** — an explicit placement signal alongside LP
  auto-queue.
- **Packaging type via string prefix** (`PL`/`SPL`/`WC`), now configurable in print settings — the
  user explicitly chose this over a `packaging_type` enum table.
- **DUMMY LP first** — get the whole flow testable on real printers before touching SQL.
- **The BPT spools in the background.** The operator needs to *know* it printed, not *watch* it.
- **`overdue_grace_min` is published by the server**, so five views cannot drift apart.
- **Why LP auto-queue stays** (§2h): after the real mint lands it stops being a *rival* placement
  signal and becomes the **recovery layer over the same source of truth**. Direct placement is a
  one-shot write to `assignments_{AREA}.json`; auto-queue re-derives from SQL every 60 s, so a lost
  or rolled-back JSON, or a rebuilt board PC, still recovers. Guard #2 of `auto_queue_from_license_plate`
  (already on the board → skip) means it can never fight the print button.
  ⚠ **One race to fix when the mint lands:** `_record_lp_suppression()` reads `_license_plate_cache`,
  which only the 60s poll populates. Once app.py is the LP writer, a GL removing a just-printed
  order records **no** suppression, and the next poll re-adds it. Fix: seed
  `_license_plate_cache[po]` at print time, in the same place as the INSERT. (Bonus: the `has_lp`
  badge then appears immediately instead of up to 60 s later.)

---

## 8. Build history

| Build | What |
|---|---|
| `.1`–`.7` | Repo bootstrap; one-operator-one-station guard; 14:10→Ca 2 shift fix; priority red = ≤1 day; worker-drag optimistic-UI races; planned-capacity setup window; same-day at-risk alert; standalone/shared viewers; **printing subsystem** (BPT + Zebra, dummy LP). |
| `.8` | **Box split fixed** — strict per-market `resolve_pack_std` shared by both engines. Silent Canon path (headless render → SumatraPDF). Blocking print-result dialog. |
| `.9` | Printing **~14× faster** ZPL + lossless ZPL compression (7029→835 KB); warm MainDatabase; simplex BPT; ticket layout rework (banner, barcodes, `PDO:`, pallet tag); per-stage timings. |
| `.10` | **BPT spools in the background** (9.91 s → 0.05 s response); persistent + prewarmed browser profile; lean headless flags; render/spool split in the logs. |
| `.11` | **`SO-` work orders** alongside `PDO`; **20-minute scan grace** before OVERDUE, with the station card turning red — applied to the control board, wall board, big screen and shared viewer. |

Also this session: `stations_only.html` reworked into a **two-area (PK + RP) wall board** with
per-area freshness badges, no selector, fitting 1920×1080 without scrolling (row-proportional flex
plus a `scale()` backstop). All missing templates and shared viewers restored to the repo.

---

## 9. Config reference (verify per deployment)

### `app.py` (top of file)
- `BUILD_VERSION` / `BUILD_NOTE` — bump on every change (glance-checkable per PC).
- `DEV_MODE` — from `BOARD_DEV_MODE`; setup-window checkbox overrides.
- `EXCEL_PLAN` / attendance paths — overridden by the setup window / `startup_config.json`.
- `PLAN_SHEET='1.FollowPlanForGL'`, `DATA_START_ROW=6`.
- `WORK_ORDER_PREFIXES = ('PDO','SO-')`, `OVERDUE_GRACE_MIN = 20`.
- SQL: `SQL_SERVER='VTN1PRDSQL002'`, `UID='svcsqllocal'`, `PWD='KetnoilocalApp'`,
  `SQL_DB_TECH='TECH_DATA'`, `SQL_DB_PROD='LOCALNPV'`. `svcsqllocal` already has INSERT on
  `[License Plate]` (the macro proves it).
- `DASHBOARD_STATUS_DIR`, `CAPACITY_FILES={'AD','PK','RP':…}`.
- `PRODUCTION_START=6.0`, `CA1_END=14.167`, `PRODUCTION_END=22.333`, `NEXT_SHIFT_OPEN_EARLY_HRS=15/60`.
- Printing: `_MAINDB_DEFAULT`, `_EMAIL_DEFAULT`, `_print_settings_defaults()` (zebra/canon/maindb/
  email/silent_bpt/sumatra_path/browser_path/pallet_prefixes/crate_prefixes) →
  `data/print_settings_{AREA}.json`.

### `bpt.py` / `boxlabel.py` / `label_print.py` / `printing.py`
- Buckets `{PVN, Poland, Other}`; `PALLET_PREFIXES=('PL','SPL')`, `CRATE_PREFIXES=('WC',)`.
- `label_print`: canvas 2007×1298 dots (6.689"×4.327" @300dpi), `ROTATE=90` (flip to 270 if upside
  down), default printer `ZDesigner ZT421-300dpi ZPL`.
- `printing._PROFILE_DIR`, `_LEAN_FLAGS`, `find_browser()`, `find_sumatra()`, `silent_ready()`.
- Deps: `openpyxl · pillow · python-barcode · pywin32(win)` (`requirements-print.txt`).

### `combined_dashboard.py`
- Own `SQL_CONN_STR` (ODBC Driver 17), `CAPACITY_FILES`, `DASHBOARD_STATUS_DIR`,
  `DEFAULT_OVERDUE_GRACE_MIN = 20` (fallback only — the authoritative value comes from each board's
  snapshot), `PORT = 8080`.
- Carries its **own** shift/break model (`SHIFTS`, `parse_capacity_file`) that overlaps `app.py`'s
  capacity code but is not identical — aggregate capacity for the chart vs per-worker roster.
  Check before assuming they can share a module.

---

## 10. Planned / not started

1. **Real LP mint + `[License Plate]` INSERT** (§5) — the next work item.
2. **Remaining macro modes:** small-box `-ThungNho`, component BOM ticket, BPT audit INSERT,
   wooden-crate dual-printer routing (the silent path makes this possible; the browser dialog never
   could), new-part Outlook-COM email (must never block printing).
3. **Printing layout confirmations:** FG revision source (G/H by FG is blank — where does it come
   from?), whether Ver/Location should populate for packaging materials, ZPL rotation (90 vs 270)
   on the real ZT421, and that the **compressed** ZPL prints identically to the old uncompressed one.
4. **`app.py` refactor.** At 4421 lines it is the structural outlier, and its own section banners
   are already stale (the block labelled "API: NG flag" is 429 lines and contains all the printing
   endpoints). Target split: `config.py` · `store.py` · `plan_excel.py` · `db.py` · `capacity.py` ·
   `etc.py` · `api_*.py` · `ui_setup.py`, leaving `app.py` ~400 lines of wiring.
   **Sequencing (agreed):** do the LP work first on the current structure — refactoring *and*
   adding production SQL writes in the same window makes a floor failure impossible to attribute.
   Then extract **incrementally, one module per change**, each shipped and run for a day. `db.py`
   falls out of the LP work naturally rather than as separate churn.
5. **Dashboard-feed cleanup:** sweep stale `.tmp` files in `dashboard_status`; trim the ~1 MB
   snapshot (drop the unassigned-PDO pool the station views do not use).
6. **D365 goods receipt** — batch-file export, or hardened in-app Selenium with ONE service account;
   kill the six plaintext D365 passwords regardless.
7. Optional: a **module build-stamp check** at startup so a half-copied update announces itself
   (much less pressing now that `git pull` is atomic).

---

## 11. Known gaps / to verify on a real deployment
1. **SQL schemas** for `Nhaplecuoingay_All` / `License Plate` / `CB_Production_BOM_Print` were given
   verbally — watch the console for SQL errors when the real writes land.
2. **Ahead/Behind SQL half** (`combined_dashboard.py`) uses ODBC Driver 17.
3. **Attendance notes / shift inference** validated against synthetic data + the 14:10 fix —
   validate on real files.
4. **Split records never self-clean** (`splits_{AREA}.json`) — harmless inert data after completion.
5. **Printing cannot be exercised off-Windows**: `print_zpl_raw` (pywin32) and SumatraPDF. The
   headless-render half *was* verified against real Chromium. Confirm on the board PC: barcodes
   scan, ZPL orientation, compressed output prints identically, both printers reachable.
6. **`combined_dashboard.py` served pages use a CDN** (Chart.js + fonts) for the Ahead/Behind half;
   the **shared** pages are fully self-contained.
7. **`stations_only.html` sizing** was tuned at 1920×1080 with 12 stations per area. More stations
   still fit (row-proportional flex + `scale()` backstop) but cards shrink; check readability at
   viewing distance on the real wall screen.
8. ⚠ **Headless `--window-size=1920,1080` reports a 993 px viewport**, which makes screenshots look
   like the layout has a gap at the bottom. It does not. Measure inside a frame that is genuinely
   1080 tall.

---

## 12. How to run
- **Control board (each area PC):** `git pull` → `python app.py` → setup window (verify plan +
  attendance paths, must validate green; optional 🧪 Dev/test checkbox) → in-app console → board at
  `http://127.0.0.1:5050/`. **Plan filename MUST start with `Packaging` / `AssyD` / `RawPart`.**
- **Printing:** `pip install -r requirements-print.txt` → put `SumatraPDF.exe` in `tools\` next to
  `app.py` → top bar **🖨 Print** to set queues, MainDatabase path and pallet prefixes → hover a PDO
  card → **🖨** → station + prod time → **Issue & queue**. See `PRINTING_README.md`.
- **Big screen:** `python combined_dashboard.py` → `http://<pc-ip>:8080/` (also starts the
  `aheadbehind_status.js` writer). **Restart it after editing any template.**
- **Shared viewers:** open `\\npvshare\...\dashboard_status\shared_rotator.html`.

---

## 13. For the next session
1. **Push works — use it.** Commit and push each change; do not fall back to handing files over in
   chat. If the user asks for files anyway, they can still be attached, but the repo is the record.
2. Restate the three load-bearing concepts (§0) and read the **safety guards (§6)** before touching
   anything. When touching workers / `remaining_qty` / splits, follow §2 + §6.
3. **The next work item is §5 — the real LP mint.** It is blocked only on the five answers in §5.7.
   Route every new SQL write through `sql_write()`, wrap the mint in `sp_getapplock`, insert before
   printing, and give dev mode an explicit dummy-plate branch (§6.17).
4. **Measure before optimising.** This session two confident performance theories were wrong and the
   instrumentation caught both. `[issue timing]` and `[bpt print]` lines exist for that reason.
5. **Verify claims against the artefact**, not against intent: the byte-identical `_gfa_hex` check,
   the independent ZPL decoder, the DOM measurement that contradicted a screenshot. Each of those
   caught a real bug or a false alarm.
6. Printing still cannot be tested off-Windows. Anything touching pywin32 or SumatraPDF needs the
   board PC.
