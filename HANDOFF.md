# Production Control System — Full Handoff

**As of:** session ending build `2026.07.30.7` · NPV (Northstar Precision Vietnam) — Packaging / Assembly / Raw Part
**Repo:** `Fonz0d1ac/Test` · **Working branch:** `claude/production-control-board-bk0blz`
**Build stamp:** `BUILD_VERSION` / `BUILD_NOTE` at the top of `app.py`, shown in the setup window + first console line.

> **PUSH IS BLOCKED (whole session).** `git push` → **HTTP 403**, GitHub MCP write → **403 "Resource
> not accessible by integration."** It is a repo-write permission on the session's GitHub integration,
> not a code problem. **Everything is committed locally on the branch; every change is delivered to the
> user as files.** The user has decided: **work locally, commit, do NOT push, hand over files.** The
> account `Fonz0d1ac` *owns* the repo, so the fix is on their side (grant the Claude GitHub App write
> access via the *Installed* GitHub Apps → Configure, or push the delivered git bundle from their own PC).
> A build-version note scheme quirk: stamps read `2026.07.30.x` for continuity even though the real
> calendar during this session was mid-August; keep bumping the `.x`.

---

## 0. TL;DR for the next session

This is a **multi-app system** plus a **ticket/label printing subsystem** added this session.

1. **`app.py` + `templates/board.html`** — the per-area **Control Board** (one instance per area:
   Packaging / Assembly / Raw Part). Flask on :5050 behind a Tkinter setup window. **Main app.**
2. **`combined_dashboard.py` + `templates/{rotator,dashboard,aheadbehind}.html`** — the **big-screen**
   app on :8080 that rotates a **Station Status** report and an **Ahead/Behind** report. **Now IN the repo.**
3. **`shared_dashboard.html` / `shared_aheadbehind.html` / `shared_rotator.html`** — **firewall-free,
   server-free** `file://` viewers that run off the shared drive and read the boards' published data.
4. **`stations_only.html`** — a big-visual single-area station map (area selector) `file://` viewer.
5. **Printing subsystem (NEW): `bpt.py`, `boxlabel.py`, `label_print.py`, `printing.py`** + the
   `app.py` print endpoints — replaces the Excel `PrintLabel` macro (BPT packaging pick ticket +
   Zebra box label + License Plate). **License Plate is a DUMMY for now; no SQL writes / email yet.**

**Three load-bearing concepts (unchanged, still sacred):**
- **`AREA` is computed once at startup from the Excel plan filename** (Packaging→PK / AssyD→AD /
  RawPart→RP). Everything derives from it. One codebase serves all three areas; **never hardcode area.**
- **`remaining_qty` is the single source of truth** for ETC / priority / report hours. Changed ONLY
  through `rebuild_final_pdo_cache()` / `inject_splits_into_cache()`.
- **`_worker_cache` holds today's FULL roster; availability is filtered LIVE** per board GET via
  **`_worker_available_now(w, now)`** = `(start − 15min) ≤ now ≤ end`.

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
 │ NEW: /api/issue_ticket prints BPT (Canon, via browser) + Zebra box label (ZPL, win32) and     │
 │      QUEUES the PDO on the chosen station; reads MainDatabase.xlsx (cached) for packing data   │
 └───────────────────────────────────────────────────────────────────────────────────────────────┘
                          │ shared folder: \\npvshare\...\dashboard_status\                 │
        {PK,AD,RP}_status.{json,js}  +  aheadbehind_status.js (written by combined_dashboard)│
              ┌───────────────────────────┼───────────────────────────────┬──────────────────┘
              ▼                            ▼                               ▼
 combined_dashboard.py :8080     shared_dashboard.html /            stations_only.html (file://,
 (rotator: /station + /aheadbehind, shared_aheadbehind.html /        single area, big-visual map)
  reads *_status.json server-side, shared_rotator.html (file://,
  publishes aheadbehind_status.js) read *_status.js + aheadbehind_status.js client-side)
```

### Deployment topology
- **3 board PCs** (one per area, Windows), each `python app.py` pointed at that area's plan Excel;
  each writes its own `{AREA}_status.{json,js}`. Board viewed locally `http://127.0.0.1:5050/`.
  Printing (Zebra + Canon) runs here (pywin32 + MainDatabase reachable). Run each copy from LOCAL disk.
- **1 big-screen PC** running `combined_dashboard.py` (needs shared folder + SQL + attendance); it now
  also publishes `aheadbehind_status.js` to the share so the shared viewers can show Ahead/Behind.
- **Anyone** opens `\\npvshare\...\dashboard_status\shared_rotator.html` (or the individual shared
  pages) — no server, no firewall rule, no CDN.
- `dashboard_status` folder: `\\npvshare\Data\06_Operation\02.Production\02.Component\06. Planning\dashboard_status`

### Repo contents (all committed on the branch; NOT pushed — 403)
```
app.py                        Control Board backend (Flask + Tkinter) + print endpoints  ~4200 lines
templates/board.html          Control Board UI (+ planned-capacity window, at-risk alert, print modals)
combined_dashboard.py         Big-screen app (station + ahead/behind) + aheadbehind_status.js writer
templates/dashboard.html      Big-screen Station Status page (+ at-risk banner/pulse)
templates/rotator.html        Big-screen rotator (fades /station ↔ /aheadbehind)
templates/aheadbehind.html    Big-screen Ahead/Behind chart (est-pace dotted line REMOVED)
shared_dashboard.html         Firewall-free station viewer (file://, reads *_status.js)
shared_aheadbehind.html       Firewall-free Ahead/Behind (Chart.js INLINED, reads aheadbehind_status.js)
shared_rotator.html           Firewall-free rotator (fades the two shared pages; centered selector)
stations_only.html            Big-visual single-area station map (area selector)
bpt.py                        BPT packaging pick ticket engine + MainDatabase reader + A5 HTML render
boxlabel.py                   Zebra box-label engine (ModeQty formulas, DUMMY LP, stack/GW math)
label_print.py                Zebra ZPL renderer (raster → rotated ZPL) — Phase-1a, sample→live data
printing.py                   win32 helpers (list printers, send RAW ZPL) — lazy imports
requirements-print.txt        openpyxl · pillow · python-barcode · pywin32(win)
PRINTING_README.md            Install + test steps for the printing feature
HANDOFF.md                    This document
.gitignore                    ignores data/, startup_config.json, worker_reset_*.json, __pycache__, *.pyc
```
`data/`, `startup_config.json`, `worker_reset_{AREA}.json`, `print_settings_{AREA}.json`, `__pycache__`
are git-ignored runtime state.

---

## 2. Core logic & pipelines (READ before changing anything)

### 2a. PDO cache pipeline (remaining_qty is sacred)
`_raw_pdo_cache` (Excel only) → `rebuild_final_pdo_cache()` → `_pdo_cache` (served).
- `read_pdos()` reads the plan Excel (filtered to this area by column G), attaches TimeStudy (SQL).
  No date-window restriction — ALL orders load. Each PDO dict carries: `id`(PO), `part`, `raw_part`,
  `desc`, `qty`, `dest`(F), `pack_type`(G, = production type), `pii_po`(Q), `vendor`(S), `ship_date`,
  `receive_date`, `days_left`, `finished_hrs`/`std_hours`, `remaining_qty`, `completed_qty`, splits, …
- `rebuild_final_pdo_cache()`: `completed = FG_completed + WIP`, `remaining = qty − completed`,
  **drop if remaining ≤ 0** (finished orders leave the active list → Finished report only).
- `inject_splits_into_cache()`: adds synthetic split children (`PDO-xxx-A/B…`), reconciles against
  the authoritative PO-level FG total (`_fg_completed_cache[parent]`) so a fully-complete PO drives
  all its splits to 0 and flushes.
- The board snapshot (`build_board_snapshot`) is **READ-ONLY** — never prunes/saves. Pruning happens
  ONLY in the 60s poll on a freshly-rebuilt cache.

### 2b. WIP — PO-level, sticky across station moves
`Nhaplecuoingay_All` holds end-of-shift partial-box progress. `get_wip_deduction_latest(po)` returns
the MOST RECENT WIP entry across whatever station last logged it; applied at the PO level so WIP
survives X→Y reassignment.

### 2c. Attendance → `read_attendance_shift_info(area, date, include_lent_away=False)`
Reads the area's attendance `.xlsm` (monthly sheets like `T7-2026`), 3 rows/worker, FULL (non-read-only)
mode so red-corner comments are readable, from a temp copy (never saved). Sets module global
**`_last_attendance_ok`** (True only when TODAY's file+sheet+column read cleanly) — guards the prune.

### 2d. Capacity model (net of breaks) — 6.73 / 8.47 / 9.08 / 6.90
`CAPACITY_HOURS_BY_TIER` (`CA`: {0:6.73, 2:8.47, 3:9.08}; `HC`: {0:6.90, 2:8.47, 3:9.08}) for full
shifts; `net_hours_for_windows(...)` for partial/noted windows. `compute_capacity_by_day(area, from, to)`
per-day for the report chart.

### 2e. Attendance time-window notes (partial-day) + shift inference
A red-corner comment on the STATUS cell spells out hours/area (e.g. `RP 6:00-11:00, PK 11:00-14:10`).
`parse_area_note(text, area)` keeps only this area's segments (times parsed `hh + mm/60`). A half-day
off-code + note → keep the worker, infer the shift from the work window.
**FIX this session:** the Ca1/Ca2 boundary compared the note start against `CA1_END` (the rounded
literal `14.167`), but a `14:10` window start parses to `14.16667` < `14.167` → mis-tagged **Ca 1**.
Now compared against **`_hm(14,10)`** (computed the same way as the note time), so a 14:10 start → **Ca 2**.

### 2f. Worker roster & live availability
`refresh_cache()` (throttled ~5 min) sets `_worker_cache = get_available_workers(AREA)` = today's FULL
roster (each with `start`/`end`). **`_worker_available_now(w, now)` = `(start − 15min) ≤ now ≤ end`** is
the single availability check, applied live in `build_board_snapshot()` and `prune_expired_workers()`.
`_active_worker_id_set()` = `start ≤ now ≤ end` (drives ETC). `NEXT_SHIFT_OPEN_EARLY_HRS = 15/60`.

### 2g. ETC
`calc_etc()` uses `remaining_qty`, `time_study`, ACTIVE worker count. Runs only 06:00–22:20 with ≥1
active worker; outside production it FREEZES (off_hours); no active worker → PAUSED. Required hours are
per SINGLE operator; ETC scales by actual operators. Shift-boundary resets in the 60s poll.

### 2h. License Plate auto-queue — SELF-HEALING
`auto_queue_from_license_plate()` places a printed PO on its station iff it is (1) active/unfinished
(`remaining > 0`), (2) not already on the board, (3) not manually removed (`lp_suppress_{AREA}.json`),
(4) on a configured station. It **does NOT skip on existing FG/WIP**. A stale License-Plate SQL query
returns `None` → the cache is KEPT and a loud warning prints.
**NOTE (this session):** the new **print action also queues the PDO on the station** (`_issue_place_pdo`,
§4d) — printing is now an explicit placement signal alongside LP auto-queue.

### 2i. Splits
Synthetic children `PDO-xxx-A/B…` in `splits_{AREA}.json`. Per-split completion via
`get_split_live_completed(parent, station, cutoff)` (cutoff bound as a real `datetime`, not the raw ISO
string — that caused SQL error 241). `inject_splits_into_cache` reconciles against the PO-level FG total.

### 2j. Dev / test mode
`DEV_MODE` (top of `app.py`, from `BOARD_DEV_MODE`; setup-window checkbox). When ON:
`write_dashboard_status()` no-ops (no big-screen/shared feed) and **`sql_write(cursor, sql, params, label)`**
logs `[dev-mode] SQL write SKIPPED …` instead of executing. **All board-originated INSERT/UPDATE/DELETE
must go through `sql_write()`.** Surfaced at startup, in `/api/board` `dev_mode`, and a banner in board.html.
(Note: **printing is NOT gated by dev mode** — it's not a SQL write, and the LP is currently a dummy.)

---

## 3. Excel plan column map (`COL`, 0-based) and per-area state files
`A`=pdo, `B`=part, `C`=raw_part, `D`=desc, `E`=qty, `F`=dest, `G`=pack_type (area filter + production
type), `H`=receive_date, `I`=ship_date, `J`=ship_mode, `K`=week (sample flag), `L`=note, `M`=docs,
`N`=vendor_name, `O`=received, `P`=progress, `Q`=pii_po, `R`=finished_hrs, `S`=vendor, `T`=station_t,
`U`=boxes, `W`=station_w. SAMPLE = the word "sample" in column K.

Per-area JSON in `data/` (git-ignored): `assignments_{AREA}.json` (**PDO placements**),
`workers_{AREA}.json` (**operator staffing**), `timers_{AREA}.json`, `splits_{AREA}.json`,
`priority_{AREA}.json`, `lp_suppress_{AREA}.json`, `notes_{AREA}.json`, `ng_{AREA}.json`,
`stations_{AREA}.json`, `worker_reset_{AREA}.json`, **`print_settings_{AREA}.json`** (NEW). Plus
`startup_config.json` (plan+attendance paths).

`{AREA}_status.{json,js}` (written every 5s): the full board snapshot `{area, area_label, stations[],
workers[], queue[](pool), board{sid:{inprogress, queue, workers, timer{…}}}, pdo_status, fg_status,
shift_windows, capacity_hours, in_production, dev_mode, ts, __written}`. The `.js` is a JSONP wrapper
(`window.NPV_STATUS["PK"]={…}`) for `file://` pages.

---

## 4. Printing subsystem (NEW this session) — replaces the Excel `PrintLabel` macro

The macro's "IN" button was a **5-job issuer** (§5 of the spec). We built the two most important paths.
Reference docs the user provided: `BPT_Production_Ticket_System_Spec.md` and the raw `macro.txt`.

### 4a. MainDatabase.xlsx — real column map (validated against the live file)
`\\npvshare\Data\06_Operation\02.Production\02.Component\01. General\MainDatabase\MainDatabase.xlsx`,
sheet `MainDatabase`. **Headers on row 2, data from row 3; the sheet is several stacked sub-tables**
(each row-1 section header is its own independent list — rows are NOT aligned across sub-tables).

| Cols | Content | Keyed by |
|---|---|---|
| B, C, D | Time study: B=Main FG, C=TS, D=OP Q'ty setup (TargetOp) | B |
| G, H | Drw Revision: G=Part, H=Rev | G (FG revision comes back **blank** for most FGs — sparse data) |
| M, N | WI location: M=Part no, N=Vị trí | M |
| P…X | Qty/Box & Weight: P=Part(FG), Q=Qty/box, R=Weight/pcs, S=Max stack/1box, T=Box weight, U=Pallet weight, V=WeightPerBox, W=Box/pallet, X=Destination(bucket) | composite **Part & PartDes** (CB col = `P&X`) |
| Y…AC | Component BOM: Y=FinishedGood, Z=Material code, AA=desc, AB=Usage, AC=Lot FG Qty | Y |
| AF…AK | **Packaging BOM: AF=Finished Good, AG=Material code, AH=Description, AI=Usage, AJ=Lot FG Qty, AK=Destination(bucket)** | AF + AK |
| BI, BJ | Type (big/small): BI=Part number, BJ=Type | BI |
| BL, BM | WH location: BL=Part number, BM=Location | BL |
| AV, AW | **Destination lookup: AV=Destination NAME, AW=Cont CODE** → CheckDes | AV |

**Market buckets are exactly `{PVN, Poland, Other}`.** `PartDes = F` if `F in ("PVN","Poland")` else
`"Other"` (macro-exact; see `bpt.market_bucket`). The composite key `Part & PartDes` selects qty/box +
weights; packaging BOM rows are filtered `AF==part AND AK==PartDes`. **`CheckDes = XLOOKUP(F, AV, AW)`**
— the plan-F **Destination NAME** (e.g. `TRICAP`) resolves to the short **code** (e.g. `TRI`) that
prints + gets barcoded. Validated: `1029887-329`/`TRICAP` → WI K20, qty/box 48, dest code TRI, 7
materials; `5144355-329`/`NP` → WI N27, TS 340, qty/box 1080, 4 materials.

### 4b. `bpt.py` — BPT packaging pick ticket (macro `ModePackPick`, prints to Canon)
- `MainDB(path)` loads the file ONCE and builds indexes: `pack_by_fg`, `qtybox`/`qtybox_any`,
  `wt_by_key[(fg,bucket)]` (Q/R/S/T/U/V/W), `wi_by_fg`, `ts_by_fg`, `targetop_by_fg`, `rev_by_part`
  (G→H), `loc_by_part` (BL→BM), `type_by_part` (BI→BJ), `des_by_raw` (AV→AW).
- `build_bpt(mdb, fg, po, dest, station, pro_time, po_qty, fg_desc)`: per box → materials with
  **pick qty = RoundUp(box_qty / AJ) * AI**; last box gets `po_qty % qpb` `(Final)`;
  **PL/SPL pallet rule** (prefix-detected as the user requested — no enum): qty 1 on the first box of
  each `box/pallet` group, else 0; **WC** prefix → `wooden_crate` routing flag.
- `render_html(ticket)` → **A5 portrait** ticket per box (approved layout): title "BPT for Packaging
  materials", **three Code128 barcodes** (ticket-code/box-count top-right, destination, PDO), meta grid,
  8-col material table (No/Part No/Description/Ver/BOM data/Qty/Location/Type), and the
  **"LẤY TEM & TIÊU CHUẨN CV" banner on the FIRST ticket of the PDO only** (box 1). Barcodes are real
  Code128 (python-barcode SVG data-URIs) — **not** the font trick.

### 4c. `boxlabel.py` + `label_print.py` — Zebra box label (macro `ModeQty`, prints to ZT421)
- `boxlabel.build_box_labels(mdb, po, part, po_qty, dest, station, pii, desc, vendor_code,
  production_type, …)` — formulas ported **verbatim** from the macro:
  `QtyLblPrint = RoundUp(POQty/QtyPerBoxSTD)`; per-box qty (last = remainder);
  stack limit (B9) / pallet GW (B12) with the partial-vs-full last-pallet branches and the
  **ModeStandard-off** branch (stack 0 + minimal weight). Weights from MainDatabase (R/S/T/U/V/W).
  Returns one label dict per box (ready for `label_print.build_zpl`) + an assembled `sql_row`.
- **License Plate is a DUMMY** (`make_dummy_lp` → `VendorCode6 + yymmdd + serial-from-200001`), and
  **revision is blank** — both per the user's request. **No `[License Plate]` INSERT is performed.**
- `label_print.py` (Phase 1a): builds the label elements (station, PII-PO, market badge `MKT_<code>`,
  PDO, item, qty, revision, WI, stack limit, pallet GW, License plate), renders to a 1-bit raster,
  rotates 90° (media feeds narrow-edge first), emits `^GF` ZPL at `^PW1298 ^LL2007`. `--rotate 270`
  flips it. Real native Code128 barcodes.

### 4c-bis. Packing standard = the single source of the box split (build `.8`)
**`bpt.resolve_pack_std(mdb, fg, bucket)` is now the ONE resolver** for qty/box +
box/pallet, used by BOTH `build_bpt` and `boxlabel.build_box_labels` — they must never
disagree, or a PDO prints N tickets and a different number of labels.
- **STRICT on the market bucket.** The old code did
  `qtybox.get((fg,bucket)) or qtybox_any.get(fg)`, where `qtybox_any` is *first row
  seen* for that FG. On a bucket miss it silently returned **another market's**
  standard: a 100/box PVN row used for a 50/box `Other` order turned a 100 pc PDO into
  **one** ticket carrying the pallet, while `boxlabel` (which had no fallback) refused
  outright. That was the reported "not splitting by boxes" bug — reproduced and fixed.
- The one fallback kept: an FG with **exactly one** packing row in the file uses it and
  reports `source='only-row (…)'`. Two or more candidates and no bucket match = hard
  miss, listing the buckets that exist.
- Pallet materials (`PL`/`SPL`) and crate (`WC`) prefixes moved to **print settings**
  (`pallet_prefixes` / `crate_prefixes`) — a new prefix is config, not code.
- Ticket now shows **Box n/N · Pallet p/P · Standard** and marks the box that opens a
  pallet; zero-qty pallet rows print greyed rather than vanishing. `.ticket:last-child`
  no longer forces a page break (killed the trailing blank sheet).

### 4c-ter. Silent BPT printing (build `.8`)
Zebra was always silent (raw ZPL via pywin32). The **paper** side now has a silent path
in `printing.py`: **headless Edge/Chrome → PDF → SumatraPDF `-print-to "<queue>"`**.
Windows cannot send a PDF to a *named* printer without UI, hence the two steps;
SumatraPDF is a single portable exe (auto-detected, or `sumatra_path` in settings).
`html_to_pdf` uses a **throwaway `--user-data-dir`** — with a shared profile, headless
attaches to the operator's already-open Edge and silently produces nothing.
Missing browser or Sumatra → `print_html_silent` returns `(False, reason)` and the
client falls back to the browser tab. `/api/printers` reports `silent_ready`/`silent_detail`.

### 4c-quater. Print performance (build `.9`) — an 11-box PDO took 30–40s
Four separate costs, all addressed; timings are now measured per stage rather than guessed
(`[issue timing]` console line + `timings` in the `/api/issue_ticket` response).
- **`label_print._gfa_hex` walked all 2.6M pixels in Python** per label (~0.19s each, and
  these are slow PCs). Now `img.tobytes()` + a 256-entry `translate` table, both C speed —
  **byte-for-byte identical output**, verified at rotate 0/90/270. Watch the row padding:
  PIL leaves the spare bits of each row at 0, and inverting turns them BLACK (a 1px stripe
  down every row) — they must be masked off.
- **ZPL ASCII compression** (`_compress_rows`): repeat counts (`G`–`Y` = 1–19, `g`–`z` =
  20–400), `,` = fill row with white, `:` = repeat previous row. **7029 KB → 835 KB** for
  11 labels, lossless — verified with an independent decoder over 200 random-noise cases,
  all-white/all-black/>419-run edge cases, and the real label at every rotation. The
  `^GFA` byte counts stay UNCOMPRESSED; decompression is transparent to the printer.
- **Logo + font caches** (`_logo_cache`, `_font_cache`). `_bold_font` was re-opening the
  TTF per text element. NB: a "faster" rewrite of `_decode_gfa` produced wrong pixels —
  it was reverted; the cache alone removes the cost, so leave the per-pixel loop be.
- **MainDatabase warm loop** (`_maindb_warm_loop`, daemon thread, refreshes just inside the
  10-min TTL). It lives on `\\npvshare`; with only a lazy TTL, whoever clicked 🖨 first
  after expiry paid for a cold network parse of a ~17k-row workbook.

Net in-sandbox: 11 labels 2.88s → 0.46s to generate, and 8.4× less data on the wire.

### 4c-quinquies. BPT layout + simplex (build `.9`)
- **Single-sided enforced**: `printing.print_pdf_silent` passes SumatraPDF
  `-print-settings "simplex,noscale"`. One ticket per box must never share a sheet.
  The browser-tab fallback can't force this — the Canon queue default applies there.
- Banner **"LẤY TEM & TIÊU CHUẨN CV"** is now a full-width band under the title (was in the
  left column of the header flex row).
- **Ticket-code** barcode larger/longer; **destination** barcode is the biggest and sits
  **left-aligned under Dest** (was small and right-aligned); **PDO** barcode enlarged with
  `write_text=False` (no caption). `code128_datauri` gained `write_text`; `module_width`
  is the knob for physical length.
- `PO No:` → **`PDO:`** throughout the ticket.
- **No `PL`/`SPL` material for the FG ⇒ no pallet tag at all** (`has_pallet_mats`), instead
  of stamping every box "no pallet".
- Destination barcode is identical on every page — rendered ONCE, not per box.

### 4c-sexies. BPT wait (build `.10`) — measured on the REAL board PC
`maindb 1.38s · zebra 0.48s · zebra_render 0.21s · bpt 7.97s · total 9.91s · 4 labels`.
**Two earlier theories died on this data — trust the instrumentation, not extrapolation:**
- Zebra spooling is NOT slow (0.48s for 4 labels, incl. 0.21s generation). An earlier
  2.25s/label reading was a first-print artefact. **Batching ZPL into one job was
  planned and then dropped — it would buy nothing.**
- BPT is ~8s and **roughly fixed regardless of box count**, so it dominates every issue.

Fixes:
- **The BPT now spools on a background thread** (`_start_bpt_job` → `/api/print_job/<id>`).
  The Zebra labels and the station placement finish in <1s, so the operator was waiting
  ~8s purely to watch a PDF reach the Canon. The result dialog opens immediately with the
  BPT line spinning and fills it in on completion. Response time 9.91s → ~0.05s.
  **Silent-path readiness is still checked synchronously**, so the client knows at once
  whether to use the browser-tab fallback.
  **A failure after the operator clicks OK re-opens as the persistent banner** — a BPT that
  never printed must never be lost.
  ⚠ `bpt_printed` is now False whenever a job is running; the tab fallback must check
  `!d.bpt_job` too or the ticket prints TWICE (that bug was introduced and caught here).
- **Persistent browser profile** (`printing._PROFILE_DIR`, under `data/`). Was a fresh
  `mkdtemp` per print — measured **12.19s cold vs 0.65s warm**, and on Windows every newly
  created file is antivirus-scanned. Still isolated from the operator's own Edge profile,
  which is why the flag exists at all (a shared profile makes headless attach to their
  running instance and silently emit nothing).
- **Prewarm** (`printing.prewarm`): once at startup (`_print_warm_startup`) and again when
  the Issue modal opens (`/api/prewarm_print`), so the cold launch happens while the
  operator is typing the station rather than after they click.
- **Lean headless flags**; dropped `--run-all-compositor-stages-before-draw` (it's for
  screenshots) and cut the virtual-time budget 8000→2000.
- **`[bpt print] render Xs · spool Ys`** splits the browser half from the Canon half.
- MainDatabase is **re-warmed in the background after a print-settings save** — the save
  invalidates the cache, which is exactly what the 1.38s `maindb` reading was.

### 4d. `app.py` integration + `printing.py`
- **MainDatabase** parsed once, cached (`get_maindb`, 10-min TTL, path from settings, reload on change).
- Endpoints: **`/api/printers`** (installed printers via win32, `[]` off-Windows) · **`/api/print_settings`**
  GET/POST (`data/print_settings_{AREA}.json`: `zebra_printer`, `canon_printer`, `maindb_path`,
  `email_recipients`) · **`/api/issue_ticket`** POST · **`/bpt/<token>`** GET (serves the A5 BPT page
  with an injected `window.print()`).
- **`/api/issue_ticket {pdo_id, station, pro_h, pro_m, do_bpt, do_boxlabel}`**: resolves the PDO from
  `_pdo_cache`, opens MainDatabase, builds the BPT + box labels, **prints the Zebra box label** (RAW
  ZPL via `printing.print_zpl_raw`), stashes the BPT in a bounded token cache (≤50) for the browser to
  fetch + print, then **`_issue_place_pdo(pdo_id, station)`** places the PDO on the station (in-progress
  if free, else queue; mirrors `/api/assign`). Returns `{ok, lps, bpt_token, placed, warnings}`.
- **UI (board.html):** a **🖨 button on each PDO card** → Issue modal (station datalist, prod time,
  which paths); a top-bar **🖨 Print** button → settings modal (Zebra + Canon printer with the
  installed-printer datalist, silent-print toggle + live readiness line, SumatraPDF path, pallet
  prefixes, MainDatabase path, email recipients).
- **Print-result dialog (`ov-printres`, build `.8`) — BLOCKING.** Every issue ends here: resolved
  packing standard (market, qty/box, box/pallet → box + pallet counts), BPT page count and where it
  went, a per-box table (box, qty, License Plate, ✓/✕ with the error), and warnings. `.ov` overlays
  in this file already ignore backdrop/Esc, so **OK is the only exit**; `forceLoad()` runs on OK, not
  before, so the board can't shift under the operator while they read. Replaced the auto-fading toast.
- **`printing.py`**: all win32 imports are **lazy** so it imports off-Windows; `list_printers()`,
  `print_zpl_raw(zpl, printer)`, `print_pdf_to(printer, pdf)` (for a future silent Canon path).
- **Two paths, two printers, two settings** — Zebra (server-side ZPL) + Canon (browser print dialog).

### 4e. Real LP mint + SQL + email — the NEXT phase (logic captured from the macro)
When wiring these, route every write through `sql_write()`:
- **LP mint:** (1) reuse — `SELECT TOP 1 [License Plate] FROM [dbo].[Nhaplecuoingay_All] a WHERE
  PII_PO=? AND PO<>? AND NOT EXISTS (SELECT 1 FROM [dbo].[FG_Database_All] b WHERE a.[License Plate]=
  b.[License Plate]) AND [License Plate] NOT IN (<UsedLP>) GROUP BY [License Plate]`. (2) else mint —
  serial `"200001"` if none, else `format(200000 + MAX(RIGHT([License Plate],5)) + 1, "000000")` where
  the max query filters `SUBSTRING(lp, len-11, 6)=yymmdd AND SUBSTRING(lp, len-5, 1)='2'`;
  `LP = Left(vendor,6) & yymmdd & serial`. Track `UsedLP` within one run. **Wrap the mint in a SQL
  transaction** to kill the documented cross-PC serial race.
- **`[dbo].[License Plate]` INSERT** columns (the `sql_row` in `boxlabel.py` is already in this order):
  `[Print time],[Work station],[PO Number],[Part number],[Box number],[Qty],[Production type],
  [License Plate],[User_ID],[Device_ID],[PII_PO],STACK_LIMIT,PALLET_GW,TS`. `User_ID/Device_ID` =
  Windows username/machine.
- **BPT audit:** INSERT `[dbo].[CB_Production_BOM_Print] (PO_No, FG, FG_Qty, Packing_Material,
  Qty_Packing_Material, Date_Print, File_type, Box_No)` per material row. Canon routing: `WC` present →
  print to BOTH `\\10.191.56.14\vtn1prt-outstock-canon01` AND `\\...\vtn1prt-tanglung`, else tanglung only.
- **New-part email:** `SELECT [Part no] FROM FG_Database_All WHERE [Part no]=?` → if empty (brand-new
  FG), send via **Outlook COM** (init `pythoncom.CoInitialize` per thread; email must never block
  printing) to the fixed recipient list (stored in `print_settings.email_recipients`, default in
  `app.py _EMAIL_DEFAULT`). Subject/body template lived in `Sheet3!A3/A4`.
- **Component BOM ticket (`ModeBOMPick`)** — not built; Canon02 `\\10.191.56.14\VTN1PRT-OUTSTOCK-CANON02`,
  INSERT `CB_PO_BOM_Ticket_Database`, uses the latest `MRP-PurchaseOrders*` file for paper numbers.
- **Small-box (`-ThungNho`)** — not built; incl. the hardcoded special-case part `1024884-329` (confirm
  with the business whether that hardcode is still needed before generalizing).

---

## 5. What was completed THIS session

**Repo bootstrap:** empty repo (local + remote) → seeded baseline; confirmed the 403 push block persists;
adopted the "commit locally, deliver files" workflow.

**Control-board bug fixes**
- **One-operator-one-station guard** (`api_workers_assign`): staffing a worker to a station now strips
  them from every other station server-side — closes double-booking (incl. the pre-assign / pool-drag
  path). (`2026.07.30.1`)
- **14:10-note → Ca 2 shift fix** (§2e float boundary). (`.2`)
- **Priority: red = due within 1 day** (`days_left ≤ 1`, was ≤ 3; urgency no longer promotes to High).
  Changed in BOTH `priority_bucket` (server) and `pc()` (client). (`.2`)
- **Worker-drag optimistic-UI races**: idempotent pool ✓ (no stacked ticks), optimistic source-tag
  removal (no lingering name), a monotonic `_wdropSeq` so the newest drop wins (no stale overwrite),
  and dropped the redundant 2nd `workers/assign` call. (`.3`)

**Control-board features**
- **Planned-capacity setup window** (report chart): dedicated modal — pick From/To → per-day table
  (operators + OT tier none/+2h/+3h), weekday/weekend quick-fill, **Copy ↓** (paste a day forward to the
  end of the range), Clear range, live per-day + total hours. Persists per area in localStorage
  `report_planned_days_{AREA}` = `{date:{ops,ot}}`. The report chart's **Actual | Planned** toggle draws
  the planned line (operators × Ca net-hours 6.73/8.47/9.08). Window opens ABOVE the report (z-index 600),
  bigger, Vietnamese-translated; "Build days" relabeled "Refresh". (`.5`/`.6`)
- **Same-day at-risk alert**: an order **due today** (`days_left==0`), **unfinished** (`remaining>0`),
  **past 14:00** → red pulsing banner + the holding station's stripe pulses red. On `board.html`,
  `dashboard.html` (big screen), and `combined_dashboard.py` (`/api/status` gate + banner list). The
  big-screen banner covers **all** at-risk orders (in-progress, queued, unassigned). (`.4`)
- **Standalone / shared viewers**: `stations_only.html` (big-visual single-area map, area selector);
  `shared_aheadbehind.html` (Chart.js **inlined**, reads `aheadbehind_status.js`) + `shared_rotator.html`
  (fades the two shared pages, **centered** selector); `combined_dashboard.py` now publishes
  `aheadbehind_status.js` (JSONP, atomic write, 60s) so the shared viewer can show Ahead/Behind with no
  server/CDN. Removed the leftover **green dotted est-pace line** from the Ahead/Behind chart.

**Printing subsystem (§4)** — BPT packaging ticket (`bpt.py`, A5) + Zebra box label (`boxlabel.py` +
`label_print.py`, DUMMY LP) + `printing.py` + full `app.py` integration (print button, Issue modal,
Print-settings modal, `/api/issue_ticket`, print-and-queue-on-station). Verified end-to-end via the
Flask test client (dummy LP, PDO placed, BPT page auto-prints, settings/printers endpoints). (`.7`)

---

## 6. SAFETY GUARDS (every guard that protects operator-entered state / production data)

1. **Snapshot is READ-ONLY.** `build_board_snapshot()` never prunes or saves. A stale id renders as an
   empty bin; the next 60s poll cleans it authoritatively. (Fixed the old "orders bounce to queue" bug.)
2. **Prune wipe-guards.** `prune_stale_assignments` no-ops unless `_last_pdo_status == 'ok'`;
   `prune_expired_workers` no-ops unless `_last_attendance_ok`. A failed/unsynced Excel read at startup
   can look identical to "empty" — the guard skips and retries rather than wiping the board.
3. **Pruning ONLY in the 60s poll**, on a freshly-rebuilt cache — never on a board GET or dashboard write.
4. **Daily worker-staffing reset** (`reset_daily_worker_staffing`): clears carried-over **operator
   staffing only** (never PDO assignments) once per calendar day (persisted in `worker_reset_{AREA}.json`)
   — closes the shift-boundary gap without touching orders.
5. **One-operator-one-station guard** (server-side, `api_workers_assign`): the authoritative anti-double-
   booking, independent of how the client dragged.
6. **`remaining_qty` single source of truth** — mutated ONLY via `rebuild_final_pdo_cache` /
   `inject_splits_into_cache`. Splits reconcile against the PO-level FG total (authoritative "is it done").
7. **Worker availability = one check** (`_worker_available_now`) used by the pool, each station's display,
   and the prune, so they never disagree.
8. **LP auto-queue keyed on `remaining_qty`, not FG/WIP presence** (FG/WIP log independently of the
   board). A stale License-Plate SQL query returns `None` → **keep the cache, warn loudly** (never wipe).
9. **Dev mode gates ALL SQL writes** (`sql_write`) + the big-screen/shared feed — one flag, surfaced
   loudly (startup line, `/api/board.dev_mode`, board banner). Every board-originated INSERT/UPDATE/DELETE
   must go through `sql_write()`.
10. **Data persistence across code updates.** Operator state lives in `data/*.json` (git-ignored) —
    the code update does NOT touch it. **Safe-update recipe: replace only `app.py` + `templates/board.html`
    (and the dashboard files on the big-screen PC); NEVER delete/overwrite `data/`, `startup_config.json`,
    `worker_reset_{AREA}.json`, `print_settings_{AREA}.json`.** (Operators — not PDOs — auto-clear once at
    the start of a new production day by design, guard #4; that is expected, not the update.)
11. **Worker-drag: newest drop wins** (`_wdropSeq`) — a slow earlier `/api/board` response can't clobber
    the board after a newer drag.
12. **Printing safety (current phase):** the **License Plate is a DUMMY** and **no SQL INSERT / email is
    performed** — so a mis-print cannot write to `[License Plate]`, `CB_Production_BOM_Print`, or send
    email. When the real mint lands, wrap it in a **SQL transaction** (cross-PC serial race) and route the
    INSERT through `sql_write()` so dev mode still protects production. Email must never block printing.
13. **MainDatabase opened read-only + cached** (10-min TTL); a bad path fails the print call cleanly
    (`"Cannot open MainDatabase"`) without crashing the board. `printing.py` win32 imports are lazy, so a
    non-Windows / no-pywin32 host degrades to a "not printed" warning instead of an error.
14. **Hardcoded secrets to remove in the rewrite** (do NOT propagate): the plan-sheet password `"12345"`,
    the SQL creds `svcsqllocal / KetnoilocalApp` (already in `app.py`/`combined_dashboard.py` config — move
    to env/secrets), and the six D365 plaintext passwords in the separate `Download_PO_received.py` macro.

---

## 7. Key decisions & rationale
- **AREA from the plan filename** — one codebase, three areas; never hardcode.
- **remaining_qty single source of truth**; **snapshot read-only; pruning only in the 60s poll**.
- **Prune must never wipe on a failed load** (wipe-guard) — destroying operator state is far worse than
  skipping a cycle.
- **Worker staffing is per-shift, reset once per day**; **availability filtered live** via one check.
- **LP auto-queue keyed on remaining_qty**; **splits reconcile against the PO-level FG total**.
- **Dev mode gates the feed + all SQL writes** — one loud flag.
- **Priority red = due within 1 day** — reserves red for genuinely imminent orders.
- **Planned capacity = per-day editor** (operators × Ca net-hour tiers), separate window, per-area; the
  report chart shows the resulting line (Actual/Planned toggle). Per-day, not cumulative (by request).
- **At-risk alert = banner + station-stripe pulse**, gated past 14:00; big-screen banner covers all
  at-risk orders.
- **Shared viewers read `.js` via `<script>` + hidden-iframe refresh** — the only CORS-free way to read
  local data from a `file://` page; **Chart.js inlined** and fonts system so the Ahead/Behind shared page
  needs no CDN. **`combined_dashboard.py` publishes the ahead/behind snapshot** so the shared viewer works
  without a server reachable through the firewall.
- **Printing = two paths, two printers.** Zebra box label = **raw ZPL** (crisp thermal + native barcodes,
  rotation handled once). BPT = the **approved A5 HTML**, printed from the browser dialog (preserves the
  layout, no chromium/PDF dependency on the board PC). **The print action queues the PDO on the station**
  (replaces LP-scan as the placement signal, per the user).
- **Packaging type via string prefix** (`PL`/`SPL`/`WC`) — the user explicitly chose to keep it simple
  (no `packaging_type` enum table).
- **DUMMY LP first** — get the whole flow testable on real printers before touching SQL.

---

## 8. Config reference (verify per deployment)
### `app.py` (top of file)
- `BUILD_VERSION`/`BUILD_NOTE` — bump on every change (glance-checkable per PC). Currently `2026.07.30.7`.
- `DEV_MODE` — from `BOARD_DEV_MODE`; setup-window checkbox overrides.
- `EXCEL_PLAN` / attendance paths — overridden by the setup window / `startup_config.json`.
- `PLAN_SHEET='1.FollowPlanForGL'`, `DATA_START_ROW=6`.
- SQL: `SQL_SERVER='VTN1PRDSQL002'`, `UID='svcsqllocal'`, `PWD='KetnoilocalApp'`, `SQL_DB_TECH='TECH_DATA'`,
  `SQL_DB_PROD='LOCALNPV'`. `svcsqllocal` already has INSERT on `[License Plate]` (macro proves it).
- `DASHBOARD_STATUS_DIR` (shared folder), `CAPACITY_FILES={'AD','PK','RP':…}`.
- `PRODUCTION_START=6.0`, `CA1_END=14.167`, `PRODUCTION_END=22.333`, `NEXT_SHIFT_OPEN_EARLY_HRS=15/60`.
- Printing: `_MAINDB_DEFAULT` (MainDatabase share path), `_EMAIL_DEFAULT` (new-part recipients),
  `_print_settings_defaults()` (zebra/canon/maindb/email) → `data/print_settings_{AREA}.json`.
### `bpt.py` / `boxlabel.py` / `label_print.py`
- Buckets `{PVN, Poland, Other}`; `CAP_TIER` net-hour tiers mirror the server model.
- `label_print`: canvas 2007×1298 dots (6.689"×4.327" @300dpi), `ROTATE=90` (flip to 270 if upside down),
  default printer `ZDesigner ZT421-300dpi ZPL`.
- Deps: `openpyxl · pillow · python-barcode · pywin32(win)` (`requirements-print.txt`).

---

## 9. In progress / planned (NOT started unless noted)
1. **Printing — real LP + SQL + email** (§4e): reuse-then-serial mint in a **SQL transaction**, the
   `[License Plate]` INSERT through `sql_write()`, and the new-part **Outlook-COM email**. `sql_row` is
   already assembled per label.
2. **Printing — remaining macro modes**: small-box `-ThungNho` (+ the `1024884-329` special case),
   component BOM ticket (`ModeBOMPick` → Canon02, `CB_PO_BOM_Ticket_Database`), silent Canon printing
   for the BPT (currently browser dialog).
3. **Printing — layout confirmations**: FG revision source (G/H by FG is blank — where does it come
   from?), whether Ver/Location should populate for packaging materials, ZPL rotation (90 vs 270) on the
   real ZT421.
4. **Dashboard-feed cleanup**: sweep stale `.tmp` files in `dashboard_status`; trim the ~1 MB snapshot
   (drop the unassigned-PDO pool the station views don't use).
5. **D365 goods receipt** — batch-file export (smallest ask) or hardened in-app Selenium with ONE service
   account; kill the six plaintext D365 passwords regardless.
6. **Report/Actual-Planned toggle + report window** still have a few English-only strings — optional VN sweep.

---

## 10. Known gaps / to verify on a real deployment
1. **SQL schemas** for `Nhaplecuoingay_All` / `License Plate` / `CB_Production_BOM_Print` were given
   verbally — watch the console for SQL errors when the real writes land.
2. **Ahead/Behind SQL half** (in `combined_dashboard.py`) uses ODBC Driver 17.
3. **Attendance notes / shift inference** validated against synthetic + the 14:10 fix — validate on real files.
4. **Split records never self-clean** (`splits_{AREA}.json`) — harmless inert data after completion.
5. **Printing**: win32 ZPL print + browser BPT print are **not exercisable off-Windows** — test on the
   board PC. Confirm barcodes scan, ZPL orientation, and that the board PC can reach `MainDatabase.xlsx`
   + both printers. The MainDatabase market-bucket rule and `CheckDes` (AV→AW) were validated on the file.
6. **`combined_dashboard.py` served pages use CDN** (Chart.js + fonts) for the Ahead/Behind half; the
   **shared** pages are fully self-contained (no CDN).

---

## 11. How to run
- **Control board (each area PC):** `python app.py` → setup window (verify plan + attendance paths, must
  validate green; optional 🧪 Dev/test checkbox) → in-app console → board `http://127.0.0.1:5050/`. Plan
  filename MUST start with `Packaging` / `AssyD` / `RawPart`.
- **Printing:** `pip install -r requirements-print.txt` on the board PC → top bar **🖨 Print** to set the
  Zebra/Canon queues + MainDatabase path → hover a PDO card → **🖨** → station + prod time → **Issue &
  queue**. See `PRINTING_README.md`.
- **Big-screen:** `python combined_dashboard.py` → `http://<pc-ip>:8080/` (also starts the
  `aheadbehind_status.js` writer).
- **Shared viewers:** open `\\npvshare\...\dashboard_status\shared_rotator.html` (or the individual pages).

## 12. For the next chat
Paste this document. **First: check whether `git push` access has been granted** (403 all session — all
work is committed locally on `claude/production-control-board-bk0blz`, delivered as files; the user has
chosen to keep working locally + hand over files). Restate the three load-bearing concepts (§0) and the
**safety guards (§6)** before touching anything. When touching workers/remaining_qty/splits, follow §2 +
§6. Route any new SQL write through `sql_write()`. **Printing resumes at the real LP mint + `[License
Plate]` INSERT + new-part email (§4e / §9.1)** — the dummy LP and the assembled `sql_row` are the
seams; wrap the mint in a SQL transaction, and keep email non-blocking. Test all printing on the actual
board PC (win32 + real printers) — it can't be exercised in the Linux sandbox.
