# Ticket / Label printing — how to run & test

Replaces the Excel `PrintLabel` macro. One **🖨** click on a PDO issues up to two
things and **queues the PDO on the chosen station** (this replaces the LP-scan
placement signal):

1. **BPT packaging pick ticket** → Canon warehouse printer, **one A5 page per box**.
   Printed **silently** when the silent path is available (see below); otherwise it
   falls back to a browser tab with the print dialog.
2. **Box label + License Plate** → Zebra ZT421, as **raw ZPL** sent server-side
   via `pywin32`. Always silent — one label per box.

Every issue ends in a **blocking result dialog** listing the packing standard it
used, the BPT page count, and each box label with its License Plate and whether
it actually reached the printer. The board does not refresh until you click **OK**.

## How a PDO splits into boxes

Both paths derive everything from the **packing standard in MainDatabase** for that
part **and that market** (`PVN` / `Poland` / `Other`, from plan column F):

```
boxes   = ceil(PDO qty / qty per box)      last box carries the remainder
pallets = ceil(boxes / box per pallet)
```

Each box gets its own BPT page and its own Zebra label. Materials whose code starts
with a **pallet prefix** (`PL`, `SPL` by default, editable in print settings) are
picked **once per pallet** — qty 1 on the box that opens each pallet group, 0 on the
rest, shown greyed so the picker can see it was considered.

> 50 pcs/box, 2 box/pallet, PDO 100 → **2 tickets**: #1 box materials **+ 1 pallet**,
> #2 box materials only. Two Zebra labels, 50 pcs each.

The market lookup is **strict**: if MainDatabase has no Qty/Box row for that part in
that market, the issue is refused and the dialog names the markets that *do* exist.
It will not quietly borrow another market's qty/box — that silently collapses the
split (a 100/box row used for a 50/box order turns a 100 pc PDO into one box).
The only fallback is unambiguous: a part with exactly one packing row in the whole
file uses it, and the dialog says so.

## Silent BPT printing (no pop-up)

The Zebra needs nothing extra — raw ZPL goes straight to the queue through `pywin32`.
The **paper** side needs two steps, because Windows can render a PDF but cannot send
one to a *named* printer without UI:

1. **Headless Edge or Chrome** renders the A5 HTML to PDF — already on every Windows PC.
2. **SumatraPDF** pushes that PDF to the Canon queue (`-print-to "<queue>" -silent`).

SumatraPDF is a single portable `.exe` — no installer, no admin rights. Drop it in a
`tools\` folder next to `app.py` (or anywhere, and set the path in print settings).
Get it from https://www.sumatrapdfreader.org → "Portable version".

If either piece is missing, printing still works: the BPT opens in a tab as before and
the result dialog says why. **🖨 Print** shows a live ✓/⚠ readiness line.

> **License Plate is a DUMMY right now** (`V-<vendor6><yymmdd><serial-from-200001>`).
> No `[License Plate]` INSERT and no email are sent yet — those are the next
> (SQL) phase. Revision on the box label is intentionally blank.

## Install (on the Windows board PC)
```
pip install -r requirements-print.txt
```
(`pywin32` is what lets the app enumerate printers and send ZPL to the Zebra.)

## Files
| File | Role |
|---|---|
| `app.py` | board + the new endpoints (`/api/issue_ticket`, `/api/print_settings`, `/api/printers`, `/bpt/<token>`) |
| `templates/board.html` | 🖨 button on each PDO card + Issue modal + **🖨 Print** settings modal (top bar) |
| `bpt.py` | BPT ticket engine + A5 HTML renderer (reads MainDatabase) |
| `boxlabel.py` | box-label engine (per-box qty, stack/GW, **dummy LP**) |
| `label_print.py` | Zebra ZPL renderer (raster → rotated ZPL) |
| `printing.py` | win32 helpers (list printers, send raw ZPL) |

## Configure (top bar → 🖨 Print)
- **Zebra printer** — pick the ZT421 queue (default `ZDesigner ZT421-300dpi ZPL`).
- **Canon printer** — the queue the BPT is sent to on the silent path.
- **Print the BPT silently** — on by default; the line under it says whether this PC
  can actually do it and what's missing if not.
- **SumatraPDF.exe path** — blank auto-detects (`tools\` next to `app.py`, `%LOCALAPPDATA%`,
  Program Files, PATH).
- **Pallet material prefixes** — codes picked once per pallet instead of per box.
  Default `PL, SPL`. **If your pallet/slip-sheet codes start with something else, set it
  here** — otherwise they'll be charged to every box.
- **MainDatabase.xlsx path** — defaults to the `\\npvshare\…\MainDatabase.xlsx`
  share; point it at a local copy for testing if needed.
- **Email recipients** — stored for the new-part email (not sent yet).

Settings are saved per area in `data/print_settings_<AREA>.json`.

## Test it
1. Start the board (`python app.py`), open `http://127.0.0.1:5050/`.
2. Top bar → **🖨 Print** → set the Zebra queue + MainDatabase path → Save.
3. Hover a PDO card → click **🖨** → enter a **Station** (+ optional prod time) →
   **Issue & queue**.
   - One **Zebra label per box** prints.
   - The **BPT prints to the Canon**, one page per box (or opens in a tab if the
     silent path isn't set up).
   - The **PDO moves onto that station** on the board.
4. The **result dialog** appears and stays until you click **OK**. Check:
   - **Packing standard** — market, qty/box, box/pallet, and the resulting box and
     pallet counts. If the box count looks wrong, this line is where to look first.
   - **BPT** — page count and where it went.
   - **Per-box table** — box no, qty, License Plate, and ✓/✕ per label.
   - **Warnings** — anything that didn't go to plan.

## If a print feels slow

The result dialog and the console both report a per-stage breakdown:

```
[issue timing] maindb 0.0s · zebra 3.1s · zebra_render 0.4s · bpt 4.2s · total 7.8s · 11 label(s)
```

- **maindb** — should be ~0. A background thread keeps MainDatabase parsed and warm,
  so a cold read off `\\npvshare` never lands on an operator's click. A non-zero value
  here means the warm loop is failing; check the console for `[MainDatabase]` lines.
- **zebra** — generating **and spooling** all labels. `zebra_render` is the generation
  part alone; the difference is time on the wire to the printer.
- **bpt** — headless render + the Canon spool. First print after a reboot is slower
  (browser cold start).

## Not done yet (next phase)
- Real **License Plate mint** (reuse `Nhaplecuoingay_All`, else `max-serial+1` in a
  SQL transaction) + the `[dbo].[License Plate]` INSERT (the audit row is already
  assembled in `boxlabel.py` as `sql_row`, ready to route through `sql_write()`).
- **New-part email** (recipients are stored; sending is TODO).
- **Wooden-crate dual routing** — a `WC` material should print the BPT to BOTH
  Canon01 and tanglung. The silent path makes this possible (the browser dialog
  never could); the second queue isn't wired up yet.
- Small-box (`-ThungNho`) + component BOM ticket paths.
