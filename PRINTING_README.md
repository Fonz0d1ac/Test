# Ticket / Label printing — how to run & test

Replaces the Excel `PrintLabel` macro. One **🖨** click on a PDO issues up to two
things and **queues the PDO on the chosen station** (this replaces the LP-scan
placement signal):

1. **BPT packaging pick ticket** → Canon warehouse printer. Rendered as the A5
   HTML ticket and opened in a **new browser tab that auto-opens the print
   dialog** — the operator just confirms the Canon printer.
2. **Box label + License Plate** → Zebra ZT421, as **raw ZPL** sent server-side
   via `pywin32`.

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
- **Canon printer** — the BPT prints from the browser dialog, so this is just for
  the record for now; set your default browser printer to the Canon to make it
  one-click.
- **MainDatabase.xlsx path** — defaults to the `\\npvshare\…\MainDatabase.xlsx`
  share; point it at a local copy for testing if needed.
- **Email recipients** — stored for the new-part email (not sent yet).

Settings are saved per area in `data/print_settings_<AREA>.json`.

## Test it
1. Start the board (`python app.py`), open `http://127.0.0.1:5050/`.
2. Top bar → **🖨 Print** → set the Zebra queue + MainDatabase path → Save.
3. Hover a PDO card → click **🖨** → enter a **Station** (+ optional prod time) →
   **Issue & queue**.
   - The **box label prints on the Zebra** (check the warnings toast if not).
   - A **new tab opens with the BPT ticket** and the print dialog — pick the Canon.
   - The **PDO moves onto that station** on the board.
4. The toast reports how many labels printed, the dummy LP(s), and any warnings
   (e.g. "no packaging BOM for this FG/market", "Zebra print failed …").

## Not done yet (next phase)
- Real **License Plate mint** (reuse `Nhaplecuoingay_All`, else `max-serial+1` in a
  SQL transaction) + the `[dbo].[License Plate]` INSERT (the audit row is already
  assembled in `boxlabel.py` as `sql_row`, ready to route through `sql_write()`).
- **New-part email** (recipients are stored; sending is TODO).
- Silent Canon printing (currently via the browser dialog).
- Small-box (`-ThungNho`) + component BOM ticket paths.
