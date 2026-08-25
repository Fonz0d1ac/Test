"""
Box label (ModeQty) path — the Zebra ZT421 production/License-Plate label.

Second of the two print paths (the first is bpt.py, the packaging pick ticket).
This computes, per box: box qty, stack limit, pallet gross weight, and a License
Plate, then hands a data dict to label_print.py to render the rotated ZPL.

Formulas are ported verbatim from the Excel macro (`cmdPrint_Click`, ModeQty):
  QtyLblPrint  = RoundUp(POQty / QtyPerBoxSTD)
  PalletQty    = RoundUp(POQty / (QtyPerBoxSTD * BoxQtyPerPalletSTD))
  QtyOfPallet  = POQty Mod (QtyPerBoxSTD * BoxQtyPerPalletSTD)
  Boxqty       = RoundUp(QtyOfPallet / QtyPerBoxSTD)
  box qty      = last box gets POQty Mod QtyPerBoxSTD, else QtyPerBoxSTD
  stack (B9) / pallet GW (B12): partial-vs-full-last-pallet formulas, gated by
    ModeStandard (unchecked = stack 0 + minimal weight estimate).
Weights come from MainDatabase (R/S/T/U/V/W) keyed by Partno&PartDes.

License Plates come from lp.py — the caller does the reuse-then-mint against SQL
and passes the finished list in as `plates=`. This engine stays pure: it computes
box splits and weights, and assembles the [dbo].[License Plate] audit row
(`sql_row`) that app.py inserts BEFORE printing. With no `plates=` it falls back
to lp.dummy_plates() for the area, so an offline/dev run still produces realistic
labels without touching SQL. Revision is intentionally blank for now.
"""

import math
import datetime
import bpt
import label_print
import lp


def _ru(x):
    """VBA Application.RoundUp(x, 0): round away from zero to an integer."""
    if x is None:
        return 0
    x = float(x)
    if x >= 0:
        return math.ceil(x - 1e-9)
    return -math.ceil(-x - 1e-9)


def _num(v, d=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return d


def make_dummy_lp(vendor6, date_text, serial_int, area="PK"):
    """DUMMY License Plate for testing (no SQL). Mirrors the real 18-char shape
    VendorCode(6) + yymmdd(6) + serial(6) so downstream/scanners see a realistic
    value.

    The area matters even here. This used to hardcode 200000 — Packaging's base —
    so an AD or RP test print produced a plate whose first serial digit claimed it
    came from Packaging. The bases live in lp.AREA_SERIAL and nowhere else."""
    return lp.format_lp(vendor6, date_text, lp.area_serial_spec(area)[1] + int(serial_int))


def box_count(std, po_qty):
    """How many boxes this PDO prints — QtyLblPrint = RoundUp(POQty / QtyPerBoxSTD).

    Exists so the caller can size the License Plate mint BEFORE any label is
    built, without re-deriving the formula. `std` is the dict that
    bpt.resolve_pack_std() returned, i.e. the SAME resolved standard the BPT
    used, which is the whole point of safety guard #16: if the two paths ever
    computed the box count differently, a PDO would print N tickets and a
    different number of labels. Returns 0 when the standard did not resolve —
    the caller reports std['error'], mints nothing, and prints nothing.
    """
    if not std or not std.get('ok'):
        return 0
    qpb = int(std.get('qty_per_box') or 0)
    if qpb <= 0:
        return 0
    return _ru(int(_num(po_qty)) / qpb)


def dummy_plates_for(area, vendor6, n, today=None):
    """n offline plates in `area`'s serial range. See lp.dummy_plates()."""
    return lp.dummy_plates(area, vendor6, n, today)


def build_box_labels(mdb, *, po, part, po_qty, dest, station, pii="", desc="",
                     vendor_code="", production_type="", ts="", mode_standard=True,
                     user="", device="", today=None, area="PK", plates=None,
                     sql_station=None):
    """Build one label per box for the ModeQty (production) path. Returns
    {ok, labels:[{box_no, box_qty, lp, d, sql_row}], ...}. `d` is ready for
    label_print.build_zpl / render_mock. `dest` is the plan-F destination NAME.

    `plates` is the list of real License Plates from lp.acquire_plates(), in box
    order. It is optional ONLY so an offline/dev run still works: without it the
    labels carry lp.dummy_plates() values for `area` and must never be inserted.

    The box count is computed here, so the caller has a chicken-and-egg problem —
    it cannot mint the right number of plates until it knows how many boxes there
    are. box_count() below answers that without building anything, which is how
    app.py orders the two steps: count → mint → build → insert → print.

    `sql_station` is the macro's station form ('PK01') for the audit row; the
    label itself keeps the board form ('PK-01') the operator reads.
    """
    partdes = bpt.market_bucket(dest)                     # PVN / Poland / Other
    # Same resolver the BPT uses — the two paths MUST agree on qty/box and
    # box/pallet, or a PDO prints N tickets and a different number of labels.
    std = bpt.resolve_pack_std(mdb, part, partdes)
    wt = std['weights']
    qpb = std['qty_per_box']                              # QtyPerBoxSTD
    R, S, T, U, V = (_num(wt.get(k)) for k in ('R', 'S', 'T', 'U', 'V'))
    W = std['box_per_pallet']                             # BoxQtyPerPalletSTD
    dest_code = mdb.des_by_raw.get(dest, "")              # CheckDes (AV->AW)
    ts = ts or mdb.ts_by_fg.get(part, "")
    wi = mdb.wi_by_fg.get(part, "")
    po_qty = int(_num(po_qty))

    if not std['ok']:
        return {'ok': False, 'error': std['error'] or 'No packing (qty/box) data for this part/market',
                'std': std, 'labels': []}

    date_text = (today or datetime.date.today()).strftime("%y%m%d")
    vendor6 = str(vendor_code or "")[:6]

    total = _ru(po_qty / qpb)                             # QtyLblPrint
    group = qpb * W                                       # QtyPerBoxSTD * BoxQtyPerPalletSTD
    qty_of_pallet = (po_qty % group) if group else 0
    boxqty = _ru(qty_of_pallet / qpb) if qpb else 0
    partial = bool(group) and (po_qty % group != 0)      # last pallet is partial

    # Stack limit (B9) / pallet gross weight (B12) — constant across the run
    # (the macro's inner j-loop only leaves the last pallet's values). ModeStandard
    # off = don't print an authoritative stack number (0) + minimal weight estimate.
    if partial:
        if mode_standard:
            if qty_of_pallet > qpb:
                b9 = _ru(S - 1.1 * ((boxqty - 1) * T + (qty_of_pallet - qpb) * R))
            else:
                b9 = _ru(S)
            b12 = _ru(qty_of_pallet * R + U + boxqty * T)
        else:
            b9, b12 = 0, _ru(qty_of_pallet * R + 2)
    else:
        if mode_standard:
            b9 = _ru(S - V * (W - 1))
            b12 = _ru(group * R + U + W * T)
        else:
            b9, b12 = 0, _ru(group * R + 2)

    now = datetime.datetime.now()
    # Real plates from lp.acquire_plates(), or dummies when the caller had no SQL
    # (dev mode / offline / live minting not switched on for this area yet). A
    # short list is a caller bug, not something to paper over with a dummy tail:
    # a dummy plate mixed into a live run would print an uninserted plate.
    if plates is None:
        plates = dummy_plates_for(area, vendor6, total, today)
        real_plates = False
    else:
        if len(plates) < total:
            return {'ok': False, 'std': std, 'labels': [],
                    'error': f'Only {len(plates)} License Plate(s) available for '
                             f'{total} box(es) — nothing printed'}
        real_plates = True
    station_sql = sql_station or str(station)
    labels = []
    for i in range(1, total + 1):
        box_qty = (po_qty % qpb) if (i == total and po_qty % qpb != 0) else qpb
        lp = plates[i - 1]
        d = dict(station=str(station), pii=str(pii), po=str(po), part=str(part),
                 desc=str(desc), qty=str(box_qty), stack=str(b9), gw=str(b12),
                 wi=str(wi), rev="", market=str(dest_code), lp=lp, pdate="", worker="")
        # Audit row for [dbo].[License Plate]. Column order matches the macro's
        # INSERT exactly, and these keys ARE the column names — lp.insert_plate_rows
        # builds its statement from them, so the schema lives only here.
        # 'Work station' takes the macro's dash-less form ('PK01'): the same
        # column is read back by app.get_license_plate_prints() for auto-queue,
        # and by whatever else on the floor reads this table.
        sql_row = {
            'Print time': now.strftime("%Y-%m-%d %H:%M:%S"), 'Work station': station_sql,
            'PO Number': po, 'Part number': part, 'Box number': i, 'Qty': box_qty,
            'Production type': production_type, 'License Plate': lp, 'User_ID': user,
            'Device_ID': device, 'PII_PO': pii, 'STACK_LIMIT': b9, 'PALLET_GW': b12,
            'TS': ts,
        }
        labels.append({'box_no': i, 'box_qty': box_qty, 'lp': lp, 'd': d, 'sql_row': sql_row})

    return {'ok': True, 'qpb': qpb, 'bpp': W, 'total': total, 'stack': b9, 'gw': b12,
            'dest_code': dest_code, 'wi': wi, 'ts': ts, 'partdes': partdes,
            'std': std, 'labels': labels, 'real_plates': real_plates}


def render_zpl(label_d, rotate=None):
    """Rotated, print-ready ZPL for one label dict (via label_print)."""
    return label_print.build_zpl(label_d, rotate=rotate)


def render_preview_png(label_d, path="boxlabel_preview.png"):
    """Rasterized PNG preview of one label (no printer needed)."""
    return label_print.render_mock(label_d, path)
