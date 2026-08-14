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

** LP is a DUMMY for testing right now ** — see make_dummy_lp(). The real mint
(reuse an unused plate from Nhaplecuoingay_All, else VendorCode+yymmdd+serial with
max-serial+1 wrapped in a SQL transaction) and the [License Plate] INSERT drop in
at the SQL phase; the audit row is already assembled below (`sql_row`) so wiring
it to sql_write() is a one-liner then. Revision is intentionally blank for now.
"""

import math
import datetime
import bpt
import label_print


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


def make_dummy_lp(vendor6, date_text, serial_int):
    """DUMMY License Plate for testing (no SQL). Mirrors the real 18-char shape
    VendorCode(6) + yymmdd(6) + serial(6, from 200001) so downstream/scanners see
    a realistic value. Replace with the real reuse-then-mint in the SQL phase."""
    return f"{vendor6}{date_text}{200000 + int(serial_int):06d}"


def build_box_labels(mdb, *, po, part, po_qty, dest, station, pii="", desc="",
                     vendor_code="", production_type="", ts="", mode_standard=True,
                     user="", device="", today=None):
    """Build one label per box for the ModeQty (production) path. Returns
    {ok, labels:[{box_no, box_qty, lp, d, sql_row}], ...}. `d` is ready for
    label_print.build_zpl / render_mock. `dest` is the plan-F destination NAME."""
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
    labels = []
    for i in range(1, total + 1):
        box_qty = (po_qty % qpb) if (i == total and po_qty % qpb != 0) else qpb
        lp = make_dummy_lp(vendor6, date_text, i)         # DUMMY for now
        d = dict(station=str(station), pii=str(pii), po=str(po), part=str(part),
                 desc=str(desc), qty=str(box_qty), stack=str(b9), gw=str(b12),
                 wi=str(wi), rev="", market=str(dest_code), lp=lp, pdate="", worker="")
        # Audit row for [dbo].[License Plate] (assembled now; INSERT via sql_write
        # in the SQL phase). Column order matches the macro's INSERT exactly.
        sql_row = {
            'Print time': now.strftime("%Y-%m-%d %H:%M:%S"), 'Work station': station,
            'PO Number': po, 'Part number': part, 'Box number': i, 'Qty': box_qty,
            'Production type': production_type, 'License Plate': lp, 'User_ID': user,
            'Device_ID': device, 'PII_PO': pii, 'STACK_LIMIT': b9, 'PALLET_GW': b12,
            'TS': ts,
        }
        labels.append({'box_no': i, 'box_qty': box_qty, 'lp': lp, 'd': d, 'sql_row': sql_row})

    return {'ok': True, 'qpb': qpb, 'bpp': W, 'total': total, 'stack': b9, 'gw': b12,
            'dest_code': dest_code, 'wi': wi, 'ts': ts, 'partdes': partdes,
            'std': std, 'labels': labels}


def render_zpl(label_d, rotate=None):
    """Rotated, print-ready ZPL for one label dict (via label_print)."""
    return label_print.build_zpl(label_d, rotate=rotate)


def render_preview_png(label_d, path="boxlabel_preview.png"):
    """Rasterized PNG preview of one label (no printer needed)."""
    return label_print.render_mock(label_d, path)
