"""
BPT — Packaging BOM Pick Ticket generator (replaces the Excel PrintLabel macro's
ModePackPick, §9 of the spec).

Pure logic + HTML rendering; no Flask, no SQL, no printing here (app.py wires
those). Reads MainDatabase.xlsx (the same master parts/packaging table the macro
opens) and produces a print-ready ticket per box for one FG + market.

Column map (validated against the live MainDatabase.xlsx — headers on row 2,
data from row 3; the sheet is several stacked sub-tables):
  Packaging BOM block : AF=Finished Good · AG=Material code · AH=Description
                        AI=Usage · AJ=Lot FG Qty · AK=Destination(market bucket)
  Qty/Box & Weight    : P=Part(FG) · Q=Qty/box · W=Boxperpallet · X=Destination
  WI location         : M=Part no(FG) · N=Vị trí
  Time study          : B=Main FG · C=TS
  Drw revision        : G=Part · H=Rev            (version lookup by material code)
  WH location         : BL=Part number · BM=Location (location lookup by material code)
  Type (big/small)    : BI=Part number · BJ=Type
Market buckets are exactly {PVN, Poland, Other}.
"""

import math
from io import BytesIO
from openpyxl.utils import column_index_from_string as _ci

# ── barcode ───────────────────────────────────────────────────────────────────
def code128_datauri(data, module_height=7.0, font_size=8, module_width=0.22,
                    write_text=True):
    """Real Code128 barcode as an inline SVG data URI (no font trick, per §13).

    `module_width` is the bar pitch — widening it makes the barcode physically
    LONGER, which is what a hand scanner wants; `module_height` makes it taller.
    `write_text=False` drops the human-readable line underneath."""
    import barcode
    from barcode.writer import SVGWriter
    import base64
    data = str(data or "")
    if not data:
        return ""
    c = barcode.get('code128', data, writer=SVGWriter())
    buf = BytesIO()
    c.write(buf, options={'module_height': module_height, 'font_size': font_size,
                          'module_width': module_width, 'quiet_zone': 1,
                          'text_distance': 3, 'write_text': write_text})
    b64 = base64.b64encode(buf.getvalue()).decode('ascii')
    return "data:image/svg+xml;base64," + b64


# ── market bucket ─────────────────────────────────────────────────────────────
def market_bucket(raw):
    """Plan-row raw market/customer code (col F) → packaging bucket used by AK/X.
    Buckets in the file are exactly PVN / Poland / Other. Anything not clearly
    PVN or Poland (e.g. 'NP', 'SHA') falls to 'Other'. Confirm the full raw→bucket
    map with the business if more markets appear."""
    r = (raw or "").strip().upper()
    if r == "PVN":
        return "PVN"
    if r == "POLAND":
        return "Poland"
    return "Other"


def _s(v):
    return str(v).strip() if v is not None else ""


class MainDB:
    """Loads MainDatabase.xlsx once and builds the lookups the BPT needs. In
    app.py this is cached and refreshed on a TTL; here it's a plain object."""

    def __init__(self, path):
        self.path = path
        self.pack_by_fg = {}      # fg -> [ {code,desc,usage,lotqty,dest} ]
        self.qtybox = {}          # (fg,bucket) -> (qty_per_box, box_per_pallet)
        self.qtybox_any = {}      # fg -> (qty_per_box, box_per_pallet)  (first seen)
        self.wt_by_key = {}       # (fg,bucket) -> {Q,R,S,T,U,V,W} std packing weights
        self.targetop_by_fg = {}  # fg -> TargetOp (col D, keyed by B)
        self.wi_by_fg = {}        # fg -> WI location
        self.ts_by_fg = {}        # fg -> time study
        self.rev_by_part = {}     # material/part code -> revision (G->H)
        self.loc_by_part = {}     # material code -> WH location (BL->BM)
        self.type_by_part = {}    # material code -> type (BI->BJ)
        self.des_by_raw = {}      # destination NAME (AV) -> dest CODE (AW), i.e. CheckDes
        self._load()

    def _load(self):
        import openpyxl
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        ws = wb['MainDatabase']
        iAF, iAG, iAH, iAI, iAJ, iAK = (_ci(c) - 1 for c in ('AF', 'AG', 'AH', 'AI', 'AJ', 'AK'))
        iP, iQ, iW, iX = (_ci(c) - 1 for c in ('P', 'Q', 'W', 'X'))
        iR, iS, iT, iU, iV = (_ci(c) - 1 for c in ('R', 'S', 'T', 'U', 'V'))
        iM, iN = _ci('M') - 1, _ci('N') - 1
        iB, iC, iD = _ci('B') - 1, _ci('C') - 1, _ci('D') - 1
        iG, iH = _ci('G') - 1, _ci('H') - 1
        iBL, iBM = _ci('BL') - 1, _ci('BM') - 1
        iBI, iBJ = _ci('BI') - 1, _ci('BJ') - 1
        iAV, iAW = _ci('AV') - 1, _ci('AW') - 1
        n = ws.max_column

        def g(row, i):
            return row[i] if i < len(row) else None

        for row in ws.iter_rows(min_row=3, values_only=True):
            # packaging BOM
            af = _s(g(row, iAF))
            if af:
                self.pack_by_fg.setdefault(af, []).append({
                    'code': _s(g(row, iAG)), 'desc': _s(g(row, iAH)),
                    'usage': g(row, iAI), 'lotqty': g(row, iAJ), 'dest': _s(g(row, iAK)),
                })
            # qty/box + box/pallet, keyed by FG (+ destination bucket)
            pcode = _s(g(row, iP))
            if pcode:
                qpb, bpp, dest = g(row, iQ), g(row, iW), _s(g(row, iX))
                self.qtybox[(pcode, dest)] = (qpb, bpp)
                self.qtybox_any.setdefault(pcode, (qpb, bpp))
                self.wt_by_key[(pcode, dest)] = {
                    'Q': g(row, iQ), 'R': g(row, iR), 'S': g(row, iS), 'T': g(row, iT),
                    'U': g(row, iU), 'V': g(row, iV), 'W': g(row, iW),
                }
            bkey = _s(g(row, iB))
            if bkey and bkey not in self.targetop_by_fg:
                self.targetop_by_fg[bkey] = g(row, iD)
            m = _s(g(row, iM))
            if m and m not in self.wi_by_fg:
                self.wi_by_fg[m] = _s(g(row, iN))
            b = _s(g(row, iB))
            if b and b not in self.ts_by_fg:
                self.ts_by_fg[b] = g(row, iC)
            gg = _s(g(row, iG))
            if gg and gg not in self.rev_by_part:
                self.rev_by_part[gg] = _s(g(row, iH))
            bl = _s(g(row, iBL))
            if bl and bl not in self.loc_by_part:
                self.loc_by_part[bl] = _s(g(row, iBM))
            bi = _s(g(row, iBI))
            if bi and bi not in self.type_by_part:
                self.type_by_part[bi] = _s(g(row, iBJ))
            av = _s(g(row, iAV))
            if av and av not in self.des_by_raw:
                self.des_by_raw[av] = _s(g(row, iAW))
        wb.close()

    def has_packaging(self, fg, bucket):
        return any(r['dest'] == bucket for r in self.pack_by_fg.get(fg, []))


def _ceil_div(a, b):
    a = float(a or 0); b = float(b or 0)
    return math.ceil(a / b) if b else 0


def _num(v, d=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return d


# Packaging materials whose quantity is per-PALLET, not per-box: one unit lands
# on the first box of each Box/Pallet group, zero on the rest. Detected by code
# prefix (the business chose prefixes over a packaging_type table). WC additionally
# flags the ticket for wooden-crate print routing. Both lists are overridable from
# print settings, so a new prefix doesn't need a code change.
PALLET_PREFIXES = ("PL", "SPL")
CRATE_PREFIXES = ("WC",)


def resolve_pack_std(mdb, fg, bucket, pallet_prefixes=None):
    """Resolve the packaging standard (qty/box, box/pallet) for ONE FG + market
    bucket, and report how it was resolved.

    STRICT on the bucket. The P…X block in MainDatabase is keyed by `Part & PartDes`,
    so each market can carry its own qty/box — and quietly falling back to "the
    first row we saw for this FG" hands back another market's standard. That is
    what collapses the box split: a 100/box PVN row used for a 50/box Other order
    turns a 100pc PDO into ONE box (one ticket, pallet included) instead of two.

    The one fallback kept is unambiguous: if the exact bucket misses but the FG has
    exactly ONE packing row in the whole file, use it and say so. With two or more
    candidates it is a hard miss, listing the buckets that exist.

    Returns {ok, qty_per_box, box_per_pallet, bucket, source, weights,
             buckets_available, error}."""
    keys = [(p, b) for (p, b) in mdb.wt_by_key if p == fg]
    have = sorted({b for (_p, b) in keys})
    row, source = mdb.wt_by_key.get((fg, bucket)), 'exact'
    if row is None and len(keys) == 1:
        row, source = mdb.wt_by_key[keys[0]], f'only-row ({keys[0][1] or "blank market"})'
    out = {'ok': False, 'qty_per_box': 0, 'box_per_pallet': 1, 'bucket': bucket,
           'source': source, 'weights': row or {}, 'buckets_available': have,
           'error': '', 'pallet_prefixes': tuple(pallet_prefixes or PALLET_PREFIXES)}
    if row is None:
        out['error'] = (f"No packing standard for {fg} in market '{bucket}'"
                        + (f" — MainDatabase has: {', '.join(b or '(blank)' for b in have)}"
                           if have else " — this FG has no rows in the Qty/Box block at all"))
        return out
    qpb = int(_num(row.get('Q')))
    bpp = int(_num(row.get('W')))
    if qpb <= 0:
        out['error'] = f"Qty/box is empty or zero for {fg} / '{bucket}' (MainDatabase col Q)"
        return out
    out.update(ok=True, qty_per_box=qpb, box_per_pallet=bpp if bpp > 0 else 1)
    return out


def is_pallet_material(code, prefixes=PALLET_PREFIXES):
    """True if this material is consumed per pallet rather than per box."""
    return str(code or "").upper().startswith(tuple(p.upper() for p in prefixes))


def build_bpt(mdb, fg, po, dest, station, pro_time, po_qty, fg_desc="",
              pallet_prefixes=None, crate_prefixes=None):
    """Build the full ticket (one entry per box) for FG+destination. `dest` is the
    PDO-report Destination value (plan col F, a NAME like 'TRICAP'/'Shakopee').
    CheckDes (the printed/barcoded destination code) = XLOOKUP(dest, AV→AW), e.g.
    'TRI'/'SHA'. The BOM market bucket (PVN/Poland/Other) is derived from dest.
    Per §9b: per-box pick qty = RoundUp(box_qty / lotqty) * usage, with the PL/SPL
    pallet rule (qty 1 only on the first box of each Box/Pallet group) and WC routing."""
    bucket = market_bucket(dest)
    dest_code = mdb.des_by_raw.get(dest) or mdb.des_by_raw.get((dest or "").strip()) or "Chua co du lieu"
    po_qty = int(po_qty or 0)
    pal_pfx = tuple(pallet_prefixes or PALLET_PREFIXES)
    crate_pfx = tuple(crate_prefixes or CRATE_PREFIXES)

    std = resolve_pack_std(mdb, fg, bucket, pal_pfx)
    qty_per_box_std = std['qty_per_box']
    box_per_pallet = std['box_per_pallet']
    wi = mdb.wi_by_fg.get(fg, "")
    ts = mdb.ts_by_fg.get(fg, "")

    mats = [r for r in mdb.pack_by_fg.get(fg, []) if r['dest'] == bucket]

    # One ticket per box: ceil(PDO qty / qty per box), last box carries the
    # remainder. Pallet-level materials appear only on the box that opens each
    # Box/Pallet group, so ticket 1 of a 2-box/pallet run carries the pallet and
    # ticket 2 carries box materials only.
    total_boxes = _ceil_div(po_qty, qty_per_box_std) if qty_per_box_std else 0
    total_pallets = _ceil_div(total_boxes, box_per_pallet) if box_per_pallet else 0
    wooden_crate = False
    boxes = []
    for xx in range(1, total_boxes + 1):
        is_final = (xx == total_boxes)
        if is_final and qty_per_box_std:
            rem = po_qty % qty_per_box_std
            box_qty = rem if rem else qty_per_box_std
        else:
            box_qty = qty_per_box_std
        opens_pallet = ((xx - 1) % box_per_pallet == 0)
        pallet_no = ((xx - 1) // box_per_pallet) + 1
        rows = []
        for i, m in enumerate(mats, start=1):
            code = m['code']
            usage = float(m['usage'] or 0)
            lotqty = float(m['lotqty'] or 0)
            # Pallet-level material: one unit on the box that opens each pallet
            # group, zero on the rest. Everything else scales with THIS box's qty.
            if is_pallet_material(code, pal_pfx):
                qty = 1 if opens_pallet else 0
            else:
                qty = _ceil_div(box_qty, lotqty) * usage
            if str(code or "").upper().startswith(tuple(p.upper() for p in crate_pfx)):
                wooden_crate = True
            ratio = f"{_fmt_num(m['usage'])}/{_fmt_num(m['lotqty'])}"
            rows.append({
                'no': i, 'code': code, 'desc': m['desc'],
                'ver': mdb.rev_by_part.get(code, ""),
                'ratio': ratio, 'qty': _fmt_num(qty),
                'location': mdb.loc_by_part.get(code, ""),
                'type': mdb.type_by_part.get(code, ""),
                'per_pallet': is_pallet_material(code, pal_pfx),
            })
        boxes.append({'box_no': xx, 'box_qty': box_qty, 'is_final': is_final,
                      'opens_pallet': opens_pallet, 'pallet_no': pallet_no,
                      'materials': rows})

    return {
        'po': po, 'fg': fg, 'fg_desc': fg_desc, 'po_qty': po_qty,
        'dest_name': dest, 'dest_text': dest_code, 'bucket': bucket,
        'wi': wi, 'station': station, 'pro_time': pro_time, 'ts': ts,
        'qty_per_box_std': qty_per_box_std, 'box_per_pallet': box_per_pallet,
        'total_boxes': total_boxes, 'total_pallets': total_pallets,
        'wooden_crate': wooden_crate, 'boxes': boxes, 'std': std,
        'material_count': len(mats),
        'has_data': bool(qty_per_box_std and mats),
    }


def _fmt_num(v):
    try:
        f = float(v)
        return str(int(f)) if f == int(f) else str(round(f, 3))
    except (TypeError, ValueError):
        return _s(v)


def _esc(s):
    return (str(s if s is not None else "")
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def render_html(ticket, min_rows=16):
    """One printable A5-portrait ticket page per box, laid out to match the Excel
    BPT (three Code128 barcodes: ticket-code/box-count, destination, PDO). The
    'LẤY TEM & TIÊU CHUẨN CV' banner prints only on the first ticket of the PDO
    (box 1)."""
    # PDO barcode: the one operators scan most, so it's the largest and carries no
    # caption (the PDO is printed right next to it anyway).
    pdo_bc = code128_datauri(ticket['po'], module_height=20, font_size=9,
                             module_width=0.52, write_text=False)
    # Destination is one short code and identical on every page of the ticket —
    # render it ONCE instead of per box.
    dest_bc = code128_datauri(ticket['dest_text'], module_height=22, font_size=13,
                              module_width=1.1)
    # Only claim a pallet if this FG actually has pallet-level materials.
    has_pallet_mats = any(m.get('per_pallet') for b in ticket['boxes'] for m in b['materials'])
    pages = []
    for box in ticket['boxes']:
        ticket_bc = code128_datauri(f"{ticket['fg']}-P-{box['box_no']}/{ticket['total_boxes']}",
                                    module_height=16, font_size=9, module_width=0.42)
        final = ' <span class="final">(Final)</span>' if box['is_final'] and ticket['total_boxes'] > 1 else ''
        # No PL/SPL material for this FG = the box needs no pallet at all, so say
        # nothing rather than labelling every box "no pallet".
        if not has_pallet_mats:
            pallet_tag = ''
        elif box['opens_pallet']:
            pallet_tag = ' <span class="ptag">pallet on this box</span>'
        else:
            pallet_tag = ' <span class="ptag dim">no pallet</span>'
        rows = list(box['materials'])
        pad = max(0, min_rows - len(rows))
        # A pallet-level material still prints on every ticket, but greyed with a
        # 0 on the boxes that don't open a pallet — so the picker sees it was
        # considered and deliberately not picked, rather than silently missing.
        mat_rows = "".join(
            f"<tr class='{'zero' if str(m['qty']) == '0' else ''}'>"
            f"<td class='c'>{m['no']}</td><td class='mono'>{_esc(m['code'])}</td><td>{_esc(m['desc'])}</td>"
            f"<td class='c'>{_esc(m['ver'])}</td><td class='mono c'>{_esc(m['ratio'])}</td>"
            f"<td class='qty'>{_esc(m['qty'])}</td><td>{_esc(m['location'])}</td><td class='c'>{_esc(m['type'])}</td></tr>"
            for m in rows
        ) + "".join("<tr><td class='c'>&nbsp;</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>"
                    for _ in range(pad))
        banner = ('<div class="banner">LẤY TEM &amp; TIÊU CHUẨN CV</div>' if box['box_no'] == 1 else '<div class="banner-sp"></div>')
        pages.append(f"""
      <div class="ticket">
        <div class="top">
          <div class="tl">
            <div class="title">BPT for Packaging materials</div>
          </div>
          <div class="tr">
            <div class="lbl">Ticket code number</div>
            <img class="ticketbc" src="{ticket_bc}" alt="ticket code">
          </div>
        </div>
        {banner}

        <table class="meta">
          <tr><td class="k">Date:</td><td>{_esc(ticket.get('date',''))}</td>
              <td class="k">PDO:</td><td class="mono">{_esc(ticket['po'])}</td>
              <td class="k">WI:</td><td>{_esc(ticket['wi'])}</td></tr>
          <tr><td class="k">Ticket issue time:</td><td>{_esc(ticket.get('issue_time',''))}</td>
              <td class="k">FG No:</td><td class="mono">{_esc(ticket['fg'])}</td>
              <td class="k">Dest:</td><td class="dest">{_esc(ticket['dest_text'])}</td></tr>
          <tr><td class="k">Pro start time:</td><td>{_esc(ticket['pro_time'])}</td>
              <td class="k">Quantity:</td><td>{_esc(ticket['po_qty'])}</td>
              <td colspan="2" class="bc"><img class="destbc" src="{dest_bc}" alt="dest"></td></tr>
          <tr><td class="k">Station:</td><td class="station">{_esc(ticket['station'])}</td>
              <td class="k">FG name:</td><td>{_esc(ticket['fg_desc'])}</td>
              <td class="k">Materials use for:</td><td class="muf"><b>{_esc(box['box_qty'])}{final}</b>&nbsp;Pcs FG</td></tr>
          <tr><td class="k">Box:</td><td class="boxno">{box['box_no']} / {ticket['total_boxes']}</td>
              <td class="k">Pallet:</td><td>{box['pallet_no']} / {ticket.get('total_pallets', 0)}{pallet_tag}</td>
              <td class="k">Standard:</td>
              <td>{ticket['qty_per_box_std']} pcs/box · {ticket['box_per_pallet']} box/pallet</td></tr>
        </table>

        <div class="pdoblock">
          <span class="k">PDO:</span> <span class="mono pdotxt">{_esc(ticket['po'])}</span>
          <img class="pdobc" src="{pdo_bc}" alt="PDO">
        </div>

        <table class="mat">
          <thead><tr><th>No</th><th>Part No.</th><th>Part description</th><th>Ver</th>
            <th>BOM data</th><th>Qty</th><th>Location</th><th>Type</th></tr></thead>
          <tbody>{mat_rows}</tbody>
        </table>
      </div>""")

    return f"""<!doctype html><html><head><meta charset="utf-8"><style>
  @page {{ size: A5 portrait; margin: 6mm; }}
  * {{ box-sizing: border-box; }}
  body {{ font-family: Arial, 'Segoe UI', sans-serif; color: #111; margin: 0; }}
  .ticket {{ page-break-after: always; }}
  /* the last ticket must NOT force a break, or every job ends on a blank sheet */
  .ticket:last-child {{ page-break-after: auto; }}
  .top {{ display: flex; align-items: flex-start; justify-content: space-between; }}
  .title {{ font-size: 15px; font-weight: 700; border-bottom: 2px solid #111; display: inline-block; padding-bottom: 1px; }}
  /* full-width band directly under the header, above the meta grid */
  .banner {{ margin: 4px 0 2px; background: #111; color: #fff; font-weight: 700;
             padding: 4px 10px; display: block; font-size: 13px; letter-spacing: .03em;
             text-align: center; }}
  .banner-sp {{ height: 6px; }}
  .tr {{ text-align: center; }}
  .tr .lbl {{ font-size: 10px; font-weight: 700; }}
  .ticketbc {{ height: 58px; }}
  table.meta {{ width: 100%; border-collapse: collapse; margin-top: 4px; font-size: 11px; }}
  table.meta td {{ padding: 2px 5px; vertical-align: middle; }}
  table.meta td.k {{ font-weight: 700; white-space: nowrap; }}
  .station {{ font-size: 13px; font-weight: 700; }}
  .boxno {{ font-size: 13px; font-weight: 700; }}
  .ptag {{ font-size: 9.5px; font-weight: 700; background: #111; color: #fff; padding: 1px 5px; border-radius: 2px; }}
  .ptag.dim {{ background: #ddd; color: #555; }}
  table.mat tr.zero td {{ color: #999; }}
  .dest {{ font-weight: 700; font-size: 13px; }}
  /* destination barcode sits under the Dest field, left-aligned with it */
  .bc {{ text-align: left; padding-top: 1px !important; }}
  .destbc {{ height: 64px; }}
  .muf {{ font-style: italic; text-align: left; white-space: nowrap; }}
  .final {{ color: #b00; }}
  .pdoblock {{ margin: 3px 0 4px; }}
  .pdotxt {{ font-size: 13px; font-weight: 700; }}
  .pdobc {{ display: block; height: 58px; margin-top: 1px; }}
  table.mat {{ width: 100%; border-collapse: collapse; margin-top: 3px; font-size: 10px; }}
  table.mat th, table.mat td {{ border: 1px solid #333; padding: 2px 4px; text-align: left; height: 16px; }}
  table.mat th {{ background: #e8e8e8; text-align: center; font-size: 9.5px; }}
  table.mat td.c {{ text-align: center; }}
  table.mat td.qty {{ text-align: left; font-weight: 700; }}
  .mono {{ font-family: 'Consolas', monospace; }}
</style></head><body>{''.join(pages)}</body></html>"""
