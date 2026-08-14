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
def code128_datauri(data, module_height=7.0, font_size=8, module_width=0.22):
    """Real Code128 barcode as an inline SVG data URI (no font trick, per §13)."""
    import barcode
    from barcode.writer import SVGWriter
    import base64
    data = str(data or "")
    if not data:
        return ""
    c = barcode.get('code128', data, writer=SVGWriter())
    buf = BytesIO()
    c.write(buf, options={'module_height': module_height, 'font_size': font_size,
                          'module_width': module_width, 'quiet_zone': 1, 'text_distance': 3})
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


def build_bpt(mdb, fg, po, dest, station, pro_time, po_qty, fg_desc=""):
    """Build the full ticket (one entry per box) for FG+destination. `dest` is the
    PDO-report Destination value (plan col F, a NAME like 'TRICAP'/'Shakopee').
    CheckDes (the printed/barcoded destination code) = XLOOKUP(dest, AV→AW), e.g.
    'TRI'/'SHA'. The BOM market bucket (PVN/Poland/Other) is derived from dest.
    Per §9b: per-box pick qty = RoundUp(box_qty / lotqty) * usage, with the PL/SPL
    pallet rule (qty 1 only on the first box of each Box/Pallet group) and WC routing."""
    bucket = market_bucket(dest)
    dest_code = mdb.des_by_raw.get(dest) or mdb.des_by_raw.get((dest or "").strip()) or "Chua co du lieu"
    po_qty = int(po_qty or 0)
    qpb, bpp = mdb.qtybox.get((fg, bucket)) or mdb.qtybox_any.get(fg) or (None, None)
    qty_per_box_std = int(qpb) if qpb else 0
    box_per_pallet = int(bpp) if bpp else 1
    wi = mdb.wi_by_fg.get(fg, "")
    ts = mdb.ts_by_fg.get(fg, "")

    mats = [r for r in mdb.pack_by_fg.get(fg, []) if r['dest'] == bucket]

    total_boxes = _ceil_div(po_qty, qty_per_box_std) if qty_per_box_std else 0
    wooden_crate = False
    boxes = []
    for xx in range(1, total_boxes + 1):
        is_final = (xx == total_boxes)
        if is_final and qty_per_box_std:
            rem = po_qty % qty_per_box_std
            box_qty = rem if rem else qty_per_box_std
        else:
            box_qty = qty_per_box_std
        rows = []
        for i, m in enumerate(mats, start=1):
            code = m['code']
            usage = float(m['usage'] or 0)
            lotqty = float(m['lotqty'] or 0)
            # PL / SPL: one unit on the first box of each pallet group, else zero
            if code.startswith("PL") or code.startswith("SPL"):
                qty = 1 if ((xx - 1) % box_per_pallet == 0) else 0
            else:
                qty = _ceil_div(box_qty, lotqty) * usage
            if code.startswith("WC"):
                wooden_crate = True
            ratio = f"{_fmt_num(m['usage'])}/{_fmt_num(m['lotqty'])}"
            rows.append({
                'no': i, 'code': code, 'desc': m['desc'],
                'ver': mdb.rev_by_part.get(code, ""),
                'ratio': ratio, 'qty': _fmt_num(qty),
                'location': mdb.loc_by_part.get(code, ""),
                'type': mdb.type_by_part.get(code, ""),
            })
        boxes.append({'box_no': xx, 'box_qty': box_qty, 'is_final': is_final, 'materials': rows})

    return {
        'po': po, 'fg': fg, 'fg_desc': fg_desc, 'po_qty': po_qty,
        'dest_name': dest, 'dest_text': dest_code, 'bucket': bucket,
        'wi': wi, 'station': station, 'pro_time': pro_time, 'ts': ts,
        'qty_per_box_std': qty_per_box_std, 'box_per_pallet': box_per_pallet,
        'total_boxes': total_boxes, 'wooden_crate': wooden_crate, 'boxes': boxes,
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
    pdo_bc = code128_datauri(ticket['po'], module_height=12, font_size=9, module_width=0.28)
    pages = []
    for box in ticket['boxes']:
        ticket_bc = code128_datauri(f"{ticket['fg']}-P-{box['box_no']}/{ticket['total_boxes']}",
                                    module_height=11, font_size=9)
        dest_bc = code128_datauri(ticket['dest_text'], module_height=14, font_size=9, module_width=0.3)
        final = ' <span class="final">(Final)</span>' if box['is_final'] and ticket['total_boxes'] > 1 else ''
        rows = list(box['materials'])
        pad = max(0, min_rows - len(rows))
        mat_rows = "".join(
            f"<tr><td class='c'>{m['no']}</td><td class='mono'>{_esc(m['code'])}</td><td>{_esc(m['desc'])}</td>"
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
            {banner}
          </div>
          <div class="tr">
            <div class="lbl">Ticket code number</div>
            <img class="ticketbc" src="{ticket_bc}" alt="ticket code">
          </div>
        </div>

        <table class="meta">
          <tr><td class="k">Date:</td><td>{_esc(ticket.get('date',''))}</td>
              <td class="k">PO No:</td><td class="mono">{_esc(ticket['po'])}</td>
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
        </table>

        <div class="pdoblock">
          <span class="k">PO No:</span> <span class="mono">{_esc(ticket['po'])}</span>
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
  .top {{ display: flex; align-items: flex-start; justify-content: space-between; }}
  .title {{ font-size: 15px; font-weight: 700; border-bottom: 2px solid #111; display: inline-block; padding-bottom: 1px; }}
  .banner {{ margin-top: 5px; background: #111; color: #fff; font-weight: 700; padding: 3px 8px; display: inline-block; font-size: 12px; letter-spacing: .02em; }}
  .banner-sp {{ height: 22px; }}
  .tr {{ text-align: center; }}
  .tr .lbl {{ font-size: 10px; font-weight: 700; }}
  .ticketbc {{ height: 40px; }}
  table.meta {{ width: 100%; border-collapse: collapse; margin-top: 4px; font-size: 11px; }}
  table.meta td {{ padding: 2px 5px; vertical-align: middle; }}
  table.meta td.k {{ font-weight: 700; white-space: nowrap; }}
  .station {{ font-size: 13px; font-weight: 700; }}
  .dest {{ font-weight: 700; font-size: 13px; }}
  .bc {{ text-align: right; }}
  .destbc {{ height: 34px; }}
  .muf {{ font-style: italic; text-align: left; white-space: nowrap; }}
  .final {{ color: #b00; }}
  .pdoblock {{ margin: 3px 0 4px; }}
  .pdobc {{ display: block; height: 42px; margin-top: 1px; }}
  table.mat {{ width: 100%; border-collapse: collapse; margin-top: 3px; font-size: 10px; }}
  table.mat th, table.mat td {{ border: 1px solid #333; padding: 2px 4px; text-align: left; height: 16px; }}
  table.mat th {{ background: #e8e8e8; text-align: center; font-size: 9.5px; }}
  table.mat td.c {{ text-align: center; }}
  table.mat td.qty {{ text-align: left; font-weight: 700; }}
  .mono {{ font-family: 'Consolas', monospace; }}
</style></head><body>{''.join(pages)}</body></html>"""
