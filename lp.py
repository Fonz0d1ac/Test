"""
License Plate mint — reuse-then-mint, plus the [License Plate] audit INSERT.

This replaces the DUMMY plate that boxlabel.make_dummy_lp() has been handing out
since the printing subsystem landed. It is the Python side of the Excel macro's
LP logic (ModeQty), ported per HANDOFF.md §5 with the two macro bugs in §5.4
deliberately NOT reproduced.

── Format (verified against production data) ──────────────────────────────────
    LicensePlate = Left(vendor,6) + yymmdd + serial(6)   e.g. V-0023 260812 200042
    serial       = area_base + MAX(last 5 digits today, same area) + 1

The FIRST DIGIT OF THE SERIAL IS THE AREA CODE — all three areas share one
[License Plate] table, and the disjoint ranges are what keeps them apart:

    AD Assembly  digit '0'  base      0  first serial of day 000001
    PK Packaging digit '2'  base 200000  first serial of day 200001
    RP Raw Part  digit '5'  base 500000  first serial of day 500001

The counter is per (area, day), SHARED ACROSS VENDORS. The SQL anchors from the
RIGHT (LEN-11, LEN-5, RIGHT(...,5)) so a vendor code shorter than 6 chars still
parses.

── Why MAX+1 and not our own counter table ────────────────────────────────────
The Excel macros are still live and mint off MAX(right(lp,5))+1 on this same
shared table. Any counter of our own would collide in both directions. MAX+1
also self-heals against unknown writers — which matters, because the legacy
Gen_other_info tool mints into AD's '0' pool with VendorCode="V-0000".
See HANDOFF.md §5.5 for the full argument, and why the area_counters design is
a phase-2 target rather than this one.

── Deviations from the macro (both intentional, HANDOFF.md §5.4) ──────────────
1. The reuse query gets an ORDER BY. The macro has none, and there are 113
   PII_PO groups with 2+ simultaneously eligible plates, so SQL Server's pick is
   arbitrary. We take the most recent production date.
2. Mint + INSERT are committed BEFORE anything prints. The macro prints first,
   so a printer exception consumed a serial that was never recorded and the next
   run re-minted the same plate onto a physical label. Our worst case is an
   orphan row and a gap in the sequence — never a duplicate plate.

── Concurrency ────────────────────────────────────────────────────────────────
The obvious fix is a trap: `SELECT MAX(right(lp,5)) WHERE SUBSTRING(...)` is
non-sargable and table-scans, so XLOCK/HOLDLOCK around it risks lock escalation
to a table-level exclusive lock on a production table — worse than the bug.
Instead:
  • sp_getapplock (LP_MINT_{AREA}_{yymmdd}) — a cheap named mutex, no table
    locks, fully serialising app-vs-app.
  • Read-back verification with retry for app-vs-MACRO. Nothing can lock out the
    macro safely, and the macro never locked either. Residual risk is small:
    areas use disjoint serial ranges, so a collision needs the app AND the macro
    minting in the same area in the same instant.

The caller owns the connection and the commit. Nothing here commits or rolls
back on its own, EXCEPT the applock release, which must survive any failure.
"""

import datetime

# ── Area serial ranges ─────────────────────────────────────────────────────────
# (first digit of the serial, base added to the per-day counter). Never hardcode
# one area here — this table is the only place the area matters in the whole mint.
AREA_SERIAL = {
    'AD': ('0', 0),
    'PK': ('2', 200000),
    'RP': ('5', 500000),
}

# Table names, kept in one place: the schemas were given verbally (HANDOFF §11.1),
# so if a name is wrong this is the single spot to fix it.
T_LP   = '[dbo].[License Plate]'
T_WIP  = '[dbo].[Nhaplecuoingay_All]'
T_FG   = '[dbo].[FG_Database_All]'

# How many times to re-read MAX and recompute when the read-back finds that
# somebody else (the macro) took one of our candidate serials in between.
MINT_RETRIES = 3
# sp_getapplock timeout, milliseconds. A mint is a few table scans; if we cannot
# get the mutex in 10s something is badly wrong and the operator should hear it
# rather than stand at a printer that has silently stalled.
APPLOCK_TIMEOUT_MS = 10000
# A day's serial counter has five digits. 99,999 boxes in one area in one day is
# impossible, but rolling past it would silently mint into another area's range,
# so it fails loudly instead.
MAX_DAILY_SERIAL = 99999


class LPError(Exception):
    """Anything that must stop the print before a label is produced."""


def area_serial_spec(area):
    """('2', 200000) for PK. Raises on an unknown area rather than defaulting —
    a default would mint into whichever area's range it defaulted to."""
    spec = AREA_SERIAL.get(str(area or '').strip().upper())
    if not spec:
        raise LPError(f'No License Plate serial range for area {area!r} '
                      f'(known: {", ".join(sorted(AREA_SERIAL))})')
    return spec


def date_text(today=None):
    """yymmdd, the middle 6 characters of the plate."""
    return (today or datetime.date.today()).strftime('%y%m%d')


def format_lp(vendor_code, dtext, serial_int):
    """Left(vendor,6) + yymmdd + serial(6). The serial is already area-based."""
    return f"{str(vendor_code or '')[:6]}{dtext}{int(serial_int):06d}"


def dummy_plates(area, vendor_code, n, today=None):
    """Offline / dev-mode plates: the real 18-char shape and the CORRECT area
    range, but never inserted and never derived from SQL. Used when dev mode is
    on, when pyodbc is missing, or when live minting hasn't been switched on for
    this area yet.

    The area range matters even for a dummy: the old make_dummy_lp() hardcoded
    200000, so an AD or RP test print produced a plate that LOOKED like a
    Packaging one. Anything that later scanned or grepped it would be wrong
    about which area it came from."""
    _digit, base = area_serial_spec(area)
    dtext = date_text(today)
    return [format_lp(vendor_code, dtext, base + i) for i in range(1, int(n) + 1)]


# ── Reuse pass ─────────────────────────────────────────────────────────────────

def find_reusable_plates(cursor, *, po, pii_po, n, log=print):
    """Plates already issued to an EARLIER PDO in the same PII_PO family that
    never made it to FG — up to n of them, most recent production date first.

    Why this exists (HANDOFF §5.3): at month-end planning splits an in-progress
    PDO (PDO-0001 75/100 → PDO-0001 closed + PDO-0002 continues). Whatever
    partial quantity was logged against the OLD PDO has to keep the SAME plate,
    or the efficiency calculation — which joins on LP — loses that WIP. PII_PO
    is the family ID that survives the split (assigned upstream in the MRP/D365
    import, not in any macro here).

    Two differences from the macro, both deliberate:
      • TOP (n) in ONE query instead of N queries accumulating a UsedLP string.
        Taking n rows in one shot makes the within-run exclusion structural —
        a plate cannot be handed to two boxes of the same run because it is only
        returned once.
      • ORDER BY MAX([Production date]) DESC. The macro has no ORDER BY at all
        (§5.4) so its pick is arbitrary whenever a family has 2+ eligible plates,
        which live data says happens for 113 families.
    """
    n = int(n)
    if n <= 0:
        return []
    pii = str(pii_po or '').strip()
    if not pii:
        # A blank PII_PO would match every other blank-PII row in the table and
        # hand this PDO a pile of unrelated plates. The macro never hit this
        # because it ran off a sheet where PII was always populated; the board
        # serves orders straight from the plan Excel, where column Q can be empty.
        log('[lp] reuse pass skipped — this PDO has no PII_PO, so there is no '
            'order family to inherit a plate from')
        return []
    cursor.execute(f"""
        SELECT TOP (?) a.[License Plate]
        FROM {T_WIP} a
        WHERE a.PII_PO = ? AND a.PO <> ?
          AND NOT EXISTS (SELECT 1 FROM {T_FG} b
                          WHERE b.[License Plate] = a.[License Plate])
        GROUP BY a.[License Plate]
        ORDER BY MAX(a.[Production date]) DESC
    """, n, pii, str(po))
    plates = [str(r[0]).strip() for r in cursor.fetchall() if r[0] and str(r[0]).strip()]
    if plates:
        log(f'[lp] reuse: {len(plates)} orphaned plate(s) from PII_PO {pii} '
            f'→ {", ".join(plates)}')
    return plates


# ── Mint pass ──────────────────────────────────────────────────────────────────

def _next_serial(cursor, digit, base, dtext):
    """area_base + MAX(last 5 digits today, same area) + 1.

    TRY_CAST rather than CAST: a single malformed plate in the table (a manual
    fix-up, a legacy row) would otherwise abort the whole mint with a conversion
    error. TRY_CAST turns that row into a NULL, which MAX ignores.

    LEN-11 / LEN-5 anchor from the right exactly as the macro does, so a vendor
    code shorter than 6 characters still parses."""
    cursor.execute(f"""
        SELECT MAX(TRY_CAST(RIGHT([License Plate], 5) AS INT))
        FROM {T_LP}
        WHERE SUBSTRING([License Plate], LEN([License Plate]) - 11, 6) = ?
          AND SUBSTRING([License Plate], LEN([License Plate]) - 5, 1) = ?
    """, dtext, digit)
    row = cursor.fetchone()
    last5 = int(row[0]) if row and row[0] is not None else 0
    if last5 >= MAX_DAILY_SERIAL:
        raise LPError(f'License Plate serial for this area is at {last5} today — '
                      f'one more would roll out of the 5-digit range and into '
                      f'another area. Nothing was minted.')
    return base + last5 + 1


def _existing(cursor, plates):
    """Which of these plates are ALREADY in [License Plate]. Empty = clear to use."""
    if not plates:
        return set()
    marks = ','.join('?' * len(plates))
    cursor.execute(f"SELECT [License Plate] FROM {T_LP} "
                   f"WHERE [License Plate] IN ({marks})", *plates)
    return {str(r[0]).strip() for r in cursor.fetchall() if r[0]}


def _applock_acquire(cursor, resource):
    """sp_getapplock with Session ownership. Session (not Transaction) because
    we do not control the caller's transaction boundaries — the caller commits
    the INSERT when it is ready. The lock is released in the finally block of
    acquire_plates(), and in the worst case dies with the connection."""
    cursor.execute("""
        DECLARE @rc int;
        EXEC @rc = sp_getapplock @Resource = ?, @LockMode = 'Exclusive',
                                 @LockOwner = 'Session', @LockTimeout = ?;
        SELECT @rc;
    """, resource, APPLOCK_TIMEOUT_MS)
    row = cursor.fetchone()
    rc = int(row[0]) if row and row[0] is not None else -999
    if rc < 0:
        # -1 timeout · -2 cancelled · -3 deadlock victim · -999 parameter/other.
        raise LPError(f'Could not take the License Plate mint lock ({resource}, '
                      f'sp_getapplock returned {rc}). Another print is minting; '
                      f'nothing was minted here. Try again in a moment.')
    return rc


def _applock_release(cursor, resource, log=print):
    try:
        cursor.execute("EXEC sp_releaseapplock @Resource = ?, @LockOwner = 'Session';",
                       resource)
        try:
            cursor.fetchall()   # sp_releaseapplock returns a result set on some drivers
        except Exception:
            pass
    except Exception as e:
        # Not fatal — the lock dies with the connection. Still worth saying out
        # loud, because a leaked session lock stalls the NEXT operator's print.
        log(f'[lp] WARNING: could not release the mint lock {resource}: {e}')


# ── The whole acquisition ──────────────────────────────────────────────────────

def acquire_plates(conn, *, area, po, pii_po, n, vendor_code, today=None, log=print):
    """Reuse-then-mint n plates for one PDO print run.

    Returns {'plates': [...n in box order...], 'reused': [...], 'minted': [...]}.
    Reused plates go to the EARLIEST boxes and minted ones follow, matching the
    macro's order.

    Holds the per-(area, day) applock across mint AND read-back so two boards
    cannot interleave. Does NOT commit — the caller inserts the audit rows on
    the same connection and commits both together, before printing.

    Raises LPError on anything that should stop the print. Nothing is printed
    with a plate this function did not vouch for."""
    n = int(n)
    if n <= 0:
        return {'plates': [], 'reused': [], 'minted': []}
    digit, base = area_serial_spec(area)
    dtext = date_text(today)
    cursor = conn.cursor()

    reused = find_reusable_plates(cursor, po=po, pii_po=pii_po, n=n, log=log)
    need = n - len(reused)
    if need <= 0:
        return {'plates': reused[:n], 'reused': reused[:n], 'minted': []}

    resource = f'LP_MINT_{str(area).upper()}_{dtext}'
    _applock_acquire(cursor, resource)
    try:
        minted = []
        for attempt in range(1, MINT_RETRIES + 1):
            start = _next_serial(cursor, digit, base, dtext)
            if start + need - 1 - base > MAX_DAILY_SERIAL:
                raise LPError(f'Minting {need} plate(s) from {start} would roll past '
                              f'the 5-digit daily serial range. Nothing was minted.')
            candidates = [format_lp(vendor_code, dtext, start + i) for i in range(need)]
            # Read-back BEFORE the insert: the macro does not take our lock, so it
            # can have consumed this exact serial between our MAX and now.
            clash = _existing(cursor, candidates)
            if not clash:
                minted = candidates
                break
            log(f'[lp] mint attempt {attempt}/{MINT_RETRIES}: {len(clash)} candidate '
                f'plate(s) already exist ({", ".join(sorted(clash))}) — the Excel '
                f'macro minted while we were reading. Re-reading MAX.')
        else:
            raise LPError(f'Could not find {need} free License Plate serial(s) after '
                          f'{MINT_RETRIES} attempts — the Excel macro is minting into '
                          f'{area} at the same time. Nothing was minted; try again.')
        span = minted[0] if len(minted) == 1 else f'{minted[0]} … {minted[-1]}'
        log(f'[lp] mint: {len(minted)} new plate(s) for {area} ({span})')
        return {'plates': reused + minted, 'reused': reused, 'minted': minted}
    finally:
        _applock_release(cursor, resource, log=log)


# ── The audit INSERT ───────────────────────────────────────────────────────────

def insert_plate_rows(conn, rows, *, sql_write, log=print):
    """INSERT the assembled [License Plate] audit rows, ONE statement per row,
    routed through the caller's sql_write() so dev mode dry-runs them like every
    other board-originated write.

    `rows` are boxlabel's sql_row dicts. The column list is taken from the dict
    keys rather than hardcoded here, so the schema lives in exactly one place
    (boxlabel.build_box_labels) and a wrong column name is fixed once.

    Returns the number of rows written (0 in dev mode). Does NOT commit."""
    if not rows:
        return 0
    cursor = conn.cursor()
    written = 0
    for r in rows:
        cols = list(r.keys())
        collist = ', '.join(f'[{c}]' for c in cols)
        marks = ', '.join('?' * len(cols))
        res = sql_write(cursor, f'INSERT INTO {T_LP} ({collist}) VALUES ({marks})',
                        tuple(r[c] for c in cols), label=f"License Plate {r.get('License Plate')}")
        if res is not None:
            written += 1
    return written


def verify_inserted(conn, plates, log=print):
    """After the commit: every plate we just wrote must appear EXACTLY ONCE.

    This is the app-vs-macro backstop. If the macro squeezed a duplicate in
    despite the read-back, the operator finds out here — before the label is
    printed — instead of two boxes carrying the same plate onto the floor."""
    if not plates:
        return True
    cursor = conn.cursor()
    marks = ','.join('?' * len(plates))
    cursor.execute(f"SELECT [License Plate], COUNT(*) FROM {T_LP} "
                   f"WHERE [License Plate] IN ({marks}) GROUP BY [License Plate]", *plates)
    counts = {str(p).strip(): int(c) for p, c in cursor.fetchall()}
    bad = [p for p in plates if counts.get(p, 0) != 1]
    if bad:
        raise LPError(f'License Plate verification failed for {", ".join(bad)} '
                      f'(expected exactly one row each, found '
                      f'{ {p: counts.get(p, 0) for p in bad} }). NOTHING was printed — '
                      f'these plates must be checked in SQL before reprinting.')
    return True
