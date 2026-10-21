"""Offline checks for the break-aware ETC (build .15): working-time calendar,
timer anchoring on reset / scan / crew signals, the snapshot's pause reasons
and the timer-debug payload. No SQL, no Excel, no Windows.

    python tests/test_timer_calendar.py

Importing app runs its headless startup (it creates data/ beside app.py, which
is git-ignored). The clock is pinned by swapping app.datetime / app.date.
"""
import os, sys, json
from datetime import datetime as _dt, date as _date, timedelta

sys.argv = ['test']
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import app  # noqa: E402

app.DEV_MODE = True   # keep the background writer off the (fake) share while we poke state

DAY = _date(2026, 10, 21)          # a Wednesday — weekday OT rules
assert DAY.weekday() < 5
NOW = [_dt(2026, 10, 21, 9, 0)]

class FakeDT(_dt):
    @classmethod
    def now(cls, tz=None):
        n = NOW[0]
        return cls(n.year, n.month, n.day, n.hour, n.minute, n.second)

class FakeDate(_date):
    @classmethod
    def today(cls):
        return cls(DAY.year, DAY.month, DAY.day)

app.datetime = FakeDT
app.date = FakeDate

def at(h, m=0, day=0):
    return FakeDT(2026, 10, 21, h, m) + timedelta(days=day)

def worker(wid, shift, ot_clock=0, windows=None):
    start, end = app._worker_window(shift, ot_clock)
    return {'id': wid, 'name': f'Op {wid}', 'shift': shift, 'start': start, 'end': end,
            'active': True, 'note': '', 'note_kind': '',
            'tier': app._worker_tier(shift, ot_clock), 'windows': windows or []}

def hm(d):
    return d.strftime('%H:%M') + ('' if d.date() == DAY else f' (+{(d.date() - DAY).days}d)')

checks = 0
def check(cond, msg):
    global checks
    checks += 1
    assert cond, msg
    print(f'  ✓ {msg}')

# ── 1. Calendar ────────────────────────────────────────────────────────────────
print('calendar')
ca1 = worker('1', 'Ca 1')                       # 06:00–14:10, breaks 08:00, 10:30–11:00, 12:10
iv = app.worker_work_intervals(ca1)
check([round(x * 60) for x in iv[0]] == [365, 480], 'Ca1 starts producing at 06:05 (after 5S), first stop 08:00')
check(round(iv[-1][1] * 60) == 14 * 60 + 5, 'Ca1 stops producing at 14:05 (5S before 14:10)')

w = app.walk_work(at(10), 1.0, [ca1])           # 60 min from 10:00, meal 10:30–11:00
check(hm(w['end']) == '11:30' and not w['estimate'], '60 min from 10:00 ends 11:30 — the 10:30 meal is skipped')

w = app.walk_work(at(10), 2.0, [ca1, worker('2', 'Ca 1')])   # 2 ops → 60 wall-min
check(hm(w['end']) == '11:30', 'two operators halve it, meal still skipped')

ca1_ot2, ca2_ot2 = worker('3', 'Ca 1', 2), worker('4', 'Ca 2', 2)
segs = app._crew_segments([ca1_ot2, ca2_ot2], DAY)
n_at = lambda t: sum(n for A, B, n in segs if A <= t < B)
check(n_at(at(12, 12)) == 0, 'OT2 12:12: Ca1 on its 12:10 break, Ca2 in its opening 5S → nobody producing')
check(n_at(at(12, 17)) == 1, 'OT2 12:17: Ca2 producing alone while Ca1 is still on break')
check(n_at(at(14, 15)) == 0, 'OT2 14:10–14:20: the shared break stops both')
check(n_at(at(15)) == 2, 'OT2 15:00: both producing')

ca1_ot3, ca2_ot3 = worker('5', 'Ca 1', 3), worker('6', 'Ca 2', 3)
segs = app._crew_segments([ca1_ot3, ca2_ot3], DAY)
check(n_at(at(14, 20)) == 0 and n_at(at(15)) == 2, 'OT3: 14:10–14:40 shared meal, both back at 14:40')

w = app.walk_work(at(13), 3.0, [ca1])           # 65 min left in Ca1, then Ca2 carries on
check(w['estimate'] and hm(w['crew_end']) == '14:05', 'running past the crew\'s shift is flagged as an estimate')
check(hm(w['end']) == '16:10', '3h from 13:00 for a Ca1 op: 65 min to 14:05, Ca2 from 14:15 → 16:10')

w = app.walk_work(at(21), 3.0, [worker('7', 'Ca 2')])
check(w['end'].date() > DAY, 'work left at 22:20 continues the next morning')

hc = worker('8', 'HC')
check(app.worker_downtime_kind(hc, 12.3) == 'break' and app.worker_downtime_kind(hc, 8.02) == '5s',
      'HC 12:00–12:40 meal is a break, 08:00–08:05 is 5S')

half = worker('9', 'Ca 1', windows=[[6.0, 11.0]])   # half day via note
check(round(app.worker_work_intervals(half)[-1][1] * 60) == 630, 'half-day note 06:00–11:00: last producing minute before the 10:30 meal')

done = app.work_done_between(at(10), at(11, 30), [ca1])
check(abs(done - 1.0) < 1e-9, 'work_done_between is the same walk measured (10:00–11:30 = 60 min)')

# ── 2. Anchors + signals ───────────────────────────────────────────────────────
print('anchors')
STATION = app._station_cache[0]
PDO = {'id': 'PDO-T1', 'part': 'P-1', 'desc': 'test', 'qty': 100, 'finished_hrs': 10.0,
       'remaining_qty': 100, 'completed_qty': 0}       # 6 min/unit
with app._cache_lock:
    app._worker_cache = [ca1, worker('2', 'Ca 1')]
    app._pdo_cache = [PDO]
    app._fg_latest_scan_cache = {}
    app._wip_cache = {}
    app._timer_log = {'date': DAY.isoformat(), 'stations': {}}
assignments = {s: {'inprogress': None, 'queue': []} for s in app._station_cache}
assignments[STATION] = {'inprogress': 'PDO-T1', 'queue': []}
workers = {STATION: ['1']}
NOW[0] = at(9)
timers = {STATION: {'started_at': at(9).isoformat(), 'fg_count': 0}}

check(app.retime_timers(timers, assignments, workers), 'a new order gets anchored')
t = timers[STATION]
check(t['anchor_at'] == at(9).isoformat() and abs(t['anchor_ph'] - 10.0) < 1e-9, 'anchor = reset time, 10 person-hours')
ev = app._timer_log['stations'][STATION][-1]
check(ev['kind'] == 'reset' and ev['reason'] == 'order started', 'logged as a reset (order started)')
first_end = app._parse_dt(ev['end'])
# 10h for one Ca1 op from 09:00: 09:00–10:30 (90) 11:00–12:10 (70) 12:20–14:05 (105) = 265 min today,
# then Ca2 carries on — so it must be an estimate past the shift, and land later than a naive 19:00.
check(ev['estimate'] and first_end > at(19), 'a 10h order on one Ca1 op runs past the shift (est.) and past a naive 19:00')

check(not app.retime_timers(timers, assignments, workers), 'nothing changed → no re-anchor, no log spam')

# FG scan: 20 units done, SQL scan time 09:50, seen at 10:00
NOW[0] = at(10)
PDO['remaining_qty'] = 80
app._fg_latest_scan_cache = {('PDO-T1', 'ST1'): at(9, 50)}
check(app.retime_timers(timers, assignments, workers), 'a scan re-anchors')
t = timers[STATION]
check(t['anchor_at'] == at(9, 50).isoformat() and abs(t['anchor_ph'] - 8.0) < 1e-9,
      'anchored at the scan\'s own SQL time (09:50) with 8 person-hours left')
ev = app._timer_log['stations'][STATION][-1]
check(ev['kind'] == 'scan' and ev['reason'] == 'remaining 100 → 80', 'logged as a scan with the qty change')

# Old double-count bug: live remaining against the ORIGINAL start. Now the end must not jump earlier
# than the honest plan just because a scan arrived.
honest = app.walk_work(at(9, 50), 8.0, [ca1])['end']
check(app._parse_dt(ev['end']) == honest, 'supposed end after the scan = walk from 09:50 with 80 units left')

# Crew change at 10:10: second op joins. 20 min done by the old crew (09:50–10:10).
NOW[0] = at(10, 10)
workers = {STATION: ['1', '2']}
check(app.retime_timers(timers, assignments, workers), 'a crew change re-anchors')
t = timers[STATION]
check(t['anchor_at'] == at(10, 10).isoformat() and abs(t['anchor_ph'] - (8.0 - 20/60)) < 1e-3,
      'work done before the change is credited at the old crew\'s pace')
check(app._timer_log['stations'][STATION][-1]['kind'] == 'crew', 'logged as a crew change')

# A shift reset replaces the dict → next poll anchors from it with its reason
NOW[0] = at(14, 10)
timers[STATION] = {'started_at': at(14, 10).isoformat(), 'fg_count': 0, 'reset_reason': 'shift change'}
app.retime_timers(timers, assignments, workers)
check(app._timer_log['stations'][STATION][-1]['reason'] == 'shift change', 'shift reset logged with its reason')

# A timer left by an older build (no anchor, started hours ago) starts its block NOW,
# not at its stale started_at against today's remaining (the old double-count).
NOW[0] = at(13)
old_timers = {STATION: {'started_at': at(7).isoformat(), 'fg_count': 0}}
app.retime_timers(old_timers, assignments, workers)
check(old_timers[STATION]['anchor_at'] == at(13).isoformat(), 'a pre-.15 timer is anchored at the first poll, not at its old start')

# ── 3. Snapshot: pause reasons ────────────────────────────────────────────────
print('snapshot')
def snap_at(t):
    NOW[0] = t
    app.save_assignments(assignments)
    app.save_workers_assigned(workers)
    timers[STATION] = {'started_at': at(9).isoformat(), 'fg_count': 0}
    app.save_timers(timers)
    return app.build_board_snapshot(app.AREA)['board'][STATION]['timer']

with app._cache_lock:
    app._station_cache = list(app._station_cache)
    tm = snap_at(at(10, 45))
    check(tm['paused'] and tm['pause_reason'] == 'break' and tm['etc'] is None,
          '10:45 (Ca1 meal): paused, reason break, no countdown')
    tm = snap_at(at(6, 2))
    check(tm['paused'] and tm['pause_reason'] == '5s', '06:02: paused for the opening 5S')
    tm = snap_at(at(9, 30))
    check(not tm['paused'] and tm['pause_reason'] is None and tm['etc'], '09:30: running with an ETC')
    tm = snap_at(at(23))
    check(tm['paused'] and tm['pause_reason'] == 'off_hours', '23:00: off hours')

# ── 4. Debug payload + feed ───────────────────────────────────────────────────
print('debug')
with app._cache_lock:
    snap_at(at(9, 2))          # the poll after a 09:00 reset
    app._timer_log = None
    app.retime_timers(timers, assignments, workers)
    app.save_timers(timers)
    d = app.build_timer_debug(app.AREA)
st = next(s for s in d['stations'] if s['id'] == STATION)
check(st['state'] == 'running' and st['block']['start'] == at(9).isoformat(timespec='seconds'),
      'debug row: running, block started 09:00')
check(st['block']['stopped_min'] > 0, 'debug row reports the break time skipped inside the block')
check(len(st['crew']) == 2 and all(c['producing'] for c in st['crew']), 'debug row lists the crew as producing')
check(any(s['shift'] == 'Ca 1' and '10:30–11:00' in s['breaks'] for s in d['schedule']),
      "today's Ca1 break table is published")
json.dumps(d, default=str)

app.DEV_MODE = False
app.write_debug_status()
app.DEV_MODE = True
p = os.path.join(app.DASHBOARD_STATUS_DIR, f'{app.AREA}_debug.js')
check(os.path.exists(p) and open(p, encoding='utf-8').read().startswith('window.NPV_DEBUG'),
      f'{app.AREA}_debug.js written to the share folder')

client = app.app.test_client()
r = client.get('/api/debug')
check(r.status_code == 200 and r.get_json()['area'] == app.AREA, '/api/debug serves the payload')
r = client.get('/debug')
check(r.status_code == 200 and b'TIMER DEBUG' in r.data, '/debug serves shared_debug.html')

print(f'\nall {checks} checks passed')
