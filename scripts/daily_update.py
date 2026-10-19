#!/usr/bin/env python3
"""Daily updater for data/latest.xlsx (sheet 'VIX期貨數據') -- runs in GitHub Actions.

  python3 scripts/daily_update.py close  [--today YYYY-MM-DD] [--file data/latest.xlsx]
      08:35 Taipei: insert row 5 for the previous US trading day (VIX futures M1-M7 settlement, K,
      Cboe VIX9D/VIX/VIX3M/VIX6M/VIX1Y closes, 0050 previous close), then recalculate with LibreOffice.
  python3 scripts/daily_update.py open   [--today YYYY-MM-DD] [--file data/latest.xlsx]
      09:05 Taipei: write the 0050 open price (BA5) of today, then recalculate.

Exit codes: 0 = updated, 10 = nothing to do (weekend / holiday / already done), 1 = failure (nothing written).
Only the target workbook is modified.
"""
import argparse, csv, datetime as dt, io, json, os, re, shutil, subprocess, sys, tempfile, time
import urllib.request, urllib.error
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parent))
import daily_row  # noqa: E402
import openpyxl  # noqa: E402

UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36'
TPE = dt.timezone(dt.timedelta(hours=8))
SKIP = 10


class Skip(Exception):
    pass


class Fail(Exception):
    pass


def log(*a):
    print(*a, flush=True)


def http(url, tries=3, timeout=45):
    last = None
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': UA, 'Accept-Language': 'en-US,en;q=0.9'})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read().decode('utf-8-sig', 'replace')
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(3 * (i + 1))
    raise Fail(f'fetch failed {url}: {last}')


def r3(x):
    return float(Decimal(str(x)).quantize(Decimal('0.001'), rounding=ROUND_HALF_UP))


# ---------------------------------------------------------------- calendars
def tw_closed(today):
    """True if the Taiwan stock market is closed on `today` (weekend or scheduled holiday)."""
    if today.weekday() >= 5:
        return True
    try:
        data = json.loads(http('https://openapi.twse.com.tw/v1/holidaySchedule/holidaySchedule'))
    except Exception as e:  # noqa: BLE001
        log('WARN holiday schedule unavailable:', e)
        return False
    roc = f'{today.year - 1911}{today.month:02d}{today.day:02d}'
    for it in data:
        if str(it.get('Date')) == roc:
            desc = (it.get('Description') or '') + (it.get('Name') or '')
            if '開始交易' in desc or '最後交易' in desc:
                continue
            log('TW holiday:', it)
            return True
    return False


def prev_weekday(d):
    d -= dt.timedelta(days=1)
    while d.weekday() >= 5:
        d -= dt.timedelta(days=1)
    return d


def nyse_holiday(d):
    """NYSE full-day closures by rule (used to tell a real US holiday from a late/missing Cboe file)."""
    y = d.year

    def nth(month, weekday, n):
        first = dt.date(y, month, 1)
        day = first + dt.timedelta(days=(weekday - first.weekday()) % 7 + 7 * (n - 1))
        return day

    def last(month, weekday):
        nxt = dt.date(y + (month == 12), month % 12 + 1, 1)
        day = nxt - dt.timedelta(days=1)
        return day - dt.timedelta(days=(day.weekday() - weekday) % 7)

    def observed(day):
        if day.weekday() == 5:
            return day - dt.timedelta(days=1)
        if day.weekday() == 6:
            return day + dt.timedelta(days=1)
        return day

    a, b, c = y % 19, y // 100, y % 100
    dd = (19 * a + b - b // 4 - ((b - (b + 8) // 25 + 1) // 3) + 15) % 30
    e = (32 + 2 * (b % 4) + 2 * (c // 4) - dd - c % 4) % 7
    m = (a + 11 * dd + 22 * e) // 451
    month, day = divmod(dd + e - 7 * m + 114, 31)
    easter = dt.date(y, month, day + 1)
    hs = {nth(1, 0, 3), nth(2, 0, 3), easter - dt.timedelta(days=2), last(5, 0), observed(dt.date(y, 6, 19)),
          observed(dt.date(y, 7, 4)), nth(9, 0, 1), nth(11, 3, 4), observed(dt.date(y, 12, 25))}
    ny = dt.date(y, 1, 1)
    if ny.weekday() != 5:
        hs.add(observed(ny))
    if d == dt.date(y, 12, 31) and dt.date(y + 1, 1, 1).weekday() == 5:
        pass
    return d in hs


# ---------------------------------------------------------------- data sources
def cboe_settlement(d):
    txt = http(f'https://www.cboe.com/us/futures/market_statistics/settlement/csv?dt={d.isoformat()}')
    rows = list(csv.DictReader(io.StringIO(txt)))
    out = []
    for r in rows:
        if r.get('Product') != 'VX' or not re.fullmatch(r'VX/[A-Z]\d', r['Symbol']):
            continue
        out.append((dt.date.fromisoformat(r['Expiration Date']), Decimal(r['Price'])))
    return sorted(out)


def futures_for(A):
    """-> (M1..M7 settlement list (3dp), M1 expiry, is_us_holiday)"""
    hol = False
    d = A
    rows = cboe_settlement(d)
    tries = 0
    if not rows:
        if not nyse_holiday(A):
            raise Fail(f'Cboe settlement file for {A} is empty but {A} is not a NYSE holiday (file not published yet?)')
    elif nyse_holiday(A):
        raise Fail(f'{A} is a NYSE holiday but Cboe returned settlement data; check manually')
    while not rows:
        hol = True
        d -= dt.timedelta(days=1)
        rows = cboe_settlement(d)
        tries += 1
        if tries > 6:
            raise Fail('no Cboe settlement data within 6 days before A')
    if hol:
        log(f'A={A} has no settlement file -> US holiday; using settlement of {d}')
    m = [(e, p) for e, p in rows if e > A][:7]
    if len(m) < 7:
        raise Fail(f'only {len(m)} monthly contracts after {A}')
    return [r3(p) for _, p in m], m[0][0], hol


def cboe_index_close(sym, A):
    txt = http(f'https://cdn.cboe.com/api/global/us_indices/daily_prices/{sym}_History.csv')
    best = None
    for r in csv.DictReader(io.StringIO(txt)):
        d = dt.datetime.strptime(r['DATE'], '%m/%d/%Y').date()
        if d <= A and (best is None or d > best[0]):
            best = (d, float(r['CLOSE']))
    if best is None:
        raise Fail(f'{sym}: no history row')
    return best


def twse_month(d):
    j = json.loads(http(f'https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date={d.year}{d.month:02d}01&stockNo=0050'))
    out = {}
    if j.get('stat') != 'OK':
        return out
    for row in j['data']:
        y, m, dd = row[0].split('/')
        day = dt.date(int(y) + 1911, int(m), int(dd))
        out[day] = {'open': float(row[3].replace(',', '')), 'close': float(row[6].replace(',', ''))}
    return out


def yahoo_0050():
    j = json.loads(http('https://query1.finance.yahoo.com/v8/finance/chart/0050.TW?range=1mo&interval=1d'))
    res = j['chart']['result'][0]
    q = res['indicators']['quote'][0]
    out = {}
    for ts, o, c in zip(res['timestamp'], q['open'], q['close']):
        day = dt.datetime.fromtimestamp(ts, TPE).date()
        out[day] = {'open': o, 'close': c}
    return out


def prev_close_0050(today):
    tw = {}
    tw.update(twse_month(today.replace(day=1) - dt.timedelta(days=1)))
    tw.update(twse_month(today))
    days = sorted(d for d in tw if d < today)
    if not days:
        raise Fail('TWSE: no trading day before today')
    d = days[-1]
    c = tw[d]['close']
    try:
        y = yahoo_0050().get(d)
        if y and y['close'] is not None and abs(y['close'] - c) > 0.011:
            raise Fail(f'0050 close mismatch on {d}: TWSE {c} vs Yahoo {y["close"]}')
    except Fail:
        raise
    except Exception as e:  # noqa: BLE001
        log('WARN yahoo cross-check skipped:', e)
    log(f'0050 previous close {c} on {d}')
    return c, d


def open_price_0050(today, wait_s=900):
    deadline = time.time() + wait_s
    while True:
        try:
            y = yahoo_0050().get(today)
            if y and y['open']:
                o = round(float(y['open']), 2)
                try:
                    t = twse_month(today).get(today)
                    if t and abs(t['open'] - o) > 0.011:
                        raise Fail(f'0050 open mismatch: TWSE {t["open"]} vs Yahoo {o}')
                except Fail:
                    raise
                except Exception as e:  # noqa: BLE001
                    log('WARN twse open cross-check skipped:', e)
                return o
        except Fail:
            raise
        except Exception as e:  # noqa: BLE001
            log('WARN yahoo:', e)
        if time.time() > deadline:
            raise Fail(f'no 0050 open price for {today} after waiting')
        log('open price not available yet, retry in 60s')
        time.sleep(60)


# ---------------------------------------------------------------- workbook helpers
def load_row5(path):
    ws = openpyxl.load_workbook(path, read_only=False).active
    return ws


def recalc(path, timeout=900):
    """Recalculate every formula with LibreOffice and store the file in place."""
    path = str(Path(path).resolve())
    with tempfile.TemporaryDirectory(prefix='lo-profile-') as pd:
        url = Path(pd).as_uri()
        subprocess.run(['soffice', '--headless', '--terminate_after_init', f'-env:UserInstallation={url}'],
                       capture_output=True, timeout=120, env={**os.environ, 'SAL_USE_VCLPLUGIN': 'svp'})
        macro_dir = Path(pd) / 'user' / 'basic' / 'Standard'
        if not macro_dir.exists():
            raise Fail('LibreOffice profile not created')
        (macro_dir / 'Module1.xba').write_text(
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
            '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">\n'
            'Sub RecalculateAndSave()\n  ThisComponent.calculateAll()\n  ThisComponent.store()\n  ThisComponent.close(True)\nEnd Sub\n'
            '</script:module>')
        before = os.stat(path)
        r = subprocess.run(['soffice', '--headless', '--norestore', f'-env:UserInstallation={url}',
                            'vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application', path],
                           capture_output=True, text=True, timeout=timeout, env={**os.environ, 'SAL_USE_VCLPLUGIN': 'svp'})
        if r.returncode != 0:
            raise Fail(f'soffice failed: {r.stderr[:300]}')
        after = os.stat(path)
        if (after.st_mtime_ns, after.st_size) == (before.st_mtime_ns, before.st_size):
            raise Fail('LibreOffice did not rewrite the file')


ERRS = ('#VALUE!', '#DIV/0!', '#REF!', '#NAME?', '#NULL!', '#NUM!', '#N/A')


def check_recalculated(path, expect):
    """no formula errors above row 60 (the 4 known errors sit at the very bottom); row 5 outputs present."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    bad = []
    total = 0
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value in ERRS:
                total += 1
                if c.row <= 60:
                    bad.append(c.coordinate)
    if bad:
        raise Fail(f'formula errors near the top: {bad[:10]}')
    if total > 6:
        raise Fail(f'unexpected number of formula errors: {total}')
    for k, v in expect.items():
        got = ws[k].value
        if isinstance(v, dt.date):
            got = got.date() if isinstance(got, dt.datetime) else got
            ok = got == v
        elif isinstance(v, float):
            ok = got is not None and abs(float(got) - v) < 1e-6
        else:
            ok = got == v
        if not ok:
            raise Fail(f'verify {k}: expected {v}, got {got}')
    if ws['AK5'].value in (None, ''):
        raise Fail('AK5 (strategy signal) is empty after recalculation')
    log(f'recalc ok: {total} known bottom errors; signals AK5={ws["AK5"].value} AM5={ws["AM5"].value} '
        f'AO5={ws["AO5"].value} AQ5={ws["AQ5"].value} AS5={ws["AS5"].value}')
    return ws


# ---------------------------------------------------------------- modes
def mode_close(a, today):
    if tw_closed(today):
        raise Skip('Taiwan market closed today')
    A = prev_weekday(today)
    ws = load_row5(a.file)
    a5, az5 = ws['A5'].value.date(), ws['AZ5'].value.date()
    if a5 >= A or az5 >= today:
        raise Skip(f'already done (A5={a5}, AZ5={az5})')
    old = {c: ws[f'{c}5'].value for c in list('BCDEFGH') + list('MNOPQ') + ['BB']}
    fut, m1exp, hol = futures_for(A)
    K = (m1exp - A).days - 1
    idx = {}
    for sym, col in zip(['VIX9D', 'VIX', 'VIX3M', 'VIX6M', 'VIX1Y'], 'MNOPQ'):
        d, v = cboe_index_close(sym, A)
        if not hol and d != A:
            raise Fail(f'{sym} history has no row for {A} (latest {d})')
        idx[col] = v
    bb, bbday = prev_close_0050(today)
    vals = dict(zip('BCDEFGH', fut))
    vals.update(idx)
    vals['K'] = K
    vals['BB'] = bb
    # sanity
    for k, v in vals.items():
        if not isinstance(v, (int, float)) or v < 0:
            raise Fail(f'bad value {k}={v}')
    for c in 'MNOPQ':
        if not 5 <= vals[c] <= 90:
            raise Fail(f'{c}={vals[c]} out of range')
    for c in 'BCDEFGH':
        o = old[c]
        if o and abs(vals[c] / float(o) - 1) > 0.25:
            raise Fail(f'{c}: {vals[c]} deviates >25% from previous {o}')
    if old['BB'] and abs(bb / float(old['BB']) - 1) > 0.08:
        raise Fail(f'BB {bb} deviates >8% from previous {old["BB"]}')
    if not 0 <= K <= 45:
        raise Fail(f'K={K} out of range')
    log(f'A={A} az={today} holiday={hol} K={K} {vals}')
    tmp = str(Path(a.file).with_suffix('.new.xlsx'))
    sets = [f'{k}={vals[k]}' for k in vals]
    daily_row.cmd_insert(SimpleNamespace(src=a.file, dst=tmp, a=A.isoformat(), az=today.isoformat(), holiday=hol, set=sets))
    recalc(tmp)
    expect = {'A5': A, 'AZ5': today, 'BB5': bb, 'K5': float(K)}
    expect.update({f'{k}5': float(v) for k, v in vals.items() if k not in ('K', 'BB')})
    check_recalculated(tmp, expect)
    if not a.dry_run:
        os.replace(tmp, a.file)
    else:
        log('dry-run: result kept at', tmp)
    return f'inserted {A}/{today}'


def mode_open(a, today):
    if tw_closed(today):
        raise Skip('Taiwan market closed today')
    ws = load_row5(a.file)
    if ws['AZ5'].value.date() != today:
        raise Fail(f'AZ5={ws["AZ5"].value} is not today; the close step has not produced today\'s row')
    if ws['BA5'].value not in (None, ''):
        raise Skip('BA5 already filled')
    bb = float(ws['BB5'].value)
    o = open_price_0050(today, a.wait)
    if not o > 0 or abs(o / bb - 1) > 0.08:
        raise Fail(f'open {o} deviates >8% from previous close {bb}')
    tmp = str(Path(a.file).with_suffix('.new.xlsx'))
    daily_row.cmd_set(SimpleNamespace(src=a.file, dst=tmp, expect_az=today.isoformat(), set=[f'BA={o}']))
    recalc(tmp)
    check_recalculated(tmp, {'AZ5': today, 'BA5': o})
    if not a.dry_run:
        os.replace(tmp, a.file)
    else:
        log('dry-run: result kept at', tmp)
    return f'open {o}'


def mode_selftest(a, today):
    """Fetch everything for A=prev weekday of `today` and print it (no workbook is written)."""
    A = prev_weekday(today)
    fut, m1, hol = futures_for(A)
    log('futures', fut, 'M1 expiry', m1, 'holiday', hol, 'K', (m1 - A).days - 1)
    for sym in ['VIX9D', 'VIX', 'VIX3M', 'VIX6M', 'VIX1Y']:
        log(sym, cboe_index_close(sym, A))
    log('prev close', prev_close_0050(today))
    log('open', open_price_0050(today, 0))
    log('tw_closed', tw_closed(today))
    for d in (dt.date(2026, 1, 1), dt.date(2026, 4, 3), dt.date(2026, 7, 3), dt.date(2026, 11, 26), dt.date(2026, 12, 25), dt.date(2026, 10, 16)):
        log('nyse_holiday', d, nyse_holiday(d))
    return 'selftest'


def main():
    p = argparse.ArgumentParser()
    p.add_argument('mode', choices=['close', 'open', 'selftest'])
    p.add_argument('--file', default='data/latest.xlsx')
    p.add_argument('--today')
    p.add_argument('--dry-run', action='store_true')
    p.add_argument('--wait', type=int, default=900, help='seconds to wait for the 0050 open price')
    a = p.parse_args()
    today = dt.date.fromisoformat(a.today) if a.today else dt.datetime.now(TPE).date()
    try:
        msg = {'close': mode_close, 'open': mode_open, 'selftest': mode_selftest}[a.mode](a, today)
        log('DONE:', msg)
        return 0
    except Skip as e:
        log('SKIP:', e)
        return SKIP
    except Fail as e:
        log('FAIL:', e)
        return 1


if __name__ == '__main__':
    sys.exit(main())
