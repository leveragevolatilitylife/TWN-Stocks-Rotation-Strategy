#!/usr/bin/env python3
"""0050 每日 OHLC -> data/0050_ohlc.csv（給網站畫 K 棒；與 data/latest.xlsx 彼此獨立）

  python3 scripts/ohlc_0050.py backfill [--start 2015-01]   從 TWSE 逐月抓歷史（約 10 分鐘），已存在的日期會被覆蓋
  python3 scripts/ohlc_0050.py update                       只抓本月與上月（每天排程用）
  python3 scripts/ohlc_0050.py validate [--xlsx data/latest.xlsx]   與 Excel 的開盤價 / 前日收盤價逐日比對

價格一律還原到「現在的股數基準」：0050 在 2025 年 6 月 1 拆 4，TWSE 歷史資料是未還原的原始價，
所以程式偵測相鄰兩天收盤價暴跌到 1/2 以下（單日漲跌幅上限 10%，只可能是拆股）就把更早的價格同除以拆股比例，
與 Excel 裡既有的 0050 價格基準一致。
CSV 欄位：date,open,high,low,close
"""
import argparse, csv, datetime as dt, json, sys, time, urllib.request
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / 'data' / '0050_ohlc.csv'
UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36'


def log(*a):
    print(*a, flush=True)


def http(url, tries=4, timeout=45):
    last = None
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={'User-Agent': UA, 'Accept-Language': 'zh-TW,zh;q=0.9'})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read().decode('utf-8-sig', 'replace')
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(3 * (i + 1))
    raise RuntimeError(f'GET failed: {url}: {last}')


def fnum(s):
    s = str(s).replace(',', '').strip()
    if s in ('', '--', '-', 'X0.00'):
        return None
    return float(s)


def parse_month(j):
    """TWSE STOCK_DAY json -> [(date, o, h, l, c)]（沒有成交的日子略過）"""
    if not isinstance(j, dict) or j.get('stat') != 'OK':
        return []
    f = j.get('fields', [])
    try:
        i_o, i_h, i_l, i_c = (f.index(k) for k in ('開盤價', '最高價', '最低價', '收盤價'))
    except ValueError:
        raise RuntimeError(f'unexpected TWSE fields: {f}')
    out = []
    for r in j.get('data', []):
        y, m, d = (int(x) for x in str(r[0]).split('/'))
        o, h, l, c = (fnum(r[i]) for i in (i_o, i_h, i_l, i_c))
        if None in (o, h, l, c):
            continue
        out.append((dt.date(y + 1911, m, d), o, h, l, c))
    return out


def fetch_month(y, m):
    url = f'https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date={y}{m:02d}01&stockNo=0050'
    return parse_month(json.loads(http(url)))


def read_csv():
    rows = {}
    if OUT.exists():
        with open(OUT, newline='', encoding='utf-8') as fh:
            for r in csv.DictReader(fh):
                rows[dt.date.fromisoformat(r['date'])] = (float(r['open']), float(r['high']), float(r['low']), float(r['close']))
    return rows


def write_csv(rows):
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with open(OUT, 'w', newline='', encoding='utf-8') as fh:
        w = csv.writer(fh)
        w.writerow(['date', 'open', 'high', 'low', 'close'])
        for d in sorted(rows):
            w.writerow([d.isoformat()] + [f'{v:.4f}'.rstrip('0').rstrip('.') for v in rows[d]])


def adjust_splits(rows):
    """相鄰兩天收盤價掉到 1/2 以下 -> 視為拆股，更早的價格同除以拆股比例"""
    ds = sorted(rows)
    for k in range(len(ds) - 1, 0, -1):
        c0, c1 = rows[ds[k - 1]][3], rows[ds[k]][3]
        if c1 > 0 and c0 / c1 > 2:
            ratio = round(c0 / c1)
            log(f'split detected between {ds[k-1]} and {ds[k]}: x1/{ratio}')
            for d in ds[:k]:
                rows[d] = tuple(round(v / ratio, 4) for v in rows[d])
    return rows


def months(start, end):
    y, m = start
    while (y, m) <= end:
        yield y, m
        m += 1
        if m > 12:
            y, m = y + 1, 1


def run(start, end, label):
    rows = read_csv()
    n_new = 0
    for y, m in months(start, end):
        got = fetch_month(y, m)
        log(f'{y}-{m:02d}: {len(got)} days')
        for d, o, h, l, c in got:
            if rows.get(d) != (o, h, l, c):
                n_new += 1
            rows[d] = (o, h, l, c)
        time.sleep(3.5)   # TWSE 會擋太密集的請求
    if not rows:
        raise SystemExit('no data')
    rows = adjust_splits(rows)
    write_csv(rows)
    log(f'{label}: {len(rows)} rows total, {n_new} fetched/changed, last {max(rows)}')


def validate(xlsx):
    import openpyxl
    rows = read_csv()
    ds = sorted(rows)
    ws = openpyxl.load_workbook(xlsx, data_only=True, read_only=True).active
    n = bad_c = bad_o = miss = 0
    worst = []
    for r in ws.iter_rows(min_row=5, values_only=True):
        d, op, pc = r[51], r[52], r[53]     # AZ（台股交易日）、BA（當日開盤）、BB（前日收盤）
        if not isinstance(d, dt.datetime) or not isinstance(pc, (int, float)):
            continue
        d = d.date()
        prev = [x for x in ds if x < d]
        if not prev:
            continue
        n += 1
        c = rows[prev[-1]][3]
        if abs(c - pc) > 0.02:
            bad_c += 1
            worst.append((d.isoformat(), 'prev_close', pc, c))
        if isinstance(op, (int, float)) and op > 0:
            if d in rows:
                if abs(rows[d][0] - op) > 0.02:
                    bad_o += 1
                    worst.append((d.isoformat(), 'open', op, rows[d][0]))
            else:
                miss += 1
    log(f'validate: {n} rows compared, prev_close mismatch {bad_c}, open mismatch {bad_o}, days missing in csv {miss}')
    for w in worst[:25]:
        log('  ', w)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('mode', choices=['backfill', 'update', 'validate'])
    ap.add_argument('--start', default='2015-01')
    ap.add_argument('--xlsx', default=str(OUT.parent / 'latest.xlsx'))
    a = ap.parse_args()
    today = dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).date()
    if a.mode == 'validate':
        return validate(a.xlsx)
    if a.mode == 'backfill':
        sy, sm = (int(x) for x in a.start.split('-'))
        run((sy, sm), (today.year, today.month), 'backfill')
    else:
        prev = (today.year, today.month - 1) if today.month > 1 else (today.year - 1, 12)
        run(prev, (today.year, today.month), 'update')


if __name__ == '__main__':
    main()
