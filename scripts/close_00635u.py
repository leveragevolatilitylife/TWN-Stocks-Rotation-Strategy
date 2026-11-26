#!/usr/bin/env python3
"""00635U 每日收盤 -> data/00635u_close.csv（給 Hybrid III 的 60MA 篩選用；與 data/latest.xlsx 彼此獨立）

  python3 scripts/close_00635u.py update    只抓本月與上月（每天排程用）
  python3 scripts/close_00635u.py backfill --start 2025-01

只需要收盤價；沿用 ohlc_0050.py 的 TWSE 抓取函式。CSV 欄位：date,close
"""
import argparse, csv, datetime as dt, json, sys, time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ohlc_0050 as o  # noqa: E402

OUT = Path(__file__).resolve().parent.parent / 'data' / '00635u_close.csv'


def fetch_month(y, m):
    url = f'https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date={y}{m:02d}01&stockNo=00635U'
    return [(d, c) for d, _o, _h, _l, c in o.parse_month(json.loads(o.http(url)))]


def read_csv():
    rows = {}
    if OUT.exists():
        with open(OUT, newline='', encoding='utf-8') as fh:
            for r in csv.DictReader(fh):
                rows[dt.date.fromisoformat(r['date'])] = float(r['close'])
    return rows


def run(start, end):
    rows = read_csv()
    changed = 0
    for y, m in o.months(start, end):
        got = fetch_month(y, m)
        o.log(f'{y}-{m:02d}: {len(got)} days')
        for d, c in got:
            if rows.get(d) != c:
                changed += 1
            rows[d] = c
        time.sleep(3.5)
    if not rows:
        raise SystemExit('no data')
    with open(OUT, 'w', newline='', encoding='utf-8') as fh:
        w = csv.writer(fh)
        w.writerow(['date', 'close'])
        for d in sorted(rows):
            w.writerow([d.isoformat(), f'{rows[d]:.4f}'.rstrip('0').rstrip('.')])
    o.log(f'{len(rows)} rows, {changed} changed, last {max(rows)}')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('mode', choices=['backfill', 'update'])
    ap.add_argument('--start', default='2025-01')
    a = ap.parse_args()
    today = dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).date()
    if a.mode == 'backfill':
        sy, sm = (int(x) for x in a.start.split('-'))
        run((sy, sm), (today.year, today.month))
    else:
        prev = (today.year, today.month - 1) if today.month > 1 else (today.year - 1, 12)
        run(prev, (today.year, today.month))


if __name__ == '__main__':
    main()
