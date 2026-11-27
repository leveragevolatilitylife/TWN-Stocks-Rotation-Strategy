#!/usr/bin/env python3
"""00635U / 00682U 每日 OHLC -> data/00635u_ohlc.csv、data/00682u_ohlc.csv
（給網站「ETF 60MA 與 BIAS」面板的 K 棒，以及 Hybrid III 的 60MA 防禦判斷使用；與 data/latest.xlsx 彼此獨立）

  python3 scripts/ohlc_etf.py update                    只抓本月與上月（每天排程用）
  python3 scripts/ohlc_etf.py backfill [--start YYYY-MM] [--stock 00635U]   逐月補歷史（每檔約數分鐘）

沿用 ohlc_0050.py 的 TWSE 抓取／拆股調整函式。價格為 TWSE 原始價（未還原息值）。CSV 欄位：date,open,high,low,close
"""
import argparse, csv, datetime as dt, json, sys, time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ohlc_0050 as o  # noqa: E402

DATA = Path(__file__).resolve().parent.parent / 'data'
STOCKS = {'00635U': '2014-04', '00682U': '2017-06'}   # 代號 -> 預設回補起始月（約上市月份）


def path(stock):
    return DATA / f'{stock.lower()}_ohlc.csv'


def read_csv(stock):
    rows = {}
    if path(stock).exists():
        with open(path(stock), newline='', encoding='utf-8') as fh:
            for r in csv.DictReader(fh):
                rows[dt.date.fromisoformat(r['date'])] = (float(r['open']), float(r['high']), float(r['low']), float(r['close']))
    return rows


def write_csv(stock, rows):
    with open(path(stock), 'w', newline='', encoding='utf-8') as fh:
        w = csv.writer(fh)
        w.writerow(['date', 'open', 'high', 'low', 'close'])
        for d in sorted(rows):
            w.writerow([d.isoformat()] + [f'{v:.4f}'.rstrip('0').rstrip('.') for v in rows[d]])


def fetch_month(stock, y, m):
    url = f'https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date={y}{m:02d}01&stockNo={stock}'
    return o.parse_month(json.loads(o.http(url)))


def run(stock, start, end):
    rows = read_csv(stock)
    changed = 0
    for y, m in o.months(start, end):
        got = fetch_month(stock, y, m)
        o.log(f'{stock} {y}-{m:02d}: {len(got)} days')
        for d, op, h, l, c in got:
            if rows.get(d) != (op, h, l, c):
                changed += 1
            rows[d] = (op, h, l, c)
        time.sleep(3.5)
    if not rows:
        raise SystemExit(f'{stock}: no data')
    rows = o.adjust_splits(rows)
    write_csv(stock, rows)
    o.log(f'{stock}: {len(rows)} rows, {changed} changed, last {max(rows)}')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('mode', choices=['backfill', 'update'])
    ap.add_argument('--start', default=None)
    ap.add_argument('--stock', default=None, choices=list(STOCKS))
    a = ap.parse_args()
    today = dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).date()
    ok = True
    for stock in ([a.stock] if a.stock else STOCKS):
        try:
            if a.mode == 'backfill':
                sy, sm = (int(x) for x in (a.start or STOCKS[stock]).split('-'))
                run(stock, (sy, sm), (today.year, today.month))
            else:
                prev = (today.year, today.month - 1) if today.month > 1 else (today.year - 1, 12)
                run(stock, prev, (today.year, today.month))
        except (SystemExit, Exception) as e:  # noqa: BLE001
            o.log(f'{stock} failed: {e}')
            ok = False
    sys.exit(0 if ok else 1)


if __name__ == '__main__':
    main()
