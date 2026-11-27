#!/usr/bin/env python3
"""把 FED 利率、美國 PCE、00635U / 00682U 的開盤價／前日收盤／60MA／乖離率寫進 data/latest.xlsx（sheet 'VIX期貨數據' 的 BQ:BZ 欄）。

  python3 -I scripts/etf_cols.py sync [FILE] [--fed fed_rates.json] [--pce us_pce.json]

欄位（以 AZ 欄的台股交易日 N 對齊每一列）：
  BQ FED 利率（目標區間中點，決議日之後的下一個台股交易日起套用；第一筆決議之前留白）
  BR 美國 PCE 年增率（%，可取得日 ≤ N 的最新值）
  BS/BT/BU/BV   00635U：開盤價(N)、收盤價(N-1)、60MA（基於前日收盤）、乖離率（基於前日收盤）
  BW/BX/BY/BZ   00682U：同上
價格來自 data/00635u_ohlc.csv、data/00682u_ohlc.csv。休市日（無 K 棒）開盤價以前日收盤代入，與 0050 的 BA 欄做法相同。
60MA／乖離率寫成與 BC／BE 欄相同形式的 Excel 公式（往下取 60 列，不足 60 筆留白）並同步寫入快取值，因此不必重新計算即可被網頁讀取。
只改 sheet1.xml 的這幾欄（樣式、其他公式、條件格式都不動）；可重複執行。FED／PCE 讀不到時略過該兩欄，其餘照常。
"""
import argparse, bisect, csv, datetime as dt, json, re, shutil, sys, urllib.request, zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SHEET = 'xl/worksheets/sheet1.xml'
RAW = 'https://raw.githubusercontent.com/leveragevolatilitylife/TWN-rotation-strategy-backtester/main/src/data/'
FIRST = 5
# 欄位定義
COLS = ['BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ']
HEAD = {
    'BQ': 'FED 利率\n(目標區間中點)', 'BR': '美國 PCE\n年增率(%)',
    'BS': '00635U\nopen price (N)', 'BT': '00635U\nclose price (N-1)', 'BU': '00635U\n60MA\n(基於前日收盤)', 'BV': '00635U\n60MA乖離\n(基於前日收盤)',
    'BW': '00682U\nopen price (N)', 'BX': '00682U\nclose price (N-1)', 'BY': '00682U\n60MA\n(基於前日收盤)', 'BZ': '00682U\n60MA乖離\n(基於前日收盤)',
}
ETFS = [('00635U', 'BS', 'BT', 'BU', 'BV'), ('00682U', 'BW', 'BX', 'BY', 'BZ')]
S_PRICE, S_NUM, S_PCT, S_HEAD, S_DASH = 86, 87, 123, 46, 58


def log(*a):
    print(*a, flush=True)


def fetch_json(name, local):
    if local:
        return json.loads(Path(local).read_text(encoding='utf-8'))
    req = urllib.request.Request(RAW + name, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode('utf-8'))


def load_ohlc(sym):
    p = ROOT / 'data' / f'{sym.lower()}_ohlc.csv'
    rows = []
    if p.exists():
        with open(p, newline='', encoding='utf-8') as fh:
            for r in csv.DictReader(fh):
                rows.append((r['date'], float(r['open']), float(r['close'])))
    rows.sort()
    return rows


def serial_to_iso(v):
    return (dt.date(1899, 12, 30) + dt.timedelta(days=int(v))).isoformat()


def num(x):
    return repr(round(float(x), 10))


def cell_num(ref, s, v):
    return f'<c r="{ref}" s="{s}" t="n"><v>{num(v)}</v></c>'


def cell_formula(ref, s, f, v):
    f = f.replace('&', '&amp;').replace('"', '&quot;').replace('<', '&lt;').replace('>', '&gt;')
    if v is None:
        return f'<c r="{ref}" s="{s}" t="str"><f aca="false">{f}</f><v></v></c>'
    return f'<c r="{ref}" s="{s}" t="n"><f aca="false">{f}</f><v>{num(v)}</v></c>'


def cell_blank(ref, s):
    return f'<c r="{ref}" s="{s}"/>'


def colnum(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


CELL = re.compile(r'<c r="([A-Z]+)(\d+)"[^>]*?(?:/>|>.*?</c>)', re.S)


def put_cells(row_xml, new):
    """把 new:{col: cell_xml} 放進 <row> 內，取代同欄既有儲存格並保持欄位順序"""
    m = re.match(r'(<row [^>]*>)(.*)(</row>)$', row_xml, re.S)
    head, body, tail = m.group(1), m.group(2), m.group(3)
    head = re.sub(r'\sspans="[^"]*"', '', head)
    cells = {}
    for cm in CELL.finditer(body):
        cells[cm.group(1)] = cm.group(0)
    cells.update(new)
    return head + ''.join(cells[c] for c in sorted(cells, key=colnum)) + tail


def sync(path, fed_path=None, pce_path=None):
    path = str(path)
    with zipfile.ZipFile(path) as z:
        xml = z.read(SHEET).decode('utf-8')
    # 取各列的 AZ（台股交易日）
    rows = []
    for rm in re.finditer(r'<row r="(\d+)"[^>]*>.*?</row>', xml, re.S):
        r = int(rm.group(1))
        if r < FIRST:
            continue
        am = re.search(r'<c r="AZ%d"[^>]*?>\s*<v>([\d.]+)</v>' % r, rm.group(0))
        if am:
            rows.append((r, serial_to_iso(float(am.group(1)))))
    if not rows:
        raise SystemExit('no data rows found')
    az = dict(rows)

    # FED / PCE
    fed = pce = None
    try:
        f = fetch_json('fed_rates.json', fed_path)
        p = fetch_json('us_pce.json', pce_path)
        fed = [(d, float(nm)) for d, nm in f['decisions']]
        pce = [(d, float(v)) for d, v in p['releases']]
    except Exception as e:  # noqa: BLE001
        log('WARN FED/PCE data unavailable, those two columns are left unchanged:', e)
    fed_d = [d for d, _ in fed] if fed else []
    pce_d = [d for d, _ in pce] if pce else []

    vals = {}   # (row) -> {col: number|None}
    for r, d in rows:
        v = {}
        if fed is not None:
            i = bisect.bisect_left(fed_d, d) - 1          # 決議日當天不算，下一個台股交易日起套用
            v['BQ'] = fed[i][1] if i >= 0 else None
            j = bisect.bisect_right(pce_d, d) - 1
            v['BR'] = pce[j][1] if j >= 0 else None
        vals[r] = v
    for sym, c_open, c_pc, c_ma, c_bias in ETFS:
        px = load_ohlc(sym)
        dates = [x[0] for x in px]
        for r, d in rows:
            k = bisect.bisect_left(dates, d) - 1          # 前一個有成交的交易日
            pc = px[k][2] if k >= 0 else None
            today_i = bisect.bisect_left(dates, d)
            if today_i < len(dates) and dates[today_i] == d:
                op = px[today_i][1]
            elif pc is not None and (not dates or d <= dates[-1]):
                op = pc                                   # 當天休市：開盤以前日收盤代入
            else:
                op = None                                 # 尚未有當天資料
            vals[r][c_open], vals[r][c_pc] = op, pc
        last = rows[-1][0]
        for r, _ in rows:
            win = [vals.get(x, {}).get(c_pc) for x in range(r, r + 60)] if r + 59 <= last else None
            ma = sum(win) / 60 if win and all(w is not None for w in win) else None
            vals[r][c_ma] = ma
            pc = vals[r][c_pc]
            vals[r][c_bias] = (pc - ma) / ma if (ma is not None and pc is not None) else None

    # 組 XML
    out_rows = {}
    for rm in re.finditer(r'<row r="(\d+)"[^>]*>.*?</row>', xml, re.S):
        r = int(rm.group(1))
        if r == 3:
            new = {c: f'<c r="{c}3" s="{S_HEAD}" t="inlineStr"><is><t xml:space="preserve">{HEAD[c].replace(chr(10), "&#10;")}</t></is></c>' for c in COLS}
        elif r == 4:
            new = {c: f'<c r="{c}4" s="{S_DASH}" t="s"><v>2</v></c>' for c in COLS}
        elif r in vals:
            v, new = vals[r], {}
            for c in COLS:
                ref = f'{c}{r}'
                if c in ('BQ', 'BR') and fed is None:
                    continue                                # 讀不到就保留原狀
                if c in ('BQ', 'BR'):
                    new[c] = cell_num(ref, S_NUM, v[c]) if v.get(c) is not None else cell_blank(ref, S_NUM)
                elif c in ('BS', 'BW'):
                    new[c] = cell_num(ref, S_PRICE, v[c]) if v.get(c) is not None else cell_blank(ref, S_PRICE)
                elif c in ('BT', 'BX'):
                    new[c] = cell_num(ref, S_NUM, v[c]) if v.get(c) is not None else cell_blank(ref, S_NUM)
                elif c in ('BU', 'BY'):
                    pc = {'BU': 'BT', 'BY': 'BX'}[c]
                    new[c] = cell_formula(ref, S_NUM, f'IF(COUNT({pc}{r}:{pc}{r+59})<60,"",AVERAGE({pc}{r}:{pc}{r+59}))', v[c])
                else:
                    pc, ma = {'BV': ('BT', 'BU'), 'BZ': ('BX', 'BY')}[c]
                    new[c] = cell_formula(ref, S_PCT, f'IF({ma}{r}="","",({pc}{r}-{ma}{r})/{ma}{r})', v[c])
            if not new:
                continue
        else:
            continue
        out_rows[r] = put_cells(rm.group(0), new)
    # 取代
    def repl(m):
        r = int(m.group(1))
        return out_rows.get(r, m.group(0))
    xml2 = re.sub(r'<row r="(\d+)"[^>]*>.*?</row>', repl, xml, flags=re.S)
    xml2 = re.sub(r'<dimension ref="A1:[A-Z]+(\d+)"/>', lambda m: f'<dimension ref="A1:BZ{m.group(1)}"/>', xml2, count=1)
    if xml2 == xml:
        log('etf_cols: no change')
        return
    tmp = path + '.etf.tmp'
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == SHEET:
                data = xml2.encode('utf-8')
            zout.writestr(item, data)
    shutil.move(tmp, path)
    last = rows[0]
    log(f'etf_cols: {len(rows)} rows synced; newest row {last[0]} az={last[1]} -> ' + ', '.join(f'{c}={vals[last[0]].get(c)}' for c in COLS))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('mode', choices=['sync'])
    ap.add_argument('file', nargs='?', default=str(ROOT / 'data' / 'latest.xlsx'))
    ap.add_argument('--fed'); ap.add_argument('--pce')
    a = ap.parse_args()
    sync(a.file, a.fed, a.pce)


if __name__ == '__main__':
    main()
