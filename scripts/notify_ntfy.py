#!/usr/bin/env python3
"""每個台股交易日 09:30 把五個策略「今天」的持倉用 ntfy 推播到手機／電腦。

資料來源：data/latest.xlsx 第 5 列（AZ = 台股交易日、BA = 當日 0050 開盤價、各策略持倉欄）。
流程：非交易日直接結束 → 等第 5 列變成「今天」且開盤價已填入（最多等 --max-wait 分鐘，每分鐘 git pull）→ 推播。
等不到就照樣推播，並在內文標出哪些資料還沒更新。

環境變數：NTFY_TOPIC（必填，GitHub secret）、NTFY_SERVER（預設 https://ntfy.sh）、NTFY_TOKEN（選填，受保護主題用）。
"""
import argparse, datetime as dt, json, os, subprocess, sys, time, urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / 'data' / 'latest.xlsx'
SITE = 'https://leveragevolatilitylife.github.io/TWN-Stocks-Rotation-Strategy/'
TPE = dt.timezone(dt.timedelta(hours=8))
# 顯示順序（Hybrid 在 Hybrid II 前面）；欄位為 0 起算：持倉欄 37/39/41/43/45（1 起算）
STRATS = [('MARS+VRS Hybrid', 38), ('MARS+VRS Hybrid II', 36), ('VIX Rotation+BTD', 40), ('MA Rotation', 42), ('MA Rotation II', 44)]
COL_DATE, COL_OPEN = 51, 52
WEEK = '一二三四五六日'


def log(*a):
    print(*a, flush=True)


def read_rows():
    import openpyxl
    ws = openpyxl.load_workbook(XLSX, data_only=True, read_only=True).active
    rows = [r for r in ws.iter_rows(min_row=5, max_row=6, values_only=True)]
    return rows[0], rows[1]


def hold(r, idx):
    v = r[idx]
    return v if v in ('00631L', '00635U') else None


def build_message(today, cur, prev):
    d = cur[COL_DATE].date() if isinstance(cur[COL_DATE], dt.datetime) else None
    op = cur[COL_OPEN]
    has_open = isinstance(op, (int, float)) and op > 0
    notes, changed = [], 0
    if d != today:
        notes.append(f'⚠ 資料尚未更新到今天（目前最新是 {d}），以下持倉可能是舊的')
    elif not has_open:
        notes.append('⚠ 今天 0050 開盤價尚未填入，MA Rotation II / Hybrid II 的持倉可能還會變')
    lines = []
    for name, idx in STRATS:
        h, p = hold(cur, idx), hold(prev, idx)
        mark = ''
        if d == today and h and p and h != p:
            mark = f'（{p} → {h} 換倉）'
            changed += 1
        lines.append(f'{name}　{h or "–"}{mark}')
    body = '\n'.join(notes + lines)
    title = f'台股持倉 {today.month}/{today.day}（{WEEK[today.weekday()]}）' + (f'・{changed} 個策略換倉' if changed else '')
    return title, body, (4 if changed or notes else 3)


def send(title, body, priority):
    topic = os.environ.get('NTFY_TOPIC', '').strip()
    if not topic:
        raise SystemExit('NTFY_TOPIC is not set (add it as a GitHub Actions secret)')
    server = (os.environ.get('NTFY_SERVER') or 'https://ntfy.sh').strip().rstrip('/')   # 沒設 secret 時 workflow 會傳空字串，所以用 or 而不是 get 的預設值
    payload = {'topic': topic, 'title': title, 'message': body, 'priority': priority,
               'tags': ['chart_with_upwards_trend'], 'click': SITE}
    hdr = {'Content-Type': 'application/json'}
    if os.environ.get('NTFY_TOKEN'):
        hdr['Authorization'] = 'Bearer ' + os.environ['NTFY_TOKEN']
    req = urllib.request.Request(server, data=json.dumps(payload).encode('utf-8'), headers=hdr)
    with urllib.request.urlopen(req, timeout=30) as r:
        log('ntfy:', r.status, r.read().decode()[:120])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--max-wait', type=int, default=25, help='等資料更新的分鐘數')
    ap.add_argument('--force', action='store_true', help='忽略非交易日與等待，立刻用目前資料推播（測試用）')
    a = ap.parse_args()
    today = dt.datetime.now(TPE).date()
    if not a.force:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        import daily_update
        kind = daily_update.tw_day_kind(today)
        if kind != 'open':
            log(f'{today} is {kind}, nothing to send')
            return
    deadline = time.time() + a.max_wait * 60
    while True:
        cur, prev = read_rows()
        d = cur[COL_DATE].date() if isinstance(cur[COL_DATE], dt.datetime) else None
        op = cur[COL_OPEN]
        ready = d == today and isinstance(op, (int, float)) and op > 0
        if ready or a.force or time.time() >= deadline:
            break
        log(f'waiting for today\'s data (row5 date={d}, open={op})')
        time.sleep(60)
        subprocess.run(['git', 'pull', '-q', '--ff-only', 'origin', 'main'], cwd=ROOT, check=False)
    title, body, pr = build_message(today, cur, prev)
    log(title); log(body)
    send(title, body, pr)


if __name__ == '__main__':
    main()
