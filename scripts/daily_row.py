#!/usr/bin/env python3
"""Daily updater for the VIX term-structure workbook (sheet 'VIX期貨數據').

  python3 -I daily_row.py insert SRC DST --a 2026-10-19 --az 2026-10-20 [--holiday] --set B=19.908 --set C=...
      Inserts one new row at row 5 (Excel "insert row" semantics), builds the new row's formulas from the
      previous row 5, re-anchors the header statistics (rows 1-4) so they include the new row, optionally
      paints the A-column date orange (US market holiday) and writes the given input values.
  python3 -I daily_row.py set SRC DST --expect-az 2026-10-20 --set BA=107.45
      Writes input values into row 5 (refuses if AZ5 is not the expected date or the cell holds a formula).

Only the XML of the sheet is edited, so styles / conditional formats / shared formulas are preserved.
"""
import re, sys, zipfile, argparse, datetime as dt
import openpyxl

INS = 5
REF = re.compile(r'(?<![A-Za-z0-9_.])(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?(?![A-Za-z0-9_(])')
PART = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')
SHEET = 'xl/worksheets/sheet1.xml'


def sh(r): return r + 1 if r >= INS else r


def part(tok, fn):
    cd, col, rd, row = PART.fullmatch(tok).groups()
    return f'{cd}{col}{rd}{fn(int(row), bool(rd))}'


def xform(f, mode):
    segs = re.split(r'("(?:[^"]|"")*")', f)
    for i in range(0, len(segs), 2):
        def rep(m):
            a, b = m.group(1), m.group(2)
            if b is None:
                if mode == 'top':
                    row = int(PART.fullmatch(a).group(4))
                    if row in (5, 6):
                        return a            # "latest" / "previous" pointers stay on row 5 / 6
                    assert row < INS, (f, a)
                return part(a, lambda r, ab: sh(r))
            if mode == 'top':
                s = part(a, lambda r, ab: r if r == INS else sh(r))   # range start at data top stays
            else:
                s = part(a, lambda r, ab: sh(r))
            return s + ':' + part(b, lambda r, ab: sh(r))
        segs[i] = REF.sub(rep, segs[i])
    return ''.join(segs)


def abs_bump(f):
    segs = re.split(r'("(?:[^"]|"")*")', f)
    for i in range(0, len(segs), 2):
        def one(tok): return part(tok, lambda r, ab: r + 1 if (ab and r >= INS) else r)
        segs[i] = REF.sub(lambda m: one(m.group(1)) + ((':' + one(m.group(2))) if m.group(2) else ''), segs[i])
    return ''.join(segs)


def xesc(s): return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
def xunesc(s): return s.replace('&lt;', '<').replace('&gt;', '>').replace('&quot;', '"').replace('&apos;', "'").replace('&amp;', '&')


def serial(d): return (d - dt.date(1899, 12, 30)).days


def parse_sets(items):
    out = {}
    for it in items or []:
        k, v = it.split('=', 1)
        float(v)
        out[k.strip().upper()] = v.strip()
    return out


FRX = re.compile(r'(<f(?:\s[^>]*?)?(?<!/)>)([^<]*)(</f>)')
REFATTR = re.compile(r'ref="([A-Z]+\d+(?::[A-Z]+\d+)?)"')
CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(/>|>(.*?)</c>)', re.S)


def strip_cached_rows(s, rows=(1, 4)):
    """drop stale cached values of formula cells in the header rows (Excel recalculates on open)"""
    def f(m):
        if int(m.group(2)) not in rows or '<f' not in (m.group(5) or ''):
            return m.group(0)
        inner = re.sub(r'<v>[^<]*</v>', '', m.group(5))
        attrs = re.sub(r' t="(e|str|b)"', '', m.group(3))
        return f'<c r="{m.group(1)}{m.group(2)}"{attrs}>{inner}</c>'
    return CELL.sub(f, s)


def write_zip(src_zip, dst, new_sheet):
    z = src_zip
    zo = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED, compresslevel=6)
    for it in z.infolist():
        n = it.filename
        if n == 'xl/calcChain.xml':
            continue
        d = z.read(n)
        if n == SHEET:
            d = new_sheet.encode('utf8')
        elif n == '[Content_Types].xml':
            d = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', '', d.decode()).encode()
        elif n == 'xl/_rels/workbook.xml.rels':
            d = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', '', d.decode()).encode()
        elif n == 'xl/workbook.xml':
            t = d.decode()
            if 'fullCalcOnLoad' not in t:
                t = re.sub(r'<calcPr ', '<calcPr fullCalcOnLoad="1" ', t, count=1)
            d = t.encode()
        zo.writestr(it, d)
    zo.close()


def cmd_insert(a):
    A = dt.date.fromisoformat(a.a); AZ = dt.date.fromisoformat(a.az)
    sets = parse_sets(a.set)
    wf = openpyxl.load_workbook(a.src).active
    assert wf['A5'].value.date() < A, f"A5={wf['A5'].value} is not older than the new date {A}"
    assert wf['A5'].value is not None and wf['AZ5'].value is not None
    new_f = {}
    for c in wf[5]:
        v = c.value
        if isinstance(v, str) and v.startswith('='):
            new_f[c.column_letter] = abs_bump(v[1:])
    orange_row = None
    for r in range(6, 3000):
        fl = wf.cell(r, 1).fill
        if fl.fill_type == 'solid' and fl.fgColor.type == 'rgb' and fl.fgColor.rgb == 'FFFFC000':
            orange_row = r; break

    z = zipfile.ZipFile(a.src)
    s = z.read(SHEET).decode('utf8')
    head, rest = s.split('<sheetData>', 1)
    body, tail = rest.split('</sheetData>', 1)

    orange_sid = None
    if a.holiday:
        assert orange_row, 'no orange (FFC000) date cell found to copy the style from'
        orange_sid = re.search(r'<c r="A%d"(?: s="(\d+)")?' % orange_row, body).group(1)

    out = []; row6 = None
    for m in re.finditer(r'<row r="(\d+)"(.*?)</row>', body, re.S):
        r = int(m.group(1)); txt = m.group(0)
        if r < INS:
            def fr(mm):
                ot = REFATTR.sub(lambda q: 'ref="' + xform(q.group(1), 'data') + '"', mm.group(1))
                return ot + xesc(xform(xunesc(mm.group(2)), 'top')) + mm.group(3)
            out.append(FRX.sub(fr, txt)); continue
        nr = r + 1
        txt = txt.replace(f'<row r="{r}"', f'<row r="{nr}"', 1)
        txt = re.sub(r'<c r="([A-Z]+)%d"' % r, lambda mm: f'<c r="{mm.group(1)}{nr}"', txt)
        def fr2(mm):
            ot = REFATTR.sub(lambda q: 'ref="' + xform(q.group(1), 'data') + '"', mm.group(1))
            return ot + xesc(xform(xunesc(mm.group(2)), 'data')) + mm.group(3)
        txt = FRX.sub(fr2, txt)
        if r == INS: row6 = txt
        out.append(txt)
    rattr = re.match(r'<row r="6"([^>]*)>', row6).group(1)
    cells = re.findall(r'<c r="([A-Z]+)6"(?: s="(\d+)")?', row6)
    values = dict(sets)
    values['A'] = str(serial(A)); values['AZ'] = str(serial(AZ))
    nc = []
    for col, sid in cells:
        if col == 'A' and orange_sid: sid = orange_sid
        sattr = f' s="{sid}"' if sid else ''
        if col in new_f:
            nc.append(f'<c r="{col}5"{sattr}><f>{xesc(new_f[col])}</f></c>')
        elif col in values:
            nc.append(f'<c r="{col}5"{sattr}><v>{values[col]}</v></c>')
        else:
            nc.append(f'<c r="{col}5"{sattr}/>')
    bad = [k for k in values if k not in {c for c, _ in cells} or k in new_f]
    assert not bad, f'cannot set {bad} (missing cell or formula cell)'
    newrow = f'<row r="5"{rattr}>' + ''.join(nc) + '</row>'
    idx = [i for i, t in enumerate(out) if t.startswith('<row r="6"')][0]
    out.insert(idx, newrow)

    def sq(rng):
        parts = []
        for tok in rng.split(' '):
            if ':' in tok:
                x, y = tok.split(':')
                parts.append(part(x, lambda r, ab: r if r == INS else sh(r)) + ':' + part(y, lambda r, ab: sh(r)))
            else:
                parts.append(part(tok, lambda r, ab: sh(r)))
        return ' '.join(parts)
    tail = re.sub(r'(<conditionalFormatting sqref=")([^"]+)(")', lambda m: m.group(1) + sq(m.group(2)) + m.group(3), tail)
    tail = re.sub(r'(<formula>)(.*?)(</formula>)', lambda m: m.group(1) + xesc(xform(xunesc(m.group(2)), 'data')) + m.group(3), tail)
    def dim(m):
        c1, r1, c2, r2 = re.fullmatch(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', m.group(1)).groups()
        return f'<dimension ref="{c1}{r1}:{c2}{int(r2) + 1}"/>'
    head = re.sub(r'<dimension ref="([^"]+)"/>', dim, head, count=1)
    news = strip_cached_rows(head + '<sheetData>' + ''.join(out) + '</sheetData>' + tail)
    write_zip(z, a.dst, news)
    print('inserted row 5:', A, '/', AZ, 'holiday' if a.holiday else '', sets)


def cmd_set(a):
    sets = parse_sets(a.set)
    z = zipfile.ZipFile(a.src)
    s = z.read(SHEET).decode('utf8')
    wf = openpyxl.load_workbook(a.src).active
    if a.expect_az:
        assert wf['AZ5'].value.date() == dt.date.fromisoformat(a.expect_az), f"AZ5={wf['AZ5'].value}, expected {a.expect_az}"
    for col, v in sets.items():
        pat = re.compile(r'<c r="%s5"( s="\d+")?(?:/>|>(.*?)</c>)' % col, re.S)
        m = pat.search(s)
        assert m, f'no cell {col}5'
        assert not (m.group(2) and '<f' in m.group(2)), f'{col}5 holds a formula'
        s = s[:m.start()] + f'<c r="{col}5"{m.group(1) or ""}><v>{v}</v></c>' + s[m.end():]
    write_zip(z, a.dst, strip_cached_rows(s))
    print('set row 5:', sets)


if __name__ == '__main__':
    p = argparse.ArgumentParser()
    sp = p.add_subparsers(dest='cmd', required=True)
    i = sp.add_parser('insert'); i.add_argument('src'); i.add_argument('dst')
    i.add_argument('--a', required=True); i.add_argument('--az', required=True)
    i.add_argument('--holiday', action='store_true'); i.add_argument('--set', action='append')
    t = sp.add_parser('set'); t.add_argument('src'); t.add_argument('dst')
    t.add_argument('--expect-az'); t.add_argument('--set', action='append', required=True)
    args = p.parse_args()
    {'insert': cmd_insert, 'set': cmd_set}[args.cmd](args)
