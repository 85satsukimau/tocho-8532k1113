#!/usr/bin/env python3
"""刀帳アプリのデータ更新スクリプト

Excel（刀剣乱舞データベース.xlsx）を読み、HTML の中のデータ部分
（<script type="application/json" id="m"> と id="u"）だけを書き換えます。
HTML の他の部分（デザインや機能）は一切変更しません。

使い方（Excel は保存して閉じてから）:
    pip install openpyxl            # 最初の1回だけ
    python update_data.py 刀剣乱舞データベース.xlsx touken_app.html

書き換え前の HTML は「touken_app.html.bak」に1世代だけ残します。
注意: Excel の列の並びを変えた（列の挿入・削除をした）場合は、このスクリプトの列番号も直す必要があります。
"""
import sys, re, json, shutil, datetime, warnings
warnings.filterwarnings('ignore')
try:
    import openpyxl
except ImportError:
    sys.exit('openpyxl が入っていません。  pip install openpyxl  を実行してください。')

if len(sys.argv) < 2:
    sys.exit('使い方: python update_data.py <Excelファイル> [HTMLファイル(省略時 touken_app.html)]')
XLSX = sys.argv[1]
HTML = sys.argv[2] if len(sys.argv) > 2 else 'touken_app.html'

try:
    wb = openpyxl.load_workbook(XLSX, data_only=True)   # data_only: Excel が保存した計算結果を読む
except PermissionError:
    sys.exit('Excel を開いたままのようです。保存して閉じてから、もう一度実行してください。')

def dt(v):
    if isinstance(v, datetime.datetime): return v.date().isoformat()
    if isinstance(v, datetime.date): return v.isoformat()
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 20000 < v < 60000:
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).isoformat()
    return None
norm = lambda s: s.replace('\u3000', ' ').strip() if isinstance(s, str) else s
def num(v):
    try: return int(float(v))
    except Exception: return None
def cv(v):
    if v is None or v == '': return None
    if isinstance(v, float) and v == int(v): return int(v)
    return v if isinstance(v, (int, str)) else str(v)

# ---------------- 刀剣男士 ----------------
w = wb['刀剣男士']; R = range(3, w.max_row + 1)
stats = lambda r: [[cv(w.cell(r, c).value) for c in range(a, a + 9)] for a in (17, 26, 35, 44)]
hdr = {c: w.cell(2, c).value for c in range(73, 83)}
rows = {norm(w.cell(r, 6).value): r for r in R if w.cell(r, 6).value}

k = wb['育成記録']; gr_src = {}
for r in range(2, k.max_row + 1):
    n = norm(k.cell(r, 2).value)
    if n:
        iv = k.cell(r, 9).value
        gr_src[n] = [num(k.cell(r, 5).value), num(k.cell(r, 7).value), num(k.cell(r, 8).value),
                     1 if k.cell(r, 6).value == '極' else 0, num(k.cell(r, 10).value),
                     (num(iv) if num(iv) is not None else (str(iv) if iv else None)),
                     [cv(k.cell(r, c).value) for c in range(13, 22)]]

sw = []; own = []; got = {}; gr = {}; tg = {}; kgot = {}; cfm = {}
tk = re.compile(r'^(.+)・特([一二三四五六七八九十]+)$')
for r in R:
    if w.cell(r, 9).value != '初' or not w.cell(r, 6).value: continue
    nn = norm(w.cell(r, 6).value)
    if tk.match(nn): continue                      # 髭切・特一 などは「段階」として扱う
    s = dict(n=nn, no=w.cell(r, 5).value, t=w.cell(r, 10).value, s=w.cell(r, 11).value,
             r=w.cell(r, 12).value, sl=w.cell(r, 13).value, d=dt(w.cell(r, 7).value), st=stats(r))
    s['get'] = {str(hdr[c]): str(w.cell(r, c).value) for c in range(73, 83)
                if w.cell(r, c).value not in (None, False, '') and hdr[c]}
    fm = []; prevN = cv(w.cell(r, 14).value); lastO = cv(w.cell(r, 15).value); rr = r + 1
    while rr <= w.max_row:
        m = tk.match(norm(w.cell(rr, 6).value or ''))
        if not m or m.group(1) != nn: break
        fm.append(dict(k='特' + m.group(2), lv=prevN, st=stats(rr)))
        if w.cell(rr, 1).value is True: cfm[nn] = '特' + m.group(2)
        prevN = cv(w.cell(rr, 14).value); lastO = cv(w.cell(rr, 15).value) or lastO; rr += 1
    if fm: s['fm'] = fm
    s['tl'] = [None if fm else cv(w.cell(r, 14).value), lastO]
    kr = rows.get(nn + '・極')
    if kr and w.cell(kr, 9).value == '極':
        s['kx'] = dict(d=dt(w.cell(kr, 7).value), st=stats(kr), u=1 if w.cell(kr, 4).value == '未' else 0)
    sw.append(s)
    if kr and dt(w.cell(kr, 8).value): kgot[nn] = dt(w.cell(kr, 8).value)
    if w.cell(r, 1).value is True: own.append(nn)
    if dt(w.cell(r, 8).value): got[nn] = dt(w.cell(r, 8).value)
    g = gr_src.get(nn) or gr_src.get(nn + '・極')
    if g: gr[nn] = g
    t = [lab for c, lab in ((60, '軽装'), (61, '祝装'), (62, '近侍曲')) if w.cell(r, c).value not in (None, False, '')]
    if t: tg[nn] = t

# ---------------- 年表 ----------------
y = wb['年表']; ev = []
def add(l, a, b, kind, text):
    a = dt(a); b = dt(b)
    if a and text: ev.append([l, a, b, kind or '', str(text)])
for r in range(2, y.max_row + 1):
    g = lambda c: y[f'{c}{r}'].value
    if g('K') and (g('H') or g('I')): add(0, g('K'), g('M'), '', ' '.join(str(x) for x in (g('H'), g('I')) if x))
    elif isinstance(g('H'), str):
        m = re.match(r'(\d{4})/(\d{1,2})/(\d{1,2})\s*(.*)', g('H'))
        if m: add(0, datetime.date(int(m[1]), int(m[2]), int(m[3])), None, '', m[4])
    add(1, g('W'), g('Y'), g('T'), g('U')); add(2, g('AJ'), g('AL'), g('AG'), g('AH'))
    add(3, g('BC'), g('BE'), g('AZ'), g('BA')); add(4, g('BR'), g('BT'), '', g('BP'))

# ---------------- 回想 ----------------
h = wb['回想']; rec = []; rc = []
for r in range(3, h.max_row + 1):
    b = h.cell(r, 2).value
    if isinstance(b, (int, float)) and (h.cell(r, 3).value or h.cell(r, 4).value or h.cell(r, 5).value):
        rec.append([int(b)] + [h.cell(r, c).value for c in (3, 4, 5, 6)])
        if h.cell(r, 1).value is True: rc.append(int(b))

# ---------------- 簡易チェック ----------------
if len(sw) < 100 or len(ev) < 100 or len(rec) < 100:
    sys.exit(f'データ件数が少なすぎます（刀剣{len(sw)}・予定{len(ev)}・回想{len(rec)}）。Excelの列構成が変わっていないか確認してください。')

MASTER = dict(sw=sw, ev=ev, rec=rec)               # 公開しても困らない部分
MINE = dict(own=own, got=got, gr=gr, tg=tg, rc=rc, kgot=kgot, cf=cfm)   # 自分のゲームデータ（非公開）
dump = lambda o: json.dumps(o, ensure_ascii=False, separators=(',', ':'), default=str).replace('</', '<\\/')

# ---------------- HTML のデータ部分だけ差し替え ----------------
html = open(HTML, encoding='utf8').read()
def swap(h, ident, obj):
    pat = re.compile(r'(<script type="application/json" id="%s">)(.*?)(</script>)' % ident, re.S)
    if not pat.search(h): sys.exit(f'HTML に id="{ident}" のデータ部分が見つかりません。')
    return pat.sub(lambda m: m.group(1) + dump(obj) + m.group(3), h, count=1)
shutil.copyfile(HTML, HTML + '.bak')
open(HTML, 'w', encoding='utf8').write(swap(swap(html, 'm', MASTER), 'u', MINE))

last = max(e[1] for e in ev)
print(f'更新しました: {HTML}（元の版は {HTML}.bak）')
print(f'  刀剣 {len(sw)}振り（所持 {len(own)}）／ 年表の予定 {len(ev)}件（最新の開始日 {last}）／ 回想 {len(rec)}件（回収済 {len(rc)}）')
