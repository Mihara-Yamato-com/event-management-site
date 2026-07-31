# -*- coding: utf-8 -*-
"""結合試験仕様書.xlsx に「エビデンス」欄（Q列）を追加し、試験エビデンス/ のファイル名を記入する。

各シートの Q列にヘッダーを作り、章と明細番号からエビデンスのファイルを引き当てて貼る。
ファイル名の規則は「<章>-<明細2桁><枝番>_final.<拡張子>」（例: 1-05a_final.png）。
二度実行しても Q列が重複しないよう、既存の Q列セルは差し替える。

使い方: python3 ツール/エビデンス欄を追加.py
"""
import zipfile, shutil, re, os, collections
from xml.sax.saxutils import escape

BASE = '/home/mihara/projects/event-management-site'
XLSX = os.path.join(BASE, '結合試験仕様書.xlsx')
EVDIR = os.path.join(BASE, '試験エビデンス')
EVREL = '試験エビデンス'
COL_Q = 'Q'
STYLE_HEADER, STYLE_BODY = '9', '8'   # P7（見出し）と O8（本文）に合わせる


def collect():
    """(章, 明細) -> [ファイル名] を作る。枝番順に並べる。"""
    ev = collections.defaultdict(list)
    for f in sorted(os.listdir(EVDIR)):
        m = re.match(r'(\d+)-(\d+)([a-z]?\d?)_final\.(png|json)$', f)
        if m:
            ev[(int(m.group(1)), int(m.group(2)))].append(f)
    for k in ev:
        ev[k].sort()
    return ev


def main():
    ev = collect()
    z = zipfile.ZipFile(XLSX)
    parts = {n: z.read(n) for n in z.namelist()}
    z.close()

    sx = parts['xl/sharedStrings.xml'].decode('utf-8')
    items = re.findall(r'<si>(.*?)</si>', sx, re.S)
    texts = [re.sub(r'<[^>]+>', '', re.sub(r'<rPh.*?</rPh>', '', s, flags=re.S)) for s in items]

    def sid(t):
        if t in texts:
            return texts.index(t)
        texts.append(t)
        items.append('<t xml:space="preserve">%s</t>' % escape(t))
        return len(texts) - 1

    hdr = sid('エビデンス')
    filled = 0

    for ch in range(1, 11):
        name = 'xl/worksheets/sheet%d.xml' % (ch + 2)
        x = parts[name].decode('utf-8')

        x = re.sub(r'(<dimension ref="A1:)[A-Z]+(\d+)"/>', r'\1Q\2"/>', x)
        if 'min="17"' not in x:
            x = x.replace('</cols>',
                          '<col min="17" max="17" width="34.6640625" style="10" customWidth="1"/></cols>')
        x = x.replace('spans="1:16"', 'spans="1:17"')
        x = x.replace('<mergeCell ref="L6:P6"/>', '<mergeCell ref="L6:Q6"/>')

        def put(rowxml, rnum, value, style):
            """行の末尾に Q セルを置く（既にあれば差し替える）。"""
            ref = '%s%d' % (COL_Q, rnum)
            cell = ('<c r="%s" s="%s" t="s"><v>%d</v></c>' % (ref, style, sid(value))
                    if value else '<c r="%s" s="%s"/>' % (ref, style))
            old = re.search(r'<c r="%s"[^>]*?(?:/>|>.*?</c>)' % ref, rowxml, re.S)
            if old:
                return rowxml[:old.start()] + cell + rowxml[old.end():]
            return rowxml.replace('</row>', cell + '</row>')

        out = []
        pos = 0
        for m in re.finditer(r'<row r="(\d+)"[^>]*>.*?</row>', x, re.S):
            rnum = int(m.group(1))
            rowxml = m.group(0)
            if rnum == 7:
                rowxml = put(rowxml, rnum, 'エビデンス', STYLE_HEADER)
            elif rnum >= 8:
                d = re.search(r'<c r="D%d"[^>]*>\s*<v>(\d+)</v>' % rnum, rowxml)
                if d:
                    files = ev.get((ch, int(d.group(1))), [])
                    val = '\n'.join(EVREL + '/' + f for f in files)
                    if files:
                        filled += 1
                    rowxml = put(rowxml, rnum, val, STYLE_BODY)
            out.append(x[pos:m.start()]); out.append(rowxml); pos = m.end()
        out.append(x[pos:])
        parts[name] = ''.join(out).encode('utf-8')

    body = ''.join('<si>%s</si>' % i for i in items)
    parts['xl/sharedStrings.xml'] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'count="%d" uniqueCount="%d">%s</sst>' % (len(items), len(items), body)).encode('utf-8')

    tmp = XLSX + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as o:
        for n, data in parts.items():
            o.writestr(n, data)
    shutil.move(tmp, XLSX)
    print('エビデンスを記入した項目: %d 件 / ファイル総数 %d' % (filled, sum(len(v) for v in ev.values())))


if __name__ == '__main__':
    main()
