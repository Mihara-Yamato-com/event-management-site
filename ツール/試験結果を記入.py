# -*- coding: utf-8 -*-
"""結合試験仕様書.xlsx の実施欄（日付・結果・AI）を書き換える。

再生成すると手動で記入済みの結果が消えるため、既存ファイルのセルだけを差し替える。
使い方: 試験結果を記入.py results.json
  results.json = [{"sheet": 4, "no": 1, "result": "〇", "date": 46232, "ai": "利用"}, ...]
  sheet は大項目の通し番号（1=イベント管理 … 10=境界値・異常系）、no は明細番号。

  記入済みの欄を空に戻すには "clear" を使う（手動で再実施する項目を未記入に戻す場合）。
  [{"sheet": 5, "no": 4, "clear": ["date", "result"]}, ...]
"""
import zipfile, shutil, sys, json, io, os, re
from xml.sax.saxutils import escape

XLSX = '/home/mihara/projects/event-management-site/結合試験仕様書.xlsx'
COL_DATE, COL_RESULT, COL_AI = 'M', 'N', 'P'


def load_sst(xml):
    items = re.findall(r'<si>(.*?)</si>', xml, re.S)
    texts = [re.sub(r'<[^>]+>', '', s) for s in items]
    return items, texts


def patch(results):
    z = zipfile.ZipFile(XLSX)
    parts = {n: z.read(n) for n in z.namelist()}
    z.close()

    sx = parts['xl/sharedStrings.xml'].decode('utf-8')
    items, texts = load_sst(sx)

    def sid(t):
        if t in texts:
            return texts.index(t)
        texts.append(t)
        items.append('<t xml:space="preserve">%s</t>' % escape(t))
        return len(texts) - 1

    # シート番号 → ファイル名（表紙・凡例の2枚ぶんずれる）
    by_sheet = {}
    for r in results:
        by_sheet.setdefault(int(r['sheet']), []).append(r)

    touched = 0
    for sheet_no, rows in by_sheet.items():
        name = 'xl/worksheets/sheet%d.xml' % (sheet_no + 2)
        xml = parts[name].decode('utf-8')
        for r in rows:
            row = 7 + int(r['no'])            # データは8行目から
            clear = set(r.get('clear', []))
            for col, val, kind, key in ((COL_DATE, r.get('date'), 'n', 'date'),
                                        (COL_RESULT, r.get('result'), 's', 'result'),
                                        (COL_AI, r.get('ai'), 's', 'ai')):
                if key not in clear and val in (None, ''):
                    continue
                ref = '%s%d' % (col, row)
                m = re.search(r'<c r="%s"(?: s="(\d+)")?[^>]*?(?:/>|>.*?</c>)' % ref, xml, re.S)
                if not m:
                    raise SystemExit('セルが見つかりません: %s シート%d' % (ref, sheet_no))
                style = m.group(1) or '0'
                if key in clear:                   # 空欄に戻す
                    cell = '<c r="%s" s="%s"/>' % (ref, style)
                elif kind == 'n':
                    cell = '<c r="%s" s="%s"><v>%s</v></c>' % (ref, style, val)
                else:
                    cell = '<c r="%s" s="%s" t="s"><v>%d</v></c>' % (ref, style, sid(val))
                xml = xml[:m.start()] + cell + xml[m.end():]
                touched += 1
        parts[name] = xml.encode('utf-8')

    body = ''.join('<si>%s</si>' % i for i in items)
    parts['xl/sharedStrings.xml'] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'count="%d" uniqueCount="%d">%s</sst>' % (len(items), len(items), body)).encode('utf-8')

    tmp = XLSX + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as out:
        for n, data in parts.items():
            out.writestr(n, data)
    shutil.move(tmp, XLSX)
    print('セル %d 個を更新しました' % touched)


if __name__ == '__main__':
    patch(json.load(io.open(sys.argv[1], encoding='utf-8')))
