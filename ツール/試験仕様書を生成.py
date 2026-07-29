# -*- coding: utf-8 -*-
"""結合試験仕様書（xlsx）を生成する。
openpyxl 等が使えないため、totaltest.xlsx の styles.xml / theme1.xml を流用し、
シートの XML を組み立てて zip し直す。列幅・スタイル・ウィンドウ枠固定は見本と同じ。
"""
import zipfile, shutil, os, sys
from xml.sax.saxutils import escape

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from 試験項目 import SECTIONS, AI_ADDED

TEMPLATE = '/home/mihara/projects/event-management-site/totaltest.xlsx'
OUT = '/home/mihara/projects/event-management-site/結合試験仕様書.xlsx'

SERVICE = 'ACTIO（イベント主催・管理サービス）'
COMPANY = '日本コムシンク株式会社'
MADE = '　作成：2026年7月29日'
TESTER = '三原'
AI_NOTE = 'AIレビューにより追記'

# ---------- 共有文字列 ----------
class SST:
    def __init__(self):
        self.items = []
        self.index = {}
    def add(self, s):
        s = '' if s is None else str(s)
        if s not in self.index:
            self.index[s] = len(self.items)
            self.items.append(s)
        return self.index[s]
    def xml(self):
        body = ''.join('<si><t xml:space="preserve">%s</t></si>' % escape(i) for i in self.items)
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'count="%d" uniqueCount="%d">%s</sst>' % (len(self.items), len(self.items), body))

sst = SST()

def cs(ref, style, text):
    """文字列セル"""
    if text is None or text == '':
        return '<c r="%s" s="%d"/>' % (ref, style)
    return '<c r="%s" s="%d" t="s"><v>%d</v></c>' % (ref, style, sst.add(text))

def cn(ref, style, num):
    """数値セル"""
    return '<c r="%s" s="%d"><v>%s</v></c>' % (ref, style, num)

def ce(ref, style):
    return '<c r="%s" s="%d"/>' % (ref, style)

HDR = ['No', '枝番', '明細', '大項目', '中項目', '小項目', '観点',
       '前提条件', '手順', '期待結果', '担当者', '日付', '結果', '備考', 'AI']

COLS = ('<cols>'
        '<col min="2" max="2" width="5.9140625" style="17" customWidth="1"/>'
        '<col min="3" max="3" width="3.9140625" style="17" customWidth="1"/>'
        '<col min="4" max="4" width="4.1640625" style="17" customWidth="1"/>'
        '<col min="5" max="5" width="22.6640625" style="17" customWidth="1"/>'
        '<col min="6" max="6" width="47.58203125" style="17" customWidth="1"/>'
        '<col min="7" max="7" width="5.6640625" style="17" customWidth="1"/>'
        '<col min="8" max="8" width="9.83203125" style="17" customWidth="1"/>'
        '<col min="9" max="9" width="35.6640625" style="17" customWidth="1"/>'
        '<col min="10" max="10" width="57.08203125" style="17" customWidth="1"/>'
        '<col min="11" max="11" width="71" style="17" customWidth="1"/>'
        '<col min="12" max="12" width="10.5" style="17" customWidth="1"/>'
        '<col min="13" max="13" width="10.75" style="17" customWidth="1"/>'
        '<col min="14" max="14" width="13" style="17" customWidth="1"/>'
        '<col min="15" max="15" width="44.58203125" style="17" customWidth="1"/>'
        '<col min="16" max="16" width="6.5" style="17" customWidth="1"/>'
        '</cols>')

WS_OPEN = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')


def sheet_section(no, name, mokuteki, zentei, rows):
    """大項目シートを組み立てる"""
    last = 7 + len(rows)
    x = [WS_OPEN]
    x.append('<dimension ref="A1:P%d"/>' % last)
    x.append('<sheetViews><sheetView zoomScale="70" zoomScaleNormal="70" workbookViewId="0">'
             '<pane xSplit="4" ySplit="7" topLeftCell="E8" activePane="bottomRight" state="frozen"/>'
             '<selection pane="bottomRight" activeCell="E8" sqref="E8"/>'
             '</sheetView></sheetViews>')
    x.append('<sheetFormatPr defaultRowHeight="18"/>')
    x.append(COLS)
    x.append('<sheetData>')

    # 1行目 タイトル
    x.append('<row r="1" spans="1:16" ht="26.5">' + cs('A1', 24, '・' + name)
             + ''.join(ce(c + '1', 24) for c in 'BCDE') + '</row>')
    # 3行目 目的
    x.append('<row r="3" spans="1:16" ht="42" customHeight="1">' + cs('B3', 1, '目的')
             + cs('C3', 23, mokuteki) + ''.join(ce(c + '3', 23) for c in 'DEF') + '</row>')
    # 4行目 前提
    x.append('<row r="4" spans="1:16" ht="42" customHeight="1">' + cs('B4', 1, '前提')
             + cs('C4', 23, zentei) + ''.join(ce(c + '4', 23) for c in 'DEF') + '</row>')
    # 6行目 帯
    x.append('<row r="6" spans="1:16">' + cs('B6', 34, 'テスト内容')
             + ''.join(ce(c + '6', 26) for c in 'CDEFGHIJ') + ce('K6', 27)
             + cs('L6', 25, 'テスト実施')
             + ''.join(ce(c + '6', 26) for c in 'MNO') + ce('P6', 27) + '</row>')
    # 7行目 見出し
    hdr = ['<row r="7" spans="1:16">', cs('B7', 1, HDR[0])]
    for i, col in enumerate('CDEFGHIJK'):
        hdr.append(cs(col + '7', 3, HDR[i + 1]))
    for i, col in enumerate('LMNO'):
        hdr.append(cs(col + '7', 4, HDR[i + 10]))
    hdr.append(cs('P7', 15, HDR[14]))
    hdr.append('</row>')
    x.append(''.join(hdr))

    # データ
    for i, (naka, kanten, zen, tejun, kitai, biko) in enumerate(rows):
        r = 8 + i
        first = (i == 0)
        ai = naka in AI_ADDED
        if ai:
            biko = (biko + ' ／ ' + AI_NOTE) if biko else AI_NOTE
        cells = [
            cn('B%d' % r, 28, no) if first else ce('B%d' % r, 29),
            cn('C%d' % r, 28, 1) if first else ce('C%d' % r, 29),
            cn('D%d' % r, 22, i + 1),
            cs('E%d' % r, 31, name) if first else ce('E%d' % r, 32),
            cs('F%d' % r, 13, naka),
            ce('G%d' % r, 13),
            cs('H%d' % r, 13, kanten),
            cs('I%d' % r, 13, zen),
            cs('J%d' % r, 13, tejun),
            cs('K%d' % r, 13, kitai),
            cs('L%d' % r, 13, TESTER),  # 担当者
            ce('M%d' % r, 2),           # 日付
            ce('N%d' % r, 13),          # 結果
            cs('O%d' % r, 13, biko),
            cs('P%d' % r, 13, '利用') if ai else ce('P%d' % r, 13),
        ]
        x.append('<row r="%d" spans="1:16">%s</row>' % (r, ''.join(cells)))

    x.append('</sheetData>')
    x.append('<mergeCells count="8">'
             '<mergeCell ref="A1:E1"/><mergeCell ref="C3:F3"/><mergeCell ref="C4:F4"/>'
             '<mergeCell ref="B6:K6"/><mergeCell ref="L6:P6"/>'
             '<mergeCell ref="B8:B%d"/><mergeCell ref="C8:C%d"/><mergeCell ref="E8:E%d"/>'
             '</mergeCells>' % (last, last, last))
    x.append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    x.append('</worksheet>')
    return ''.join(x)


def sheet_cover():
    x = [WS_OPEN, '<dimension ref="A1:L32"/>',
         '<sheetViews><sheetView showGridLines="0" tabSelected="1" zoomScale="85" '
         'zoomScaleNormal="85" workbookViewId="0"/></sheetViews>',
         '<sheetFormatPr defaultRowHeight="18"/>',
         '<cols><col min="1" max="1" width="60" customWidth="1"/></cols>',
         '<sheetData>']
    body = {8: (11, SERVICE), 12: (9, '結合試験仕様書兼実施項目書'),
            14: (8, '対象: index.html（単一ファイル・file:// 起動）'),
            15: (8, '試験項目: 全%d項目（大項目%d）' % (sum(len(x[3]) for x in SECTIONS), len(SECTIONS))),
            16: (8, '動作保証: Google Chrome のみ'),
            29: (8, COMPANY), 30: (8, MADE)}
    for r in range(1, 33):
        if r in body:
            st, t = body[r]
            x.append('<row r="%d" spans="1:12" ht="23" customHeight="1">%s</row>' % (r, cs('A%d' % r, st, t)))
        else:
            x.append('<row r="%d" spans="1:12">%s</row>' % (r, ce('A%d' % r, 8)))
    x.append('</sheetData>')
    x.append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    x.append('</worksheet>')
    return ''.join(x)


def simple_sheet(title, note, header, rows, widths):
    """凡例・NG項目一覧のような単純な表"""
    last = 4 + len(rows)
    x = [WS_OPEN, '<dimension ref="A1:%s%d"/>' % (chr(ord('A') + len(header)), max(last, 5)),
         '<sheetViews><sheetView zoomScale="85" zoomScaleNormal="85" workbookViewId="0">'
         '<pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/>'
         '</sheetView></sheetViews>',
         '<sheetFormatPr defaultRowHeight="18"/>']
    cols = ['<cols>']
    for i, w in enumerate(widths):
        cols.append('<col min="%d" max="%d" width="%s" style="17" customWidth="1"/>' % (i + 2, i + 2, w))
    cols.append('</cols>')
    x.append(''.join(cols))
    x.append('<sheetData>')
    x.append('<row r="1" spans="1:12" ht="26.5">%s</row>' % cs('B1', 24, title))
    if note:
        x.append('<row r="2" spans="1:12">%s</row>' % cs('B2', 13, note))
    hdr = ['<row r="4" spans="1:12">']
    for i, h in enumerate(header):
        hdr.append(cs(chr(ord('B') + i) + '4', 3, h))
    hdr.append('</row>')
    x.append(''.join(hdr))
    for i, row in enumerate(rows):
        r = 5 + i
        cells = ''.join(cs(chr(ord('B') + j) + str(r), 13, v) for j, v in enumerate(row))
        x.append('<row r="%d" spans="1:12">%s</row>' % (r, cells))
    x.append('</sheetData>')
    x.append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    x.append('</worksheet>')
    return ''.join(x)


# ---------- 凡例 ----------
LEGEND_ROWS = [
    ('書式の見方', 'No / 枝番 / 明細', '試験項目の通し番号・枝番号・明細番号。行を特定するための番号です。'),
    ('', '大項目 / 中項目 / 小項目', 'テスト対象の分類。大項目＝機能のまとまり（シート名）、中項目＝1件ごとの試験項目です。小項目は本仕様書では使用しません。'),
    ('', '観点', 'そのテストで「何を確かめるか」の種類（下の観点セクションを参照）。'),
    ('', '前提条件', 'テストを始める前に整えておく状態・準備。'),
    ('', '手順', '実際に行う操作の手順。'),
    ('', '期待結果', '操作後に「こうなっていれば正しい」という状態。'),
    ('', '結果（〇 / ×）', '〇＝期待どおり動いた。×＝期待どおりにならなかった（不具合）。×はNG項目一覧に転記します。'),
    ('', '備考', '補足。先頭の「ID:」は対応する機能ID・入力仕様IDで、続けてその仕様にした理由や、実施するうえでの注意を書いています。'),
    ('', 'AI', 'AI を利用した項目に「利用」と記入します。本仕様書では、AI のレビューで追記した項目に記入し、備考にも「AIレビューにより追記」と残しています。'),
    ('観点', '操作', 'ボタンや入力など、利用者の操作が意図どおり動くか。'),
    ('', '表示', '画面に出る内容・レイアウトが正しいか。'),
    ('', '状態遷移', '操作の前後で状態が正しく変わるか（未受付→参加確定 など）。'),
    ('', '機能間結合', '複数の機能をまたいだときに整合が取れているか（CSV出力・JSON入出力 など）。'),
    ('', '異常系', '誤った入力や想定外の操作をしたときに、適切に止めて案内できるか。'),
    ('', '境界値', '上限・下限・ちょうどの値で正しく振る舞うか。'),
    ('', 'モード切替', '主催者と参加者など、立場によって表示が変わるか。'),
    ('用語', 'F-xx / S-xx / D-xx', '機能ID・画面ID・入力仕様ID。備考欄に「ID:」として記載しています。機能一覧.md / 画面一覧.md / 仕様決定シート.md と対応し、どの要件を確かめている項目かを追えます。'),
    ('', '券種', '参加枠のこと（一般・学生 など）。券種ごとに定員を決められます。本サービスでは無料のものだけを扱います。'),
    ('', 'イベントの状態', '公開前（下書き）／公開中／終了。終了は、終了日時を過ぎると自動的に切り替わります。'),
    ('', '申込の状態', '申し込み済み／参加確定（当日の受付が済んだ状態）／キャンセル済み。'),
    ('', 'シリーズ', '第1回・第2回と続く連続開催のまとまり。回をまたいだ集客の変化を見るために使います。'),
    ('', 'マイチケット', '参加者が自分の申込を確認・キャンセルする画面。メールアドレスとパスワードで開きます。'),
    ('', '受付コード', '申込ごとに発行される8桁の英数字。当日受付・申込確認・キャンセルに使います。'),
    ('', 'デモデータ', '試験の前提条件を1操作で作る機能（設定・データ管理）。特記なき場合これを投入した状態から始めます。'),
    ('前提', '動作環境', 'index.html をダブルクリックし file:// で開く。Google Chrome のみを動作保証対象とします。'),
    ('', '素材', '画像・資料は 試験データ/ 配下のファイルを使います。'),
]

NG_HEADER = ['No', '状態', '対象機能', '明細No', '試験項目', '観点', '期待結果', '実際の挙動（NGの内容）', '対応・備考']

# ---------- 生成 ----------
sheets = []   # (name, xml)
sheets.append(('表紙', sheet_cover()))
sheets.append(('凡例・用語集', simple_sheet(
    '★ 凡例・用語集（この仕様書の見方と用語の説明）',
    '※ システムに詳しくない方が各シートを読む際の手引きです。',
    ['分類', '用語・項目', 'この仕様書での意味'], LEGEND_ROWS, ['16', '28', '95'])))

for i, (name, mokuteki, zentei, rows) in enumerate(SECTIONS, start=1):
    sheets.append((name, sheet_section(i, name, mokuteki, zentei, rows)))

sheets.append(('NG項目一覧', simple_sheet(
    '★ NG項目一覧（試験でNGだった点と対応）',
    '※ 試験の実施中に記入します。修正後は再試験の結果もここに残します。',
    NG_HEADER, [], ['6', '14', '20', '8', '40', '8', '40', '46', '46'])))

# ---------- パッケージ ----------
tz = zipfile.ZipFile(TEMPLATE)
styles = tz.read('xl/styles.xml')
theme = tz.read('xl/theme/theme1.xml')

ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
for i in range(1, len(sheets) + 1):
    ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % i)
ct.append('<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>')
ct.append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
ct.append('<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>')
ct.append('</Types>')

wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>']
for i, (name, _) in enumerate(sheets, start=1):
    wb.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(name), i, i))
wb.append('</sheets></workbook>')

rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
for i in range(1, len(sheets) + 1):
    rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i, i))
n = len(sheets)
rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' % (n + 1))
rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (n + 2))
rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' % (n + 3))
rels.append('</Relationships>')

root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
             '</Relationships>')

with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ''.join(ct))
    z.writestr('_rels/.rels', root_rels)
    z.writestr('xl/workbook.xml', ''.join(wb))
    z.writestr('xl/_rels/workbook.xml.rels', ''.join(rels))
    for i, (_, xml) in enumerate(sheets, start=1):
        z.writestr('xl/worksheets/sheet%d.xml' % i, xml)
    z.writestr('xl/theme/theme1.xml', theme)
    z.writestr('xl/styles.xml', styles)
    z.writestr('xl/sharedStrings.xml', sst.xml())

print('生成: %s' % OUT)
print('  シート %d 枚 / 共有文字列 %d 件 / %.0f KB'
      % (len(sheets), len(sst.items), os.path.getsize(OUT) / 1024))
