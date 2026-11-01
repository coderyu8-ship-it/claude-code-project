import copy, sys
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
FONT = 'Noto Sans JP'
INK = RGBColor(0x2A, 0x2F, 0x33); SUB = RGBColor(0x5F, 0x66, 0x6B); TEAL = RGBColor(0x14, 0x5F, 0x6B)
TERRA = RGBColor(0xD9, 0x76, 0x4A); CREAM = RGBColor(0xF1, 0xE9, 0xDC); LINE = RGBColor(0xE6, 0xDD, 0xCF); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GOLD = TEAL
W, MX = 12192000, 762000
CW = W - 2 * MX
exec(open('deck_helpers.py').read())

SRC, OUT = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
S = list(prs.slides)

def notes(s, body):
    ns = s.notes_slide
    if ns.notes_text_frame is None:
        src = S[4].notes_slide
        body_ph = [ph for ph in src.placeholders if ph.placeholder_format.type == 2][0]
        el = copy.deepcopy(body_ph._element)
        ns.shapes._spTree.append(el)
    ns.notes_text_frame.text = body.strip()


def shp(slide, i):
    return list(slide.shapes)[i]

def delete(x):
    x._element.getparent().remove(x._element)

def settext(x, t):
    r = x.text_frame.paragraphs[0].runs
    r[0].text = t
    for extra in r[1:]:
        extra._r.getparent().remove(extra._r)

# 指示メモを削除

delete(shp(S[0], 5))
delete(shp(S[17], 22))

def header_slide(label, title, lead, speech):
    """料金一覧スライドの見出し部分（0〜7番の図形）を複製して新しいスライドを作る。"""
    s = prs.slides.add_slide(S[17].slide_layout)
    for x in list(s.shapes):
        delete(x)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = RGBColor(0xFA, 0xF6, 0xEF)
    for i in range(8):
        s.shapes._spTree.append(copy.deepcopy(shp(S[17], i)._element))
    settext(shp(s, 2), label)
    settext(shp(s, 3), title)
    if lead:
        text(s, MX, 1200000, CW, 300000, lead, 14, True, TEAL)
    notes(s, speech)
    return s

# ---- 6枚目（新規）：LPのファーストビュー
fv = header_slide('SERVICE SITE', 'サービスサイト', 'U-ライティングのサービスページ（https://www.u-promotion.com/）', """
このあたりの考え方は、私のサービスサイトにもまとめています。
（サイトの最初の画面を見せて、「この後、広告づくりで大事にしている原則を5つお話しします」と次につなげる）
""")
ph = box(fv, MX, 1650000, CW, 4550000, fill=CREAM, line=TEAL)
ph.line.dash_style = 4
tf = ph.text_frame
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = 'ここにLPのファーストビュー（最初に見える画面）のスクリーンショットを貼る'
font(r, 14, True, TEAL)

# ---- 18枚目：カワリミの料金一覧＋紹介料
k = S[17]
settext(shp(k, 3), '料金一覧：制作代行「カワリミ」')
for i in range(19, 13, -1):
    delete(shp(k, i))
delete([x for x in k.shapes if x.has_text_frame and x.text_frame.text.startswith('※ 価格は税込')][0])
for x in list(k.shapes)[8:14]:
    x.top = x.top - 120000
    if x.height == 1150000:
        x.height = Emu(950000)
ref = box(k, 6250000, 1430000, 5180000, 950000, fill=TEAL)
text(k, 6550000, 1510000, 2400000, 300000, '紹介料', 12, True, WHITE)
text(k, 6550000, 1800000, 2600000, 600000, '契約金額の10%', 22, True, WHITE)
text(k, 9000000, 1580000, 2300000, 900000, ['契約金額10万円以上が対象', '着金確認後にお支払い'], 11, False, WHITE, spacing=1.5)
tbl_shape = [x for x in k.shapes if x.has_table][0]
tbl = tbl_shape.table
rows_add = [('Web制作', 'HPライティング', '1ページ 20,000円〜'),
            ('紙媒体', 'パンフレット原稿', '1ページ 10,000円〜'),
            ('SNS・広告', '有料広告用バナー制作（10枚セット）', '30,000円〜'),
            ('その他', '動画・メルマガ原稿', '1本 10,000円〜')]
tr_tpl = tbl._tbl.tr_lst[-1]
for cat, item, price in rows_add:
    tr = copy.deepcopy(tr_tpl)
    tbl._tbl.append(tr)
    row = tbl.rows[len(tbl.rows) - 1]
    for j, v in enumerate((cat, item, price)):
        rs = row.cells[j].text_frame.paragraphs[0].runs
        rs[0].text = v
        for extra in rs[1:]:
            extra._r.getparent().remove(extra._r)
for i, row in enumerate(tbl.rows):
    row.height = Emu(285000)
    for j, c in enumerate(row.cells):
        c.margin_top = c.margin_bottom = Emu(15000)
        if i:
            c.fill.solid()
            c.fill.fore_color.rgb = CREAM if i % 2 == 0 else WHITE
        for rr in c.text_frame.paragraphs[0].runs:
            rr.font.size = Pt(10.5 if i else 10)
tbl_shape.top = Emu(2520000)
text(k, MX, 6130000, CW, 220000, '※ 価格は税込・目安です。内容やボリュームに応じて事前にお見積りします。「ミチスジ」顧問プラン契約中は20% OFF。', 9, False, SUB)
k.notes_slide.notes_text_frame.text = ("制作代行カワリミの料金はこちらです。単発で5,000円から、LPやHPの制作は10万円からです。\n"
    "今日のお話だと、【○○】が近いと思います。正式な金額は、伺った内容をもとにお見積りしますね。\n"
    "（紹介を目指す相手には）ご紹介からカワリミの契約につながった場合は、契約金額の10%を紹介料としてお支払いしています。\n"
    "（全部説明しない。相手に関係する行だけ指差す）")

# ---- 19枚目（新規）：ミチスジの料金一覧＋紹介料
m = header_slide('PRICE', '料金一覧：集客顧問「ミチスジ」', None, """
集客顧問ミチスジは、ご自身で動きたい方の横で道筋を示すサービスです。
まずは1回の相談から、方向性を資料で欲しい方は提案プラン、続けて伴走が必要なら顧問プランと、段階的に使えます。
添削だけのご依頼もできます。
（紹介を目指す相手には）顧問プランは、契約が続く限り、毎月の契約金額の20%を紹介料としてお支払いしています。
""")
plans = [('相談プラン', '1回', '14,800円〜', '最大2時間のZoom相談で、現状の添削と優先順位をご提案'),
         ('提案プラン', '単発', '49,800円〜', '現状分析・競合リサーチをもとに、施策の提案書を作成・プレゼン'),
         ('顧問プラン', '月額', '49,800円〜', '月最大4時間のZoom相談、添削、施策の分析・仮説検証。カワリミ20%OFF')]
cw = 3400000
for i, (name, unit, price, desc) in enumerate(plans):
    x = MX + i * (cw + 234000)
    box(m, x, 1550000, cw, 1800000, fill=WHITE, line=TEAL if i == 2 else LINE, lw=1.5 if i == 2 else 1)
    box(m, x, 1550000, cw, 50800, fill=TEAL)
    text(m, x + 250000, 1720000, cw - 500000, 400000, name, 17, True)
    text(m, x + 250000, 2200000, 700000, 400000, unit, 11, False, SUB)
    text(m, x + 800000, 2100000, cw - 1050000, 500000, price, 22, True, TERRA)
    text(m, x + 250000, 2750000, cw - 500000, 800000, desc, 11, False, SUB, spacing=1.45)
text(m, MX, 3500000, 5300000, 300000, '添削・監修メニュー', 12, True, TEAL)
edits = [('薬機法ライティング監修（記事・LP）', '10,000円〜'), ('LP添削', '40,000円〜'), ('HP添削', '1ページ 10,000円〜'), ('スピーチ原稿・自己紹介シート添削', '5,000円〜'), ('略歴・GAINSシート制作サポート', '20,000円〜')]
for i, (h_, price) in enumerate(edits):
    y = 3820000 + i * 450000
    box(m, MX, y, 5300000, 410000, fill=CREAM if i % 2 == 0 else WHITE, line=LINE)
    text(m, MX + 200000, y + 75000, 3300000, 300000, h_, 11, True)
    text(m, MX + 3300000, y + 75000, 1800000, 300000, price, 12, True, TERRA, PP_ALIGN.RIGHT)
box(m, 6250000, 3500000, 5180000, 2500000, fill=TEAL)
text(m, 6550000, 3750000, 4600000, 300000, '紹介料（顧問プラン）', 12, True, WHITE)
text(m, 6550000, 4150000, 4600000, 700000, '毎月、契約金額の20%', 24, True, WHITE)
text(m, 6550000, 5000000, 4600000, 700000, ['契約が続く限り、毎月お支払い', '着金確認後にお支払い'], 11, False, WHITE, spacing=1.5)
text(m, MX, 6130000, CW, 220000, '※ 価格は税込・目安です。契約の縛りはありません。', 9, False, SUB)

# ---- 並べ替え：FVを6枚目、ミチスジ料金を19枚目に
lst = prs.slides._sldIdLst
els = list(lst)
fv_el, m_el = els[-2], els[-1]
for e in els:
    lst.remove(e)
order = els[:5] + [fv_el] + els[5:18] + [m_el] + els[18:22]
for e in order:
    lst.append(e)

# ---- ページ番号を通し番号に（表紙は除く）
for n, s in enumerate(prs.slides, 1):
    for x in s.shapes:
        if x.has_text_frame and x.text_frame.text.strip().isdigit() and x.left == 11010000 and x.top == 6430000:
            settext(x, str(n))

prs.save(OUT)
print('saved', OUT, len(prs.slides))
