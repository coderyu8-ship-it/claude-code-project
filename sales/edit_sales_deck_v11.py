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
exec(open(sys.argv[3]).read())
prs = Presentation(sys.argv[1])
S = list(prs.slides)

def shp(s, i):
    return list(s.shapes)[i]

def set_paras(x, lines):
    tf = x.text_frame
    ps = list(tf.paragraphs)
    while len(ps) < len(lines):
        ps[-1]._p.addnext(copy.deepcopy(ps[-1]._p)); ps = list(tf.paragraphs)
    for p in ps[len(lines):]:
        p._p.getparent().remove(p._p)
    for p, t in zip(tf.paragraphs, lines):
        rs = p.runs
        rs[0].text = t
        for r in rs[1:]:
            r._r.getparent().remove(r._r)

# ---- ご紹介のお願い（22枚目）
k = S[21]
set_paras(shp(k, 4), ['こんな声を聞いたら、「広告の赤ペン先生」を思い出してください'])
for i, t in zip((12, 16, 20, 24, 28), ['「集客がうまくいかない」', '「売上を伸ばしたい」「事業を発展させたい」',
                                      '「アピールしているのに、サービスの魅力が伝わらない」', '「サービスに興味を持ってもらえない」',
                                      '「AIや制作会社に頼んでも、集客につながらない」']):
    set_paras(shp(k, i), [t])
set_paras(shp(k, 29), ['「売り方を見直してくれる、広告の赤ペン先生がいるよ」'])
k.notes_slide.notes_text_frame.text = ("（紹介を目指す場合、または契約の話のあとに）最後に一つ、お願いがあります。\n"
    "お会いした方から、こんな声を聞いたら私を思い出してください。\n"
    "そのときは「売り方を見直してくれる、広告の赤ペン先生がいるよ」と紹介していただけると嬉しいです。\n"
    "お会いした方には、ぜひ「今の悩み」と「今後の目標」を聞いてみてください。そこにヒントがあります。")

# ---- ご紹介いただきたい方（23枚目）
k = S[22]
set_paras(shp(k, 11), ['集客・売上に課題を抱える事業者'])
set_paras(shp(k, 12), ['● 美容サロン・脱毛サロン・整体院', '● 医療クリニック', '● 健康食品・健康雑貨・美容機器・化粧品の販売者',
                       '● 制作会社・マーケティング会社', '● 経営・資金調達・補助金コンサルタント'])
set_paras(shp(k, 16), ['お互いに紹介し合える方'])
set_paras(shp(k, 17), ['● 医療・美容機器ディーラー', '● 店舗物件メインの不動産仲介', '● ビジネスマッチングサービス運営者',
                       '● 経営者コミュニティ・シェアサロン運営者', '● 税理士・会計士・財務コンサルタント'])
for x in k.shapes:
    if x.height == 4450000:
        x.height = Emu(4150000)
text(k, MX, 5950000, CW, 300000, '※ 個人クリエイター（ライター・デザイナー・動画編集者）の方、金融投資・スピリチュアル商材・MLMが主力の方は、ご紹介の対象外とさせてください。', 10, False, SUB)
k.notes_slide.notes_text_frame.text = ("特に、美容・健康系で集客や売上に課題を抱えている事業者さんとは相性がいいです。\n"
    "協業できるパートナーさんとして、特に医療・美容機器ディーラー、店舗物件の不動産仲介、ビジネスマッチングサービス運営者の方とつながりたいです。\n"
    "（対象外の方は、角が立たないよう「お役に立てる範囲が限られてしまうので」と伝える）")

# ---- 付録（21枚目）を削除
lst = prs.slides._sldIdLst
sid = list(lst)[20]
prs.part.drop_rel(sid.rId)
lst.remove(sid)

# ---- ページ番号
for n, s in enumerate(prs.slides, 1):
    for x in s.shapes:
        if x.has_text_frame and x.text_frame.text.strip().isdigit() and x.left == 11010000 and x.top == 6430000:
            set_paras(x, [str(n)])
prs.save(sys.argv[2])
print('saved', len(prs.slides))
