"""U-ライティング 1on1営業資料 v5（理念 → 広告づくりの原則 → サービスと強み）

usage: python3 build_v3.py <assets_dir> <out.pptx>
"""
import sys
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import Image

A, OUT = sys.argv[1], sys.argv[2]
FONT = 'Noto Sans JP'
INK = RGBColor(0x2A, 0x2F, 0x33)      # 文字：墨色
SUB = RGBColor(0x5F, 0x66, 0x6B)      # 補足文字
TEAL = RGBColor(0x14, 0x5F, 0x6B)     # 主色：ディープティール
TERRA = RGBColor(0xD9, 0x76, 0x4A)    # アクセント：テラコッタ（数字・価格・おすすめだけ）
IVORY = RGBColor(0xFA, 0xF6, 0xEF)    # 背景：アイボリー
CREAM = RGBColor(0xF1, 0xE9, 0xDC)    # 強調ボックス：サンド
LINE = RGBColor(0xE6, 0xDD, 0xCF)     # 罫線：ベージュ
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GOLD = TEAL                            # 旧コードの構造色をティールに置換
W, H = 12192000, 6858000
MX = 762000          # 左右マージン
CW = W - 2 * MX      # コンテンツ幅

prs = Presentation()
prs.slide_width, prs.slide_height = W, H
BLANK = prs.slide_layouts[6]
page = [0]


# ------------------------------------------------------------ helpers
def font(run, size, bold=False, color=INK):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rpr = run._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set('typeface', FONT)


def text(s, x, y, w, h, lines, size=14, bold=False, color=INK,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.25):
    """lines: str | list of str | list of (str, size, bold, color)."""
    tb = s.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(lines, str):
        lines = [lines]
    for i, ln in enumerate(lines):
        if isinstance(ln, str):
            ln = (ln, size, bold, color)
        t, sz, b, c = ln
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run()
        r.text = t
        font(r, sz, b, c)
    return tb


def box(s, x, y, w, h, fill=WHITE, line=None, shape=MSO_SHAPE.RECTANGLE, lw=1):
    b = s.shapes.add_shape(shape, Emu(x), Emu(y), Emu(w), Emu(h))
    b.shadow.inherit = False
    if fill is None:
        b.fill.background()
    else:
        b.fill.solid()
        b.fill.fore_color.rgb = fill
    if line is None:
        b.line.fill.background()
    else:
        b.line.color.rgb = line
        b.line.width = Pt(lw)
    return b


def pic(s, path, x, y, w=None, h=None, border=False):
    iw, ih = Image.open(path).size
    if w and not h:
        h = int(w * ih / iw)
    if h and not w:
        w = int(h * iw / ih)
    p = s.shapes.add_picture(path, Emu(x), Emu(y), Emu(w), Emu(h))
    if border:
        p.line.color.rgb = LINE
        p.line.width = Pt(1)
    return p


def notes(s, body):
    s.notes_slide.notes_text_frame.text = body.strip()


def new(dark=False):
    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = INK if dark else WHITE
    page[0] += 1
    return s


def base(label, title, speech):
    s = new()
    s.shapes.add_picture(f'{A}/topbar.png', 0, 0, Emu(W), Emu(142875))
    text(s, MX, 470000, CW, 260000, label, 11, True, GOLD)
    text(s, MX, 740000, CW, 600000, title, 26, True)
    box(s, MX, 6420000, CW, 9525, fill=LINE)
    text(s, MX, 6500000, 4000000, 200000, '© U-Writing All Rights Reserved.', 8, False, SUB)
    text(s, W - MX - 600000, 6470000, 600000, 260000, str(page[0]), 12, True, INK, PP_ALIGN.RIGHT)
    notes(s, speech)
    return s


def section(no, title, sub, speech):
    s = new(dark=True)
    text(s, MX, 2300000, 3000000, 900000, no, 54, True, TERRA)
    box(s, MX, 3250000, 900000, 38100, fill=GOLD)
    text(s, MX, 3450000, CW, 700000, title, 36, True, WHITE)
    text(s, MX, 4250000, CW, 500000, sub, 15, False, RGBColor(0xCF, 0xCF, 0xCF))
    notes(s, speech)
    return s


def tag(s, x, y, label, fill=TEAL, color=WHITE, size=10, w=None):
    w = w or (len(label) * 150000 + 300000)
    b = box(s, x, y, w, 300000, fill=fill)
    tf = b.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    font(r, size, True, color)
    return b


def card(s, x, y, w, h, head, body, num=None, fill=WHITE, head_size=16, body_size=12, accent=True):
    box(s, x, y, w, h, fill=fill, line=LINE)
    if accent:
        box(s, x, y, w, 50800, fill=GOLD)
    ty = y + 280000
    if num:
        text(s, x + 280000, ty, w - 560000, 300000, num, 11, True, GOLD)
        ty += 330000
    text(s, x + 280000, ty, w - 560000, 800000, head, head_size, True)
    text(s, x + 280000, ty + (head_size * 12700 * 2 + 250000 if '\n' in head else head_size * 12700 + 280000),
         w - 560000, h - (ty - y) - 400000, body, body_size, False, SUB, spacing=1.45)


def quote(s, x, y, w, h, q, who):
    box(s, x, y, w, h, fill=CREAM)
    text(s, x + 300000, y + 150000, 600000, 600000, '“', 54, True, TERRA)
    text(s, x + 300000, y + 750000, w - 600000, h - 1300000, q, 14, False, INK, spacing=1.5)
    text(s, x + 300000, y + h - 500000, w - 600000, 300000, who, 11, True, SUB, PP_ALIGN.RIGHT)


# ---- A案デザイン：アイボリー地＋ティールの骨格＋テラコッタの強調 ---------------------
def new(dark=False):
    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = IVORY
    page[0] += 1
    return s


def base(label, title, speech, plain=False):
    s = new()
    box(s, 0, 0, Emu(W), Emu(76200), fill=TEAL)
    box(s, MX, 600000, 76200, 820000, fill=TEAL)
    text(s, MX + 230000, 560000, CW, 260000, label, 10.5, True, TEAL)
    text(s, MX + 230000, 840000, CW - 230000, 600000, title, 26, True, INK)
    box(s, MX, 6380000, CW, 9525, fill=LINE)
    text(s, MX, 6480000, 4000000, 200000, '© U-Writing All Rights Reserved.', 8, False, SUB)
    pg = box(s, W - MX - 420000, 6430000, 420000, 300000, fill=TEAL, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, W - MX - 420000, 6430000, 420000, 300000, str(page[0]), 11, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    notes(s, speech)
    return s


def section(no, title, sub, speech):
    s = base(no or 'SECTION', title, speech)
    text(s, MX, 1900000, CW, 500000, sub, 16, False, SUB)
    return s
# ================================================================ slides
# 1 表紙 ---------------------------------------------------------------
s = new()
box(s, W - 4200000, 0, 4200000, H, fill=TEAL)
_ph = Image.open(f'{A}/profile.jpg').convert('RGB')
_sz = min(_ph.size)
_ph = _ph.crop(((_ph.size[0] - _sz) // 2, 0, (_ph.size[0] - _sz) // 2 + _sz, _sz)).resize((900, 900))
from PIL import ImageDraw as _D
_m = Image.new('L', (900, 900), 0)
_D.Draw(_m).ellipse((0, 0, 899, 899), fill=255)
_c = Image.new('RGBA', (900, 900), (0, 0, 0, 0))
_c.paste(_ph, (0, 0), _m)
_c.save(f'{A}/profile_circle.png')
ring = box(s, W - 3650000, 1350000, 3100000, 3100000, fill=IVORY, shape=MSO_SHAPE.OVAL)
pic(s, f'{A}/profile_circle.png', W - 3550000, 1450000, w=2900000)
text(s, W - 4200000, 4750000, 4200000, 700000, [('集客プランナー / ライター', 12, True, WHITE), ('上原 龍之介', 22, True, WHITE)], align=PP_ALIGN.CENTER)
box(s, 0, 0, W - 4200000, 76200, fill=TERRA)
text(s, MX, 1450000, 6500000, 300000, 'U-ライティング サービスのご案内', 13, True, TEAL)
box(s, MX, 1880000, 700000, 50800, fill=TERRA)
text(s, MX, 2150000, 6600000, 1800000, ['マーケティング迷子を、', '減らす。'], 40, True, INK, spacing=1.2)
text(s, MX, 4050000, 6600000, 400000, '薬機法を守りながら“売れる”広告づくりと集客支援', 15, False, SUB)
text(s, MX, 5600000, 6600000, 300000, 'U-Writing', 11, True, TEAL)
notes(s, """
【画面共有はまだしない】①アイスブレイクは顔を見て話す。

▼アイスブレイク（1つ選ぶ）
・紹介で会う場合：「○○さんとはどういうご縁なんですか？」（共通の知人の話）
・マッチングサービス経由：「プロフィールで△△を拝見して、気になっていたんです」
・交流会で会った場合：「先日の□□でお話しした件の続きなんですが」

▼自分から連絡した場合の理由（必ず具体的に一言）
・人をつなぐ方へ：「○○さんの周りにいる美容・健康系の事業者さんのお役に立てると思って、お会いしたかったんです」
・見込み客へ：「○○さんの（Instagram／LP／チラシ）を拝見して、△△の伝え方がもったいないと感じたんです」
　「もったいない」はOK。「間違っている」はNG。
""")


# 2 本日の流れ -----------------------------------------------------------
s = base('AGENDA', '本日の流れ', """
今日はまず○○さんのお仕事を伺って、そのあと私の話をさせてください。
私は「マーケティング迷子を減らす」をテーマにしているので、その観点でお役に立てることがないか、一緒に探せたら嬉しいです。
""")
steps = [('01', '○○様のお仕事', 'お仕事の内容や、集客の今を\n伺います'),
         ('02', '私の仕事と想い', 'なぜこの仕事をしているか、\nどう広告を作るかをお話しします'),
         ('03', 'お役に立てること', 'サービスと強み、\n実績をご紹介します'),
         ('04', '次の一歩', '今日のお話をもとに\n進め方を決めます')]
cw, gap = 2520000, 196000
for i, (n, h_, b_) in enumerate(steps):
    x = MX + i * (cw + gap)
    box(s, x, 1900000, cw, 3600000, fill=CREAM if i == 0 else WHITE, line=LINE)
    text(s, x + 280000, 2150000, cw, 700000, n, 40, True, TERRA)
    text(s, x + 280000, 3000000, cw - 400000, 500000, h_, 18, True)
    text(s, x + 280000, 3650000, cw - 400000, 1200000, b_.split('\n'), 12, False, SUB, spacing=1.5)
text(s, MX, 5700000, CW, 400000, '※ 最初は○○様のお話を中心に伺います', 12, False, SUB)

# 3 伺いたいこと（ヒアリング補助） ---------------------------------------
s = base('HEARING', 'まずは○○様のことを教えてください', """
【任意。共有しなくてもよい】ヒアリングの手元メモとして使う。
提案したくなっても我慢。相手の課題は一言でメモしておく。

▼見込み客の場合に追加（課題の深さ・お金・時期）
・その課題で、どれくらい損をしていそうですか？
・今、集客にどれくらいかけていますか？
・いつまでに解決したいですか？

▼③に移るタイミング（どちらか）
・相手の課題を要約して「そういうことですね」と同意が取れたとき
・開始から20〜25分が経ったとき
「お話を伺って、○○さんは△△のところで少し迷子になっているように感じました。私がなぜこの仕事をしているかも含めて、少しお話しさせてください。」
""")
qs = [('お仕事のこと', ['どんなお客様に、何を提供されていますか？', '今のお仕事を始めたきっかけは？']),
      ('集客のこと', ['「これで合ってるのかな」と感じていることは？', 'やってみたけど、効果が分からなかったものは？', '逆に、うまくいったと感じたのは？']),
      ('お知り合いのこと', ['普段どんな業種の方とつながっていますか？', '集客に困っていそうな方はいますか？'])]
cw, gap = 3400000, 234000
for i, (h_, items) in enumerate(qs):
    x = MX + i * (cw + gap)
    card(s, x, 1800000, cw, 4300000, h_, ['・' + q for q in items], num=f'0{i+1}', body_size=13)

# 5 自己紹介 -------------------------------------------------------------
s = base('PROFILE', '上原 龍之介', """
上原龍之介です。美容・健康分野を中心に、薬機法を守りながら“売れる”広告を作っています。
ライター歴は6年以上で、今は公式LINEやMeta広告の運用まで支援しています。
（30秒で終える。略歴は聞かれたら答える）

→ 理念へのつなぎ
私がこの仕事を続けている理由を、少しだけお話しさせてください。
""")
pic(s, f'{A}/profile.jpg', MX, 1650000, w=3900000)
x0 = MX + 3900000 + 500000
text(s, x0, 1700000, 6000000, 300000, '集客プランナー / ライター', 13, True, GOLD)
text(s, x0, 2050000, 6300000, 1100000, ['美容・健康分野で、', '薬機法を守りながら“売れる”広告を作っています。'], 20, True, spacing=1.35)
facts = [('6年以上', 'ライター歴'), ('LP・LINE・広告', '企画から運用まで'), ('薬機法・景表法', '配慮した表現設計')]
fw = 2000000
for i, (big, small) in enumerate(facts):
    x = x0 + i * (fw + 150000)
    box(s, x, 3450000, fw, 1050000, fill=CREAM)
    text(s, x, 3600000, fw, 450000, big, 16, True, INK, PP_ALIGN.CENTER)
    text(s, x, 4080000, fw, 300000, small, 11, False, SUB, PP_ALIGN.CENTER)
text(s, x0, 4800000, 6300000, 1300000, [
    ('略歴', 11, True, GOLD),
    ('ゲーム制作の専門学校を卒業 → ゲームアプリのメールサポート（2年）', 11, False, SUB),
    ('→ 建築積算（3.5年） → フリーランスライターとして独立', 11, False, SUB),
    ('→ 現在：ライター兼集客プランナー', 11, False, SUB)], spacing=1.5)

# 8 マーケティング迷子 -----------------------------------------------------
s = base('MY PHILOSOPHY', '仕事の理念：マーケティング迷子を減らす', """
（自己紹介のあと、ここで自分の想いを話す。2分以内）
私がこの仕事を続けている理由を、少しだけお話しさせてください。
いろんな事業者さんとお話しする中で、「何十万円もかけてホームページを作ったけど、問い合わせが来ない」「高い講座やコンサルを受けたけど、成果につながらなかった」という話を本当によく聞くんです。
それでも皆さん諦めずに、自力で頑張ろうとする。でも自力だと時間がかかるし、専門家に頼むとまたお金がかかる。
頑張っているのに、何にお金と時間を使えばいいか分からなくなっている。
私はこれを「マーケティング迷子」と呼んでいます。
（ここで一拍置く）

【注意】「支援業者は高すぎる」という言い方は絶対にしない。同席者や紹介先にコンサル・制作会社がいる。
""")
voices_ = [('高いお金をかけたのに、\n成果が出ない', '数十万円かけてHP・LPを作った、講座やコンサルを受けた。それでも問い合わせが増えない。'),
           ('自力で頑張るが、\n結果まで遠回りに', 'もう一度自分で集客に取り組むものの、手探りのため結果が出るまで時間がかかる。'),
           ('専門家に頼みたいが、\n費用が重い', '早く結果を出すには専門家の力が必要。でも、また大きなお金がかかってしまう。')]
cw = 3400000
for i, (h_, b_) in enumerate(voices_):
    card(s, MX + i * (cw + 234000), 1700000, cw, 2350000, h_, b_, num=f'よく聞く声 0{i+1}', head_size=15, body_size=12)
box(s, MX, 4300000, CW, 1750000, fill=TEAL)
text(s, MX, 4450000, CW, 400000, '頑張っているのに、何にお金と時間を使えばいいか分からない。私はこれを', 15, False, WHITE, PP_ALIGN.CENTER)
text(s, MX, 4900000, CW, 700000, '「マーケティング迷子」', 34, True, TERRA, PP_ALIGN.CENTER)
text(s, MX, 5550000, CW, 400000, 'と呼んでいます。一人でも減らしたい。それが私の仕事の理由です。', 15, False, WHITE, PP_ALIGN.CENTER)

# 10 私の約束 --------------------------------------------------------------
s = base('MY PROMISE', 'だから私は、こう決めています', """
広告は一種の「投資」です。いきなり大金をつぎ込まず、まずは少額から試すのが大原則。
だから私は、本当に必要なものだけを作ると決めています。
ご相談を受けても、作らなくていいと思ったら正直にそう言いますし、
まずは1回の相談だけ、という使い方もできるようにしています。誇大広告もしません。
マーケティング迷子を一人でも減らしたい。それが、私の仕事の理由です。

→ 次は「広告づくりの考え方」へ：
では、実際に私がどう広告を作っているかをお話しします。
""")
promises = [('本当に必要なものだけを作る', 'ご依頼の前に、それが本当に必要かを一緒に考えます。'),
            ('作らなくていいなら、正直に言う', '不要だと判断したものは、おすすめしません。'),
            ('1回の相談から使える', '契約の縛りはなく、必要なときだけご利用いただけます。'),
            ('誇大広告はしない', '薬機法・景表法を守り、消費者とお客様が笑顔でつながる広告を作ります。')]
cw, ch = 5180000, 1850000
for i, (h_, b_) in enumerate(promises):
    x = MX + (i % 2) * (cw + 308000)
    y = 1750000 + (i // 2) * (ch + 200000)
    box(s, x, y, cw, ch, fill=WHITE, line=LINE)
    box(s, x, y, 76200, ch, fill=GOLD)
    text(s, x + 350000, y + 300000, 800000, 500000, f'0{i+1}', 24, True, TERRA)
    text(s, x + 1100000, y + 330000, cw - 1350000, 500000, h_, 17, True)
    text(s, x + 1100000, y + 900000, cw - 1350000, 800000, b_, 12, False, SUB, spacing=1.5)


# 広告づくりの考え方：導入 ---------------------------------------------------
s = base('MY METHOD', '売れる広告は、センスではなく「論理」で作る', """
では、実際に私がどう広告を作っているかをお話しします。
広告は投資なので、なんとなくの感覚では作りません。誰に、何を、どこで、どう届けるか。順番に論理を組み立ててから作ります。
私のサイトでは「集客がうまくいかない代表的な理由」として載せているんですが、裏返すと、これが私の広告づくりの5つの原則です。

【注意】診断やクイズにはしない。あくまで「私の作り方」として話す。
相手がヒアリングで話したことに近い原則があれば、「先ほどの△△のお話は、まさにここなんですが」と軽く触れる程度に。
""")
principles = [('誰に', '「商品のすごさ」ではなく、「使った後の変化」を伝える'),
              ('何を', '「どこでも言える魅力」ではなく、「あなたに頼む理由」を伝える'),
              ('どう', '「雰囲気」ではなく、「行動したくなる言葉」で伝える'),
              ('どこで', '「なんとなくの場所」ではなく、「お客様がいる場所」で伝える'),
              ('どう直す', '「流行り」ではなく、「仮説と数字」で方針を決める')]
for i, (k, v) in enumerate(principles):
    y = 1650000 + i * 880000
    box(s, MX, y, CW, 760000, fill=WHITE, line=LINE)
    box(s, MX, y, 76200, 760000, fill=GOLD)
    text(s, MX + 300000, y + 110000, 900000, 540000, f'0{i+1}', 22, True, TERRA, anchor=MSO_ANCHOR.MIDDLE)
    tag(s, MX + 1150000, y + 230000, k, fill=TEAL, size=11, w=1300000)
    text(s, MX + 2750000, y + 110000, CW - 3000000, 540000, v, 17, True, INK, anchor=MSO_ANCHOR.MIDDLE)

# 広告づくりの考え方：原則①〜⑤（LPスクショ） -----------------------------------
method = [
    ('lp_r1.png', '「商品のすごさ」ではなく「使った後の変化」を伝える',
     '商品の内容や素晴らしさばかりを伝えてしまう',
     '見込み客が知りたいのは、商品そのものより「使った自分がどう変わるか」。まずターゲットの悩みと、使った後の理想の姿を言葉にします。',
     '薬機法で効果を言えない商品でも、「体験」や「変化」なら伝えられる。',
     'CASE 01：寝具を「朝の目覚め」の体験で伝えて目標の279%',
     """例えば美容サロンだと、「最新の美顔器を導入」と書いてしまいがちです。でもお客様が知りたいのは、機械のすごさより「自分の肌がどう変わるか」なんです。
だから私は、まずターゲットの悩みと、使った後の理想の姿を言葉にするところから始めます。"""),
    ('lp_r2.png', '「どこでも言える魅力」ではなく「あなたに頼む理由」を伝える',
     '類似サービスと変わらない魅力を伝えてしまう',
     '「スピード」「安心」「丁寧」はどこでも言える。経歴・実績・想いを掘り起こして、「あなただけの理由」を言葉にします。',
     'ご本人が当たり前だと思っていることほど、強みになる。',
     'CASE 04：求める人物像と代表の想いを記事にして、1週間で内定',
     """「圧倒的スピード」「納得のコスト」「万全のアフターケア」。どの会社にも書けてしまうんですよね。
大事なのは「なぜあなたに頼むのか」。それはたいてい、ご本人の経歴や想いの中に埋まっていて、ご自身では気づいていないことが多いです。私はヒアリングで、そこを掘り起こします。"""),
    ('lp_r3.png', '「雰囲気」ではなく「行動したくなる言葉」で伝える',
     'デザインや雰囲気を重視して広告を作ってしまう',
     'おしゃれでも、言葉が弱ければ人は動かない。先に「読んだ人に何をしてほしいか」と、そのための言葉を決め、デザインはそれを届ける器にします。',
     '※商品によっては、デザイン重視が正解の場合もあります。',
     '',
     """雰囲気はすごく素敵なんです。でも見た人が「で、何をすればいいの？」となってしまう。
私は先に「読んだ人に何をしてほしいか」を決めて、そのための言葉を作ります。デザインは、その言葉を届けるための器だと考えています。

【注意】相手がジムやサロンでも名指しで否定しない。"""),
    ('lp_r4.png', '「なんとなくの場所」ではなく「お客様がいる場所」で伝える',
     'ターゲットが少ない場所で集客を続けてしまう',
     'どんなに良い広告でも、必要とする人がいない場所では反応は出ない。お客様が実際にいる媒体・地域・タイミングを選んでから作ります。',
     'LP・Meta広告・LINEなど、媒体の組み合わせまで設計する。',
     'CASE 03：広告・LP・LINEをつなぎ直して、月の売上が最大150万円',
     """シャッター街で不動産投資のチラシを配っても、反応は出ないですよね。極端な例ですが、同じことがSNSや広告でもよく起きています。
だから私は、作る前に「お客様がどこにいるか」を決めます。LPだけでなく、広告やLINEとのつなぎ方まで考えます。"""),
    ('lp_r5.png', '「流行り」ではなく「仮説と数字」で方針を決める',
     '何も考えずに集客方針を決めてしまう',
     '流行りを追うだけだと、結果を見ても何を直せばいいか分からない。「これで何人来るはず」という仮説を立てて始め、数字を見て一つずつ直します。',
     '作って終わりではなく、改善まで伴走する。',
     '',
     """SNSもLPも新しいツールも、全部頑張っているのに、どれが効いているのか分からない。まさに「マーケティング迷子」の状態です。
私は始める前に「これで何人来るはず」という仮説を立てます。そうすれば、結果を見て何を直せばいいかが分かります。

→ ここから「なので私は、作る（制作）と考える（コンサル）の両方をやっています」とサービスにつなげる。"""),
]
for i, (img, title_, ng, how, point, case_, speech) in enumerate(method):
    s = base(f'MY METHOD  原則 0{i+1}', title_, speech)
    iw, ih = Image.open(f'{A}/{img}').size
    h_ = 4500000
    w_ = int(h_ * iw / ih)
    if w_ > 5400000:
        w_ = 5400000
        h_ = int(w_ * ih / iw)
    pic(s, f'{A}/{img}', MX, 1600000 + (4500000 - h_) // 2, w=w_, border=True)
    rx = MX + w_ + 350000
    rw = W - MX - rx
    text(s, rx, 1650000, rw, 300000, 'よくある失敗', 11, True, SUB)
    text(s, rx, 1950000, rw, 500000, ng, 14, True, SUB, spacing=1.35)
    box(s, rx, 2700000, rw, 2300000, fill=CREAM)
    box(s, rx, 2700000, 76200, 2300000, fill=GOLD)
    text(s, rx + 250000, 2830000, rw - 400000, 300000, '私はこう作る', 11, True, GOLD)
    text(s, rx + 250000, 3150000, rw - 400000, 1800000, how, 13, False, INK, spacing=1.55)
    text(s, rx, 5150000, rw, 450000, point, 12, True, INK, spacing=1.35)


# サービス ---------------------------------------------------------------
s = base('SERVICE', '提供しているサービス', """
この考え方で、2つのサービスを提供しています。
手を動かす時間がない方には、私が代わりに作る「カワリミ」。
自分で動きたいけど方向性に迷っている方には、横で道筋を示す「ミチスジ」です。
○○さんのお話だと、【カワリミ／ミチスジ】が合いそうだと思っています。

【サービスは1つに絞って言い切る】
""")
svc = [('icon_kawarimi.png', '制作代行サービス', 'カワリミ', '作る時間がない方に（制作）',
        ['LP・HP制作／ライティング', '公式LINE構築・運用', 'Meta広告運用', 'チラシ・営業資料など'], '単発 5,000円〜', INK),
       ('icon_michisuji.png', '集客顧問サービス', 'ミチスジ', '方向性に迷っている方に（コンサル）',
        ['月最大4時間のZoom相談', '制作物への添削・アドバイス', '集客施策の分析・仮説検証', '1回だけの相談もOK'], '月額 49,800円〜', GOLD)]
cw = 5180000
for i, (ic, kind, name, who, items, price, col) in enumerate(svc):
    x = MX + i * (cw + 308000)
    box(s, x, 1650000, cw, 4500000, fill=WHITE, line=LINE)
    box(s, x, 1650000, cw, 1000000, fill=col)
    text(s, x + 350000, 1780000, 3500000, 300000, kind, 12, True, WHITE)
    text(s, x + 350000, 2080000, 3500000, 500000, name, 26, True, WHITE)
    pic(s, f'{A}/{ic}', x + cw - 1250000, 1450000, w=1100000)
    tag(s, x + 350000, 2900000, who, fill=CREAM, color=INK, size=13, w=cw - 700000)
    text(s, x + 350000, 3450000, cw - 700000, 1600000, ['✓ ' + t for t in items], 13, False, SUB, spacing=1.6)
    box(s, x + 350000, 5250000, cw - 700000, 9525, fill=LINE)
    text(s, x + 350000, 5400000, cw - 700000, 500000, price, 20, True, TERRA, PP_ALIGN.RIGHT)

# 強み ---------------------------------------------------------------------
s = base('OUR STRENGTHS', 'U-ライティングの2つの強み', """
私の強みは2つです。
1つ目は、制作とコンサルの両方をやっていること。
制作会社だと「作ること」が前提になるし、コンサルだと「作るのは別の会社」になる。私は両方やっているので、○○さんの状況に合わせて「今は作らずに相談だけ」「ここだけ作り直す」といった提案ができます。
2つ目は、差別化まで狙える薬機法対策です。法律を守るだけなら、表現を弱めれば済みます。でもそれだと、他社と同じ言葉になってしまう。私は守りながら、「体験」や「変化」の言葉で他社との違いを作ります。
""")
st = [('制作とコンサル、\nどちらも対応できる',
       'だから、お客様の要望に合わせた提案ができる',
       ['「今は作らず、相談だけ」', '「LPの冒頭だけ作り直す」', '「方針を決めてから、必要なものだけ作る」'],
       '作ることが前提の制作会社でも、作らないコンサルでもない。必要なものだけを選べます。'),
      ('差別化まで狙える\n薬機法対策',
       '守るだけでなく、他社との違いを作る',
       ['効果効能は言わず、「体験」「変化」を言葉にする', '単なる言い換えではなく、ベネフィットを設計', '景表法にも配慮した表現'],
       '法律を意識するほど、どこも同じ表現になりがち。だからこそ差別化のチャンスになります。')]
cw = 5180000
for i, (h_, sub, items, foot) in enumerate(st):
    x = MX + i * (cw + 308000)
    box(s, x, 1650000, cw, 4500000, fill=WHITE, line=LINE)
    box(s, x, 1650000, cw, 50800, fill=GOLD)
    text(s, x + 350000, 1900000, 2000000, 400000, f'STRENGTH 0{i+1}', 11, True, GOLD)
    text(s, x + 350000, 2250000, cw - 700000, 1000000, h_.split('\n'), 22, True, spacing=1.2)
    text(s, x + 350000, 3300000, cw - 700000, 400000, sub, 14, True, GOLD)
    box(s, x + 350000, 3800000, cw - 700000, 1450000, fill=CREAM)
    text(s, x + 550000, 3950000, cw - 1100000, 1250000, ['・' + t for t in items], 13, False, INK, spacing=1.55)
    text(s, x + 350000, 5400000, cw - 700000, 650000, foot, 11, False, SUB, spacing=1.45)

# 16-19 事例 ---------------------------------------------------------------
cases = [
    ('CASE 01', 'クラファン｜ページ企画・ライティング担当', '機能性寝具（羽毛ふとん）',
     '薬機法上、「快眠」「疲労回復」などの効果を直接うたえず、魅力の伝え方に悩んでいた。',
     '法に触れない範囲で「睡眠体験」「朝の目覚め」の変化に焦点を当てたストーリーを設計。写真撮影のディレクションも実施。',
     '279%', '目標30万円に対し、公開1ヶ月で83.9万円', 'case1.png',
     '“快眠”と言えない寝具を、“朝の目覚め”の体験として伝えて、目標の2.8倍になりました。'),
    ('CASE 02', 'クラファン｜ページ企画・ライティング担当', '健康雑貨（ツボ押しシート）',
     'ニッチな商品で、短期間で認知を広げて支援を集める必要があった。',
     'ターゲットの「痛み」「悩み」に寄り添う構成案を作成。利用シーンが浮かぶ撮影ディレクションで、共感を呼ぶページに。',
     '219万円', '公開3日で約50万円、最終的に目標比109%', 'case2.png',
     '公開3日で約50万円、最終的に219万円まで伸びました。'),
    ('CASE 03', '広告運用・LP改善・LINE配信', 'オンライン講座',
     '広告からの流入はあるが、LPでの離脱が多く成約率が伸びない。月商も横ばいが続いていた。',
     'Meta広告の運用に加え、LPの構成とデザインを改修。LINEのステップ配信も見直し、教育から成約までの導線を最適化。',
     '最大150万円', '月平均50万円の売上から改善', None,
     '広告・LP・LINEをつなぎ直して、月平均50万円の売上が、最大で150万円まで伸びました。'),
    ('CASE 04', '採用コンテンツ制作', '建設会社（防水・塗装）の採用',
     '当初の依頼はパンフレット制作。ヒアリングすると、HPの情報不足による採用のミスマッチが本当の課題だった。',
     'パンフレット制作は中止を提案。代わりに、求める人物像と代表の想いを伝える記事をHPに掲載。',
     '1週間', '記事公開から1週間以内に、新卒1名が入社決定', 'case4.png',
     'パンフレットのご依頼でしたが、本当に必要なのは別のものだと分かって、記事に切り替えました。公開1週間で内定が出ています。\n（「必要なものだけを作る」の具体例として話せる）'),
]
for no, tg, title, prob, act, big, small, img, line in cases:
    s = base(f'{no}  導入事例', title, f"""
実際の例を一つだけお見せしますね。
{line}

【事例は相手に近いものを1件だけ】
・美容・健康の物販 → CASE 01 / 02
・講座・スクール・コーチ → CASE 03
・採用に悩む会社 → CASE 04
クラファン（01・02）は「ページ企画・ライティング担当」。聞かれたら担当範囲を正直に答える。
""")
    tag(s, MX, 1450000, tg, w=len(tg) * 140000 + 300000)
    lw_ = 5300000
    text(s, MX, 1950000, lw_, 300000, '課題', 11, True, GOLD)
    text(s, MX, 2250000, lw_, 900000, prob, 12, False, INK, spacing=1.5)
    text(s, MX, 3200000, lw_, 300000, 'やったこと', 11, True, GOLD)
    text(s, MX, 3500000, lw_, 1000000, act, 12, False, INK, spacing=1.5)
    box(s, MX, 4700000, lw_, 1500000, fill=CREAM)
    box(s, MX, 4700000, 76200, 1500000, fill=GOLD)
    text(s, MX + 300000, 4850000, 1000000, 300000, '成果', 11, True, GOLD)
    text(s, MX + 300000, 5100000, lw_ - 500000, 650000, big, 36, True, INK)
    text(s, MX + 300000, 5800000, lw_ - 500000, 300000, small, 12, False, SUB)
    rx, rw = MX + lw_ + 400000, CW - lw_ - 400000
    if img:
        iw, ih = Image.open(f'{A}/{img}').size
        h_ = min(4600000, int(rw * ih / iw))
        w_ = int(h_ * iw / ih)
        pic(s, f'{A}/{img}', rx + (rw - w_) // 2, 1500000 + (4700000 - h_) // 2, w=w_, border=True)
    else:
        for k, (st, desc) in enumerate([('Meta広告運用', '見込み客を集める'), ('LP改修', '離脱を減らし、成約率を上げる'), ('LINEステップ配信', '教育から成約までの導線を整える')]):
            y = 1600000 + k * 1250000
            box(s, rx, y, rw, 1000000, fill=WHITE, line=LINE)
            text(s, rx + 250000, y + 180000, 700000, 600000, f'0{k+1}', 26, True, TERRA)
            text(s, rx + 1000000, y + 200000, rw - 1200000, 400000, st, 16, True)
            text(s, rx + 1000000, y + 580000, rw - 1200000, 300000, desc, 12, False, SUB)
        tri = box(s, rx + rw // 2 - 250000, 5330000, 500000, 250000, fill=GOLD, shape=MSO_SHAPE.ISOSCELES_TRIANGLE)
        tri.rotation = 180
        text(s, rx, 5650000, rw, 450000, '月の売上 最大150万円', 18, True, INK, PP_ALIGN.CENTER)

# 20 お客様の声 --------------------------------------------------------------
s = base('VOICE', 'お客様の声', """
お客様からは、こんな声をいただいています。
特に多いのが「ヒアリングが丁寧」という声で、今日のように最初にしっかりお話を伺うのは、そのためなんです。
""")
voices = [('提案力の高さ', 'こちらの要望も丁寧に聞いてくださり、的確な提案をいただけました。リーズナブルな価格にもかかわらず、しっかりとした内容でとても助かりました。', '健康装具の販売者 Sさん'),
          ('ヒアリング力の高さ', 'これまでの広告課題やお客様の反応を丁寧にヒアリングしてくださり、納得感のある広告に仕上げていただけました。次回もぜひお願いしたいと思える方です。', '脱毛・ホワイトニングサロン Yさん'),
          ('薬機法も対策可', '薬機法が関わる案件でも、安心して上原さんにすべてお任せできました。丁寧な対応と、高い完成度の原稿にとても満足しています。', 'デザイン制作会社 Hさん')]
cw = 3400000
for i, (h_, q, who) in enumerate(voices):
    x = MX + i * (cw + 234000)
    box(s, x, 1700000, cw, 4400000, fill=CREAM)
    tag(s, x + 250000, 1950000, h_, fill=GOLD, w=len(h_) * 170000 + 300000)
    text(s, x + 250000, 2450000, 500000, 600000, '“', 48, True, TERRA)
    text(s, x + 250000, 3050000, cw - 500000, 2300000, q, 13, False, INK, spacing=1.6)
    text(s, x + 250000, 5550000, cw - 500000, 300000, who, 11, True, SUB, PP_ALIGN.RIGHT)

# 21 その他の実績 ------------------------------------------------------------
s = base('WORKS', 'その他の実績', """
他にも、こんな実績があります。全部はお話ししませんが、○○さんに近いところだと――
（相手の業種に近いものを1〜2件だけ指差す。読み上げない）
・美容・化粧品・サプリ → 化粧品のLINEシナリオ「257通中44件成約、成約率17.12%」
・講座・セミナー → セミナー文章改修「0人→5人」、Meta広告「1人約14,000円で集客」
・採用 → 採用記事「公開1週間で新卒1人入社」
・物販 → クラファン3件
""")
groups = [('LINE・広告', [('17.12%', '化粧品の販促LINEシナリオ', '257通中44件が成約・単価2,000円'),
                        ('約14,000円', '講座の集客にMeta広告を活用', '1人あたりの集客単価'),
                        ('約6,750円', 'セールス代行LPのライティング', '約54,000円で8件のLINE登録')]),
          ('文章・記事', [('0人→5人', 'コーチングのセミナー文章を改修', '集客数'),
                        ('新卒1人', '防水・塗装会社の採用記事', '公開1週間以内に入社決定'),
                        ('60→70サイト', '転職サービスのPR記事を企画', '記事の平均転載数')]),
          ('クラファン', [('約30万円', 'PCガジェット', '公開3日以内（ライティング担当）'),
                       ('2件', '機能性寝具・健康雑貨', 'CASE 01・02でご紹介')])]
cw = 3400000
for i, (g, items) in enumerate(groups):
    x = MX + i * (cw + 234000)
    tag(s, x, 1600000, g, fill=TEAL, w=cw)
    for k, (big, what, sub) in enumerate(items):
        y = 2050000 + k * 1380000
        box(s, x, y, cw, 1250000, fill=WHITE, line=LINE)
        text(s, x + 220000, y + 150000, cw - 440000, 450000, big, 20, True, TERRA)
        text(s, x + 220000, y + 600000, cw - 440000, 300000, what, 12, True)
        text(s, x + 220000, y + 900000, cw - 440000, 300000, sub, 10, False, SUB)
text(s, MX, 6150000, CW, 250000, '※ サービスサイト（https://www.u-promotion.com/）掲載の実績より', 9, False, SUB)

# 料金一覧（制作代行・顧問プラン・各種制作） --------------------------------------
s = base('PRICE', '料金一覧', """
料金はこちらです。制作代行は単発で5,000円から、顧問プランは月額49,800円からです。
各種制作の目安は下の表のとおりです。今日のお話だと、【○○】が近いと思います。正式な金額は、伺った内容をもとにお見積りしますね。
顧問プランの契約中は、制作代行が20%引きになります。
（全部説明しない。相手に関係する行だけ指差す）
""")
plans = [('制作代行「カワリミ」', '単発', '5,000円〜', '広告物の制作・運用を代行'),
         ('集客顧問「ミチスジ」顧問プラン', '月額', '49,800円〜', '月4時間までのZoom相談')]
cw = 5180000
for i, (name, unit, price, desc) in enumerate(plans):
    x = MX + i * (cw + 308000)
    box(s, x, 1550000, cw, 1150000, fill=CREAM)
    box(s, x, 1550000, 76200, 1150000, fill=GOLD if i else INK)
    text(s, x + 300000, 1680000, cw - 600000, 350000, name, 15, True)
    text(s, x + 300000, 2100000, cw - 2900000, 350000, desc, 11, False, SUB)
    text(s, x + cw - 2600000, 2020000, 600000, 400000, unit, 12, False, SUB)
    text(s, x + cw - 2100000, 1900000, 1850000, 600000, price, 24, True, TERRA, PP_ALIGN.RIGHT)
menu = [('Web制作', 'LP制作・HP制作（Wix／ペライチ／WordPress）', '100,000円〜'),
        ('Web制作', 'LPライティング（構成・原稿・ワイヤー作成）', '80,000円〜'),
        ('SNS・広告', '公式LINE構築＋運用（配信代行含む）', '初月50,000円〜／月額30,000円〜'),
        ('SNS・広告', 'Meta広告運用（Facebook／Instagram）', '初期30,000円〜／月額 広告費の20%'),
        ('紙媒体', 'チラシ制作（片面／両面）', '30,000円〜／50,000円〜'),
        ('その他', 'スピーチ原稿／営業資料', '5,000円〜／50,000円〜')]
tbl = s.shapes.add_table(len(menu) + 1, 3, Emu(MX), Emu(2900000), Emu(CW), Emu(3000000)).table
for j, wd in enumerate((1700000, 5300000, 3668000)):
    tbl.columns[j].width = Emu(wd)
for i, row_ in enumerate([('カテゴリー', '各種制作メニュー', '価格（税込）')] + menu):
    tbl.rows[i].height = Emu(400000)
    for j, v in enumerate(row_):
        c = tbl.cell(i, j)
        c.fill.solid()
        c.fill.fore_color.rgb = TEAL if i == 0 else (CREAM if i % 2 == 0 else WHITE)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        c.margin_left = c.margin_right = Emu(150000)
        c.margin_top = c.margin_bottom = Emu(30000)
        p = c.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT if (j == 2 and i) else PP_ALIGN.LEFT
        r = p.add_run()
        r.text = v
        font(r, 11 if i == 0 else 12, i == 0 or j == 2, WHITE if i == 0 else (GOLD if j == 2 else INK))
text(s, MX, 6080000, CW, 300000, '※ 価格は税込・目安です。内容やボリュームに応じて事前にお見積りします。「ミチスジ」顧問プラン契約中は、制作代行の費用が20% OFF。', 9, False, SUB)

# 24 次の一歩 ---------------------------------------------------------------
s = base('NEXT STEP', '次の一歩', """
【必ず具体的な行動で終える。「またよろしくお願いします」で終わらせない】

▼契約を目指す場合：その場で提案・見積もりを送る日を決める（無料）
今日伺った内容で、必要なものだけのご提案とお見積りをお送りします。○日までにお送りしてよいですか？
・まだ迷っている相手には、相談プラン（14,800円〜）を選択肢として出す。

▼紹介を目指す場合 → 次のセクション（ご紹介のお願い）へ
""")
text(s, MX, 1600000, CW, 500000, '本日のお話をもとに、必要なものだけをご提案します。', 18, True)
nexts = [('おすすめ', 'ご提案・お見積り', '無料', '今日伺った課題に対して、制作・顧問のご提案と見積りをお送りします。'),
         ('まだ迷っている方', '相談プラン', '14,800円〜', '最大2時間のZoom相談で、現状と優先順位を一緒に整理します。'),
         ('お知り合いに', 'ご紹介', '', '集客や広告でお困りの方がいれば、ぜひおつなぎください。')]
cw = 3400000
for i, (lab, h_, price, b_) in enumerate(nexts):
    x = MX + i * (cw + 234000)
    box(s, x, 2350000, cw, 3600000, fill=CREAM if i == 0 else WHITE, line=GOLD if i == 0 else LINE, lw=2 if i == 0 else 1)
    tag(s, x + 250000, 2600000, lab, fill=GOLD if i == 0 else INK, w=len(lab) * 170000 + 300000)
    text(s, x + 250000, 3150000, cw - 500000, 500000, f'0{i+1}  {h_}', 20, True)
    if price:
        text(s, x + 250000, 3750000, cw - 500000, 500000, price, 22, True, TERRA)
    text(s, x + 250000, 4450000, cw - 500000, 1300000, b_, 13, False, SUB, spacing=1.5)

# 紹介：5つの理由をそのまま紹介のきっかけに ------------------------------------
s = base('REFERRAL KEYWORDS', 'この5つで悩んでいる方がいたら、思い出してください', """
（紹介を目指す場合、または契約の話のあとに）最後に一つ、お願いがあります。
今日お話しした5つのよくある失敗、どれかで悩んでいる方が周りにいたら、私を思い出してください。
「チラシやLPを作ったけど反応がない」「何から手をつければいいか分からない」と言っている方は、たいていこのどれかです。
合言葉は「マーケティング迷子」です。
""")
voices_kw = [('01', '「商品の良さは分かってもらえるはずなのに、売れない」'),
             ('02', '「同業と何が違うのか、うまく説明できない」'),
             ('03', '「おしゃれなHPを作ったのに、問い合わせが来ない」'),
             ('04', '「SNSを頑張っているのに、反応がない」'),
             ('05', '「何をやっても効いているのか分からない」')]
for i, (n, q) in enumerate(voices_kw):
    y = 1600000 + i * 780000
    box(s, MX, y, CW, 680000, fill=CREAM)
    box(s, MX, y, 76200, 680000, fill=GOLD)
    text(s, MX + 300000, y + 100000, 700000, 480000, n, 18, True, TERRA, anchor=MSO_ANCHOR.MIDDLE)
    text(s, MX + 1100000, y + 100000, CW - 1400000, 480000, q, 16, True, INK, anchor=MSO_ANCHOR.MIDDLE)
text(s, MX, 5600000, CW, 500000, [('合言葉は「マーケティング迷子」', 20, True, TERRA)], align=PP_ALIGN.CENTER)

# 27 ご紹介いただきたい方 ----------------------------------------------------
s = base('WHO TO INTRODUCE', 'ご紹介いただきたい方', """
特に、美容サロン、整体院、医療クリニック、健康食品を扱っている事業者さんとは相性がいいです。
協業できるパートナーさんも探しています。
""")
cols = [('お客様として', '美容・健康分野で集客に悩む事業者', ['美容サロン・脱毛サロン', '整体院・治療院', '医療クリニック', '健康食品・化粧品・健康雑貨の販売者', '講座・スクール・コーチ']),
        ('協業パートナーとして', 'ライティング・薬機法対応を任せたい方', ['Web制作会社・デザイン会社', '経営・財務コンサルタント', '医療・美容機器ディーラー', 'レンタルスペース・店舗物件の不動産仲介', '経営者コミュニティ・マッチングサービス運営'])]
cw = 5180000
for i, (h_, sub, items) in enumerate(cols):
    x = MX + i * (cw + 308000)
    box(s, x, 1650000, cw, 4450000, fill=WHITE, line=LINE)
    box(s, x, 1650000, cw, 50800, fill=GOLD if i == 0 else INK)
    text(s, x + 350000, 1950000, cw - 700000, 450000, h_, 20, True)
    text(s, x + 350000, 2450000, cw - 700000, 300000, sub, 12, False, SUB)
    text(s, x + 350000, 3000000, cw - 700000, 2900000, ['● ' + t for t in items], 14, False, INK, spacing=1.7)

# 28 紹介パートナー制度 ------------------------------------------------------
s = base('REFERRAL PROGRAM', '紹介パートナー制度', """
ご紹介がご契約につながった場合は、紹介料をお支払いしています。
顧問プランなら契約が続く限り毎月20%、制作代行なら10万円以上のご契約で10%です。
商談やご提案はすべて私が行うので、つないでいただくだけで大丈夫です。
""")
fees = [('集客顧問「ミチスジ」顧問プラン', '20%', '契約金額の20%を毎月', '契約が続く限り、着金確認後にお支払い'),
        ('制作代行「カワリミ」', '10%', '契約金額の10%', '契約金額10万円以上の場合。着金確認後にお支払い')]
for i, (h_, pct, a_, b_) in enumerate(fees):
    y = 1700000 + i * 1650000
    box(s, MX, y, 5300000, 1450000, fill=CREAM)
    text(s, MX + 300000, y + 220000, 4800000, 350000, h_, 13, True)
    text(s, MX + 300000, y + 600000, 1700000, 750000, pct, 40, True, TERRA)
    text(s, MX + 2000000, y + 680000, 3200000, 300000, a_, 13, True)
    text(s, MX + 2000000, y + 1000000, 3200000, 300000, b_, 10, False, SUB)
flow = [('ご紹介・おつなぎ', 'お困りの方をご紹介ください'), ('商談・ご提案', 'ヒアリングと提案はすべて私が行います'),
        ('ご契約・業務開始', 'お客様と直接契約します'), ('紹介料のお支払い', '入金確認後にお支払いします')]
fx = MX + 5700000
for k, (h_, b_) in enumerate(flow):
    y = 1700000 + k * 850000
    box(s, fx, y, CW - 5700000, 720000, fill=WHITE, line=LINE)
    o = box(s, fx + 200000, y + 160000, 400000, 400000, fill=TEAL, shape=MSO_SHAPE.OVAL)
    text(s, fx + 200000, y + 160000, 400000, 400000, str(k + 1), 12, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    text(s, fx + 800000, y + 120000, 3800000, 300000, h_, 13, True)
    text(s, fx + 800000, y + 400000, 3800000, 300000, b_, 10, False, SUB)
text(s, MX, 5250000, CW, 700000, ['メリット：自社サービスの付加価値として提案できる／リスクなく新たな収益源になる／専門的な薬機法対応を任せられる'], 11, False, SUB)

# 29 クロージング ------------------------------------------------------------
s = base('THANK YOU', '本日はありがとうございました', '')
box(s, MX, 1750000, CW, 1900000, fill=CREAM)
box(s, MX, 1750000, 76200, 1900000, fill=GOLD)
text(s, MX, 2050000, CW, 400000, '最後に教えてください', 14, False, SUB, PP_ALIGN.CENTER)
text(s, MX, 2550000, CW, 700000, '私から○○様にご紹介できるのは、どんな方ですか？', 24, True, INK, PP_ALIGN.CENTER)
box(s, 3600000, 4100000, 4992000, 1650000, fill=WHITE, line=LINE)
text(s, 3850000, 4300000, 3200000, 400000, [('U-ライティング 上原 龍之介', 13, True, INK), ('サービスサイト', 10, False, SUB), ('https://www.u-promotion.com/', 11, True, GOLD)], spacing=1.5)
pic(s, f'{A}/qr.png', 7300000, 4350000, w=1150000)
notes(s, """
▼紹介を目指す場合：その場で名前を出してもらう
今日のお話で、思い浮かんだ方はいらっしゃいますか？
・出なければ「1週間後に一度ご連絡してもいいですか？」とフォローの約束を取る。

▼最後に必ず
逆に、私から○○さんにご紹介できるのはどんな方ですか？

▼打ち合わせ後（当日中）
・お礼の連絡と、決まった次の行動を文字で送る。
・相手から聞いた「紹介してほしい人」を控え、こちらから先に紹介できる人を探す。
""")



# 14 ミチスジ ---------------------------------------------------------------
s = base('APPENDIX  MICHISUJI', '集客顧問「ミチスジ」― 迷子にならない道筋を、一緒に', """
ミチスジは、ご自身で動きたい方の横で、道筋を示し続けるサービスです。
いきなり顧問でなくても大丈夫です。まずは1回の相談から、必要なら提案書、続けるなら顧問、と段階的に使えます。
顧問中は、制作代行が20%引きになります。
""")
plans = [('相談プラン', '14,800円〜', '1回限り', ['最大2時間のZoom相談', '現状の広告物を添削', 'まず何から手をつけるか、優先順位をご提案'], 'まず1回、現状を整理したい方'),
         ('提案プラン', '49,800円〜', '単発', ['現状分析・競合リサーチ', '具体的な施策の提案書を作成', '提案プレゼン・質疑応答'], '集客の方向性を資料で欲しい方'),
         ('顧問プラン', '49,800円〜', '月額', ['月最大4時間のZoom相談', '制作物への添削・アドバイス', '施策の分析・仮説検証', 'カワリミ 20%OFF'], '継続して伴走してほしい方')]
cw = 3400000
for i, (name, price, unit, items, who) in enumerate(plans):
    x = MX + i * (cw + 234000)
    y = 2300000 - i * 250000
    hh = 3850000 + i * 250000
    box(s, x, y, cw, hh, fill=CREAM if i < 2 else WHITE, line=GOLD if i == 2 else LINE, lw=1.5 if i == 2 else 1)
    text(s, x + 250000, y + 230000, cw - 500000, 300000, f'STEP {i+1}', 11, True, GOLD)
    text(s, x + 250000, y + 520000, cw - 500000, 450000, name, 20, True)
    text(s, x + 250000, y + 1050000, cw - 500000, 500000, [(f'{unit}  ', 11, False, SUB)], 11)
    text(s, x + 900000, y + 960000, cw - 1150000, 500000, price, 20, True, TERRA)
    text(s, x + 250000, y + 1600000, cw - 500000, 1700000, ['・' + t for t in items], 12, False, INK, spacing=1.5)
    text(s, x + 250000, y + hh - 450000, cw - 500000, 300000, '→ ' + who, 11, True, SUB)
text(s, MX, 1600000, 5000000, 400000, '小さく始めて、必要なら次へ。', 14, True, INK)

prs.save(OUT)
print('saved', OUT, 'slides:', len(prs.slides))
