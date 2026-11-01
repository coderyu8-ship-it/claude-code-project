import sys
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
s = prs.slides[19]

text(s, MX, 4680000, 6000000, 300000, 'ミチスジのサポート内容', 13, True, TEAL)
items = [('Zoom相談', '顧問は月最大4時間。集客施策の壁打ち・進捗確認・次の一手の相談'),
         ('添削・アドバイス', 'チラシ・LP・SNS投稿など、お手持ちの制作物を添削して改善点をお伝えします'),
         ('分析・仮説検証', '施策の結果を分析し、「なぜうまくいかなかったか」「次はどうするか」を一緒に考えます'),
         ('メッセージ相談', 'Zoomの合間も、気になったことをメッセージで気軽に相談できます（無料）')]
cw, gap = 2550000, 156000
for i, (h_, b_) in enumerate(items):
    x = MX + i * (cw + gap)
    box(s, x, 4990000, cw, 960000, fill=WHITE, line=LINE)
    box(s, x, 4990000, 50800, 960000, fill=TEAL)
    text(s, x + 180000, 5050000, cw - 300000, 280000, h_, 12, True, INK)
    text(s, x + 180000, 5340000, cw - 300000, 620000, b_, 9.5, False, SUB, spacing=1.35)

ns = s.notes_slide
body = ("集客顧問ミチスジは、ご自身で動きたい方の横で道筋を示すサービスです。\n"
        "サポート内容は4つ。Zoomでの相談、制作物の添削、施策の分析と仮説検証、それからZoomの合間のメッセージ相談です。\n"
        "まずは1回の相談プラン、方向性を資料で欲しい方は提案プラン、続けて伴走が必要なら顧問プランと、段階的に使えます。\n"
        "（紹介を目指す相手には）顧問プランは、契約が続く限り、毎月の契約金額の20%を紹介料としてお支払いしています。")
if ns.notes_text_frame is not None:
    ns.notes_text_frame.text = body
prs.save(sys.argv[2])
print('saved')
