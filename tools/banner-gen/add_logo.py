"""Gemini で作った画像（コピー・CTA入り、左下は空き）に、正規ロゴを重ねて書き出す。

ロゴは AI に描かせず、必ずこのスクリプトで正規データを貼る（ESTD.1582 の誤記を防ぐ）。

  # Downloads にある最新の Gemini 画像を raw/ に移す
  py -3.12 tools/banner-gen/add_logo.py pull --raw output/tamaruya/20261029/raw --name c2_text_1x1.jpg

  # ロゴを重ねて書き出す
  py -3.12 tools/banner-gen/add_logo.py logo \
    --image output/tamaruya/20261029/raw/c2_text_1x1.jpg --size 1x1 \
    --out output/tamaruya/20261029/tamaruya_20261029_c2-friday_1x1.png
"""
import argparse
import shutil
from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[2]
LOGO = ROOT / "brands" / "tamaruya" / "assets" / "logo_3631_tamaruya.jpg"
DOWNLOADS = Path.home() / "Downloads"
SIZES = {"1x1": (1080, 1080), "4x5": (1080, 1350), "9x16": (1080, 1920)}


def pull(raw, name):
    """Downloads の最新の Gemini_Generated_Image_* を raw/ に移す。"""
    files = sorted(DOWNLOADS.glob("Gemini_Generated_Image_*"), key=lambda p: p.stat().st_mtime)
    if not files:
        raise SystemExit("Downloads に Gemini_Generated_Image_* がありません（保存途中なら少し待つ）")
    raw = Path(raw)
    raw.mkdir(parents=True, exist_ok=True)
    dest = raw / (name or files[-1].name)
    shutil.move(str(files[-1]), dest)
    return dest


def cover(img, w, h):
    scale = max(w / img.width, h / img.height)
    img = img.resize((round(img.width * scale), round(img.height * scale)), Image.LANCZOS)
    left, top = (img.width - w) // 2, (img.height - h) // 2
    return img.crop((left, top, left + w, top + h))


def logo_plate(width):
    """正規ロゴの文字部分を切り出し、黒の角丸プレートに載せる。"""
    logo = Image.open(LOGO).convert("RGB")
    bbox = logo.convert("L").point(lambda v: 255 if v > 60 else 0).getbbox()
    pad = round((bbox[3] - bbox[1]) * 0.35)
    logo = logo.crop((bbox[0] - pad, bbox[1] - pad, bbox[2] + pad, bbox[3] + pad))
    logo = logo.resize((width, round(logo.height * width / logo.width)), Image.LANCZOS)
    mask = Image.new("L", logo.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, *logo.size], radius=round(logo.height * 0.2), fill=235)
    return logo, mask


def add_logo(image, size, out, logo_width=0.40, margin=0.035):
    w, h = SIZES[size]
    base = cover(Image.open(image).convert("RGB"), w, h)
    # 9:16 は下14%を UI に隠れる領域として避ける
    safe = round(h * 0.14) if size == "9x16" else 0
    logo, mask = logo_plate(round(w * logo_width))
    m = round(w * margin)
    base.paste(logo, (m, h - safe - m - logo.height), mask)
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    base.save(out, "PNG")
    return out


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("pull")
    s.add_argument("--raw", required=True)
    s.add_argument("--name")
    s = sub.add_parser("logo")
    s.add_argument("--image", required=True)
    s.add_argument("--size", choices=SIZES, default="1x1")
    s.add_argument("--out", required=True)
    s.add_argument("--logo-width", type=float, default=0.40, help="画像幅に対するロゴ幅の比率")
    a = p.parse_args()
    if a.cmd == "pull":
        print(pull(a.raw, a.name))
    else:
        print(add_logo(a.image, a.size, a.out, a.logo_width))


if __name__ == "__main__":
    main()
