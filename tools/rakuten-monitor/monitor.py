"""楽天ショップのトップページを取得し、前回スナップショットとの差分をメールで送る。

標準ライブラリのみで動く。GitHub Actions から毎日実行する想定。
必要な環境変数: SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, MAIL_TO（MAIL_FROM は任意）
DRY_RUN=1 ならメールを送らず本文を標準出力に出す。
"""
import json
import os
import re
import smtplib
import sys
import time
import urllib.request
from datetime import datetime, timedelta, timezone
from email.mime.text import MIMEText
from html.parser import HTMLParser
from pathlib import Path

SHOPS = [
    "https://www.rakuten.co.jp/sc-rkt/",
    "https://www.rakuten.co.jp/hattendo-shop/",
    "https://www.rakuten.co.jp/tsujirihei/",
    "https://www.rakuten.co.jp/o-ohsho/",
    "https://www.rakuten.co.jp/lindt-chocolate/",
]
SNAPSHOT_DIR = Path(__file__).parent / "snapshots"
JST = timezone(timedelta(hours=9))
# UA は指定しない（Python 標準）。ブラウザ偽装や独自 UA は Akamai のブロックページ（約40バイト）が返る
MAX_LIST = 30  # 1 カテゴリあたりメールに載せる件数の上限


def clean(text):
    return re.sub(r"\s+", " ", text or "").strip()


class ShopParser(HTMLParser):
    """商品リンク・バナー画像の alt・見出し・お知らせ的テキストを抜き出す。"""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.title = ""
        self.items = {}      # 商品URL -> 商品名
        self.banners = set()  # バナー画像の alt
        self.headings = set()
        self._in_title = False
        self._link = None
        self._link_text = []
        self._heading = None
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in ("script", "style", "noscript"):
            self._skip += 1
        elif tag == "title":
            self._in_title = True
        elif tag == "a":
            href = (a.get("href") or "").split("?")[0].split("#")[0]
            if "item.rakuten.co.jp" in href:
                self._link = href.rstrip("/") + "/"
                self._link_text = []
        elif tag == "img":
            alt = clean(a.get("alt"))
            if self._link and alt:
                self._link_text.append(alt)
            elif alt and len(alt) >= 4:
                self.banners.add(alt)
        elif tag in ("h1", "h2", "h3"):
            self._heading = []

    def handle_endtag(self, tag):
        if tag in ("script", "style", "noscript"):
            self._skip = max(0, self._skip - 1)
        elif tag == "title":
            self._in_title = False
        elif tag == "a" and self._link:
            # ランキング順位は日々変わるので除く（誤検知防止）
            name = re.sub(r"^\d+\s*位\s*", "", clean(" ".join(self._link_text)))
            if name and len(name) > len(self.items.get(self._link, "")):
                self.items[self._link] = name
            elif self._link not in self.items:
                self.items[self._link] = ""
            self._link = None
        elif tag in ("h1", "h2", "h3") and self._heading is not None:
            h = clean("".join(self._heading))
            if h:
                self.headings.add(h)
            self._heading = None

    def handle_data(self, data):
        if self._skip:
            return
        if self._in_title:
            self.title += data
        if self._link is not None:
            self._link_text.append(data)
        if self._heading is not None:
            self._heading.append(data)


def fetch(url):
    req = urllib.request.Request(url, headers={"Accept-Language": "ja"})
    last = None
    for wait in (0, 5, 15):
        time.sleep(wait)
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                raw = r.read()
                charset = r.headers.get_content_charset()
            break
        except Exception as e:  # noqa: BLE001
            last = e
    else:
        raise last
    if not charset:
        m = re.search(rb'charset=["\']?([\w-]+)', raw[:4000], re.I)
        charset = m.group(1).decode() if m else "utf-8"
    return raw.decode(charset, errors="replace")


def snapshot(url):
    p = ShopParser()
    p.feed(fetch(url))
    return {
        "url": url,
        "title": clean(p.title),
        "items": dict(sorted(p.items.items())),
        "banners": sorted(p.banners),
        "headings": sorted(p.headings),
    }


def diff(old, new):
    lines = []
    if old["title"] != new["title"]:
        lines.append(f"■ ページタイトル変更\n  旧: {old['title']}\n  新: {new['title']}")

    def section(label, added, removed):
        for mark, vals in (("追加", added), ("削除", removed)):
            if vals:
                shown = vals[:MAX_LIST]
                more = f"\n  …ほか {len(vals) - MAX_LIST} 件" if len(vals) > MAX_LIST else ""
                lines.append(f"■ {label}の{mark}（{len(vals)}件）\n" + "\n".join(f"  ・{v}" for v in shown) + more)

    oi, ni = old["items"], new["items"]
    section("掲載商品",
            [f"{ni[u] or '(名称なし)'}  {u}" for u in ni if u not in oi],
            [f"{oi[u] or '(名称なし)'}  {u}" for u in oi if u not in ni])
    renamed = [f"{oi[u]} → {ni[u]}  {u}" for u in ni if u in oi and oi[u] and ni[u] and oi[u] != ni[u]]
    section("商品名・表示テキスト変更", renamed, [])
    for key, label in (("banners", "バナー/画像"), ("headings", "見出し")):
        o, n = set(old[key]), set(new[key])
        section(label, sorted(n - o), sorted(o - n))
    return lines


def send(subject, body):
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"] = os.environ.get("MAIL_FROM") or os.environ["SMTP_USER"]
    msg["To"] = os.environ["MAIL_TO"]
    with smtplib.SMTP(os.environ["SMTP_HOST"], int(os.environ.get("SMTP_PORT", "587"))) as s:
        s.starttls()
        s.login(os.environ["SMTP_USER"], os.environ["SMTP_PASSWORD"])
        s.send_message(msg)


def main():
    SNAPSHOT_DIR.mkdir(exist_ok=True)
    today = datetime.now(JST).strftime("%Y-%m-%d")
    blocks, changed, errors = [], 0, 0
    for url in SHOPS:
        shop = url.rstrip("/").rsplit("/", 1)[-1]
        path = SNAPSHOT_DIR / f"{shop}.json"
        try:
            new = snapshot(url)
        except Exception as e:  # noqa: BLE001
            errors += 1
            blocks.append(f"【{shop}】取得失敗: {e}\n{url}")
            continue
        if not new["items"] and not new["banners"]:
            errors += 1
            blocks.append(f"【{shop}】ページ内容が取得できませんでした（ブロック/構造変更の可能性）。前回データは保持します。\n{url}")
            continue
        if path.exists():
            lines = diff(json.loads(path.read_text("utf-8")), new)
            if lines:
                changed += 1
                blocks.append(f"【{shop}】更新あり  {new['title']}\n{url}\n\n" + "\n\n".join(lines))
            else:
                blocks.append(f"【{shop}】更新なし\n{url}")
        else:
            blocks.append(f"【{shop}】初回取得（比較対象なし。明日から差分を通知します）\n商品リンク {len(new['items'])} 件 / バナー {len(new['banners'])} 件\n{url}")
        path.write_text(json.dumps(new, ensure_ascii=False, indent=1) + "\n", "utf-8")

    status = f"更新あり {changed}店舗" if changed else "更新なし"
    if errors:
        status += f" / 取得失敗 {errors}店舗"
    subject = f"[楽天ショップ監視] {today} {status}"
    body = f"{today} 時点の楽天ショップ更新チェック結果です。\n\n" + ("\n\n" + "-" * 40 + "\n\n").join(blocks) + "\n"
    if os.environ.get("DRY_RUN"):
        print(subject + "\n\n" + body)
    else:
        send(subject, body)
        print(subject)
    return 0


if __name__ == "__main__":
    sys.exit(main())
