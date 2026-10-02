#!/usr/bin/env python3
"""GA4 の商品ページ表示回数を「LPリンクの遷移数」シートへ転記する。

サブコマンド:
  discover  クエリ文字列付きのページパスを集計表示する（カテゴリ判定パラメータの特定用）
  verify    指定期間の API 値とシートに手入力済みの値を突き合わせる（書き込みなし）
  run       指定日（既定: 前日）の値をシートに書き込む

認証（どちらか）:
  - GA_OAUTH_CLIENT_ID / GA_OAUTH_CLIENT_SECRET / GA_OAUTH_REFRESH_TOKEN（本人の OAuth 認可）
  - GA_SERVICE_ACCOUNT_JSON（サービスアカウントの JSON キー本文）
"""

import argparse
import calendar
import datetime as dt
import json
import os
import re
import sys
from collections import defaultdict
from pathlib import Path
from urllib.parse import parse_qs, quote, urlsplit
from zoneinfo import ZoneInfo

from google.auth.transport.requests import AuthorizedSession
from google.oauth2 import credentials as user_credentials
from google.oauth2 import service_account

CONFIG_PATH = Path(__file__).with_name("config.json")
SCOPES = [
    "https://www.googleapis.com/auth/analytics.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]
GA_URL = "https://analyticsdata.googleapis.com/v1beta/properties/{}:runReport"
SHEETS_URL = "https://sheets.googleapis.com/v4/spreadsheets/{}"

HEADER_ROWS = 2  # 1行目: カテゴリ（結合セル）, 2行目: 商品名 + 末尾行にスラッグ
CUMULATIVE_ROW = 3
FIRST_DATA_COL = 2  # B 列
MONTH_TOTAL_RE = re.compile(r"^\d{1,2}月合計$")


def load_config():
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def session():
    oauth = [os.environ.get(k, "").strip() for k in
             ("GA_OAUTH_CLIENT_ID", "GA_OAUTH_CLIENT_SECRET", "GA_OAUTH_REFRESH_TOKEN")]
    raw = os.environ.get("GA_SERVICE_ACCOUNT_JSON")
    if all(oauth):
        # サービスアカウント鍵を作れない環境向け: 本人の OAuth リフレッシュトークンで認証する
        client_id, client_secret, refresh_token = oauth
        creds = user_credentials.Credentials.from_authorized_user_info(
            {"client_id": client_id, "client_secret": client_secret, "refresh_token": refresh_token},
            scopes=SCOPES)
    elif raw:
        creds = service_account.Credentials.from_service_account_info(json.loads(raw), scopes=SCOPES)
    else:
        missing = [k for k, v in zip(("GA_OAUTH_CLIENT_ID", "GA_OAUTH_CLIENT_SECRET", "GA_OAUTH_REFRESH_TOKEN"), oauth) if not v]
        sys.exit("認証情報がありません。未設定: " + ", ".join(missing))
    return AuthorizedSession(creds)


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def check(resp):
    if resp.status_code >= 400:
        sys.exit(f"API エラー {resp.status_code}: {resp.text[:2000]}")
    return resp.json()


# ---------- GA4 ----------

def fetch_pageviews(sess, cfg, start, end):
    """{(YYYY-MM-DD, pagePathPlusQueryString): views} を返す。"""
    body = {
        "dateRanges": [{"startDate": start.isoformat(), "endDate": end.isoformat()}],
        "dimensions": [{"name": "date"}, {"name": "pagePathPlusQueryString"}],
        "metrics": [{"name": cfg["metric"]}],
        "limit": 100000,
    }
    data = check(sess.post(GA_URL.format(cfg["property_id"]), json=body))
    out = {}
    for row in data.get("rows", []):
        d, path = (v["value"] for v in row["dimensionValues"])
        out[(f"{d[:4]}-{d[4:6]}-{d[6:]}", path)] = int(row["metricValues"][0]["value"])
    if data.get("rowCount", 0) > len(out):
        sys.exit(f"GA の行数が上限を超えました（{data['rowCount']} 行）。期間を短くしてください")
    return out


def path_slugs(path):
    segs = [s for s in urlsplit(path).path.split("/") if s]
    return {re.sub(r"\.[a-z]+$", "", s) for s in segs}


def aggregate(pageviews, cfg, columns):
    """{date: {列番号: views}} を返す。columns は [(列番号, カテゴリ名, スラッグ)]。"""
    param = cfg["category_param"]
    value_to_cat = {v: k for k, v in cfg["categories"].items()}
    result = defaultdict(lambda: defaultdict(int))
    for (date, path), views in pageviews.items():
        values = parse_qs(urlsplit(path).query).get(param, [])
        cats = {value_to_cat[v] for v in values if v in value_to_cat}
        if not cats:
            continue
        slugs = path_slugs(path)
        for col, cat, slug in columns:
            if cat in cats and slug in slugs:
                result[date][col] += views
    return result


# ---------- Sheets ----------

class Sheet:
    def __init__(self, sess, cfg):
        self.sess = sess
        self.sid = cfg["spreadsheet_id"]
        self.name = cfg["sheet_name"]
        meta = check(sess.get(SHEETS_URL.format(self.sid), params={"fields": "sheets.properties"}))
        props = next((s["properties"] for s in meta["sheets"] if s["properties"]["title"] == self.name), None)
        if props is None:
            sys.exit(f"シート「{self.name}」が見つかりません")
        self.sheet_id = props["sheetId"]
        self.row_count = props["gridProperties"]["rowCount"]
        self.reload()

    def reload(self):
        rng = quote(f"'{self.name}'!A1:ZZ{self.row_count}", safe="")
        data = check(self.sess.get(
            f"{SHEETS_URL.format(self.sid)}/values/{rng}",
            params={"valueRenderOption": "FORMATTED_VALUE"},
        ))
        self.rows = data.get("values", [])

    def cell(self, r, c):
        row = self.rows[r - 1] if r - 1 < len(self.rows) else []
        return row[c - 1] if c - 1 < len(row) else ""

    def columns(self, cfg):
        """[(列番号, カテゴリ名, スラッグ)]。1行目の結合セルは左端にだけ値があるので前方補完する。"""
        cols, cat = [], ""
        width = max(len(self.rows[0]), len(self.rows[1]))
        for c in range(FIRST_DATA_COL, width + 1):
            cat = self.cell(1, c).strip() or cat
            product = self.cell(2, c).strip()
            if not product:
                continue
            slug = product.splitlines()[-1].strip()
            if cat not in cfg["categories"]:
                sys.exit(f"{col_letter(c)}1 のカテゴリ「{cat}」が config.json にありません")
            cols.append((c, cat, slug))
        return cols

    def find_date_row(self, date):
        label = date.strftime("%Y/%m/%d")
        for i, row in enumerate(self.rows, start=1):
            if row and row[0].strip() == label:
                return i
        return None

    def last_used_row(self):
        for i in range(len(self.rows), 0, -1):
            if self.rows[i - 1] and self.rows[i - 1][0].strip():
                return i
        return 0

    def month_total_rows(self):
        return [i for i, row in enumerate(self.rows, start=1) if row and MONTH_TOTAL_RE.match(row[0].strip())]

    def batch_values(self, data):
        body = {"valueInputOption": "USER_ENTERED", "data": data}
        check(self.sess.post(f"{SHEETS_URL.format(self.sid)}/values:batchUpdate", json=body))

    def batch_update(self, requests):
        check(self.sess.post(f"{SHEETS_URL.format(self.sid)}:batchUpdate", json={"requests": requests}))

    def grid(self, r1, r2, c1, c2):
        return {"sheetId": self.sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r2,
                "startColumnIndex": c1 - 1, "endColumnIndex": c2}


def add_month_block(sheet, year, month, last_col):
    """シート末尾に「◯月合計」行と日付行を追加し、書式を直前の月から写す。"""
    totals = sheet.month_total_rows()
    if not totals:
        sys.exit("書式のコピー元になる「◯月合計」行が見つかりません")
    src_total = totals[-1]
    ndays = calendar.monthrange(year, month)[1]
    top = sheet.last_used_row() + 1
    bottom = top + ndays

    requests = []
    if bottom > sheet.row_count:
        requests.append({"appendDimension": {"sheetId": sheet.sheet_id, "dimension": "ROWS",
                                             "length": bottom - sheet.row_count + 50}})
    requests += [
        {"copyPaste": {"source": sheet.grid(src_total, src_total, 1, last_col),
                       "destination": sheet.grid(top, top, 1, last_col), "pasteType": "PASTE_FORMAT"}},
        {"copyPaste": {"source": sheet.grid(src_total + 1, src_total + 1, 1, last_col),
                       "destination": sheet.grid(top + 1, bottom, 1, last_col), "pasteType": "PASTE_FORMAT"}},
    ]
    sheet.batch_update(requests)

    total_row = [f"{month}月合計"] + [
        f"=sum({col_letter(c)}{top + 1}:{col_letter(c)}{bottom})" for c in range(FIRST_DATA_COL, last_col + 1)
    ]
    dates = [[f"{year}/{month:02d}/{d:02d}"] for d in range(1, ndays + 1)]
    sheet.batch_values([
        {"range": f"'{sheet.name}'!A{top}:{col_letter(last_col)}{top}", "values": [total_row]},
        {"range": f"'{sheet.name}'!A{top + 1}:A{bottom}", "values": dates},
    ])
    if bottom > sheet.row_count:
        sheet.row_count = bottom + 50
    sheet.reload()
    print(f"{year}/{month:02d} のブロックを {top}〜{bottom} 行目に追加しました")


def ensure_cumulative(sheet, last_col):
    """累計行を、全ての「◯月合計」行の合計にする。"""
    formulas = [f'=SUMIF($A$4:$A,"*月合計",{col_letter(c)}$4:{col_letter(c)})'
                for c in range(FIRST_DATA_COL, last_col + 1)]
    sheet.batch_values([{
        "range": f"'{sheet.name}'!{col_letter(FIRST_DATA_COL)}{CUMULATIVE_ROW}:{col_letter(last_col)}{CUMULATIVE_ROW}",
        "values": [formulas],
    }])


# ---------- commands ----------

def parse_date(s):
    return dt.date.fromisoformat(s)


def cmd_discover(args, cfg):
    sess = session()
    pv = fetch_pageviews(sess, cfg, parse_date(args.start), parse_date(args.end))
    by_path, keys = defaultdict(int), defaultdict(int)
    for (_, path), v in pv.items():
        if "?" not in path:
            continue
        by_path[path] += v
        for k in parse_qs(urlsplit(path).query):
            keys[k] += v
    print("== クエリパラメータ名（表示回数の合計）==")
    for k, v in sorted(keys.items(), key=lambda x: -x[1]):
        print(f"{v:8d}  {k}")
    print(f"\n== クエリ付きパス 上位 {args.top} ==")
    for p, v in sorted(by_path.items(), key=lambda x: -x[1])[: args.top]:
        print(f"{v:8d}  {p}")


def require_mapping(cfg):
    if not cfg["category_param"] or not all(cfg["categories"].values()):
        sys.exit("config.json の category_param / categories が未設定です。discover で特定してください")


def cmd_verify(args, cfg):
    require_mapping(cfg)
    sess = session()
    sheet = Sheet(sess, cfg)
    cols = sheet.columns(cfg)
    start, end = parse_date(args.start), parse_date(args.end)
    agg = aggregate(fetch_pageviews(sess, cfg, start, end), cfg, cols)
    mismatches = cells = 0
    d = start
    while d <= end:
        r = sheet.find_date_row(d)
        if r is None:
            print(f"{d}: シートに行がありません")
        else:
            for c, cat, slug in cols:
                manual = sheet.cell(r, c).strip()
                if manual == "":
                    continue
                cells += 1
                api = agg[d.isoformat()].get(c, 0)
                if int(float(manual.replace(",", ""))) != api:
                    mismatches += 1
                    print(f"{d} {col_letter(c)} {cat}/{slug}: シート={manual} API={api}")
        d += dt.timedelta(days=1)
    print(f"\n比較 {cells} セル / 不一致 {mismatches} セル")
    sys.exit(1 if mismatches else 0)


def cmd_run(args, cfg):
    require_mapping(cfg)
    tz = ZoneInfo(cfg["timezone"])
    target = parse_date(args.date) if args.date else dt.datetime.now(tz).date() - dt.timedelta(days=1)
    sess = session()
    sheet = Sheet(sess, cfg)
    cols = sheet.columns(cfg)
    last_col = cols[-1][0]

    row = sheet.find_date_row(target)
    if row is None:
        if args.dry_run:
            print(f"[dry-run] {target:%Y/%m} のブロックを追加します")
        else:
            add_month_block(sheet, target.year, target.month, last_col)
            row = sheet.find_date_row(target)
            if row is None:
                sys.exit(f"{target} の行を作成できませんでした")

    agg = aggregate(fetch_pageviews(sess, cfg, target, target), cfg, cols)[target.isoformat()]
    values = [agg.get(c, 0) for c in range(FIRST_DATA_COL, last_col + 1)]
    print(f"{target} → {row} 行目: 合計 {sum(values)} / " + ", ".join(
        f"{col_letter(c)}={agg.get(c, 0)}" for c, _, _ in cols))
    if args.dry_run:
        return
    sheet.batch_values([{
        "range": f"'{sheet.name}'!{col_letter(FIRST_DATA_COL)}{row}:{col_letter(last_col)}{row}",
        "values": [values],
    }])
    ensure_cumulative(sheet, last_col)
    print("書き込み完了")


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    d = sub.add_parser("discover")
    d.add_argument("--start", required=True)
    d.add_argument("--end", required=True)
    d.add_argument("--top", type=int, default=50)
    v = sub.add_parser("verify")
    v.add_argument("--start", required=True)
    v.add_argument("--end", required=True)
    r = sub.add_parser("run")
    r.add_argument("--date", help="YYYY-MM-DD（既定: 前日）")
    r.add_argument("--dry-run", action="store_true")
    args = p.parse_args()
    cfg = load_config()
    {"discover": cmd_discover, "verify": cmd_verify, "run": cmd_run}[args.cmd](args, cfg)


if __name__ == "__main__":
    main()
