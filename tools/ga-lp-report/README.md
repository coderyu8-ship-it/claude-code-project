# GA4 → 「LPリンクの遷移数」シート 自動転記

GA4（プロパティ 499270573）の商品ページ表示回数を、LP カテゴリ × 商品ごとに
スプレッドシート「田丸屋様_Meta広告の数値管理」の「LPリンクの遷移数」シートへ書き込む。

## 仕組み

- 商品ページ URL のクエリパラメータ（`config.json` の `category_param`）でカテゴリを判定
- 2 行目見出しの末尾行（例 `nama-udon-8p`）を URL パスの一部と照合して商品を判定
- 対象日の行（A 列が `YYYY/MM/DD`）の B 列以降を上書き。該当日がなければ
  「◯月合計」行と日付行を末尾に自動生成し、書式は直前の月から写す
- 3 行目「累計」は全ての「◯月合計」行の合計に揃える

## 準備

1. 認可するアカウントが GA4 プロパティの閲覧権限とスプレッドシートの編集権限を持っていること
2. 環境変数 `GA_OAUTH_CLIENT_ID` `GA_OAUTH_CLIENT_SECRET` `GA_OAUTH_REFRESH_TOKEN` を設定
   （サービスアカウントを使う場合は `GA_SERVICE_ACCOUNT_JSON` に JSON キー本文）。コミット禁止
3. `pip install -r tools/ga-lp-report/requirements.txt`

## 使い方

```bash
cd tools/ga-lp-report
python3 sync.py discover --start 2026-09-14 --end 2026-09-30   # カテゴリ判定パラメータの特定
python3 sync.py verify   --start 2026-09-14 --end 2026-10-01   # 手入力値との突き合わせ
python3 sync.py run --dry-run                                   # 前日分を書き込まずに確認
python3 sync.py run                                             # 前日分を書き込む
python3 sync.py run --date 2026-10-05                           # 日付指定
```
