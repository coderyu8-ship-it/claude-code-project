# 楽天ショップ更新チェック

`.github/workflows/rakuten-monitor.yml` が毎日 08:53 JST に `monitor.py` を実行し、
対象ショップのトップページを前日のスナップショット（`snapshots/`）と比較して、
店舗ごとの更新内容（掲載商品・バナー・見出し・タイトルの追加/削除/変更）を
`uehara@tstyle.marketing` にメールする。更新がない日も「更新なし」で送る。

## 初期設定（GitHub リポジトリの Settings → Secrets and variables → Actions）

| Secret | 例（Gmail の場合） |
| --- | --- |
| `SMTP_HOST` | `smtp.gmail.com` |
| `SMTP_PORT` | `587` |
| `SMTP_USER` | 送信元の Gmail アドレス |
| `SMTP_PASSWORD` | Gmail の **アプリパスワード**（通常のパスワードは不可。2段階認証が必要） |
| `MAIL_FROM` | 任意。未設定なら `SMTP_USER` |

設定後、Actions タブから「楽天ショップ更新チェック」を手動実行（Run workflow）すると初回スナップショットが作られる。差分通知は2回目以降。

## 注意

- ワークフローはデフォルトブランチに入って初めてスケジュール実行される。
- 60日間リポジトリに活動がないと GitHub がスケジュールを止めるが、毎日スナップショットをコミットするので通常は止まらない。
- 対象店舗の追加・削除は `monitor.py` の `SHOPS` を編集する。
- ローカル確認: `DRY_RUN=1 python tools/rakuten-monitor/monitor.py`（メールを送らず本文を表示）
