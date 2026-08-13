# claude-code-project

TypeScript プロジェクトのテスト基盤（Vitest + v8 カバレッジ + CI 閾値チェック）と、
その上で動くサンプルのコアモジュール。

## セットアップ

```bash
npm install
```

Node.js 20 以上が必要です（CI は 22 で実行）。

## コマンド

| コマンド | 内容 |
| --- | --- |
| `npm test` | テストを1回実行 |
| `npm run test:watch` | ウォッチモードで実行 |
| `npm run test:coverage` | カバレッジ付きで実行（閾値未達なら失敗） |
| `npm run coverage:report` | ファイル別カバレッジを低い順に一覧表示 |
| `npm run coverage:report -- --min 80` | 加えて、80%未満のファイルがあれば exit 1 |
| `npm run typecheck` | 型チェックのみ |
| `npm run build` | `dist/` へビルド |

カバレッジ計測後、`coverage/index.html` をブラウザで開くと行単位の未カバー箇所を確認できます。

## カバレッジの計測方針

設定は `vitest.config.ts` にあります。

- **`all: true`** — テストから import されていないファイルも計測対象に含めます。
  これが無いと、テストが1つも無いモジュールはレポートに現れず、全体の数値が
  実態より高く出てしまいます。
- **グローバル閾値 80%**（statements / branches / functions / lines）— 未達なら
  `npm run test:coverage` が失敗します。
- **ファイル単位の下限 80%** — `scripts/coverage-report.mjs` が担当します。
  グローバル閾値だけでは、弱いモジュールが全体平均に埋もれて検出できません。
  実際、未テストのモジュールを1つ追加した状態で全体は 97.84%（＝グローバル閾値は通過）
  でしたが、ファイル単位のゲートはこれを検出しました。

閾値は、スイートの成長に合わせて引き上げてください。赤いビルドを緑にするために
下げるのは避けてください。

## CI

`.github/workflows/ci.yml` が push（main）と pull request で実行され、
型チェック → ビルド → カバレッジ付きテスト → ファイル単位の下限チェック
の順に走ります。カバレッジレポートは artifact として保存されます（保持14日）。

## 構成

```
src/core/money.ts      通貨つき金額の演算（整数マイナー単位、丸め、按分）
src/core/discount.ts   割引ルール（定率／定額、上限、順次適用）
src/core/cart.ts       カートの明細・小計・割引・税・合計
src/index.ts           公開エントリポイント
tests/                 上記に対応するテスト
scripts/               カバレッジ集計スクリプト
```

`src/core/` 以下はテスト基盤を動かすためのサンプル実装です。実際の対象コードに
差し替える場合も、`vitest.config.ts` の `include` は `src/**/*.ts` のままで動きます。
