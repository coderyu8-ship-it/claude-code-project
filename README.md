# claude-code-project

Claude Code から Google Gemini を呼び出せるようにした作業用リポジトリ。

Claude Code 本体のモデルを Gemini に置き換えることはできないため、
**Gemini を Claude が使える「道具」として追加する** 構成になっている。
Claude が調べ物や下調べに Gemini を呼び、その知見を統合して成果物を出す。

## セットアップ

```bash
# 1. API キーを取得（Google AI Studio）
#    https://aistudio.google.com/apikey

# 2. .env を作ってキーを設定
cp .env.example .env
$EDITOR .env        # GEMINI_API_KEY=... を記入

# 3. Claude Code を再起動する（.mcp.json は起動時に読まれる）
```

`npm install` は不要。MCP サーバーは依存パッケージなしで書かれている（Node.js 18 以上が必要）。

初回起動時、Claude Code がプロジェクトの MCP サーバーを信頼するか確認してくる。
承認すると `gemini` サーバーが有効になる。`/mcp` コマンドで接続状態を確認できる。

## 使えるツール

| ツール | 説明 |
| --- | --- |
| `ask_gemini` | Gemini に質問する。セカンドオピニオン、要約、アイデア出し |
| `gemini_research` | Google 検索グラウンディング付きの調査。**出典 URL 付きで返る** |

Claude はこれらを自動で呼ぶ（判断基準は [CLAUDE.md](CLAUDE.md) に記載）。
明示的に使わせたい場合は「Gemini でも調べて」と伝えればよい。

## 設定

`.env` で変更できる項目:

| 変数 | 既定値 | 説明 |
| --- | --- | --- |
| `GEMINI_API_KEY` | （必須） | Google AI Studio の API キー |
| `GEMINI_MODEL` | `gemini-2.5-pro` | 既定モデル。速度優先なら `gemini-2.5-flash` |
| `GEMINI_TIMEOUT_MS` | `120000` | API 呼び出しのタイムアウト（ミリ秒） |

## 構成

```
.mcp.json                  MCP サーバーの登録（Claude Code が読む）
CLAUDE.md                  Claude への常時ルール（いつ Gemini を使うか）
.env                       API キー（gitignore 済み・コミット禁止）
tools/gemini-mcp/index.js  MCP サーバー本体
```

## トラブルシューティング

**`/mcp` に gemini が出てこない**
Claude Code を再起動したか確認する。プロジェクトスコープの MCP サーバーは
初回に信頼の確認が入るので、承認していない場合は承認する。

**「GEMINI_API_KEY が設定されていません」と返る**
`.env` がリポジトリ直下にあるか、`GEMINI_API_KEY=` に値が入っているかを確認する。
環境変数として直接設定してもよい。

**「API key not valid」と返る**
キーが誤っているか失効している。AI Studio で再発行する。

**プロキシ環境（Claude Code on the web など）**
`HTTPS_PROXY` が設定されている場合は自動的に `curl` 経由で通信する（Node の `fetch` は
プロキシを見ないため）。特別な設定は不要。
