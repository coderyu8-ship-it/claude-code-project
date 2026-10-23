# 依頼：Meta広告の数値をスプレッドシートへ自動入力する設定

あなたは私のChromeを操作して、以下の手順1〜4を順番に進めてください。
各手順の完了時に、何をしたかを一言報告してから次へ進んでください。

## 守ること
- 発行したアクセストークンの文字列を、チャット・メモ・シートのセルに書かないでください。保存先は手順3のスクリプトプロパティだけです。
- 「権限の承認」「システムユーザーの作成」「トークンの生成」の画面では、ボタンを押す前に私に確認してください。
- 画面が手順と違う、またはエラーが出たら、推測で進めずに止まって、画面の状況を報告してください。
- 既存のシートのセル・数式・タブには触らないでください。

## 手順1：Apps Script にコードを貼る
1. 次のシートを開く：https://docs.google.com/spreadsheets/d/1IWlkOrVyqNjxx_WXbFP_N9HuceOdEkRiDDPRsIgshgo/edit
2. メニュー「拡張機能 → Apps Script」を開く。
3. 左の歯車「プロジェクトの設定」で「『appsscript.json』マニフェスト ファイルをエディタで表示する」にチェックを入れる。
4. エディタに戻り、`appsscript.json` の中身をすべて消して、末尾の【appsscript.json】を貼る。
5. `コード.gs` の中身をすべて消して、末尾の【Code.gs】を貼る。
6. 保存（Ctrl+S）する。

## 手順2：Meta のアクセストークンを発行する
1. https://business.facebook.com/settings を開く。
2. 「ユーザー → システムユーザー」で「追加」を押す。名前は `sheet-sync`、役割は「管理者」にする（作成前に私に確認）。
3. 作ったシステムユーザーで「アセットを割り当て」を押す。広告アカウント `825232673399116` を選び、「レポートを表示」以上の権限を付ける。
4. 「トークンを生成」を押す。アプリを選ぶ画面が出て、選べるアプリが無い場合は止まって報告する。
5. 権限 `ads_read` にチェックを入れて生成する（生成前に私に確認）。有効期限を選べる場合は「無期限」にする。

## 手順3：トークンを保存する
1. Apps Script のタブに戻り、「プロジェクトの設定 → スクリプト プロパティ → スクリプト プロパティを追加」を押す。
2. プロパティ名は `META_ACCESS_TOKEN`、値は手順2で生成したトークンにして保存する。
3. Meta 側のトークン表示画面を閉じる。

## 手順4：action_type の一覧を取る
1. エディタ上部の関数の選択欄で `listActionTypes` を選び、「実行」を押す。
2. 初回は権限の承認画面が出る（承認前に私に確認）。「詳細 → 安全ではないページに移動」が出るのは自作スクリプトでは正常。
3. 下に出る「実行ログ」の内容を、そのまま全文チャットに貼る。
4. エラーが出た場合は、エラー文を全文貼る。

ここで終了です。`syncMetaAds` と `installTrigger` はまだ実行しないでください。

---

【appsscript.json】
```json
{
  "timeZone": "Asia/Tokyo",
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets.currentonly",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/script.scriptapp"
  ]
}
```

【Code.gs】
```javascript
/**
 * Meta広告 → 「Meta広告の数値」タブ 日次自動入力
 *
 * 入力するのは B,C,D,J,K,M 列のみ。率・単価・合計行は既存の数式に任せる。
 * スクリプトプロパティ:
 *   META_ACCESS_TOKEN … システムユーザーのアクセストークン（必須）
 */

const CONFIG = {
  API_VERSION: 'v23.0',
  AD_ACCOUNT_ID: 'act_825232673399116',
  SHEET_NAME: 'Meta広告の数値',
  DATE_COLUMN: 1,          // A列
  DAYS_BACK: 1,            // 1 = 前日分のみ。3 にすると直近3日を毎回上書き（Metaの遅延計上に追従）
  TRIGGER_HOUR: 9,

  // 列番号 → 取得元。type: 'field' は insights の直接フィールド、'action' は actions 配列の action_type。
  // ★ ACTION 系の action_type は listActionTypes() の結果を見て書き換えること。
  COLUMNS: {
    2:  { type: 'field',  key: 'impressions' },          // B 表示回数 IMP
    3:  { type: 'field',  key: 'inline_link_clicks' },   // C クリック数 CL（リンククリック）
    4:  { type: 'action', key: 'REPLACE_ME_商品詳細' },   // D 商品詳細ページ遷移数
    10: { type: 'action', key: 'REPLACE_ME_v2' },         // J 決済入力ページ遷移(v2)
    11: { type: 'action', key: 'REPLACE_ME_v4' },         // K 購入後ページ遷移(v4)
    13: { type: 'field',  key: 'spend' },                // M 広告費用 COST
  },
};

/** 毎朝のトリガーから呼ばれる本体 */
function syncMetaAds() {
  assertConfigured_();
  const sheet = getSheet_();
  const rowByDate = buildRowIndex_(sheet);
  const missing = [];

  for (let i = CONFIG.DAYS_BACK; i >= 1; i--) {
    const ymd = ymdDaysAgo_(i);
    const row = rowByDate[ymd];
    if (!row) { missing.push(ymd); continue; }

    const insight = fetchInsight_(ymd);
    Object.keys(CONFIG.COLUMNS).forEach(col => {
      sheet.getRange(row, Number(col)).setValue(extract_(insight, CONFIG.COLUMNS[col]));
    });
  }

  if (missing.length) {
    // 月ブロック（合計行＋日付行）は手動で用意する運用のため、行は自動追加しない
    throw new Error('シートに日付行がありません: ' + missing.join(', '));
  }
}

/** 初回確認用：前日の action_type 一覧をログに出す。COLUMNS の REPLACE_ME を埋めるのに使う */
function listActionTypes() {
  const insight = fetchInsight_(ymdDaysAgo_(1));
  const actions = (insight && insight.actions) || [];
  if (!actions.length) { Logger.log('前日の actions が空です。CVのあった日付で ymdDaysAgo_ の引数を変えて試してください。'); return; }
  actions.forEach(a => Logger.log(a.action_type + ' = ' + a.value));
}

/** 毎朝 9 時台のトリガーを作成（重複作成しない） */
function installTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncMetaAds')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncMetaAds').timeBased().everyDays(1).atHour(CONFIG.TRIGGER_HOUR).create();
}

// ---- internals ----

function fetchInsight_(ymd) {
  const token = PropertiesService.getScriptProperties().getProperty('META_ACCESS_TOKEN');
  if (!token) throw new Error('スクリプトプロパティ META_ACCESS_TOKEN が未設定です');

  const params = {
    level: 'account',
    fields: 'impressions,inline_link_clicks,clicks,spend,actions',
    time_range: JSON.stringify({ since: ymd, until: ymd }),
    access_token: token,
  };
  const qs = Object.keys(params).map(k => k + '=' + encodeURIComponent(params[k])).join('&');
  const url = 'https://graph.facebook.com/' + CONFIG.API_VERSION + '/' + CONFIG.AD_ACCOUNT_ID + '/insights?' + qs;

  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const body = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) {
    throw new Error('Meta API エラー (' + ymd + '): ' + JSON.stringify(body.error || body));
  }
  return (body.data && body.data[0]) || null; // 配信なしの日は null
}

function extract_(insight, spec) {
  if (!insight) return 0;
  if (spec.type === 'field') return Number(insight[spec.key] || 0);
  const hit = (insight.actions || []).find(a => a.action_type === spec.key);
  return hit ? Number(hit.value) : 0;
}

function buildRowIndex_(sheet) {
  const tz = Session.getScriptTimeZone();
  const values = sheet.getRange(1, CONFIG.DATE_COLUMN, sheet.getLastRow(), 1).getValues();
  const index = {};
  values.forEach((v, i) => {
    if (v[0] instanceof Date) index[Utilities.formatDate(v[0], tz, 'yyyy-MM-dd')] = i + 1;
  });
  return index;
}

function ymdDaysAgo_(n) {
  const d = new Date();
  d.setDate(d.getDate() - n);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getSheet_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('タブが見つかりません: ' + CONFIG.SHEET_NAME);
  return sheet;
}

function assertConfigured_() {
  const unset = Object.values(CONFIG.COLUMNS).filter(c => c.key.indexOf('REPLACE_ME') === 0);
  if (unset.length) throw new Error('COLUMNS の action_type が未設定です。listActionTypes() を実行して埋めてください');
}
```
