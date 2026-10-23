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
