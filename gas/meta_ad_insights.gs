/**
 * Meta広告 広告（バナー）別の実績を、週ごとに新しいスプレッドシートとして書き出す。
 *
 * 構成:
 *   このスクリプトは「スタンドアロン」の Apps Script プロジェクト1つに置く。
 *   毎週トリガーで createWeeklyReport() が走り、REPORT_FOLDER_ID のフォルダに
 *   「Meta広告バナー週次_YYYY-MM-DD」というスプレッドシートを新規作成する。
 *
 * 前提:
 *   - スクリプトプロパティ META_ACCESS_TOKEN に ads_read 権限のシステムユーザートークン
 *   - システムユーザーに広告アカウントが割り当て済み
 *
 * 初回セットアップ:
 *   1. スクリプトプロパティに META_ACCESS_TOKEN を登録
 *   2. createWeeklyReport() を手動実行して権限を承認・出力を確認
 *   3. installWeeklyTrigger() を1回だけ実行（毎週月曜 7時台に自動実行）
 */

const AD_ACCOUNT_ID = 'act_825232673399116';
const API_VERSION = 'v21.0';
const REPORT_FOLDER_ID = '1boFSaU9VM3VTTvQLNSYN5lfRMtQ5jLXq';
const FILE_PREFIX = 'Meta広告バナー週次_';

// v2 / v4 に対応する action_type。初回実行後「アクション一覧」シートを見て確定させる。
const ACTION_V2 = 'offsite_conversion.fb_pixel_add_to_cart';
const ACTION_V4 = 'offsite_conversion.fb_pixel_add_payment_info';

// 1ファイルに作るシート（期間ごと）
const PERIODS = [
  { sheet: '直近7日', preset: 'last_7d' },
  { sheet: '直近30日', preset: 'last_30d' },
];

const HEADER = [
  'バナー', 'バナー照合キー', '広告ID', '広告名', 'キャンペーン', '広告セット', '配信状態',
  'IMP', 'リーチ', 'フリークエンシー', 'クリック(全体)', 'リンククリック',
  'CTR(全体)', 'リンクCTR', 'CPC', 'CPM', '広告費',
  'v2(カート)', 'v4(決済入力後)', 'CPA(v2)', 'CPA(v4)', '画像URL',
];
const COL_SPEND = HEADER.indexOf('広告費');

function createWeeklyReport() {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const ss = SpreadsheetApp.create(FILE_PREFIX + today);
  DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(REPORT_FOLDER_ID));

  const creatives = fetchCreatives_();
  let actionTotals = {};

  PERIODS.forEach(p => {
    const result = buildRows_(p.preset, creatives);
    if (p.preset === 'last_30d') actionTotals = result.actionTotals;
    writeSheet_(ss, p.sheet, HEADER, result.rows, p.preset, formatAdsSheet_);
  });

  const actionRows = Object.keys(actionTotals)
    .sort((a, b) => actionTotals[b] - actionTotals[a])
    .map(k => [k, actionTotals[k], k === ACTION_V2 ? 'v2' : k === ACTION_V4 ? 'v4' : '']);
  writeSheet_(ss, 'アクション一覧', ['action_type', '合計(直近30日)', '割当'], actionRows, 'last_30d');

  // SpreadsheetApp.create が作る空の「シート1」を削除
  const blank = ss.getSheets().find(s => s.getLastRow() === 0);
  if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);

  Logger.log(ss.getUrl());
}

function installWeeklyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'createWeeklyReport')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('createWeeklyReport')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).inTimezone('Asia/Tokyo')
    .create();
}

function buildRows_(preset, creatives) {
  const insights = fetchAll_(`${AD_ACCOUNT_ID}/insights`, {
    level: 'ad',
    date_preset: preset,
    fields: [
      'campaign_name', 'adset_name', 'ad_id', 'ad_name',
      'impressions', 'reach', 'frequency', 'clicks', 'inline_link_clicks',
      'ctr', 'inline_link_click_ctr', 'cpc', 'cpm', 'spend', 'actions',
    ].join(','),
    limit: 500,
  });

  const actionTotals = {};
  const rows = insights.map(r => {
    const actions = toActionMap_(r.actions);
    Object.keys(actions).forEach(k => { actionTotals[k] = (actionTotals[k] || 0) + actions[k]; });

    const spend = Number(r.spend || 0);
    const v2 = actions[ACTION_V2] || 0;
    const v4 = actions[ACTION_V4] || 0;
    const c = creatives[r.ad_id] || {};
    const img = c.imageUrl || c.thumbnailUrl || '';

    return [
      img ? `=IMAGE("${img}")` : '',
      bannerKey_(r.ad_name),
      r.ad_id, r.ad_name, r.campaign_name, r.adset_name, c.status || '',
      Number(r.impressions || 0), Number(r.reach || 0), Number(r.frequency || 0),
      Number(r.clicks || 0), Number(r.inline_link_clicks || 0),
      Number(r.ctr || 0) / 100, Number(r.inline_link_click_ctr || 0) / 100,
      Number(r.cpc || 0), Number(r.cpm || 0), spend,
      v2, v4, v2 ? spend / v2 : '', v4 ? spend / v4 : '',
      img,
    ];
  });

  rows.sort((a, b) => b[COL_SPEND] - a[COL_SPEND]);
  return { rows, actionTotals };
}

/** 広告ID → { status, imageUrl, thumbnailUrl } */
function fetchCreatives_() {
  const ads = fetchAll_(`${AD_ACCOUNT_ID}/ads`, {
    fields: 'id,effective_status,creative{image_url,thumbnail_url,object_story_spec}',
    thumbnail_width: 600,
    thumbnail_height: 600,
    limit: 200,
  });
  const map = {};
  ads.forEach(ad => {
    const cr = ad.creative || {};
    const linkData = (cr.object_story_spec || {}).link_data || {};
    map[ad.id] = {
      status: ad.effective_status,
      imageUrl: cr.image_url || linkData.picture || '',
      thumbnailUrl: cr.thumbnail_url || '',
    };
  });
  return map;
}

/** 広告名から「verN」を抜き出し、ドライブの verN.jpg と照合できるようにする */
function bannerKey_(adName) {
  const m = String(adName || '').match(/ver\s*\.?\s*(\d+)/i);
  return m ? `ver${m[1]}` : '';
}

function toActionMap_(actions) {
  const map = {};
  (actions || []).forEach(a => { map[a.action_type] = Number(a.value || 0); });
  return map;
}

function fetchAll_(path, params) {
  const token = PropertiesService.getScriptProperties().getProperty('META_ACCESS_TOKEN');
  if (!token) throw new Error('スクリプトプロパティ META_ACCESS_TOKEN が未設定です');

  const query = Object.keys(params)
    .map(k => `${k}=${encodeURIComponent(params[k])}`)
    .concat(`access_token=${encodeURIComponent(token)}`)
    .join('&');
  let url = `https://graph.facebook.com/${API_VERSION}/${path}?${query}`;

  const out = [];
  while (url) {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const body = JSON.parse(res.getContentText());
    if (body.error) throw new Error(`Meta API エラー: ${body.error.message}`);
    out.push(...(body.data || []));
    url = body.paging && body.paging.next;
  }
  return out;
}

function formatAdsSheet_(sheet, rowCount) {
  if (!rowCount) return;
  sheet.setColumnWidth(1, 140);
  sheet.setRowHeights(2, rowCount, 140);
  sheet.getRange(2, 13, rowCount, 2).setNumberFormat('0.00%');
  sheet.getRange(2, 15, rowCount, 3).setNumberFormat('¥#,##0');
  sheet.getRange(2, 20, rowCount, 2).setNumberFormat('¥#,##0');
}

function writeSheet_(ss, name, header, rows, preset, format) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (rows.length) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  sheet.getRange(1, header.length + 2)
    .setValue(`取得: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')} / ${preset}`);
  sheet.setFrozenRows(1);
  if (format) format(sheet, rows.length);
}
