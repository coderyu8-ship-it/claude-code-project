/**
 * Meta広告 広告（バナー）別の実績を、週ごとに新しいスプレッドシートとして書き出す。
 *
 * 構成:
 *   このスクリプトは Apps Script プロジェクト1つに置く。
 *   毎週トリガーで createWeeklyReport() が走り、REPORT_FOLDER_ID のフォルダに
 *   「YYYYMMDD時点_Meta広告バナー分析」というスプレッドシートを新規作成する。
 *   期間は 1週間 / 2週間 / 3週間 / 1ヶ月（いずれも前日まで）。
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
const FILE_SUFFIX = '時点_Meta広告バナー分析';
// バナー画像の置き場所。広告名の「verN」と同じ名前の画像（verN.jpg）をセルに埋め込む
const BANNER_FOLDER_ID = '1XrrEhzGZZOHLRj6hFgUQf0gnCLfuxRuD';

// v2 / v4 に対応する action_type。初回実行後「アクション一覧」シートを見て確定させる。
const ACTION_V2 = 'offsite_conversion.fb_pixel_add_to_cart';
const ACTION_V4 = 'offsite_conversion.fb_pixel_add_payment_info';

// 1ファイルに作るシート（期間ごと）。days 日前〜前日を集計する
const PERIODS = [
  { sheet: '1週間', days: 7 },
  { sheet: '2週間', days: 14 },
  { sheet: '3週間', days: 21 },
  { sheet: '1ヶ月', days: 30 },
];

// 判定ルール。広告費が MIN_SPEND 未満は母数不足として判定しない。
// CPA(v2) を期間内の全体 CPA(v2) と比べ、WIN_RATIO 以下なら「好調」、LOSE_RATIO 以上なら「停止候補」
const MIN_SPEND = 5000;
const WIN_RATIO = 0.8;
const LOSE_RATIO = 1.2;

const HEADER = [
  'バナー', 'バナー照合キー', '広告ID', '広告名', 'キャンペーン', '広告セット', '配信状態',
  'IMP', 'リーチ', 'フリークエンシー', 'クリック(全体)', 'リンククリック',
  'CTR(全体)', 'リンクCTR', 'CPC', 'CPM', '広告費',
  'v2(カート)', 'v4(決済入力後)', 'CPA(v2)', 'CPA(v4)', '判定', '判定理由', '画像URL',
];
// 黄色で強調する列
const HIGHLIGHT_COLUMNS = ['広告名', 'IMP', 'クリック(全体)', 'CTR(全体)', '広告費', 'v2(カート)', 'v4(決済入力後)'];
const HIGHLIGHT_COLOR = '#fff2cc';

// 1行目: バナー画像フォルダのURL / 2行目: 見出し / 3行目〜: データ
const HEADER_ROW = 2;
const DATA_ROW = 3;
const BANNER_FOLDER_URL = `https://drive.google.com/drive/folders/${BANNER_FOLDER_ID}`;

const COL_SPEND = HEADER.indexOf('広告費');
const COL_V2 = HEADER.indexOf('v2(カート)');
const COL_JUDGE = HEADER.indexOf('判定');

function createWeeklyReport() {
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  const ss = SpreadsheetApp.create(today + FILE_SUFFIX);
  DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(REPORT_FOLDER_ID));

  const creatives = fetchCreatives_();
  let actionTotals = {};

  const summary = {};

  PERIODS.forEach(p => {
    const range = timeRange_(p.days);
    const result = buildRows_(range, creatives);
    actionTotals = result.actionTotals; // 最後（最長期間）の値が残る
    result.rows.forEach(r => {
      const key = r[1] || r[3]; // 照合キーが無ければ広告名
      summary[key] = summary[key] || { img: r[0], cells: {} };
      summary[key].cells[p.sheet] = r[COL_JUDGE];
    });
    writeSheet_(ss, p.sheet, HEADER, result.rows, `${range.since}〜${range.until}`, formatAdsSheet_);
  });

  const summaryRows = Object.keys(summary).map(k =>
    [summary[k].img, k].concat(PERIODS.map(p => summary[k].cells[p.sheet] || '配信なし')));
  writeSheet_(ss, '判定サマリー', ['バナー', 'バナー/広告名'].concat(PERIODS.map(p => p.sheet)),
    summaryRows, `足切り 広告費¥${MIN_SPEND}未満は保留`, formatSummarySheet_);

  const actionRows = Object.keys(actionTotals)
    .sort((a, b) => actionTotals[b] - actionTotals[a])
    .map(k => [k, actionTotals[k], k === ACTION_V2 ? 'v2' : k === ACTION_V4 ? 'v4' : '']);
  writeSheet_(ss, 'アクション一覧', ['action_type', '合計(1ヶ月)', '割当'], actionRows, '1ヶ月');
  ss.setActiveSheet(ss.getSheetByName('判定サマリー'));
  ss.moveActiveSheet(1);

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

/** days 日前〜前日（JST）の time_range */
function timeRange_(days) {
  const fmt = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  const until = new Date(Date.now() - 24 * 3600 * 1000);
  const since = new Date(until.getTime() - (days - 1) * 24 * 3600 * 1000);
  return { since: fmt(since), until: fmt(until) };
}

function buildRows_(range, creatives) {
  const insights = fetchAll_(`${AD_ACCOUNT_ID}/insights`, {
    level: 'ad',
    time_range: JSON.stringify(range),
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
    const key = bannerKey_(r.ad_name);
    const banner = bannerImage_(r.ad_id, c.thumbnailUrl || c.imageUrl || '', key);

    return [
      banner.cell,
      key,
      r.ad_id, r.ad_name, r.campaign_name, r.adset_name, c.status || '',
      Number(r.impressions || 0), Number(r.reach || 0), Number(r.frequency || 0),
      Number(r.clicks || 0), Number(r.inline_link_clicks || 0),
      Number(r.ctr || 0) / 100, Number(r.inline_link_click_ctr || 0) / 100,
      Number(r.cpc || 0), Number(r.cpm || 0), spend,
      v2, v4, v2 ? spend / v2 : '', v4 ? spend / v4 : '',
      '', '',
      banner.url,
    ];
  });

  judge_(rows);
  rows.sort((a, b) => b[COL_SPEND] - a[COL_SPEND]);
  return { rows, actionTotals };
}

/** 期間内の全体 CPA(v2) を基準に各広告を判定し、判定・判定理由列を埋める */
function judge_(rows) {
  const totalSpend = rows.reduce((s, r) => s + r[COL_SPEND], 0);
  const totalV2 = rows.reduce((s, r) => s + r[COL_V2], 0);
  const baseCpa = totalV2 ? totalSpend / totalV2 : 0;

  rows.forEach(r => {
    const spend = r[COL_SPEND];
    const v2 = r[COL_V2];
    let verdict;
    let reason;
    if (spend < MIN_SPEND) {
      verdict = '保留';
      reason = `広告費¥${Math.round(spend)}が足切り未満`;
    } else if (!v2) {
      verdict = '停止候補';
      reason = `¥${Math.round(spend)}消化でv2が0件`;
    } else {
      const ratio = baseCpa ? (spend / v2) / baseCpa : 1;
      verdict = ratio <= WIN_RATIO ? '好調' : ratio >= LOSE_RATIO ? '停止候補' : '継続';
      reason = `CPA(v2)が全体比${Math.round(ratio * 100)}%`;
    }
    r[COL_JUDGE] = verdict;
    r[COL_JUDGE + 1] = reason;
  });
}

const imageCache_ = {};
let bannerFiles_ = null;

/** ドライブのバナーフォルダを「verN」→ ファイル の対応表にする（1回だけ読む） */
function loadBannerFiles_() {
  if (bannerFiles_) return bannerFiles_;
  bannerFiles_ = {};
  const it = DriveApp.getFolderById(BANNER_FOLDER_ID).getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (!/^image\//.test(f.getMimeType())) continue;
    const key = bannerKey_(f.getName());
    if (key) bannerFiles_[key] = f;
  }
  return bannerFiles_;
}

/**
 * バナー画像セルの中身と、その画像の URL。
 * ドライブの verN.jpg を優先し、無い広告だけ Meta のサムネイルを使う。
 * 画像データ自体をセルに埋め込むので、ドライブの共有設定に関係なく表示される。
 */
function bannerImage_(adId, metaUrl, key) {
  const cacheKey = key || adId;
  if (cacheKey in imageCache_) return imageCache_[cacheKey];

  let blob = null;
  let url = '';
  const file = key && loadBannerFiles_()[key];
  if (file) {
    blob = file.getSize() > 1.5 * 1024 * 1024 ? file.getThumbnail() : file.getBlob();
    url = file.getUrl();
  } else if (metaUrl) {
    const res = UrlFetchApp.fetch(metaUrl, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) blob = res.getBlob();
    url = metaUrl;
  }

  const cell = blob
    ? SpreadsheetApp.newCellImage()
      .setSourceUrl(`data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`)
      .setAltTextTitle(cacheKey)
      .build()
    : '';
  imageCache_[cacheKey] = { cell, url };
  return imageCache_[cacheKey];
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
  HIGHLIGHT_COLUMNS.forEach(name => {
    sheet.getRange(HEADER_ROW, HEADER.indexOf(name) + 1, rowCount + 1, 1).setBackground(HIGHLIGHT_COLOR);
  });
  if (!rowCount) return;
  sheet.setColumnWidth(1, 140);
  sheet.setRowHeights(DATA_ROW, rowCount, 140);
  sheet.getRange(DATA_ROW, 13, rowCount, 2).setNumberFormat('0.00%');
  sheet.getRange(DATA_ROW, 15, rowCount, 3).setNumberFormat('¥#,##0');
  sheet.getRange(DATA_ROW, 20, rowCount, 2).setNumberFormat('¥#,##0');
}

function formatSummarySheet_(sheet, rowCount) {
  if (!rowCount) return;
  sheet.setColumnWidth(1, 140);
  sheet.setRowHeights(DATA_ROW, rowCount, 140);
  const range = sheet.getRange(DATA_ROW, 3, rowCount, PERIODS.length);
  const rule = (text, color) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text).setBackground(color).setRanges([range]).build();
  sheet.setConditionalFormatRules([
    rule('好調', '#d9ead3'), rule('停止候補', '#f4cccc'), rule('保留', '#eeeeee'),
  ]);
}

function writeSheet_(ss, name, header, rows, note, format) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.getRange(1, 1, 1, 3).setValues([[
    'バナー画像フォルダ',
    BANNER_FOLDER_URL,
    `取得: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')} / ${note}`,
  ]]);
  sheet.getRange(HEADER_ROW, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (rows.length) sheet.getRange(DATA_ROW, 1, rows.length, header.length).setValues(rows);
  sheet.setFrozenRows(HEADER_ROW);
  if (format) format(sheet, rows.length);
}
