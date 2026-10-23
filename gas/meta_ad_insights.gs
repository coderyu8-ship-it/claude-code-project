/**
 * Meta広告 広告（バナー）別の直近30日実績をシートに書き出す。
 *
 * 前提:
 *   - スクリプトプロパティ META_ACCESS_TOKEN に ads_read 権限のシステムユーザートークン
 *   - システムユーザーに広告アカウントが割り当て済み
 *
 * 出力シート:
 *   - 広告別_直近30日 : 広告ごとの実績 + クリエイティブ画像URL + バナー照合キー(verN)
 *   - アクション一覧   : 取得できた action_type の合計（v2/v4 の対応付け確認用）
 */

const AD_ACCOUNT_ID = 'act_825232673399116';
const API_VERSION = 'v21.0';
const DATE_PRESET = 'last_30d';
const SHEET_ADS = '広告別_直近30日';
const SHEET_ACTIONS = 'アクション一覧';

// v2 / v4 に対応する action_type。初回実行後「アクション一覧」を見て確定させる。
// 例: 'offsite_conversion.fb_pixel_add_to_cart' / 'offsite_conversion.fb_pixel_custom.XXXX'
const ACTION_V2 = 'offsite_conversion.fb_pixel_add_to_cart';
const ACTION_V4 = 'offsite_conversion.fb_pixel_add_payment_info';

function exportAdInsights() {
  const insights = fetchAll_(`${AD_ACCOUNT_ID}/insights`, {
    level: 'ad',
    date_preset: DATE_PRESET,
    fields: [
      'campaign_name', 'adset_name', 'ad_id', 'ad_name',
      'impressions', 'reach', 'frequency', 'clicks', 'inline_link_clicks',
      'ctr', 'inline_link_click_ctr', 'cpc', 'cpm', 'spend', 'actions',
    ].join(','),
    limit: 500,
  });

  const creatives = fetchCreatives_();
  const actionTotals = {};

  const header = [
    'バナー照合キー', '広告ID', '広告名', 'キャンペーン', '広告セット', '配信状態',
    'IMP', 'リーチ', 'フリークエンシー', 'クリック(全体)', 'リンククリック',
    'CTR(全体)', 'リンクCTR', 'CPC', 'CPM', '広告費',
    'v2(カート)', 'v4(決済入力後)', 'CPA(v2)', 'CPA(v4)',
    '画像URL', 'サムネイルURL',
  ];

  const rows = insights.map(r => {
    const actions = toActionMap_(r.actions);
    Object.keys(actions).forEach(k => { actionTotals[k] = (actionTotals[k] || 0) + actions[k]; });

    const spend = Number(r.spend || 0);
    const v2 = actions[ACTION_V2] || 0;
    const v4 = actions[ACTION_V4] || 0;
    const c = creatives[r.ad_id] || {};

    return [
      bannerKey_(r.ad_name),
      r.ad_id, r.ad_name, r.campaign_name, r.adset_name, c.status || '',
      Number(r.impressions || 0), Number(r.reach || 0), Number(r.frequency || 0),
      Number(r.clicks || 0), Number(r.inline_link_clicks || 0),
      Number(r.ctr || 0) / 100, Number(r.inline_link_click_ctr || 0) / 100,
      Number(r.cpc || 0), Number(r.cpm || 0), spend,
      v2, v4, v2 ? spend / v2 : '', v4 ? spend / v4 : '',
      c.imageUrl || '', c.thumbnailUrl || '',
    ];
  });

  rows.sort((a, b) => b[15] - a[15]); // 広告費の降順

  writeSheet_(SHEET_ADS, header, rows, sheet => {
    if (!rows.length) return;
    sheet.getRange(2, 12, rows.length, 2).setNumberFormat('0.00%');
    sheet.getRange(2, 14, rows.length, 3).setNumberFormat('¥#,##0');
    sheet.getRange(2, 19, rows.length, 2).setNumberFormat('¥#,##0');
  });

  const actionRows = Object.keys(actionTotals)
    .sort((a, b) => actionTotals[b] - actionTotals[a])
    .map(k => [k, actionTotals[k], k === ACTION_V2 ? 'v2' : k === ACTION_V4 ? 'v4' : '']);
  writeSheet_(SHEET_ACTIONS, ['action_type', '合計', '割当'], actionRows);
}

/** 広告ID → { status, imageUrl, thumbnailUrl } */
function fetchCreatives_() {
  const ads = fetchAll_(`${AD_ACCOUNT_ID}/ads`, {
    fields: 'id,name,effective_status,creative{id,image_url,thumbnail_url,object_story_spec}',
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

function writeSheet_(name, header, rows, format) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (rows.length) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  sheet.getRange(1, header.length + 2).setValue(`更新: ${new Date().toLocaleString('ja-JP')} / ${DATE_PRESET}`);
  sheet.setFrozenRows(1);
  if (format) format(sheet);
}
