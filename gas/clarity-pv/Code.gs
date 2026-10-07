/**
 * Clarity の LP セッション数を「LPヒートマップ」シートの PV数 列へ毎日記入する。
 *
 * 設定（Apps Script エディタ > プロジェクトの設定 > スクリプト プロパティ）:
 *   CLARITY_TOKEN : Clarity の API トークン（設定 > データのエクスポート で発行）
 *   LP_URL        : 対象 LP の URL の一部（例: "/lp/gift"）。この文字列を含む URL を合算する
 *
 * 初回だけ createDailyTrigger() を手動実行すると、毎日 0 時台に recordYesterdayPv() が動く。
 */

const SHEET_NAME = 'LPヒートマップ';
const DATE_COL = 1; // A列: 日付
const PV_COL = 2;   // B列: PV数
const API_URL =
  'https://www.clarity.ms/export-data/api/v1/project-live-insights?numOfDays=1&dimension1=URL';

function recordYesterdayPv() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('CLARITY_TOKEN');
  const lpUrl = props.getProperty('LP_URL');
  if (!token || !lpUrl) throw new Error('CLARITY_TOKEN と LP_URL をスクリプト プロパティに設定してください');

  const sessions = fetchLpSessions_(token, lpUrl);

  // numOfDays=1 は「実行時点から過去24時間」。0 時台に実行し、前日分として記録する
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const row = findDateRow_(yesterday);
  if (!row) throw new Error('日付の行が見つかりません: ' + formatDate_(yesterday));

  const cell = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME).getRange(row, PV_COL);
  if (cell.getValue() !== '') {
    console.log('既に値があるため上書きしません: ' + formatDate_(yesterday));
    return;
  }
  cell.setValue(sessions);
  console.log(formatDate_(yesterday) + ' PV数=' + sessions);
}

function fetchLpSessions_(token, lpUrl) {
  const res = UrlFetchApp.fetch(API_URL, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Clarity API エラー ' + res.getResponseCode() + ': ' + res.getContentText());
  }
  const traffic = JSON.parse(res.getContentText()).find(m => m.metricName === 'Traffic');
  if (!traffic) throw new Error('Traffic 指標がレスポンスにありません');

  return traffic.information
    .filter(i => String(i.URL || i.Url || '').indexOf(lpUrl) !== -1)
    .reduce((sum, i) => sum + Number(i.totalSessionCount || 0), 0);
}

function findDateRow_(date) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const values = sheet.getRange(1, DATE_COL, sheet.getLastRow(), 1).getValues();
  const target = formatDate_(date);
  for (let i = 0; i < values.length; i++) {
    const v = values[i][0];
    if (v instanceof Date && formatDate_(v) === target) return i + 1;
    if (typeof v === 'string' && v.trim() === target) return i + 1;
  }
  return null;
}

function formatDate_(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
}

function createDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'recordYesterdayPv')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('recordYesterdayPv').timeBased().everyDays(1).atHour(0).create();
}
