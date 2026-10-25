/**
 * 要因分析（meta_ad_insights.gs と同じ Apps Script プロジェクトに別ファイルとして置く）
 *
 * 1. 内部要因「施策の効果」
 *    GAS を置いたスプレッドシートの「施策ログ」シートに書いた自社の施策ごとに、
 *    実施日の前7日間と後7日間の数値を比べ、何が起きたかをまとめる。
 *
 * 外部要因（社会の風潮・出来事）の調査と、数値の考察は GAS では行わない。
 * 毎週木曜に Claude の定期タスクがこのレポートを読み、ウェブ調査のうえで
 * 週次の分析ドキュメントを同じフォルダに作る。
 */

const ACTION_LOG_SHEET = '施策ログ';
const ACTION_LOG_HEADER = ['実施日', '対象（verN / 全体）', '施策内容', '狙い・期待する変化'];
const ACTION_WINDOW_DAYS = 7;      // 施策の前後それぞれ何日間で比べるか
const ACTION_LOOKBACK_DAYS = 60;   // 基準日から何日前までの施策を評価するか
const EFFECT_THRESHOLD = 0.1;      // CPA(v2) が ±10% 以上動いたら改善/悪化とみなす

const DAY_MS = 24 * 3600 * 1000;

/** レポート作成から呼ぶ。失敗してもレポート本体は止めない */
function writeFactorAnalysis_(ss, baseDate) {
  try {
    writeActionEffects_(ss, baseDate);
  } catch (e) {
    writeMessageSheet_(ss, '施策の効果', `施策の効果を作れませんでした: ${e.message}`);
  }
}

// ---------------------------------------------------------------- 内部要因

function writeActionEffects_(ss, baseDate) {
  const actions = readActionLog_(baseDate);
  const header = [
    '実施日', '対象', '施策内容', '狙い', '比較期間（前）', '比較期間（後）',
    '広告費 前', '広告費 後', 'CTR 前', 'CTR 後', 'CPM 前', 'CPM 後',
    'v2 前', 'v2 後', 'CPA(v2) 前', 'CPA(v2) 後', 'CPA(v2) 変化',
    'v4 前', 'v4 後', '結果', '何が起きたか',
  ];
  if (!actions.length) {
    writeMessageSheet_(ss, '施策の効果',
      `「${ACTION_LOG_SHEET}」シートに、直近${ACTION_LOOKBACK_DAYS}日の施策がありません。`
      + '施策を実施したら、実施日・対象・施策内容・狙いを1行ずつ書いてください。');
    return;
  }

  const lastDay = new Date(baseDate.getTime() - DAY_MS); // 集計の最終日（基準日の前日）
  const rows = actions.map(a => {
    const before = {
      since: ymd_(new Date(a.date.getTime() - ACTION_WINDOW_DAYS * DAY_MS)),
      until: ymd_(new Date(a.date.getTime() - DAY_MS)),
    };
    const afterEnd = new Date(Math.min(a.date.getTime() + (ACTION_WINDOW_DAYS - 1) * DAY_MS, lastDay.getTime()));
    const after = { since: ymd_(a.date), until: ymd_(afterEnd) };
    const afterDays = Math.round((afterEnd - a.date) / DAY_MS) + 1;

    const b = aggregate_(fetchAdInsights_(before), a.target);
    const f = aggregate_(fetchAdInsights_(after), a.target);
    const verdict = judgeEffect_(b, f, afterDays);

    return [
      ymd_(a.date), a.target, a.content, a.intent,
      `${before.since}〜${before.until}`, `${after.since}〜${after.until}（${afterDays}日）`,
      b.spend, f.spend, b.ctr, f.ctr, b.cpm, f.cpm,
      b.v2, f.v2, b.cpa || '', f.cpa || '', b.cpa && f.cpa ? f.cpa / b.cpa - 1 : '',
      b.v4, f.v4, verdict.result, verdict.note,
    ];
  });

  writeSheet_(ss, '施策の効果', header, rows,
    `施策の前後${ACTION_WINDOW_DAYS}日間を比較。外部要因（季節・競合など）の影響も含まれる`);
  const sheet = ss.getSheetByName('施策の効果');
  const n = rows.length;
  sheet.getRange(DATA_ROW, 7, n, 2).setNumberFormat('¥#,##0');
  sheet.getRange(DATA_ROW, 9, n, 2).setNumberFormat('0.00%');
  sheet.getRange(DATA_ROW, 11, n, 2).setNumberFormat('¥#,##0');
  sheet.getRange(DATA_ROW, 15, n, 2).setNumberFormat('¥#,##0');
  sheet.getRange(DATA_ROW, 17, n, 1).setNumberFormat('+0%;-0%');
  sheet.getRange(DATA_ROW, 3, n, 1).setWrap(true);
  sheet.getRange(DATA_ROW, header.length, n, 1).setWrap(true);
}

/** 施策ログを読む。シートが無ければ見出し付きで作る */
function readActionLog_(baseDate) {
  const host = SpreadsheetApp.getActiveSpreadsheet();
  if (!host) return [];
  let sheet = host.getSheetByName(ACTION_LOG_SHEET);
  if (!sheet) {
    sheet = host.insertSheet(ACTION_LOG_SHEET);
    sheet.appendRow(ACTION_LOG_HEADER);
    sheet.appendRow(['2026/10/20', 'ver7', '（記入例・削除してOK）メインコピーを「法人ギフト向け」に変更', 'CTR 改善']);
    sheet.getRange(1, 1, 1, ACTION_LOG_HEADER.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return [];
  }

  const from = baseDate.getTime() - ACTION_LOOKBACK_DAYS * DAY_MS;
  return sheet.getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[2] && !String(r[2]).startsWith('（記入例'))
    .map(r => ({
      date: r[0] instanceof Date ? r[0] : parseDate_(r[0]),
      target: String(r[1] || '全体').trim(),
      content: String(r[2]).trim(),
      intent: String(r[3] || '').trim(),
    }))
    .filter(a => a.date.getTime() >= from && a.date.getTime() < baseDate.getTime())
    .sort((x, y) => x.date - y.date);
}

/** 対象（verN または 全体）に当てはまる広告の合計 */
function aggregate_(insights, target) {
  const key = bannerKey_(target);
  const rows = insights.filter(r => !key || bannerKey_(r.ad_name) === key);
  const t = { spend: 0, imp: 0, clicks: 0, v2: 0, v4: 0 };
  rows.forEach(r => {
    const a = toActionMap_(r.actions);
    t.spend += Number(r.spend || 0);
    t.imp += Number(r.impressions || 0);
    t.clicks += Number(r.clicks || 0);
    t.v2 += a[ACTION_V2] || 0;
    t.v4 += a[ACTION_V4] || 0;
  });
  t.ctr = t.imp ? t.clicks / t.imp : 0;
  t.cpm = t.imp ? t.spend / t.imp * 1000 : 0;
  t.cpa = t.v2 ? t.spend / t.v2 : 0;
  return t;
}

/** 施策前後の比較から結果と「何が起きたか」を文章にする */
function judgeEffect_(b, f, afterDays) {
  if (afterDays < 3) return { result: '計測中', note: `施策後のデータが${afterDays}日分しかない` };
  if (b.spend < MIN_SPEND || f.spend < MIN_SPEND) {
    return { result: '判定保留', note: `広告費が¥${MIN_SPEND}未満の期間があり母数不足` };
  }

  const pct = (x, y) => (y ? Math.round((x / y - 1) * 100) : null);
  const facts = [];
  const ctr = pct(f.ctr, b.ctr);
  const cpm = pct(f.cpm, b.cpm);
  if (ctr !== null) facts.push(`CTR ${ctr >= 0 ? '+' : ''}${ctr}%`);
  if (cpm !== null) facts.push(`CPM ${cpm >= 0 ? '+' : ''}${cpm}%`);
  facts.push(`v2 ${b.v2}→${f.v2}件`);

  let result;
  if (!b.v2 || !f.v2) {
    result = f.v2 > b.v2 ? '改善' : f.v2 < b.v2 ? '悪化' : '変化なし';
  } else {
    const change = f.cpa / b.cpa - 1;
    result = change <= -EFFECT_THRESHOLD ? '改善' : change >= EFFECT_THRESHOLD ? '悪化' : '変化なし';
  }

  // どこが効いた／効かなかったかの読み
  let reading = '';
  if (cpm !== null && Math.abs(cpm) >= 20 && (ctr === null || Math.abs(ctr) < 10)) {
    reading = 'CTR はほぼ横ばいで CPM が大きく動いている。施策より外部要因（競合・季節）の影響が大きい可能性';
  } else if (ctr !== null && ctr >= 10 && result !== '改善') {
    reading = 'CTR は上がったが v2 につながっていない。バナーと LP・商品の訴求がずれている可能性';
  } else if (ctr !== null && ctr <= -10) {
    reading = 'CTR が下がっている。施策で訴求が弱まったか、同じ人への表示が増えて飽きられた可能性';
  } else if (result === '改善') {
    reading = 'CTR・v2 の改善が CPA に反映されている。施策が効いた可能性が高い';
  }
  return { result, note: facts.join(' / ') + (reading ? `。${reading}` : '') };
}

// ---------------------------------------------------------------- 共通

function fetchAdInsights_(range) {
  return fetchAll_(`${AD_ACCOUNT_ID}/insights`, {
    level: 'ad',
    time_range: JSON.stringify(range),
    fields: 'ad_id,ad_name,impressions,clicks,spend,actions',
    limit: 500,
  });
}

function ymd_(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function writeMessageSheet_(ss, name, message) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.getRange(1, 1).setValue(message);
}
