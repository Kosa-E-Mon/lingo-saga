// ============================================================
// 🎮 英会話RPG - GAS WebApp (code.gs)
// スプレッドシートの「拡張機能 > Apps Script」に貼り付け
// 初回のみ setup() を実行してからデプロイしてください
// ============================================================

const SHEET_SESSIONS = 'sessions';
const SHEET_CONFIG   = 'config';

const PARAM_KEYS = [
  'pronunciation','vocabulary','grammar','fluency','listening',
  'expression','naturalness','initiative','pragmatics'
];

const HEADERS = [
  'date','topic','difficulty','minutes','level','cefr','class',
  ...PARAM_KEYS,
  'exp_gained','feeling','phrases_learned','good_points','improve_points'
];

const DIFF_MULT = { 'Easy': 0.8, 'Normal': 1.0, 'Hard': 1.5 };

// 累計EXPによる冒険者ランク（絶対に下がらない）
const RANKS = [
  { rank: 'G', need: 0 },
  { rank: 'F', need: 1000 },
  { rank: 'E', need: 3500 },
  { rank: 'D', need: 8500 },
  { rank: 'C', need: 18500 },
  { rank: 'B', need: 38500 },
  { rank: 'A', need: 78500 },
  { rank: 'S', need: 158500 }
];

const LEVEL_TIERS = [
  { min: 1,   max: 99,  cefr: 'A1' },
  { min: 100, max: 299, cefr: 'A2-B1' },
  { min: 300, max: 499, cefr: 'B1-B2' },
  { min: 500, max: 699, cefr: 'B2-C1' },
  { min: 700, max: 899, cefr: 'C1-C2' },
  { min: 900, max: 999, cefr: 'C2+' }
];

// ============================================================
// 初期セットアップ（1回だけ手動実行）
// ============================================================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sh = ss.getSheetByName(SHEET_SESSIONS);
  if (!sh) sh = ss.insertSheet(SHEET_SESSIONS);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }

  let cf = ss.getSheetByName(SHEET_CONFIG);
  if (!cf) cf = ss.insertSheet(SHEET_CONFIG);
  if (cf.getLastRow() === 0) {
    cf.getRange(1, 1, 3, 2).setValues([
      ['key', 'value'],
      ['target_level', 700],
      ['target_label', '全国通訳案内士']
    ]);
  }
}

// ============================================================
// GET: ?action=getDashboard / ?action=getSessions
// ============================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getDashboard';
  try {
    if (action === 'getSessions') {
      return jsonOut({ status: 'success', data: readSessions() });
    }
    return jsonOut({ status: 'success', data: buildDashboard() });
  } catch (err) {
    return jsonOut({ status: 'error', message: String(err) });
  }
}

// ============================================================
// POST: セッション追加（relay.htmlから送信）
// EXPはここで自動計算する（GPTには計算させない）
// ============================================================
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);

    // パラメーター平均から出来栄え係数を算出
    const vals = PARAM_KEYS.map(k => Number(p[k]) || 0);
    const paramAvg = vals.reduce((a, b) => a + b, 0) / PARAM_KEYS.length;

    const minutes = Number(p.minutes) || 10;
    const mult = DIFF_MULT[p.difficulty] || 1.0;
    const exp = Math.round(minutes * 10 * mult * (paramAvg / 50));

    const row = HEADERS.map(h => {
      if (h === 'exp_gained') return exp;
      if (h === 'date') return p.date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
      return p[h] !== undefined ? p[h] : '';
    });

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESSIONS);
    sh.appendRow(row);

    const dash = buildDashboard();
    return jsonOut({
      status: 'success',
      exp_gained: exp,
      dashboard: dash
    });
  } catch (err) {
    return jsonOut({ status: 'error', message: String(err) });
  }
}

// ============================================================
// セッション読み込み
// ============================================================
function readSessions() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESSIONS);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).getValues();
  return values.map(r => {
    const o = {};
    HEADERS.forEach((h, i) => {
      let v = r[i];
      if (v instanceof Date) v = Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
      o[h] = v;
    });
    return o;
  });
}

// ============================================================
// ダッシュボードデータ生成（全計算の中心）
// ============================================================
function buildDashboard() {
  const sessions = readSessions();
  const config = readConfig();
  const n = sessions.length;

  // --- 累計EXP・ランク（下がらない） ---
  const expTotal = sessions.reduce((a, s) => a + (Number(s.exp_gained) || 0), 0);
  let rank = RANKS[0], next = null;
  for (let i = 0; i < RANKS.length; i++) {
    if (expTotal >= RANKS[i].need) rank = RANKS[i];
    else { next = RANKS[i]; break; }
  }
  const nextRankExp = next ? (next.need - expTotal) : 0;

  // --- 実力レベル（トリム平均・上下する） ---
  const recent10 = sessions.slice(-10);
  const levels = recent10.map(s => Number(s.level) || 0);
  let currentLevel = 0;
  let tutorial = n < 10;
  if (n === 0) {
    currentLevel = 0;
  } else if (tutorial) {
    currentLevel = Math.round(levels.reduce((a, b) => a + b, 0) / levels.length);
  } else {
    const sorted = [...levels].sort((a, b) => a - b);
    const trimmed = sorted.slice(2, 8); // 上下2個ずつ除外 → 6個
    currentLevel = Math.round(trimmed.reduce((a, b) => a + b, 0) / trimmed.length);
  }
  const tier = LEVEL_TIERS.find(t => currentLevel >= t.min && currentLevel <= t.max);

  // --- 直近10回のパラメーター平均 ---
  const out = {
    session_count: String(n),
    exp_total: String(expTotal),
    current_level: String(currentLevel),
    cefr: tier ? tier.cefr : '--',
    class: n > 0 ? (sessions[n - 1].class || 'Traveler') : 'Traveler',
    rank: rank.rank,
    next_rank_exp: String(nextRankExp),
    tutorial: tutorial,
    tutorial_count: String(Math.min(n, 10)),
    target_level: String(config.target_level || 700),
    target_label: config.target_label || '',
    remaining: String(Math.max(0, (Number(config.target_level) || 700) - currentLevel))
  };

  const avgs = {};
  PARAM_KEYS.forEach(k => {
    const vals = recent10.map(s => Number(s[k]) || 0).filter(v => v > 0);
    avgs[k] = vals.length ? (vals.reduce((a, b) => a + b, 0) / vals.length) : 0;
    out['avg_' + k] = avgs[k].toFixed(1);
  });

  // --- 弱点（平均が低い順に3つ） ---
  out.weak_params = Object.keys(avgs)
    .sort((a, b) => avgs[a] - avgs[b])
    .slice(0, 3)
    .join(', ');

  return out;
}

// ============================================================
// データ全消去（Apps Script画面から手動実行する管理用関数）
// テンプレート整備時や、最初からやり直したいときに使う
// ヘッダー行は残し、セッションデータだけ削除する
// ============================================================
function resetAllData() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    '⚠️ データ全消去',
    'すべてのセッションデータを削除します。この操作は取り消せません。実行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SESSIONS);
  if (sh && sh.getLastRow() > 1) {
    sh.deleteRows(2, sh.getLastRow() - 1);
  }
  ui.alert('✅ 削除しました。冒険はまっさらな状態から始まります。');
}

// ============================================================
// ユーティリティ
// ============================================================
function readConfig() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  const o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .forEach(r => { if (r[0]) o[r[0]] = r[1]; });
  return o;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
