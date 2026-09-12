/**
 * 夫のタイプ診断 回答ロガー＋ヒアリングシート紐づけ（Apps Script）v2
 *
 * できること
 *  1) 診断アプリから POST(JSON) を受けて「夫タイプ診断」タブに1行追記（名前・メール・結果・全回答）
 *  2) メール（→名前）で専用ヒアリングシートの回答タブと照合し、その行の「診断（自動照合）」列に結果を書く
 *     - 診断が後から来た場合: doPost の中で即照合
 *     - ヒアリングシートが後から来た場合: フォーム送信トリガー（onHearingSubmit）で照合
 *     → 順序に関係なく、最終的にヒアリングシートの行に診断が乗る
 *
 * 置き場所: フォーム回答スプシ（下の SPREADSHEET_ID）の 拡張機能 → Apps Script
 * 初回だけ:
 *   a) HEARING_TAB を、専用ヒアリングシートの回答タブ名（例「フォームの回答 2」）に合わせる
 *   b) エディタで setup() を1回実行（権限承認＋フォーム送信トリガー作成）
 *   c) testAppend() を1回実行（「夫タイプ診断」タブが作られテスト行が入る）
 *   d) デプロイ → 新しいデプロイ → ウェブアプリ（実行ユーザー: 自分／アクセス: 全員）→ URL を index.html の CONFIG.GAS_URL へ
 */

const SPREADSHEET_ID = '1VYpivGlAL2qDK7Au3oR10l5NPytZCJclWdfEBX9gzBE'; // フォーム回答スプシ
const LOG_SPREADSHEET_ID = '';        // 診断の回答を別スプシに溜めたいときだけID。空なら同じスプシの新タブ
const LOG_TAB = '夫タイプ診断';         // 診断の回答が溜まるタブ
const HEARING_TAB = 'PS個別アプリあり'; // 専用ヒアリングシートの回答タブ名（2026-09-12 確認済み。タブ名を変えたらここも変える）
const LINK_COL_TITLE = '診断（自動照合）'; // ヒアリングシート側に足す列の見出し（無ければ末尾に自動作成）

const HEADER = [
  'タイムスタンプ', '名前', 'メール', '照合', '主タイプ', '副タイプ', '判定不能', '安全フラグ',
  '怒ったときの行動', '気持ち', '困る一言',
  '点数 貝', '点数 チワワ', '点数 小学生', '点数 栄養失調',
  'Q1 不機嫌の最初', 'Q2 話しかけた返事', 'Q3 LINE', 'Q4 話し合い', 'Q5 子ども', 'Q6 休日',
  'Q7 戻り方', 'Q8 外での様子', 'Q9 決め事', 'Q10 大変なとき', 'Q11 夫から話す', 'Q12 誘うと',
  'Q13 この1ヶ月', 'Q14 パターン', '結果URL', 'UA'
];
const COL = {}; HEADER.forEach((h, i) => COL[h] = i + 1);

/* ---------- ユーティリティ ---------- */
function normEmail_(s) {
  return String(s || '').trim().toLowerCase()
    .replace(/[Ａ-Ｚａ-ｚ０-９＠．＿－]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
}
function normName_(s) {
  return String(s || '').replace(/[\s　]/g, '').replace(/(さん|様)$/, '').toLowerCase();
}
function logSheet_() {
  const ss = SpreadsheetApp.openById(LOG_SPREADSHEET_ID || SPREADSHEET_ID);
  let sh = ss.getSheetByName(LOG_TAB);
  // 旧版（v1: メール列なし）の見出しのタブが残っていたら退避して作り直す（列ズレ防止）
  if (sh && sh.getLastRow() > 0) {
    const h = sh.getRange(1, 1, 1, 4).getValues()[0];
    if (h[2] !== 'メール' || h[3] !== '照合') {
      let n = 1; while (ss.getSheetByName(LOG_TAB + '_v1_' + n)) n++;
      sh.setName(LOG_TAB + '_v1_' + n);
      sh = null;
    }
  }
  if (!sh) sh = ss.insertSheet(LOG_TAB);
  if (sh.getLastRow() === 0) { sh.appendRow(HEADER); sh.setFrozenRows(1); }
  return sh;
}
function hearingSheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(HEARING_TAB);
}
function findHeaderCol_(sh, title) {
  const hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  const i = hdr.findIndex(h => String(h).trim() === title);
  return i >= 0 ? i + 1 : 0;
}
function findHeaderColLike_(sh, words) {
  const hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  const i = hdr.findIndex(h => words.some(w => String(h).indexOf(w) >= 0));
  return i >= 0 ? i + 1 : 0;
}
function ensureLinkCol_(sh) {
  let c = findHeaderCol_(sh, LINK_COL_TITLE);
  if (!c) { c = sh.getLastColumn() + 1; sh.getRange(1, c).setValue(LINK_COL_TITLE); }
  return c;
}
/** 診断の要約テキスト（ヒアリングシート側に書く中身） */
function summary_(d) {
  const L = [];
  L.push('【夫のタイプ診断】' + (d.undetermined ? '型が絞れなかった' : (d.main || '')) + (d.sub ? '（副: ' + d.sub + '）' : ''));
  if (d.mind) L.push('気持ち: ' + d.mind);
  L.push('怒ったときの行動チェック: ' + ((d.safety || []).length ? d.safety.join('、') : 'なし'));
  if (d.f) L.push('困る一言: ' + d.f);
  if (d.score) L.push('点数: 貝' + (d.score.kai || 0) + ' チワワ' + (d.score.chi || 0) + ' 小学生' + (d.score.sho || 0) + ' 栄養失調' + (d.score.eiyo || 0));
  const answers = d.answers || [];
  if (answers.length) { L.push('--- 回答の記録 ---'); answers.forEach((a, i) => { if (a) L.push('Q' + (i + 1) + ': ' + a); }); }
  if (d.q13) L.push('Q13: ' + d.q13);
  if (d.q14) L.push('Q14: ' + d.q14);
  L.push('（診断日時: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') + '）');
  return L.join('\n');
}

/* ---------- 1) 診断アプリからの受信 ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents || '{}');
    const sh = logSheet_();
    const answers = d.answers || [];
    const row = [
      new Date(), d.name || '', d.email || '', '',
      d.main || '', d.sub || '', d.undetermined ? '絞れず' : '', d.safe ? 'あり' : '',
      (d.safety || []).join('、'), d.mind || '', d.f || '',
      (d.score && d.score.kai) || 0, (d.score && d.score.chi) || 0, (d.score && d.score.sho) || 0, (d.score && d.score.eiyo) || 0
    ];
    for (let i = 0; i < 12; i++) row.push(answers[i] || '');
    row.push(d.q13 || '', d.q14 || '', d.resultUrl || '', d.ua || '');
    sh.appendRow(row);
    const logRow = sh.getLastRow();
    // 診断が後から来たケース: ヒアリングシートに同じ人がいれば今すぐ書く
    const linked = writeToHearing_(d.email, d.name, summary_(d));
    sh.getRange(logRow, COL['照合']).setValue(linked ? '済' : '未');
    return ContentService.createTextOutput(JSON.stringify({ ok: true, linked: linked }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}
function doGet() { return ContentService.createTextOutput('ok'); }

/** ヒアリングシートの回答タブから、メール→名前の順で同じ人の行を探して書く。書けたら true */
function writeToHearing_(email, name, text) {
  const hs = hearingSheet_();
  if (!hs || hs.getLastRow() < 2) return false;
  const emailCol = findHeaderColLike_(hs, ['メールアドレス', 'メール']);
  const nameCol = findHeaderColLike_(hs, ['お名前', '名前']);
  const last = hs.getLastRow();
  const emails = emailCol ? hs.getRange(2, emailCol, last - 1, 1).getValues().map(r => normEmail_(r[0])) : [];
  const names = nameCol ? hs.getRange(2, nameCol, last - 1, 1).getValues().map(r => normName_(r[0])) : [];
  let idx = -1;
  const ne = normEmail_(email), nn = normName_(name);
  if (ne) idx = emails.lastIndexOf(ne);               // 最新の行を優先
  if (idx < 0 && nn) idx = names.lastIndexOf(nn);      // メールで見つからなければ名前（完全一致のみ）
  if (idx < 0) return false;
  const linkCol = ensureLinkCol_(hs);
  hs.getRange(idx + 2, linkCol).setValue(text);
  return true;
}

/* ---------- 2) ヒアリングシートが後から送られたケース（フォーム送信トリガー） ---------- */
function onHearingSubmit(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== HEARING_TAB) return; // 旧シートや別フォームの送信は無視
    const r = e.range.getRow();
    const emailCol = findHeaderColLike_(sheet, ['メールアドレス', 'メール']);
    const nameCol = findHeaderColLike_(sheet, ['お名前', '名前']);
    const email = emailCol ? sheet.getRange(r, emailCol).getValue() : '';
    const name = nameCol ? sheet.getRange(r, nameCol).getValue() : '';
    const hit = findLatestLog_(email, name);
    if (!hit) return;
    const linkCol = ensureLinkCol_(sheet);
    sheet.getRange(r, linkCol).setValue(hit.text);
    logSheet_().getRange(hit.row, COL['照合']).setValue('済');
  } catch (err) {
    console.error('onHearingSubmit: ' + err);
  }
}
/** 診断ログから、メール→名前で最新の1件を探す */
function findLatestLog_(email, name) {
  const sh = logSheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const vals = sh.getRange(2, 1, last - 1, HEADER.length).getValues();
  const ne = normEmail_(email), nn = normName_(name);
  let idx = -1;
  for (let i = vals.length - 1; i >= 0; i--) { if (ne && normEmail_(vals[i][COL['メール'] - 1]) === ne) { idx = i; break; } }
  if (idx < 0) for (let i = vals.length - 1; i >= 0; i--) { if (nn && normName_(vals[i][COL['名前'] - 1]) === nn) { idx = i; break; } }
  if (idx < 0) return null;
  const v = vals[idx];
  const d = {
    main: v[COL['主タイプ'] - 1], sub: v[COL['副タイプ'] - 1], undetermined: !!v[COL['判定不能'] - 1],
    safety: String(v[COL['怒ったときの行動'] - 1] || '').split('、').filter(Boolean),
    mind: v[COL['気持ち'] - 1], f: v[COL['困る一言'] - 1],
    score: { kai: v[COL['点数 貝'] - 1], chi: v[COL['点数 チワワ'] - 1], sho: v[COL['点数 小学生'] - 1], eiyo: v[COL['点数 栄養失調'] - 1] },
    answers: HEADER.filter(h => /^Q(\d|1[0-2]) /.test(h)).map(h => v[COL[h] - 1]),
    q13: v[COL['Q13 この1ヶ月'] - 1], q14: v[COL['Q14 パターン'] - 1]
  };
  return { row: idx + 2, text: summary_(d) };
}

/* ---------- 初回セットアップ・テスト ---------- */
/** 1回だけ実行: フォーム送信トリガーを作る（既にあれば作らない） */
function setup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onHearingSubmit');
  if (!exists) ScriptApp.newTrigger('onHearingSubmit').forSpreadsheet(ss).onFormSubmit().create();
  logSheet_();
  Logger.log('setup done. HEARING_TAB=' + HEARING_TAB + ' exists=' + !!hearingSheet_());
}
/** 手動テスト: 1行入り、同じメール/名前の人がヒアリングシートにいれば紐づく */
function testAppend() {
  const fake = { postData: { contents: JSON.stringify({
    name: 'テスト 花子', email: 'test@example.com', main: '貝型', sub: 'でかいチワワ型', undetermined: false, safe: false,
    safety: [], mind: 'わからない', f: 'テスト送信', score: { kai: 14, chi: 2, sho: 0, eiyo: 0 },
    answers: ['貝｜黙る。部屋に行く。返事が「あー」「うん」になる'], q13: '黙る・無視', q14: '黙る→怒る→黙る、を繰り返す',
    resultUrl: '', ua: 'test' }) } };
  Logger.log(doPost(fake).getContent());
}
/** 未照合の診断をまとめて再照合（後から手で直したいとき用） */
function relinkAll() {
  const sh = logSheet_();
  const last = sh.getLastRow();
  if (last < 2) return;
  const vals = sh.getRange(2, 1, last - 1, HEADER.length).getValues();
  let n = 0;
  vals.forEach((v, i) => {
    if (v[COL['照合'] - 1] === '済') return;
    const hit = findLatestLog_(v[COL['メール'] - 1], v[COL['名前'] - 1]);
    if (hit && writeToHearing_(v[COL['メール'] - 1], v[COL['名前'] - 1], hit.text)) { sh.getRange(i + 2, COL['照合']).setValue('済'); n++; }
  });
  Logger.log('relinked: ' + n);
}
