/**
 * 夫のタイプ診断 回答ロガー（Apps Script Webアプリ）
 *
 * 置き場所: フォーム回答スプシ（1VYpivGlAL2qDK7Au3oR10l5NPytZCJclWdfEBX9gzBE）の
 *          拡張機能 → Apps Script に新規プロジェクトとして貼る
 * デプロイ: デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *          実行ユーザー「自分」／アクセスできるユーザー「全員」→ ウェブアプリのURLを
 *          index.html の CONFIG.GAS_URL に貼る
 * 動作: 診断アプリから POST(JSON) を受けて「夫タイプ診断」タブに1行追記する
 *       タブが無ければ自動で作る（ヘッダー付き）
 * 確認: ウェブアプリURLをブラウザで開くと "ok" が返る（doGet）
 */

const SPREADSHEET_ID = '1VYpivGlAL2qDK7Au3oR10l5NPytZCJclWdfEBX9gzBE';
const SHEET_NAME = '夫タイプ診断';

const HEADER = [
  'タイムスタンプ', '名前', '主タイプ', '副タイプ', '判定不能', '安全フラグ',
  '怒ったときの行動', '気持ち', '困る一言',
  '点数 貝', '点数 チワワ', '点数 小学生', '点数 栄養失調',
  'Q1 不機嫌の最初', 'Q2 話しかけた返事', 'Q3 LINE', 'Q4 話し合い', 'Q5 子ども', 'Q6 休日',
  'Q7 戻り方', 'Q8 外での様子', 'Q9 決め事', 'Q10 大変なとき', 'Q11 夫から話す', 'Q12 誘うと',
  'Q13 この1ヶ月', 'Q14 パターン', '結果URL', 'UA'
];

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents || '{}');
    const answers = d.answers || [];
    const row = [
      new Date(),
      d.name || '',
      d.main || '',
      d.sub || '',
      d.undetermined ? '絞れず' : '',
      d.safe ? 'あり' : '',
      (d.safety || []).join('、'),
      d.mind || '',
      d.f || '',
      (d.score && d.score.kai) || 0,
      (d.score && d.score.chi) || 0,
      (d.score && d.score.sho) || 0,
      (d.score && d.score.eiyo) || 0
    ];
    for (let i = 0; i < 12; i++) row.push(answers[i] || '');
    row.push(d.q13 || '', d.q14 || '', d.resultUrl || '', d.ua || '');
    getSheet_().appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function doGet() {
  return ContentService.createTextOutput('ok');
}

/** 手動テスト用: エディタから実行すると1行入る（初回の権限承認にも使う） */
function testAppend() {
  const fake = {
    postData: { contents: JSON.stringify({
      name: 'テスト', main: '貝型', sub: 'でかいチワワ型', undetermined: false, safe: false,
      safety: [], mind: 'わからない', f: 'テスト送信', score: { kai: 14, chi: 2, sho: 0, eiyo: 0 },
      answers: ['貝｜黙る。部屋に行く。返事が「あー」「うん」になる'], q13: '黙る・無視', q14: '黙る→怒る→黙る、を繰り返す',
      resultUrl: '', ua: 'test'
    }) }
  };
  Logger.log(doPost(fake).getContent());
}
