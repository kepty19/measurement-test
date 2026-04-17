/**
 * 現状把握テスト — 「Result」1シートにスコア＋氏名キー集計、「raw data」に単語・文法のローデータを記録する。
 * スピーキングは LINE 側で取得するためここには書き込まない。
 *
 * セットアップ:
 * 1. このスプレッドシートに紐づけてプロジェクトを作成し、本ファイルを貼り付ける。
 * 2. SPREADSHEET_ID をこのブックの ID に合わせる。
 * 3. デプロイ → 新しいデプロイ → 種類: ウェブアプリ
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員（匿名ユーザーを含む）
 * 4. 発行された URL をフロントの config.js の SUBMIT_URL に設定。
 * 5. （任意）SECRET を設定し、フロントの SUBMIT_SECRET と一致させる。
 *
 * 旧「Results」「集計」シートを使っていた場合は、不要なら削除してよい。
 */

var SPREADSHEET_ID = '19GdI5qQWc-VyLQEgRiJfaxxTLwrtxU6ofx4tZYPax0M';
var SECRET = '';

var SHEET_RESULT = 'Result';
var SHEET_RAW = 'raw data';

function getSs_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function ensureResultHeaders_(sh) {
  if (sh.getLastRow() >= 1) return;
  sh.appendRow([
    '実施日時',
    '氏名',
    '単語_正答数',
    '単語_問題数',
    '単語_正答率',
    '文法_正答数',
    '文法_問題数',
    '文法_正答率',
    '受験回数_累計',
    '単語_平均正答率_累計',
    '文法_平均正答率_累計',
  ]);
}

function ensureRawHeaders_(sh) {
  if (sh.getLastRow() >= 1) return;
  sh.appendRow(['実施日時', '氏名', 'セクション', '問題番号', '問題文', '正解', '回答', '合否']);
}

function buildAggregates_(resultSheet, name, vRate, gRate) {
  var data = resultSheet.getDataRange().getValues();
  if (data.length < 2) {
    return { n: 1, avgV: numOr_(vRate), avgG: numOr_(gRate) };
  }
  var header = data[0];
  var idxName = header.indexOf('氏名');
  var idxVr = header.indexOf('単語_正答率');
  var idxGr = header.indexOf('文法_正答率');
  if (idxName < 0) idxName = 1;
  if (idxVr < 0) idxVr = 4;
  if (idxGr < 0) idxGr = 7;

  var sumV = 0;
  var sumG = 0;
  var count = 0;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idxName]).trim() === name) {
      count++;
      sumV += Number(data[r][idxVr]) || 0;
      sumG += Number(data[r][idxGr]) || 0;
    }
  }
  count++;
  sumV += Number(vRate) || 0;
  sumG += Number(gRate) || 0;
  return {
    n: count,
    avgV: Math.round((sumV / count) * 100) / 100,
    avgG: Math.round((sumG / count) * 100) / 100,
  };
}

function appendRawData_(ss, iso, name, vocabRaw, grammarRaw) {
  var raw = ss.getSheetByName(SHEET_RAW);
  if (!raw) raw = ss.insertSheet(SHEET_RAW);
  ensureRawHeaders_(raw);

  var vmax = Math.min(20, vocabRaw.length);
  var gmax = Math.min(20, grammarRaw.length);

  for (var i = 0; i < vmax; i++) {
    var item = vocabRaw[i];
    raw.appendRow([
      iso,
      name,
      '単語',
      i + 1,
      item.english || '',
      item.expectedJapanese || '',
      item.userAnswer || '',
      item.isCorrect ? '○' : '×',
    ]);
  }
  for (var j = 0; j < gmax; j++) {
    var g = grammarRaw[j];
    raw.appendRow([
      iso,
      name,
      '文法',
      j + 1,
      g.question || '',
      g.correctAnswer || '',
      g.chosen || '',
      g.isCorrect ? '○' : '×',
    ]);
  }
}

function numOr_(x) {
  var n = Number(x);
  return isNaN(n) ? 0 : Math.round(n * 100) / 100;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ ok: false, error: 'empty body' });
    }
    var body = JSON.parse(e.postData.contents);
    if (SECRET && body.secret !== SECRET) {
      return jsonResponse_({ ok: false, error: 'unauthorized' });
    }

    var name = String(body.name || '').trim();
    if (!name) {
      return jsonResponse_({ ok: false, error: 'name required' });
    }

    var v = body.scores && body.scores.vocabulary ? body.scores.vocabulary : {};
    var g = body.scores && body.scores.grammar ? body.scores.grammar : {};

    var now = new Date();
    var iso = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");

    var vocabRaw = body.raw && body.raw.vocabulary ? body.raw.vocabulary : [];
    var grammarRaw = body.raw && body.raw.grammar ? body.raw.grammar : [];

    var ss = getSs_();
    var resultSheet = ss.getSheetByName(SHEET_RESULT);
    if (!resultSheet) resultSheet = ss.insertSheet(SHEET_RESULT);
    ensureResultHeaders_(resultSheet);

    var agg = buildAggregates_(resultSheet, name, v.ratePercent, g.ratePercent);

    resultSheet.appendRow([
      iso,
      name,
      num_(v.correct),
      num_(v.total),
      num_(v.ratePercent),
      num_(g.correct),
      num_(g.total),
      num_(g.ratePercent),
      agg.n,
      agg.avgV,
      agg.avgG,
    ]);

    appendRawData_(ss, iso, name, vocabRaw, grammarRaw);

    return jsonResponse_({ ok: true });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

function num_(x) {
  if (x === null || x === undefined || x === '') return '';
  var n = Number(x);
  return isNaN(n) ? '' : n;
}

function doGet() {
  return jsonResponse_({ ok: true, message: 'measurement-test endpoint' });
}

function jsonResponse_(obj) {
  var out = ContentService.createTextOutput(JSON.stringify(obj));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}
