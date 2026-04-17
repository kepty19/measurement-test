/**
 * 現状把握テスト — 結果をスプレッドシートに記録し、「集計」シートを氏名キーで更新する。
 *
 * セットアップ:
 * 1. このスプレッドシートに紐づけてプロジェクトを作成し、本ファイルを貼り付ける。
 * 2. SPREADSHEET_ID をこのブックの ID に設定（通常は自動で SpreadsheetApp.getActiveSpreadsheet() でも可）。
 * 3. デプロイ → 新しいデプロイ → 種類: ウェブアプリ
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員（匿名ユーザーを含む）
 * 4. 発行された URL をフロントの SUBMIT_URL に設定。
 * 5. （任意）SECRET を設定し、フロントの SUBMIT_SECRET と一致させる。
 */

var SPREADSHEET_ID = '19GdI5qQWc-VyLQEgRiJfaxxTLwrtxU6ofx4tZYPax0M';
/** 空文字なら検証しない */
var SECRET = '';

var SHEET_RESULTS = 'Results';
var SHEET_SUMMARY = '集計';

function getSs_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function ensureResultsHeaders_(sh) {
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
    'スピーキング_完了数',
    'スピーキング_目標数',
    'スピーキング_完了率',
    '単語_回答JSON',
    '文法_回答JSON',
    'スピーキング_JSON',
  ]);
}

function ensureSummaryHeaders_(sh) {
  if (sh.getLastRow() >= 1) return;
  sh.appendRow([
    '氏名',
    '最新実施日時',
    '受験回数',
    '単語_平均正答率',
    '文法_平均正答率',
    'スピーキング_平均完了率',
    '単語_最新正答率',
    '文法_最新正答率',
    'スピーキング_最新完了率',
  ]);
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

    var ss = getSs_();
    var results = ss.getSheetByName(SHEET_RESULTS);
    if (!results) results = ss.insertSheet(SHEET_RESULTS);
    ensureResultsHeaders_(results);

    var name = String(body.name || '').trim();
    if (!name) {
      return jsonResponse_({ ok: false, error: 'name required' });
    }

    var now = new Date();
    var iso = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");

    var v = body.scores && body.scores.vocabulary ? body.scores.vocabulary : {};
    var g = body.scores && body.scores.grammar ? body.scores.grammar : {};
    var s = body.scores && body.scores.speaking ? body.scores.speaking : {};

    var vocabJson = JSON.stringify(body.raw && body.raw.vocabulary ? body.raw.vocabulary : {});
    var grammarJson = JSON.stringify(body.raw && body.raw.grammar ? body.raw.grammar : {});
    var speakJson = JSON.stringify(body.raw && body.raw.speaking ? body.raw.speaking : {});

    results.appendRow([
      iso,
      name,
      num_(v.correct),
      num_(v.total),
      num_(v.ratePercent),
      num_(g.correct),
      num_(g.total),
      num_(g.ratePercent),
      num_(s.completed),
      num_(s.target),
      num_(s.ratePercent),
      vocabJson,
      grammarJson,
      speakJson,
    ]);

    refreshSummaryForName_(ss, name);

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

/** デプロイ確認用（ブラウザで URL を開いて応答が返れば有効） */
function doGet() {
  return jsonResponse_({ ok: true, message: 'measurement-test endpoint' });
}

function jsonResponse_(obj) {
  var out = ContentService.createTextOutput(JSON.stringify(obj));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}

/**
 * Results から氏名が一致する行を集計し、集計シートの該当行を更新（なければ追加）。
 */
function refreshSummaryForName_(ss, name) {
  var results = ss.getSheetByName(SHEET_RESULTS);
  if (!results || results.getLastRow() < 2) return;

  var data = results.getDataRange().getValues();
  var header = data[0];
  var idxName = header.indexOf('氏名');
  var idxTime = header.indexOf('実施日時');
  var idxVr = header.indexOf('単語_正答率');
  var idxGr = header.indexOf('文法_正答率');
  var idxSr = header.indexOf('スピーキング_完了率');
  if (idxName < 0) idxName = 1;

  var rows = [];
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idxName]).trim() === name) rows.push(data[r]);
  }
  if (!rows.length) return;

  var n = rows.length;
  var last = rows[rows.length - 1];
  var sumV = 0;
  var sumG = 0;
  var sumS = 0;
  for (var i = 0; i < rows.length; i++) {
    sumV += Number(rows[i][idxVr]) || 0;
    sumG += Number(rows[i][idxGr]) || 0;
    sumS += Number(rows[i][idxSr]) || 0;
  }

  var summary = ss.getSheetByName(SHEET_SUMMARY);
  if (!summary) summary = ss.insertSheet(SHEET_SUMMARY);
  ensureSummaryHeaders_(summary);

  var sData = summary.getDataRange().getValues();
  var rowIndex = -1;
  for (var sr = 1; sr < sData.length; sr++) {
    if (String(sData[sr][0]).trim() === name) {
      rowIndex = sr + 1;
      break;
    }
  }

  var newRow = [
    name,
    last[idxTime],
    n,
    Math.round((sumV / n) * 100) / 100,
    Math.round((sumG / n) * 100) / 100,
    Math.round((sumS / n) * 100) / 100,
    last[idxVr],
    last[idxGr],
    last[idxSr],
  ];

  if (rowIndex > 0) {
    summary.getRange(rowIndex, 1, rowIndex, newRow.length).setValues([newRow]);
  } else {
    summary.appendRow(newRow);
  }
}
