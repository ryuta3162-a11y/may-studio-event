/**
 * JOYFIT24経堂 5月スタジオイベント LP 用 Google Apps Script
 *
 * 手順（概要）:
 * 1. 対象スプレッドシートを開く
 *    https://docs.google.com/spreadsheets/d/1RPUw0slNCit9ZwJgINGfv89oc2Hxw8zzAZyMt6g_QuY/edit
 * 2. gid=799051187 のシート（タブ）名を確認し、下の SHEET_NAME を一致させる
 * 3. 1行目に HEADER_ROW の列名をその順序で貼り付ける（既存列がある場合は Code.gs 側を合わせる／メール・希望内容列を追加）
 * 4. 本ファイルを「拡張機能」>「Apps Script」に貼り付け
 * 5. 「デプロイ」>「新しいデプロイ」> 種類「ウェブアプリ」
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員（または社内ポリシーに合わせて変更）
 * 6. 発行されたURLを index.html の webAppUrl に貼り付け
 */

var SPREADSHEET_ID = '1RPUw0slNCit9ZwJgINGfv89oc2Hxw8zzAZyMt6g_QuY';
/** シートのタブ名（gid799051187のシート名に合わせて変更） */
var SHEET_NAME = 'シート1';

/** 1行目のヘッダー（この順でデータを書き込みます） */
var HEADER_ROW = [
  'タイムスタンプ',
  'フォーム種別',
  '店舗名',
  '氏名',
  'メールアドレス',
  '電話番号',
  '参加希望レッスン',
  '参加希望日',
  '希望開始時間',
  '参加希望日時まとめ',
  '希望内容（自由記入）',
  '備考',
  '同意'
];

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, message: 'kyodo event gas alive' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return jsonResponse({ result: 'error', message: 'サーバーが混雑しています。しばらくしてから再度お試しください。' });
  }

  try {
    var p = normalizeParams_(e);
    validateParams_(p);

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error('シートが見つかりません: ' + SHEET_NAME);
    }

    ensureHeaderRow_(sheet);

    var row = [
      new Date(),
      p.formType,
      p.store_name,
      p.name,
      p.email,
      p.phone,
      p.event_slot,
      p.visit_date,
      p.visit_time,
      p.visit_datetime,
      p.lesson_memo,
      p.remarks,
      p.consent
    ];

    sheet.appendRow(row);

    return jsonResponse({ result: 'success' });
  } catch (err) {
    return jsonResponse({ result: 'error', message: err.message || String(err) });
  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function normalizeParams_(e) {
  var p = {};
  if (!e || !e.parameter) {
    return p;
  }
  var raw = e.parameter;
  p.formType = String(raw.formType || '');
  p.store_name = String(raw.store_name || '');
  p.name = String(raw.name || '');
  p.email = String(raw.email || '');
  p.phone = String(raw.phone || '');
  p.event_slot = String(raw.event_slot || '');
  p.visit_date = String(raw.visit_date || '');
  p.visit_time = String(raw.visit_time || '');
  p.visit_datetime = String(raw.visit_datetime || '');
  p.lesson_memo = String(raw.lesson_memo || '');
  p.remarks = String(raw.remarks || '');
  p.consent = String(raw.consent || '');
  return p;
}

function validateParams_(p) {
  if (!p.name) throw new Error('氏名は必須です。');
  if (!p.email || !p.email.trim()) throw new Error('メールアドレスは必須です。');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(p.email.trim())) {
    throw new Error('メールアドレスの形式が正しくありません。');
  }
  if (!p.phone) throw new Error('電話番号は必須です。');
  var hasSlot = !!(p.event_slot && String(p.event_slot).trim());
  var hasMemo = !!(p.lesson_memo && String(p.lesson_memo).trim());
  if (!hasSlot && !hasMemo) {
    throw new Error('参加希望のレッスンを選択するか、「希望内容」にご記入ください。');
  }
  if (hasSlot && !p.visit_date) throw new Error('参加希望日が取得できませんでした。もう一度お試しください。');
  if (!p.consent) throw new Error('同意が必要です。');
}

/**
 * 1行目が空、または先頭セルがタイムスタンプでない場合にヘッダーを書き込みます。
 * 既存データがあるシートでは手動でヘッダーを合わせることを推奨します。
 */
function ensureHeaderRow_(sheet) {
  var first = sheet.getRange(1, 1, 1, HEADER_ROW.length).getValues()[0];
  var empty = first.every(function (c) { return c === '' || c === null; });
  if (empty) {
    sheet.getRange(1, 1, 1, HEADER_ROW.length).setValues([HEADER_ROW]);
    return;
  }
  if (String(first[0]) !== HEADER_ROW[0]) {
    // 既存列と異なる場合は上書きしない（データ破壊防止）
    // 必要なら手動で1行目を HEADER_ROW に合わせてください。
  }
}
