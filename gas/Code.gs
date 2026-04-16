/**
 * JOYFIT24経堂 5月スタジオイベント LP 用 Google Apps Script
 *
 * 手順（概要）:
 * 1. 対象スプレッドシートを開く
 *    https://docs.google.com/spreadsheets/d/1Eba3Uvn4lRK5z4hshHdsOav6yMrpvQov6bT_5TWINag/edit
 * 2. gid=799051187 のシート（タブ）名を確認し、下の SHEET_NAME を一致させる
 * 3. 1行目に HEADER_ROW の列名をその順序で貼り付ける（既存列がある場合は Code.gs 側を合わせる／メール・希望内容列を追加）
 * 4. 本ファイルを「拡張機能」>「Apps Script」に貼り付け
 * 5. 「デプロイ」>「新しいデプロイ」> 種類「ウェブアプリ」
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員（または社内ポリシーに合わせて変更）
 * 6. 発行されたURLを index.html の webAppUrl に貼り付け
 */

var SPREADSHEET_ID = '1Eba3Uvn4lRK5z4hshHdsOav6yMrpvQov6bT_5TWINag';
/** シートのタブ名（gid799051187のシート名に合わせて変更） */
var SHEET_NAME = '5月スタジオイベント';
/** スプレッドシートURL（案内用） */
var SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1Eba3Uvn4lRK5z4hshHdsOav6yMrpvQov6bT_5TWINag/edit';

/** 通知先管理者 */
var ADMIN_EMAILS = [
  'r-kusaka@okamoto-group.co.jp',
  'jf-kyoudou@okamoto-group.co.jp',
  'yuka-hachiya@okamoto-group.co.jp'
];

/** 申請完了メールの返信先（差出人として使いたいアドレス） */
var REPLY_TO_EMAIL = 'jf-kyoudou@okamoto-group.co.jp';
/** 表示名 */
var SENDER_NAME = 'JOYFIT24経堂';
/** Gmailエイリアスが設定済みなら有効化（true） */
var USE_GMAIL_ALIAS_FROM = false;
/** Gmailエイリアスで送る場合の from アドレス */
var GMAIL_ALIAS_FROM = 'jf-kyoudou@okamoto-group.co.jp';

/** リマインド送信対象（日数前） */
var REMINDER_DAYS_BEFORE = 1;
/** リマインド一括送信上限（1回実行あたり） */
var REMINDER_BATCH_LIMIT = 100;
/** スクリプトのタイムゾーン */
var TZ = 'Asia/Tokyo';

/** 1行目のヘッダー（この順でデータを書き込みます） */
var HEADER_ROW = [
  'タイムスタンプ',
  '申込ID',
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
  '同意',
  '申請完了メール送信日時',
  '管理者通知送信日時',
  'リマインド送信状態',
  'リマインド送信日時'
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
    var requestId = buildRequestId_();

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      throw new Error('シートが見つかりません: ' + SHEET_NAME);
    }

    ensureHeaderRow_(sheet);

    var row = [
      new Date(),
      requestId,
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
      p.consent,
      '',
      '',
      '未送信',
      ''
    ];

    var rowNumber = sheet.getLastRow() + 1;
    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);

    var confirmedAt = '';
    var adminNotifiedAt = '';
    try {
      sendApplicantCompletedMail_(p, requestId);
      confirmedAt = formatDateTime_(new Date());
    } catch (mailErr) {
      Logger.log('applicant mail error: ' + mailErr);
    }
    try {
      sendAdminNotificationMail_(p, requestId, rowNumber);
      adminNotifiedAt = formatDateTime_(new Date());
    } catch (adminErr) {
      Logger.log('admin mail error: ' + adminErr);
    }

    if (confirmedAt) {
      sheet.getRange(rowNumber, getHeaderIndex_('申請完了メール送信日時')).setValue(confirmedAt);
    }
    if (adminNotifiedAt) {
      sheet.getRange(rowNumber, getHeaderIndex_('管理者通知送信日時')).setValue(adminNotifiedAt);
    }

    return jsonResponse({
      result: 'success',
      requestId: requestId
    });
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
  if (!hasSlot) throw new Error('参加希望のレッスンを1つ選択してください。');
  if (hasSlot && !p.visit_date) throw new Error('参加希望日が取得できませんでした。もう一度お試しください。');
  if (hasSlot && isLateForSameDayEntry_(p.visit_date, p.visit_time)) {
    throw new Error('当日申込みはレッスン開始30分前までです。別の時間帯をお選びください。');
  }
  if (!p.consent) throw new Error('同意が必要です。');
}

function isLateForSameDayEntry_(visitDate, visitTime) {
  if (!visitDate || !visitTime) return false;
  var now = new Date();
  var todayYmd = formatYmd_(now);
  if (String(visitDate) !== todayYmd) return false;

  var parts = String(visitTime).split(':');
  if (parts.length < 2) return false;
  var lessonH = parseInt(parts[0], 10);
  var lessonM = parseInt(parts[1], 10);
  if (isNaN(lessonH) || isNaN(lessonM)) return false;
  var lessonTotal = lessonH * 60 + lessonM;
  var nowH = parseInt(Utilities.formatDate(now, TZ, 'H'), 10);
  var nowM = parseInt(Utilities.formatDate(now, TZ, 'm'), 10);
  var nowTotal = nowH * 60 + nowM;
  return nowTotal > (lessonTotal - 30);
}

/**
 * 前日リマインド送信（時間主導トリガーで1日1回推奨）
 * 例: 毎日 18:00 実行
 */
function sendReminderBatch() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('シートが見つかりません: ' + SHEET_NAME);
  ensureHeaderRow_(sheet);

  var headerMap = buildHeaderMap_(sheet);
  var last = sheet.getLastRow();
  if (last < 2) return;
  var values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();

  var sentCount = 0;
  var targetDate = new Date();
  targetDate.setDate(targetDate.getDate() + REMINDER_DAYS_BEFORE);
  var targetYmd = formatYmd_(targetDate);

  for (var i = 0; i < values.length; i++) {
    if (sentCount >= REMINDER_BATCH_LIMIT) break;
    var row = values[i];
    var rowNo = i + 2;
    var status = String(row[headerMap['リマインド送信状態']] || '');
    var email = String(row[headerMap['メールアドレス']] || '').trim();
    var visitDate = String(row[headerMap['参加希望日']] || '').trim();
    if (!email || !visitDate) continue;
    if (status === '送信済') continue;
    if (visitDate !== targetYmd) continue;

    var p = {
      name: String(row[headerMap['氏名']] || ''),
      store_name: String(row[headerMap['店舗名']] || 'JOYFIT24経堂'),
      event_slot: String(row[headerMap['参加希望レッスン']] || ''),
      visit_date: visitDate,
      visit_time: String(row[headerMap['希望開始時間']] || ''),
      visit_datetime: String(row[headerMap['参加希望日時まとめ']] || ''),
      lesson_memo: String(row[headerMap['希望内容（自由記入）']] || ''),
      phone: String(row[headerMap['電話番号']] || ''),
      email: email
    };
    var requestId = String(row[headerMap['申込ID']] || '');

    try {
      sendApplicantReminderMail_(p, requestId);
      sheet.getRange(rowNo, headerMap['リマインド送信状態'] + 1).setValue('送信済');
      sheet.getRange(rowNo, headerMap['リマインド送信日時'] + 1).setValue(formatDateTime_(new Date()));
      sentCount += 1;
    } catch (err) {
      Logger.log('reminder send failed row ' + rowNo + ': ' + err);
      sheet.getRange(rowNo, headerMap['リマインド送信状態'] + 1).setValue('送信失敗');
    }
  }
}

function setupReminderTrigger() {
  var fn = 'sendReminderBatch';
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === fn) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger(fn).timeBased().everyDays(1).atHour(18).create();
}

function sendApplicantCompletedMail_(p, requestId) {
  var to = p.email.trim();
  var subject = 'JOYFIT24経堂　5月無料体験会申請が完了しました。';
  var selected = formatLessonForApplicant_(p.event_slot, p.lesson_memo);
  var visit = p.visit_datetime || p.visit_date || '未指定';

  var body =
    p.name + ' 様\n' +
    'この度は、JOYFIT24経堂 ホットスタジオレッスン無料体験会へお申し込みいただきありがとうございます。\n' +
    '以下内容で受け付けました。\n\n' +
    '■お申し込み内容\n' +
    selected + '\n' +
    '希望日時: ' + visit + '\n' +
    '※レッスン30分前にご来館下さい。\n\n' +
    '※当日申込みはレッスン開始30分前までです。\n\n' +
    '■当日の流れ\n' +
    '1) インターホンを鳴らしてお呼び出しください。スタッフが開錠します。\n' +
    '2) 2階で体験チケットとタオル・マットをお渡しし、3階ロッカールームでご準備いただきます。\n' +
    '3) レッスン開始10分前からスタジオに入館できます。体験チケットをインストラクターへお渡しください。\n\n' +
    '※レッスンにより満員となる場合がございます。あらかじめご了承ください。\n' +
    '※タオルレンタルには限りがあるため、可能であればご持参ください。\n\n' +
    '※本メールは自動送信です。\n' +
    '※内容変更・キャンセル等こちらのメールへご返信ください';

  var htmlBody = buildApplicantCompletedHtml_(p, selected, visit, requestId);
  sendMail_(to, subject, body, htmlBody);
}

function sendApplicantReminderMail_(p, requestId) {
  var to = p.email.trim();
  var subject = '【前日リマインド】明日はJOYFIT24経堂 無料体験レッスンです';
  var selected = p.event_slot && p.event_slot.trim() ? p.event_slot : '（自由記入）' + p.lesson_memo;
  var body =
    p.name + ' 様\n\n' +
    '明日は、JOYFIT24経堂 ホットスタジオ無料体験会のご予約日です。\n\n' +
    '■申込ID\n' + requestId + '\n\n' +
    '■ご予約内容\n' +
    '参加希望: ' + selected + '\n' +
    '希望日時: ' + (p.visit_datetime || p.visit_date || '未指定') + '\n\n' +
    '※当日申込みはレッスン開始30分前までです。\n\n' +
    '■当日の流れ\n' +
    '1) インターホンを鳴らしてお呼び出しください。スタッフが開錠します。\n' +
    '2) 2階で体験チケットとタオル・マットをお渡しし、3階ロッカールームでご準備いただきます。\n' +
    '3) レッスン開始10分前からスタジオに入館できます。体験チケットをインストラクターへお渡しください。\n\n' +
    '※レッスンにより満員となる場合がございます。あらかじめご了承ください。\n' +
    '■持ち物\n' +
    '動きやすいウェア、飲み物\n' +
    '※フェイスタオル・バスタオル・ヨガマットは無料貸出（数に限りあり）\n' +
    '※タオルレンタルには限りがあるため、可能であればご持参ください。\n\n' +
    'ご来館をお待ちしております。';
  sendMail_(to, subject, body);
}

function sendAdminNotificationMail_(p, requestId, rowNumber) {
  var to = ADMIN_EMAILS.join(',');
  var subject = '【新規申込】JOYFIT24経堂 ホットスタジオ無料体験会';
  var selected = p.event_slot && p.event_slot.trim() ? p.event_slot : '（自由記入）' + p.lesson_memo;
  var body =
    '新しい申込がありました。\n\n' +
    '■申込ID\n' + requestId + '\n' +
    '■受付時刻\n' + formatDateTime_(new Date()) + '\n\n' +
    '■申込者情報\n' +
    '氏名: ' + p.name + '\n' +
    'メール: ' + p.email + '\n' +
    '電話番号: ' + p.phone + '\n\n' +
    '■申込内容\n' +
    '参加希望: ' + selected + '\n' +
    '希望日時: ' + (p.visit_datetime || p.visit_date || '未指定') + '\n' +
    (p.remarks ? '備考: ' + p.remarks + '\n' : '') +
    '\n' +
    'シート行: ' + rowNumber + '\n' +
    'スプレッドシート: ' + SPREADSHEET_URL;
  sendMail_(to, subject, body);
}

function sendMail_(to, subject, body, htmlBody) {
  var options = {
    name: SENDER_NAME,
    replyTo: REPLY_TO_EMAIL
  };
  if (htmlBody) options.htmlBody = htmlBody;
  if (USE_GMAIL_ALIAS_FROM) {
    options.from = GMAIL_ALIAS_FROM;
    GmailApp.sendEmail(to, subject, body, options);
    return;
  }
  options.to = to;
  options.subject = subject;
  options.body = body;
  MailApp.sendEmail(options);
}

function formatLessonForApplicant_(eventSlot, lessonMemo) {
  var slot = String(eventSlot || '').trim();
  if (!slot) return '（自由記入）' + String(lessonMemo || '');
  var parts = slot.split('｜');
  if (parts.length >= 5) {
    return parts[0] + '｜' + parts[2] + '｜' + parts[3] + '｜' + parts[4];
  }
  return slot;
}

function escHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function buildApplicantCompletedHtml_(p, selected, visit, requestId) {
  var name = escHtml_(p.name);
  var selectedEsc = escHtml_(selected);
  var visitEsc = escHtml_(visit);

  return (
    '<div style="margin:0;padding:24px;background:#f4f4f5;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Helvetica,Arial,sans-serif;color:#171717;">' +
      '<div style="max-width:680px;margin:0 auto;background:#ffffff;border:1px solid #e5e7eb;border-radius:18px;overflow:hidden;box-shadow:0 12px 30px rgba(0,0,0,.06);">' +
        '<div style="background:linear-gradient(135deg,#c21642 0%,#9d1135 100%);padding:20px 24px;color:#fff;">' +
          '<div style="font-size:14px;opacity:.9;letter-spacing:.04em;">JOYFIT24経堂</div>' +
          '<h1 style="margin:8px 0 0;font-size:22px;line-height:1.4;">5月無料体験会 申請完了</h1>' +
        '</div>' +
        '<div style="padding:22px 24px 26px;">' +
          '<p style="margin:0 0 10px;font-size:16px;font-weight:700;color:#111827;">' + name + ' 様</p>' +
          '<p style="margin:0 0 14px;font-size:14px;line-height:1.8;color:#374151;">この度は、JOYFIT24経堂 ホットスタジオレッスン無料体験会へお申し込みいただきありがとうございます。<br>以下内容で受け付けました。</p>' +
          '<h2 style="margin:16px 0 8px;font-size:15px;color:#111827;">■お申し込み内容</h2>' +
          '<p style="margin:0 0 6px;font-size:14px;line-height:1.8;color:#111827;">' + selectedEsc + '</p>' +
          '<p style="margin:0 0 6px;font-size:14px;line-height:1.8;color:#111827;">希望日時: ' + visitEsc + '</p>' +
          '<p style="margin:0 0 14px;font-size:13px;line-height:1.7;color:#c21642;font-weight:700;">※レッスン30分前にご来館下さい。</p>' +
          '<p style="margin:0 0 14px;font-size:13px;line-height:1.7;color:#c21642;font-weight:700;">※当日申込みはレッスン開始30分前までです。</p>' +
          '<h2 style="margin:16px 0 8px;font-size:15px;color:#111827;">■当日の流れ</h2>' +
          '<ol style="margin:0;padding-left:18px;font-size:14px;line-height:1.9;color:#374151;">' +
            '<li>インターホンを鳴らしてお呼び出しください。スタッフが開錠します。</li>' +
            '<li>2階で体験チケットとタオル・マットをお渡しし、3階ロッカールームでご準備いただきます。</li>' +
            '<li>レッスン開始10分前からスタジオに入館できます。体験チケットをインストラクターへお渡しください。</li>' +
          '</ol>' +
          '<div style="margin-top:12px;padding:12px 13px;border-radius:10px;background:#fff8f8;border:1px solid #ffd6dd;font-size:13px;line-height:1.8;color:#7f1d1d;">' +
            '※レッスンにより満員となる場合がございます。あらかじめご了承ください。<br>' +
            '※タオルレンタルには限りがあるため、可能であればご持参ください。' +
          '</div>' +
          '<div style="margin-top:16px;padding-top:12px;border-top:1px solid #e5e7eb;font-size:12px;line-height:1.8;color:#6b7280;">' +
            '※本メールは自動送信です。<br>' +
            '※内容変更・キャンセル等こちらのメールへご返信ください' +
          '</div>' +
        '</div>' +
      '</div>' +
    '</div>'
  );
}

function buildRequestId_() {
  var d = new Date();
  return 'KYODO-' + Utilities.formatDate(d, TZ, 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}

function formatDateTime_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy/MM/dd HH:mm:ss');
}

function formatYmd_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function buildHeaderMap_(sheet) {
  var h = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < h.length; i++) {
    map[String(h[i])] = i;
  }
  return map;
}

function getHeaderIndex_(headerName) {
  for (var i = 0; i < HEADER_ROW.length; i++) {
    if (HEADER_ROW[i] === headerName) return i + 1;
  }
  throw new Error('header not found: ' + headerName);
}

/**
 * 1行目が空、または先頭セルがタイムスタンプでない場合にヘッダーを書き込みます。
 * 既存データがあるシートでは手動でヘッダーを合わせることを推奨します。
 */
function ensureHeaderRow_(sheet) {
  var col = Math.max(sheet.getLastColumn(), HEADER_ROW.length);
  var first = sheet.getRange(1, 1, 1, col).getValues()[0];
  var empty = first.every(function (c) { return c === '' || c === null; });
  if (empty) {
    sheet.getRange(1, 1, 1, HEADER_ROW.length).setValues([HEADER_ROW]);
    return;
  }
  if (String(first[0]) !== HEADER_ROW[0]) {
    // 既存列と異なる場合は上書きしない（データ破壊防止）
    // 必要なら手動で1行目を HEADER_ROW に合わせてください。
    return;
  }
  // 既存ヘッダーに不足列がある場合は右側を補完
  for (var i = 0; i < HEADER_ROW.length; i++) {
    if (!first[i]) {
      sheet.getRange(1, i + 1).setValue(HEADER_ROW[i]);
    }
  }
}
