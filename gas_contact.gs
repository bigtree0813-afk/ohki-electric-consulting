/**
 * お問い合わせフォーム - Google Apps Script
 *
 * 【セットアップ手順】
 * 既存の診断用GASプロジェクトにこのコードを追加ファイルとして貼り付け
 * （Apps Scriptエディタで「＋」→「スクリプト」→ファイル名を「contact」に）
 * ※ doPost は既存のものを拡張するため、既存の doPost を下記の統合版に差し替えてください
 */

// ===== 設定 =====
var CONTACT_SHEET_NAME = 'お問い合わせ';
var CONTACT_REPLY_FROM_NAME = '電気工事 研修設計コンサルティング';
var CONTACT_REPLY_FROM_EMAIL = 'info@ohki-electric.com';
var NOTIFICATION_EMAIL = 'kenji@ohki-electric.com'; // 通知先メールアドレス
var SPREADSHEET_ID = '1QYcByw6S1h28ixODGvMV13ZFhsd8f7xLkg65lV9JYoA'; // スプレッドシートのID

// ===== お問い合わせ処理 =====
function handleContact(data) {
  saveContactToSheet(data);
  sendContactReplyEmail(data);
  sendContactNotificationEmail(data);
}

// ===== スプレッドシートへの記録 =====
function saveContactToSheet(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONTACT_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONTACT_SHEET_NAME);
    sheet.appendRow(['日時', 'お名前（会社名）', 'メールアドレス', 'ご相談内容']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#fff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(4, 400);
  }

  sheet.appendRow([
    new Date(),
    data.name,
    data.email,
    data.message || ''
  ]);
}

// ===== 問い合わせ者への自動返信メール =====
function sendContactReplyEmail(data) {
  var subject = '【ご相談ありがとうございます】電気工事 研修設計コンサルティング';

  var html = ''
    + '<!DOCTYPE html>'
    + '<html><head><meta charset="utf-8"></head>'
    + '<body style="margin:0;padding:0;background:#F9F7F4;font-family:\'Hiragino Kaku Gothic ProN\',\'Noto Sans JP\',sans-serif;">'
    + '<div style="max-width:600px;margin:0 auto;padding:24px;">'

    // ヘッダー
    + '<div style="background:#1a3a5c;padding:32px 24px;text-align:center;border-radius:8px 8px 0 0;">'
    + '<h1 style="color:#fff;font-size:18px;margin:0 0 8px;">お問い合わせを受け付けました</h1>'
    + '<p style="color:#8fb8d9;font-size:13px;margin:0;">電気工事 研修設計コンサルティング</p>'
    + '</div>'

    // 本文
    + '<div style="background:#fff;padding:32px 24px;border-radius:0 0 8px 8px;">'
    + '<p style="font-size:14px;color:#444;line-height:1.8;margin:0 0 20px;">'
    + data.name + ' 様<br><br>'
    + 'このたびはお問い合わせいただき、ありがとうございます。<br>'
    + '内容を確認のうえ、1〜2営業日以内にご返信いたします。<br><br>'
    + 'まだ具体的でなくても構いません。まずはお気軽にお話しください。'
    + '</p>'

    // 受信内容の確認
    + '<div style="background:#F9F7F4;padding:20px;border-radius:8px;margin-bottom:20px;">'
    + '<p style="font-size:12px;color:#888;margin:0 0 8px;">お送りいただいた内容：</p>'
    + '<p style="font-size:13px;color:#444;margin:0;line-height:1.7;white-space:pre-wrap;">' + (data.message || '（記載なし）') + '</p>'
    + '</div>'

    + '<p style="font-size:13px;color:#888;line-height:1.7;margin:0;">'
    + 'ご不明な点がございましたら、このメールへの返信またはメール（info@ohki-electric.com）にてお気軽にご連絡ください。'
    + '</p>'
    + '</div>'

    // フッター
    + '<div style="text-align:center;padding:24px;color:#999;font-size:11px;">'
    + '<p>電気工事 研修設計コンサルティング</p>'
    + '<p>大木 健司</p>'
    + '<p>✉ info@ohki-electric.com</p>'
    + '</div>'

    + '</div>'
    + '</body></html>';

  var options = {
    name: CONTACT_REPLY_FROM_NAME,
    htmlBody: html,
    replyTo: CONTACT_REPLY_FROM_EMAIL
  };

  GmailApp.sendEmail(data.email, subject, '', options);
}

// ===== 自分への通知メール =====
function sendContactNotificationEmail(data) {
  var subject = '【新規お問い合わせ】' + data.name;

  var body = '新しいお問い合わせが届きました。\n\n'
    + 'お名前（会社名）: ' + data.name + '\n'
    + 'メールアドレス: ' + data.email + '\n\n'
    + 'ご相談内容:\n' + (data.message || '（記載なし）') + '\n\n'
    + 'スプレッドシートにも記録済みです。\n'
    + '1〜2営業日以内に返信してください。';

  GmailApp.sendEmail(NOTIFICATION_EMAIL, subject, body);
}
