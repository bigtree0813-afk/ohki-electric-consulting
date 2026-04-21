/**
 * 教育体制診断 - Google Apps Script
 *
 * 【セットアップ手順】
 * 1. Google スプレッドシートを新規作成
 * 2. 「拡張機能」→「Apps Script」を開く
 * 3. このコードを貼り付け
 * 4. REPLY_FROM_EMAIL を自分のメールアドレスに変更
 * 5. 「デプロイ」→「新しいデプロイ」→ 種類: ウェブアプリ
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 6. デプロイURLをコピーして checklist.html の GAS_URL に設定
 */

// ===== 設定 =====
var REPLY_FROM_EMAIL = 'info@ohki-electric.com'; // 自動返信の送信元（Gmailのエイリアスに登録済みの場合のみ変更可）
var REPLY_FROM_NAME = '電気工事 研修設計コンサルティング';
var SHEET_NAME = '診断結果';
var CHECKLIST_NOTIFICATION_EMAIL = 'kenji@ohki-electric.com';
var CHECKLIST_SPREADSHEET_ID = '1QYcByw6S1h28ixODGvMV13ZFhsd8f7xLkg65lV9JYoA';

// ===== Webアプリのエントリーポイント =====
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // typeフィールドで診断とお問い合わせを振り分け
    if (data.type === 'contact') {
      // お問い合わせフォーム
      handleContact(data);
    } else {
      // 診断チェックリスト（既存処理）
      saveToSheet(data);
      sendReplyEmail(data);
      sendNotificationEmail(data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// CORS対応（preflight request）
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ result: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== スプレッドシートへの記録 =====
function saveToSheet(data) {
  var ss = SpreadsheetApp.openById(CHECKLIST_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      '日時', '会社名', 'お名前', 'メールアドレス',
      '総合スコア', '判定',
      'A.育成計画', 'B.指導体制', 'C.安全・資格', 'D.技術継承',
      'チェック項目詳細'
    ]);
    // ヘッダー行の書式設定
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#fff');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date(),
    data.company,
    data.name,
    data.email,
    data.score,
    data.level,
    data.scoreA,
    data.scoreB,
    data.scoreC,
    data.scoreD,
    data.checkedItems ? data.checkedItems.join(' / ') : ''
  ]);
}

// ===== 診断者への自動返信メール =====
function sendReplyEmail(data) {
  var subject = '【診断結果】御社の新人育成体制スコア：' + data.score + '/20（' + data.level + '）';

  var body = buildReplyEmailBody(data);

  var options = {
    name: REPLY_FROM_NAME,
    htmlBody: body,
    replyTo: REPLY_FROM_EMAIL
  };

  GmailApp.sendEmail(data.email, subject, '', options);
}

// ===== 自分への通知メール =====
function sendNotificationEmail(data) {
  var subject = '【新規診断】' + data.company + ' ' + data.name + '様（' + data.score + '/20）';

  var body = '新しい診断結果が届きました。\n\n'
    + '会社名: ' + data.company + '\n'
    + 'お名前: ' + data.name + '\n'
    + 'メール: ' + data.email + '\n'
    + '総合スコア: ' + data.score + '/20（' + data.level + '）\n'
    + 'A.育成計画: ' + data.scoreA + '/5\n'
    + 'B.指導体制: ' + data.scoreB + '/5\n'
    + 'C.安全・資格: ' + data.scoreC + '/5\n'
    + 'D.技術継承: ' + data.scoreD + '/5\n\n'
    + 'チェック項目:\n' + (data.checkedItems ? data.checkedItems.join('\n') : 'なし') + '\n\n'
    + 'スプレッドシートにも記録済みです。';

  GmailApp.sendEmail(CHECKLIST_NOTIFICATION_EMAIL, subject, body);
}

// ===== 自動返信メールのHTML本文 =====
function buildReplyEmailBody(data) {
  // カテゴリ別のアドバイス
  var adviceA = getAdvice('A', data.scoreA);
  var adviceB = getAdvice('B', data.scoreB);
  var adviceC = getAdvice('C', data.scoreC);
  var adviceD = getAdvice('D', data.scoreD);

  // 総合判定の詳細メッセージ
  var overallAdvice = getOverallAdvice(data.score, data.level);

  var html = ''
    + '<!DOCTYPE html>'
    + '<html><head><meta charset="utf-8"></head>'
    + '<body style="margin:0;padding:0;background:#F9F7F4;font-family:\'Hiragino Kaku Gothic ProN\',\'Noto Sans JP\',sans-serif;">'
    + '<div style="max-width:600px;margin:0 auto;padding:24px;">'

    // ヘッダー
    + '<div style="background:#1a3a5c;padding:32px 24px;text-align:center;border-radius:8px 8px 0 0;">'
    + '<h1 style="color:#fff;font-size:18px;margin:0 0 8px;">新人育成体制 診断結果</h1>'
    + '<p style="color:#8fb8d9;font-size:13px;margin:0;">電気工事 研修設計コンサルティング</p>'
    + '</div>'

    // スコア表示
    + '<div style="background:#fff;padding:32px 24px;text-align:center;border-bottom:1px solid #e8e8e8;">'
    + '<p style="color:#888;font-size:12px;margin:0 0 8px;">' + data.company + ' 様の診断結果</p>'
    + '<div style="font-size:48px;font-weight:800;color:#D4A574;line-height:1;">' + data.score + '<span style="font-size:18px;color:#888;font-weight:500;"> / 20</span></div>'
    + '<div style="display:inline-block;margin-top:12px;padding:6px 20px;border-radius:20px;font-size:14px;font-weight:700;color:#fff;background:' + getLevelColor(data.level) + ';">' + data.level + '</div>'
    + '</div>'

    // 総合所見
    + '<div style="background:#fff;padding:24px;border-bottom:1px solid #e8e8e8;">'
    + '<h2 style="font-size:15px;color:#1a3a5c;margin:0 0 12px;border-left:3px solid #D4A574;padding-left:12px;">総合所見</h2>'
    + '<p style="font-size:14px;color:#444;line-height:1.8;margin:0;">' + overallAdvice + '</p>'
    + '</div>'

    // カテゴリ別スコア
    + '<div style="background:#fff;padding:24px;border-bottom:1px solid #e8e8e8;">'
    + '<h2 style="font-size:15px;color:#1a3a5c;margin:0 0 16px;border-left:3px solid #D4A574;padding-left:12px;">カテゴリ別スコアと改善ヒント</h2>'

    + buildCategorySection('A. 育成計画・カリキュラム', data.scoreA, 5, '#2563a0', adviceA)
    + buildCategorySection('B. 指導体制', data.scoreB, 5, '#D4A574', adviceB)
    + buildCategorySection('C. 安全教育・資格取得支援', data.scoreC, 5, '#2e8b57', adviceC)
    + buildCategorySection('D. 技術継承・育成文化', data.scoreD, 5, '#7b5ea7', adviceD)

    + '</div>'

    // CTA
    + '<div style="background:#1a3a5c;padding:32px 24px;text-align:center;border-radius:0 0 8px 8px;">'
    + '<h3 style="color:#fff;font-size:16px;margin:0 0 12px;">具体的な改善策を知りたい方へ</h3>'
    + '<p style="color:#8fb8d9;font-size:13px;margin:0 0 20px;line-height:1.7;">御社の診断結果をもとに、専門家が具体的な改善策をご提案します。<br>初回のヒアリング（約30分・オンライン）は無料です。</p>'
    + '<a href="mailto:info@ohki-electric.com?subject=' + encodeURIComponent('無料ヒアリング希望（診断スコア：' + data.score + '/20）') + '" style="display:inline-block;background:#D4A574;color:#0f2b45;font-weight:700;font-size:14px;padding:14px 32px;border-radius:8px;text-decoration:none;">無料ヒアリングに申し込む</a>'
    + '</div>'

    // フッター
    + '<div style="text-align:center;padding:24px;color:#999;font-size:11px;">'
    + '<p>電気工事 研修設計コンサルティング</p>'
    + '<p>✉ info@ohki-electric.com</p>'
    + '<p style="margin-top:8px;">このメールは診断結果の送信にのみ使用しています。</p>'
    + '</div>'

    + '</div>'
    + '</body></html>';

  return html;
}

// カテゴリセクションのHTML生成
function buildCategorySection(title, score, max, color, advice) {
  var barWidth = (score / max * 100);
  return ''
    + '<div style="margin-bottom:20px;">'
    + '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">'
    + '<span style="font-size:13px;font-weight:700;color:#1a3a5c;">' + title + '</span>'
    + '<span style="font-size:13px;font-weight:700;color:' + color + ';">' + score + ' / ' + max + '</span>'
    + '</div>'
    + '<div style="background:#eee;border-radius:4px;height:8px;overflow:hidden;">'
    + '<div style="background:' + color + ';height:100%;width:' + barWidth + '%;border-radius:4px;"></div>'
    + '</div>'
    + '<p style="font-size:13px;color:#666;margin:8px 0 0;line-height:1.7;">' + advice + '</p>'
    + '</div>';
}

// 判定レベルに応じた色
function getLevelColor(level) {
  switch (level) {
    case '充実': return '#2e8b57';
    case '標準的': return '#e6a817';
    case '要改善': return '#d4762c';
    case '要対策': return '#c0392b';
    default: return '#888';
  }
}

// カテゴリ別アドバイス
function getAdvice(cat, score) {
  if (cat === 'A') {
    if (score >= 4) return '育成計画が体系的に整備されています。定期的な見直しで精度をさらに高めましょう。';
    if (score >= 2) return '計画の骨格はありますが、到達目標や学習順序の明確化に改善の余地があります。まずは「入社から独り立ちまでのロードマップ」を1枚にまとめることから始めてみてください。';
    return '育成が計画なしで進んでいる状態です。教える内容と順番を文書化するだけで、指導の質と効率が大きく変わります。';
  }
  if (cat === 'B') {
    if (score >= 4) return '指導体制がしっかり構築されています。指導者同士の情報共有を強化すると、さらに効果的です。';
    if (score >= 2) return '指導担当者はいるものの、教え方が属人化している可能性があります。「教え方マニュアル」の整備と、指導者への研修がカギです。';
    return '指導が個人の善意や経験に依存しています。まずはメンター制度の導入と、最低限の指導ガイドラインの作成を優先しましょう。';
  }
  if (cat === 'C') {
    if (score >= 4) return '安全教育と資格支援が充実しています。ヒヤリハットの教材化で、さらに実践的な教育が可能です。';
    if (score >= 2) return '基本的な安全教育は実施していますが、資格取得支援や事故事例の教育活用に伸びしろがあります。資格取得の学習スケジュールを会社として用意すると合格率が上がります。';
    return '安全教育の体系化が急務です。入社時の安全教育プログラムの整備と、資格取得支援制度の導入を検討してください。';
  }
  if (cat === 'D') {
    if (score >= 4) return '技術継承と育成文化が根付いています。この強みを採用ブランディングにも活かしましょう。';
    if (score >= 2) return 'ベテランの技術を次世代に伝える仕組みが不十分です。まずは「よくある失敗」「勘所」を言語化する取り組みから始めてみてください。';
    return '技術継承が危機的な状況です。ベテランの退職前に、核となる技術やノウハウを記録・マニュアル化する取り組みが急がれます。';
  }
  return '';
}

// 総合アドバイス
function getOverallAdvice(score, level) {
  if (score >= 16) {
    return '業界でもトップクラスの育成体制です。今の仕組みを維持しつつ、定期的な見直しと改善を続けることで、採用力の強化や業界内でのブランディングにもつながります。';
  }
  if (score >= 11) {
    return '基本的な基盤はできていますが、いくつかの領域に改善の余地があります。特にスコアが低いカテゴリから優先的に取り組むことで、育成の効率と質が大きく向上します。部分的な改善（カリキュラムの体系化、指導マニュアルの整備など）だけでも効果が見込めます。';
  }
  if (score >= 6) {
    return '育成が属人的になっている可能性が高い状態です。「誰が教えるか」によって新人の成長にばらつきが出ていませんか？まずは育成カリキュラムの文書化と、指導者向けマニュアルの整備から始めることをおすすめします。仕組み化することで、指導者の負担も軽減できます。';
  }
  return '体系的な育成の仕組みがほとんど機能していない状態です。新人の早期離職や、ベテランの退職による技術の途絶リスクが高まっています。育成体制の構築は、一度つくれば毎年使える「資産」になります。早めの対策が将来の大きなコスト削減につながります。';
}
