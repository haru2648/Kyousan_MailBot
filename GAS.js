// ============================================================
// Google Apps Script - Gmail → Discord 通知Bot
// スクリプトプロパティ（歯車アイコン > スクリプトプロパティ）:
//   CLOUDFLARE_URL    : WorkerのエンドポイントURL
//   SPREADSHEET_ID    : スプレッドシートID
//   GAS_SECRET        : WorkerのGAS_SECRETと同じ文字列
//   ERROR_WEBHOOK_URL : エラー通知先のDiscord Webhook URL（任意）
//
// スプレッドシート「設定」シート:
//   B1 : 最終メッセージID（自動更新）
//   B2 : 自分のメールアドレス（手動で設定）
//   B3 : デフォルトメンション先（キーワード未マッチ時。例: <@123456> or @here）
// ============================================================

// -------------------------------------------------------
// HMAC-SHA256 署名を生成してHex文字列で返す
// ※ タイムスタンプ + ペイロードを連結して署名する（リプレイ攻撃対策）
// -------------------------------------------------------
function computeHmacHex(payload, secret) {
  const rawBytes = Utilities.computeHmacSha256Signature(
    Utilities.newBlob(payload, 'text/plain', 'UTF-8').getBytes(),
    Utilities.newBlob(secret,  'text/plain', 'UTF-8').getBytes()
  );
  return rawBytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// -------------------------------------------------------
// WorkerへのリクエストをHMAC署名付きで送信する汎用関数
// action: 'notify'（メール通知）or 'summary'（日次サマリー）
// payload: action='notify' のときのみ必要
// -------------------------------------------------------
function sendToWorker(action, payload) {
  const props      = PropertiesService.getScriptProperties();
  const url        = props.getProperty('CLOUDFLARE_URL');
  const secret     = props.getProperty('GAS_SECRET');
  const bodyStr    = payload ? JSON.stringify(payload) : '{}';
  const timestamp  = String(Date.now());
  const signTarget = timestamp + bodyStr;
  const signature  = computeHmacHex(signTarget, secret);

  const response = UrlFetchApp.fetch(url, {
    method:             'post',
    contentType:        'application/json',
    headers: {
      'X-GAS-Signature': signature,
      'X-GAS-Timestamp': timestamp,
      'X-GAS-Action':    action
    },
    payload:            bodyStr,
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Worker returned HTTP ${code}: ${response.getContentText()}`);
  }
}

// -------------------------------------------------------
// エラーをDiscord Webhookへ通知する（失敗してもログのみ）
// ERROR_WEBHOOK_URL が未設定の場合は何もしない
// -------------------------------------------------------
function notifyError(subject, errorMessage) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('ERROR_WEBHOOK_URL');
  if (!webhookUrl) return;

  try {
    UrlFetchApp.fetch(webhookUrl, {
      method:             'post',
      contentType:        'application/json',
      payload:            JSON.stringify({
        content: `⚠️ **Discord通知失敗**\n件名: ${subject}\nエラー: ${errorMessage}`
      }),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('エラー通知のWebhook送信自体に失敗:', e.message);
  }
}

// -------------------------------------------------------
// メインのトリガー関数（時間ベーストリガーで呼び出す）
// -------------------------------------------------------
function checkMailsAndNotify() {
  const props          = PropertiesService.getScriptProperties();
  const CLOUDFLARE_URL = props.getProperty('CLOUDFLARE_URL');
  const SS_ID          = props.getProperty('SPREADSHEET_ID');
  const GAS_SECRET     = props.getProperty('GAS_SECRET');

  if (!CLOUDFLARE_URL || !SS_ID || !GAS_SECRET) {
    throw new Error('設定エラー: スクリプトプロパティを確認してください (CLOUDFLARE_URL / SPREADSHEET_ID / GAS_SECRET)');
  }

  const ss           = SpreadsheetApp.openById(SS_ID);
  const mappingSheet = ss.getSheetByName('管理シート');
  const configSheet  = ss.getSheetByName('設定');
  if (!mappingSheet || !configSheet) {
    throw new Error('シート名エラー: 「管理シート」と「設定」シートが存在するか確認してください');
  }

  const lastMessageId  = String(configSheet.getRange('B1').getValue()).trim() || '0';
  const myEmail        = String(configSheet.getRange('B2').getValue()).trim().toLowerCase();
  const defaultMention = String(configSheet.getRange('B3').getValue()).trim() || '@here';

  if (!myEmail) {
    throw new Error('設定エラー: 設定シートのB2に自分のメールアドレスを入力してください');
  }

  const mappingData = mappingSheet.getDataRange().getValues();
  const rules = [];
  for (let i = 1; i < mappingData.length; i++) {
    const keyword = String(mappingData[i][0] ?? '').trim();
    const rawId   = String(mappingData[i][2] ?? '').trim();
    if (!keyword || !rawId) continue;
    const formattedId = rawId.startsWith('<') ? rawId : `<@${rawId.replace(/[^0-9]/g, '')}>`;
    rules.push({ keyword, discordId: formattedId });
  }

  const threads = GmailApp.getInboxThreads(0, 50);
  if (threads.length === 0) {
    console.log('受信トレイにメールがありません。');
    return;
  }

  if (!lastMessageId || lastMessageId === '0') {
    const latestId = threads[0].getMessages().slice(-1)[0].getId();
    configSheet.getRange('B1').setValue(latestId);
    console.log('初回実行: 最新メッセージIDを記録しました。');
    return;
  }

  const newMessages = [];
  let latestId = lastMessageId;

  for (const thread of threads) {
    const messages  = thread.getMessages();
    const latestMsg = messages[messages.length - 1];
    const msgId     = latestMsg.getId();

    if (msgId > lastMessageId) {
      if (latestMsg.getFrom().toLowerCase().includes(myEmail)) {
        console.log('送信元が自分自身のためスキップ: ' + latestMsg.getFrom());
      } else {
        newMessages.push(latestMsg);
      }
      if (msgId > latestId) latestId = msgId;
    }
    // break しない（取りこぼし対策）
  }

  if (newMessages.length === 0) {
    console.log('通知すべき新しい外部メールはありません。');
    configSheet.getRange('B1').setValue(latestId);
    return;
  }

  newMessages.reverse();

  let successCount = 0;
  const errors     = [];

  for (const msg of newMessages) {
    const subject     = msg.getSubject() || '(件名なし)';
    const fromAddress = msg.getFrom();
    const plainBody   = msg.getPlainBody() ?? '';

    const bodySnippet = plainBody
      .split('\n')
      .map(l => l.trim())
      .filter(l => l.length > 0)
      .slice(0, 5)
      .join('\n');

    const searchText = (subject + ' ' + fromAddress + ' ' + plainBody).toLowerCase();
    const mentionSet = new Set();
    for (const rule of rules) {
      if (searchText.includes(rule.keyword.toLowerCase())) mentionSet.add(rule.discordId);
    }
    const mentionText = mentionSet.size > 0 ? [...mentionSet].join(' ') : defaultMention;

    const payload = {
      content: `${mentionText} 新着メールの通知です`,
      embeds: [{
        title:       subject,
        description: `**送信者:** ${fromAddress}\n\n**文頭:**\n\`\`\`\n${bodySnippet}\n\`\`\``,
        color:       16711680
      }],
      components: [{
        type: 1,
        components: [
          { type: 2, style: 1, label: '対応開始', custom_id: 'btn_start'    },
          { type: 2, style: 3, label: '対応完了', custom_id: 'btn_complete' }
        ]
      }]
    };

    try {
      sendToWorker('notify', payload);
      successCount++;
    } catch (e) {
      console.error(`Discord送信失敗（件名: ${subject}）: ${e.message}`);
      errors.push({ subject, error: e.message });
      notifyError(subject, e.message);
    }
  }

  configSheet.getRange('B1').setValue(latestId);

  if (errors.length > 0) {
    const errSummary = errors.map(e => `  - 「${e.subject}」: ${e.error}`).join('\n');
    console.warn(`${errors.length} 件の送信に失敗しました:\n${errSummary}`);
  }
  console.log(`完了: ${successCount} 件送信成功, ${errors.length} 件失敗`);
}

// -------------------------------------------------------
// 日次サマリー通知（毎朝トリガーで呼び出す）
// Workerがフォーラムチャンネルのスレッドを集計し、
// BotとしてリマインドチャンネルへDirect投稿する
//
// 【トリガー設定手順】
//   GASエディタ > トリガー（時計アイコン）> トリガーを追加
//   実行する関数 : sendDailySummary
//   イベントのソース : 時間主導型
//   時間の種類 : 日付ベースのタイマー
//   時刻 : 毎朝9時など希望の時間を設定
// -------------------------------------------------------
function sendDailySummary() {
  try {
    sendToWorker('summary', null);
    console.log('日次サマリーリクエストを送信しました。');
  } catch (e) {
    console.error('日次サマリーの送信に失敗:', e.message);
  }
}

// -------------------------------------------------------
// テスト用関数（動作確認後に削除してOK）
// -------------------------------------------------------
function testSendToDiscord() {
  const payload = {
    content: 'テスト通知',
    embeds: [{
      title:       'テスト件名',
      description: '本文テスト',
      color:       16711680
    }],
    components: [{
      type: 1,
      components: [
        { type: 2, style: 1, label: '対応開始', custom_id: 'btn_start'    },
        { type: 2, style: 3, label: '対応完了', custom_id: 'btn_complete' }
      ]
    }]
  };

  try {
    sendToWorker('notify', payload);
    console.log('成功');
  } catch(e) {
    console.log('失敗:', e.message);
  }
}

function testDailySummary() {
  sendDailySummary();
}

function debugSummary() {
  const props = PropertiesService.getScriptProperties();
  const url    = props.getProperty('CLOUDFLARE_URL');
  const secret = props.getProperty('GAS_SECRET');
  const bodyStr   = '{}';
  const timestamp = String(Date.now());
  const signTarget = timestamp + bodyStr;

  const rawBytes = Utilities.computeHmacSha256Signature(
    Utilities.newBlob(signTarget, 'text/plain', 'UTF-8').getBytes(),
    Utilities.newBlob(secret,     'text/plain', 'UTF-8').getBytes()
  );
  const signature = rawBytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');

  const response = UrlFetchApp.fetch(url, {
    method:             'post',
    contentType:        'application/json',
    headers: {
      'X-GAS-Signature': signature,
      'X-GAS-Timestamp': timestamp,
      'X-GAS-Action':    'summary'
    },
    payload:            bodyStr,
    muteHttpExceptions: true
  });

  console.log('HTTP:', response.getResponseCode());
  console.log('Body:', response.getContentText());
}