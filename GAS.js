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
//
// スプレッドシート「管理シート」:
//   A列 : キーワード
//   C列 : メンション先Discord ID
//   D列 : 投稿先チャンネルID（空欄の場合はデフォルトの CHANNEL_ID に投稿）
//   E列 : デフォルトチャンネルにも重複投稿するか（チェックボックス）
//         空欄/TRUE = D列の専用チャンネルに加えてデフォルトチャンネルにも投稿する（従来の挙動・デフォルト）
//         FALSE     = 専用チャンネルのみに投稿し、デフォルトチャンネルへは投稿しない
//         ※D列が空欄の行では意味を持たない
//
// 【複数チャンネルへの投稿について】
//   D列にチャンネルIDが設定されているキーワードにマッチした場合:
//     - E列が空欄またはTRUEなら、そのチャンネルに加えてデフォルトチャンネルにも投稿します（従来通り）。
//     - E列がFALSEなら、専用チャンネルのみに投稿します。
//   同じメールが複数の専用チャンネルルールにマッチし、E列の値が食い違う場合は、
//   いずれか1つでも「重複する」側があればデフォルトチャンネルにも投稿します（OR判定）。
//   例（E=FALSE）: 「寄附」にマッチ → 寄付チャンネルのみに投稿
//   例（E=空欄） : 「請求書」にマッチ → 請求書チャンネル + デフォルトチャンネルの両方に投稿（従来通り）
//
// スプレッドシート「処理ログ」シート:
//   通知対象になったメールの送信結果のみ記録（デバッグ用）
//   列: 日時 / メッセージID / 件名 / 送信者 / 判定結果
//
// スプレッドシート「ステータス管理」シート:
//   通知した外部メールのスレッドごとに、返信要否を自動判定して記録する
//   列: スレッドID / 件名 / 送信者 / 検知日時 / 返信要否（要返信/返信済み） / 最終チェック日時
//   判定基準: スレッドの最新メッセージの送信者が「設定」シートB2のアドレスなら「返信済み」
//   「返信済み」から ARCHIVE_AFTER_DAYS 日経過した行は自動で「ステータス管理_アーカイブ」へ移動する
//
// スプレッドシート「ステータス管理_アーカイブ」シート:
//   「ステータス管理」から自動アーカイブされた行の保管先（列は同じ＋アーカイブ日時）
// ============================================================

// 「返信済み」になってからこの日数が経過した行を自動でアーカイブする
const ARCHIVE_AFTER_DAYS = 60;

// -------------------------------------------------------
// HMAC-SHA256 署名を生成してHex文字列で返す
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
// エラーをDiscord Webhookへ通知する
// -------------------------------------------------------
function notifyError(subject, errorMessage) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('ERROR_WEBHOOK_URL');
  if (!webhookUrl) return;

  try {
    UrlFetchApp.fetch(webhookUrl, {
      method:             'post',
      contentType:        'application/json',
      payload:            JSON.stringify({
        content: `⚠️ **Discord通知失敗**\n件名: ${subject}\nエラー: ${errorMessage}\n\n※このメールは次回実行時に自動で再試行されます。`
      }),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('エラー通知のWebhook送信自体に失敗:', e.message);
  }
}

// -------------------------------------------------------
// 処理ログを「処理ログ」シートに記録する
// -------------------------------------------------------
function logProcessing(ss, msgId, subject, from, status) {
  try {
    let sheet = ss.getSheetByName('処理ログ');
    if (!sheet) {
      sheet = ss.insertSheet('処理ログ');
      sheet.appendRow(['日時', 'メッセージID', '件名', '送信者', '判定結果']);
    }
    sheet.appendRow([new Date(), msgId, subject, from, status]);
  } catch (e) {
    console.error('処理ログの記録に失敗:', e.message);
  }
}

// -------------------------------------------------------
// 「ステータス管理」シートに返信要否を記録・更新する
// 同じスレッドが既に記録されている場合は「要返信」に戻して検知日時を更新する
// （一度返信済みになったスレッドに新着が来た＝再度返信が必要、とみなす）
// -------------------------------------------------------
function upsertReplyStatus(ss, threadId, subject, fromAddress) {
  try {
    let sheet = ss.getSheetByName('ステータス管理');
    if (!sheet) {
      sheet = ss.insertSheet('ステータス管理');
      sheet.appendRow(['スレッドID', '件名', '送信者', '検知日時', '返信要否', '最終チェック日時']);
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === threadId) {
        sheet.getRange(i + 1, 4).setValue(new Date());
        sheet.getRange(i + 1, 5).setValue('要返信');
        return;
      }
    }
    sheet.appendRow([threadId, subject, fromAddress, new Date(), '要返信', '']);
  } catch (e) {
    console.error('ステータス管理の更新に失敗:', e.message);
  }
}

// -------------------------------------------------------
// 「ステータス管理」シートの「要返信」行を再チェックし、
// スレッドの最新メッセージが自分からの送信なら「返信済み」に更新する
// -------------------------------------------------------
function updateReplyStatuses(ss, myEmail) {
  const sheet = ss.getSheetByName('ステータス管理');
  if (!sheet) return; // まだ1件も記録されていない場合は何もしない

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const threadId = String(data[i][0] ?? '').trim();
    const status   = String(data[i][4] ?? '').trim();
    if (!threadId || status !== '要返信') continue;

    try {
      const thread = GmailApp.getThreadById(threadId);
      if (!thread) continue;

      const messages = thread.getMessages();
      const lastMsg  = messages[messages.length - 1];
      const replied  = lastMsg.getFrom().toLowerCase().includes(myEmail);

      if (replied) {
        sheet.getRange(i + 1, 5).setValue('返信済み');
      }
      sheet.getRange(i + 1, 6).setValue(new Date());
    } catch (e) {
      console.error(`スレッド ${threadId} の返信チェックに失敗: ${e.message}`);
    }
  }
}

// -------------------------------------------------------
// 「ステータス管理」シートの行を「ステータス管理_アーカイブ」シートへ追記する
// -------------------------------------------------------
function appendToArchive(ss, rowValues) {
  let archiveSheet = ss.getSheetByName('ステータス管理_アーカイブ');
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet('ステータス管理_アーカイブ');
    archiveSheet.appendRow(['スレッドID', '件名', '送信者', '検知日時', '返信要否', '最終チェック日時', 'アーカイブ日時']);
  }
  archiveSheet.appendRow([...rowValues, new Date()]);
}

// -------------------------------------------------------
// 「ステータス管理」シートのうち、「返信済み」になってから
// ARCHIVE_AFTER_DAYS 日以上経過した行を「ステータス管理_アーカイブ」へ移動する
// -------------------------------------------------------
function archiveResolvedStatuses(ss) {
  const sheet = ss.getSheetByName('ステータス管理');
  if (!sheet) return;

  try {
    const data = sheet.getDataRange().getValues();
    const now  = new Date();
    const rowsToDelete = [];

    for (let i = 1; i < data.length; i++) {
      const status    = String(data[i][4] ?? '').trim();
      const checkedAt = data[i][5];
      if (status !== '返信済み' || !(checkedAt instanceof Date)) continue;

      const daysSince = (now - checkedAt) / (1000 * 60 * 60 * 24);
      if (daysSince >= ARCHIVE_AFTER_DAYS) {
        appendToArchive(ss, data[i]);
        rowsToDelete.push(i + 1); // シート上の実際の行番号（ヘッダー分+1）
      }
    }

    // 後ろの行から削除することで、削除時の行番号ズレを防ぐ
    for (let j = rowsToDelete.length - 1; j >= 0; j--) {
      sheet.deleteRow(rowsToDelete[j]);
    }
  } catch (e) {
    console.error('ステータス管理のアーカイブに失敗:', e.message);
  }
}

// -------------------------------------------------------
// メール1件分のDiscord通知ペイロードを生成して返す
// -------------------------------------------------------
function buildPayload(targetChannelId, mentionText, subject, fromAddress, bodySnippet, isSpam) {
  const titlePrefix = isSpam ? '⚠️[迷惑メール判定] ' : '';
  return {
    targetChannelId: targetChannelId || null,
    content: `${mentionText} 新着メールの通知です${isSpam ? '（迷惑メールフォルダより）' : ''}`,
    embeds: [{
      title:       `${titlePrefix}${subject}`,
      description: `**送信者:** ${fromAddress}\n\n**文頭:**\n\`\`\`\n${bodySnippet}\n\`\`\``,
      color:       isSpam ? 16744272 : 16711680
    }],
    components: [{
      type: 1,
      components: [
        { type: 2, style: 1, label: '対応開始', custom_id: 'btn_start'    },
        { type: 2, style: 3, label: '対応完了', custom_id: 'btn_complete' }
      ]
    }]
  };
}

// -------------------------------------------------------
// メインのトリガー関数（時間ベーストリガーで呼び出す）
//
// 【B1（最終メッセージID）の更新方針】
//   通知対象メールが1件でも送信失敗した場合、B1は更新しない。
//   全件成功した場合のみ、B1を最新IDまで進める。
//
// 【複数チャンネルへの投稿】
//   D列にチャンネルIDが設定されたキーワードにマッチした場合、
//   そのチャンネルとデフォルトチャンネルの両方に投稿する。
//   キーワード未マッチ or D列が空欄の場合はデフォルトチャンネルのみに投稿。
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

  // 新着メールの有無に関わらず、毎回「要返信」スレッドの再チェックとアーカイブ整理を行う
  updateReplyStatuses(ss, myEmail);
  archiveResolvedStatuses(ss);

  const mappingData = mappingSheet.getDataRange().getValues();
  const rules = [];
  for (let i = 1; i < mappingData.length; i++) {
    const keyword   = String(mappingData[i][0] ?? '').trim();
    const rawId     = String(mappingData[i][2] ?? '').trim();
    const channelId = String(mappingData[i][3] ?? '').trim();
    if (!keyword || !rawId) continue;
    const formattedId = rawId.startsWith('<') ? rawId : `<@${rawId.replace(/[^0-9]/g, '')}>`;

    // E列: デフォルトチャンネルへの重複投稿可否。空欄・未設定は true（従来通り重複する）扱いとし、
    // 明示的にチェックを外した（false）場合のみ重複しない。手入力文字列 'FALSE' にも寛容に対応する。
    const dupRaw = mappingData[i][4];
    const duplicateToDefault = !(dupRaw === false || String(dupRaw ?? '').trim().toUpperCase() === 'FALSE');

    rules.push({ keyword, discordId: formattedId, channelId, duplicateToDefault });
  }

  // 受信トレイと迷惑メールフォルダを取得してマージ
  const inboxThreads = GmailApp.getInboxThreads(0, 50);
  const spamThreads  = GmailApp.search('in:spam', 0, 50);

  const seenThreadIds = new Set();
  const allThreads    = [];
  for (const thread of [...inboxThreads, ...spamThreads]) {
    const tid = thread.getId();
    if (!seenThreadIds.has(tid)) {
      seenThreadIds.add(tid);
      allThreads.push({ thread, isSpam: !inboxThreads.includes(thread) });
    }
  }

  if (allThreads.length === 0) {
    console.log('受信トレイ・迷惑メールフォルダにメールがありません。');
    return;
  }

  if (!lastMessageId || lastMessageId === '0') {
    const latestId = inboxThreads.length > 0
      ? inboxThreads[0].getMessages().slice(-1)[0].getId()
      : allThreads[0].thread.getMessages().slice(-1)[0].getId();
    configSheet.getRange('B1').setValue(latestId);
    console.log('初回実行: 最新メッセージIDを記録しました。');
    return;
  }

  const newMessages = [];
  let scanLatestId  = lastMessageId;

  for (const { thread, isSpam } of allThreads) {
    const messages  = thread.getMessages();
    const latestMsg = messages[messages.length - 1];
    const msgId     = latestMsg.getId();

    if (msgId > lastMessageId) {
      if (!latestMsg.getFrom().toLowerCase().includes(myEmail)) {
        newMessages.push({ msg: latestMsg, isSpam, threadId: thread.getId() });
      }
      if (msgId > scanLatestId) scanLatestId = msgId;
    }
  }

  if (newMessages.length === 0) {
    console.log('通知すべき新しい外部メールはありません。');
    configSheet.getRange('B1').setValue(scanLatestId);
    return;
  }

  newMessages.sort((a, b) => a.msg.getId() < b.msg.getId() ? -1 : 1);

  let allSucceeded = true;
  let totalSuccess = 0;
  let totalFail    = 0;

  for (const { msg, isSpam, threadId } of newMessages) {
    const subject     = msg.getSubject() || '(件名なし)';
    const fromAddress = msg.getFrom();
    const plainBody   = msg.getPlainBody() ?? '';

    const bodySnippet = plainBody
      .split('\n')
      .map(l => l.trim())
      .filter(l => l.length > 0)
      .slice(0, 10)          // 行数を10に変更
      .join('\n')
      .slice(0, 500);        // 500文字で切り捨て（安全策）

    const searchText = (subject + ' ' + fromAddress + ' ' + plainBody).toLowerCase();

    // キーワードマッチング
    // mentionsByChannel: { チャンネルID → Set<メンションID> }
    // チャンネルIDが空 = デフォルトチャンネルとして '' をキーに使う
    const mentionsByChannel = new Map();
    let duplicateToDefault = false; // このメールで専用チャンネルにマッチしたルールのいずれかがtrueならtrue（OR判定）

    for (const rule of rules) {
      if (searchText.includes(rule.keyword.toLowerCase())) {
        // D列にチャンネルIDがある場合 → そのチャンネルに追加
        if (rule.channelId) {
          if (!mentionsByChannel.has(rule.channelId)) {
            mentionsByChannel.set(rule.channelId, new Set());
          }
          mentionsByChannel.get(rule.channelId).add(rule.discordId);
          if (rule.duplicateToDefault) duplicateToDefault = true;
        }
        // D列にチャンネルIDがない場合 → デフォルトチャンネルに追加
        if (!rule.channelId) {
          if (!mentionsByChannel.has('')) {
            mentionsByChannel.set('', new Set());
          }
          mentionsByChannel.get('').add(rule.discordId);
        }
      }
    }

    // 投稿先チャンネルの決定
    // - 専用チャンネル（D列あり）にマッチした場合: 専用チャンネル + デフォルトチャンネルの両方
    // - デフォルトチャンネル（D列なし）のみにマッチした場合: デフォルトチャンネルのみ
    // - どこにもマッチしなかった場合: デフォルトチャンネルのみ（デフォルトメンション使用）
    const postTargets = []; // [{ channelId: null or string, mentionText: string }]

    // 専用チャンネルへの投稿
    for (const [channelId, mentions] of mentionsByChannel.entries()) {
      if (channelId === '') continue; // デフォルトチャンネルは後で処理
      postTargets.push({
        channelId,
        mentionText: [...mentions].join(' ')
      });
    }

    // デフォルトチャンネルへの投稿
    // - D列なしルールがマッチした場合は必ずそのメンションで投稿（変更なし）
    // - 専用チャンネルのみにマッチした場合は、E列(duplicateToDefault)が空欄/TRUEのときだけデフォルトにも投稿
    // - どこにもマッチしなかった場合は、必ずデフォルトメンションでデフォルトチャンネルに投稿（変更なし）
    const defaultMentions = mentionsByChannel.get('');         // D列なしルールのメンション
    const hasSpecialChannel = postTargets.length > 0;          // 専用チャンネルへの投稿があるか

    if (defaultMentions && defaultMentions.size > 0) {
      // D列なしのキーワードにマッチ → そのメンションでデフォルトチャンネルに投稿
      postTargets.push({ channelId: null, mentionText: [...defaultMentions].join(' ') });
    } else if (hasSpecialChannel) {
      if (duplicateToDefault) {
        // 専用チャンネルのみにマッチし、E列が空欄/TRUE → デフォルトメンションでデフォルトチャンネルにも投稿
        postTargets.push({ channelId: null, mentionText: defaultMention });
      }
      // E列がFALSEの場合はデフォルトチャンネルへの投稿をスキップ（専用チャンネルのみに投稿）
    } else {
      // どこにもマッチしなかった → デフォルトメンションでデフォルトチャンネルのみ
      postTargets.push({ channelId: null, mentionText: defaultMention });
    }

    // 各投稿先に送信
    let msgSucceeded = true;
    const channelLabels = [];

    for (const { channelId, mentionText } of postTargets) {
      const payload = buildPayload(channelId, mentionText, subject, fromAddress, bodySnippet, isSpam);
      const label   = channelId ? `専用チャンネル(${channelId})` : 'デフォルトチャンネル';

      try {
        sendToWorker('notify', payload);
        totalSuccess++;
        channelLabels.push(`${label}:成功`);
      } catch (e) {
        console.error(`Discord送信失敗（件名: ${subject} / ${label}）: ${e.message}`);
        notifyError(`${subject}（${label}）`, e.message);
        channelLabels.push(`${label}:失敗`);
        msgSucceeded  = false;
        allSucceeded  = false;
        totalFail++;
      }
    }

    const spamLabel = isSpam ? '・迷惑メール' : '';
    logProcessing(ss, msg.getId(), subject, fromAddress,
      `${msgSucceeded ? '送信成功' : '一部失敗'}（${channelLabels.join(' / ')}${spamLabel}）`
    );
    upsertReplyStatus(ss, threadId, subject, fromAddress);
  }

  if (allSucceeded) {
    configSheet.getRange('B1').setValue(scanLatestId);
  } else {
    console.warn('一部のメール送信に失敗したため、B1は更新しません。次回実行時に再試行します。');
  }

  console.log(`完了: ${totalSuccess} 件送信成功, ${totalFail} 件失敗`);
}

// -------------------------------------------------------
// 日次サマリー通知（毎朝トリガーで呼び出す）
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
  const payload = buildPayload(null, '@here', 'テスト件名', 'test@example.com', 'テスト本文', false);
  try {
    sendToWorker('notify', payload);
    console.log('成功');
  } catch(e) {
    console.log('失敗:', e.message);
  }
}

function testDonationChannel() {
  // 寄付チャンネルへの投稿テスト
  const payloadDonation = buildPayload('1526086972282769551', '<@&寄付金ロールID>', 'テスト件名【寄付のお願い】', 'test@example.com', 'テスト本文', false);
  // デフォルトチャンネルへの投稿テスト
  const payloadDefault  = buildPayload(null, '@here', 'テスト件名【寄付のお願い】', 'test@example.com', 'テスト本文', false);

  try {
    sendToWorker('notify', payloadDonation);
    console.log('寄付チャンネル: 成功');
    sendToWorker('notify', payloadDefault);
    console.log('デフォルトチャンネル: 成功');
  } catch(e) {
    console.log('失敗:', e.message);
  }
}

function testDailySummary() {
  sendDailySummary();
}

function debugChannelId() {
  const SS_ID       = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss          = SpreadsheetApp.openById(SS_ID);
  const mappingSheet = ss.getSheetByName('管理シート');
  const mappingData  = mappingSheet.getDataRange().getValues();

  for (let i = 1; i < mappingData.length; i++) {
    const keyword   = String(mappingData[i][0] ?? '').trim();
    const channelId = String(mappingData[i][3] ?? '').trim();
    if (keyword) {
      console.log(`行${i+1} キーワード:「${keyword}」 D列チャンネルID:「${channelId}」 文字数:${channelId.length}`);
    }
  }
}