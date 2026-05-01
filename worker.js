// ============================================================
// Cloudflare Worker - Discord メール通知Bot
// 環境変数 (Dashboard > Settings > Variables and Secrets):
//   PUBLIC_KEY         : Discord Bot の公開鍵
//   BOT_TOKEN          : Discord Bot トークン
//   CHANNEL_ID         : メール通知用フォーラムチャンネルID
//   GUILD_ID           : DiscordサーバーID
//   TAG_UNRESOLVED     : 【未対応】タグID
//   TAG_PROGRESS       : 【対応中】タグID
//   TAG_COMPLETED      : 【完了】タグID
//   GAS_SECRET         : GASと共有するHMACシークレット
//   SUMMARY_CHANNEL_ID : 日次サマリー投稿先のリマインドチャンネルID
// ============================================================

// -------------------------------------------------------
// 定数時間での文字列比較（タイミング攻撃対策）
// -------------------------------------------------------
function timingSafeEqual(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return diff === 0;
}

// -------------------------------------------------------
// HMAC-SHA256 を Hex 文字列で返す
// -------------------------------------------------------
async function computeHmac(body, secret) {
  const encoder = new TextEncoder();
  const cryptoKey = await crypto.subtle.importKey(
    "raw", encoder.encode(secret), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]
  );
  const sigBuffer = await crypto.subtle.sign("HMAC", cryptoKey, encoder.encode(body));
  return Array.from(new Uint8Array(sigBuffer))
    .map(b => b.toString(16).padStart(2, '0'))
    .join('');
}

// -------------------------------------------------------
// Discord Ed25519 署名の検証
// -------------------------------------------------------
async function verifyDiscordSignature(body, signature, timestamp, publicKeyHex) {
  try {
    const encoder       = new TextEncoder();
    const timestampData = encoder.encode(timestamp);
    const bodyData      = encoder.encode(body);
    const message       = new Uint8Array(timestampData.length + bodyData.length);
    message.set(timestampData);
    message.set(bodyData, timestampData.length);

    const sigBytes    = new Uint8Array(signature.match(/.{1,2}/g).map(b => parseInt(b, 16)));
    const pubKeyBytes = new Uint8Array(publicKeyHex.match(/.{1,2}/g).map(b => parseInt(b, 16)));

    let key;
    try {
      key = await crypto.subtle.importKey("raw", pubKeyBytes, { name: "Ed25519" }, false, ["verify"]);
    } catch {
      key = await crypto.subtle.importKey(
        "raw", pubKeyBytes,
        { name: "NODE-ED25519", namedCurve: "NODE-ED25519", public: true },
        false, ["verify"]
      );
    }

    return await crypto.subtle.verify({ name: key.algorithm.name || "Ed25519" }, key, sigBytes, message);
  } catch {
    return false;
  }
}

// -------------------------------------------------------
// Discord API へのラッパー（エラーログ付き）
// -------------------------------------------------------
async function discordFetch(path, method, token, body) {
  const res = await fetch(`https://discord.com/api/v10${path}`, {
    method,
    headers: {
      'Authorization': `Bot ${token}`,
      'Content-Type':  'application/json'
    },
    body: body ? JSON.stringify(body) : undefined
  });
  if (!res.ok) {
    const text = await res.text();
    console.error(`Discord API error [${method} ${path}]: ${res.status}`, text);
  }
  return res;
}

// -------------------------------------------------------
// GASリクエストの認証（HMAC + タイムスタンプ検証）
// 成功: null / 失敗: Response
// -------------------------------------------------------
async function verifyGasRequest(bodyText, gasSignature, gasTimestamp, secret) {
  if (!gasSignature || !gasTimestamp || !secret) {
    return new Response('Forbidden', { status: 403 });
  }
  const tsMs = Number(gasTimestamp);
  if (!Number.isFinite(tsMs) || Math.abs(Date.now() - tsMs) > 5 * 60 * 1000) {
    console.warn(`GAS timestamp rejected: diff=${Math.abs(Date.now() - tsMs)}ms`);
    return new Response('Forbidden', { status: 403 });
  }
  const computed = await computeHmac(gasTimestamp + bodyText, secret);
  if (!timingSafeEqual(computed, gasSignature)) {
    console.warn('GAS HMAC mismatch');
    return new Response('Forbidden', { status: 403 });
  }
  return null; // 認証OK
}

// -------------------------------------------------------
// ボタン定義
// -------------------------------------------------------
const BTN_START    = { type: 2, style: 1, label: "対応開始",    custom_id: "btn_start"    };
const BTN_COMPLETE = { type: 2, style: 3, label: "対応完了",    custom_id: "btn_complete" };
const BTN_RESET    = { type: 2, style: 2, label: "未対応に戻す", custom_id: "btn_reset"    };

const STATE = {
  btn_start: {
    targetTag:     'TAG_PROGRESS',
    titleStatus:   '【対応中】',
    color:         16776960,
    components:    [{ type: 1, components: [BTN_COMPLETE, BTN_RESET] }],
    shouldArchive: false
  },
  btn_complete: {
    targetTag:     'TAG_COMPLETED',
    titleStatus:   '【完了】',
    color:         5763719,
    components:    [{ type: 1, components: [BTN_RESET] }],
    shouldArchive: true
  },
  btn_reset: {
    targetTag:     'TAG_UNRESOLVED',
    titleStatus:   '【未対応】',
    color:         16711680,
    components:    [{ type: 1, components: [BTN_START, BTN_COMPLETE] }],
    shouldArchive: false
  }
};

// -------------------------------------------------------
// 日次サマリー処理
// サーバー全体のアクティブスレッドとアーカイブスレッドを取得し
// タグ別に集計してリマインドチャンネルへBotとして投稿する
// -------------------------------------------------------
async function handleSummary(env) {
  if (!env.SUMMARY_CHANNEL_ID) {
    console.warn('SUMMARY_CHANNEL_ID が未設定のためサマリーをスキップします。');
    return new Response('OK', { status: 200 });
  }

  // アクティブスレッド: GUILD_ID経由で取得（フォーラムチャンネルIDでは404になるため）
  // アーカイブスレッド: チャンネルID経由で取得
  const [activeRes, archivedRes] = await Promise.all([
    discordFetch(`/guilds/${env.GUILD_ID}/threads/active`,              'GET', env.BOT_TOKEN, null),
    discordFetch(`/channels/${env.CHANNEL_ID}/threads/archived/public`, 'GET', env.BOT_TOKEN, null),
  ]);

  if (!activeRes.ok || !archivedRes.ok) {
    return new Response('Failed to fetch threads', { status: 500 });
  }

  const [activeData, archivedData] = await Promise.all([activeRes.json(), archivedRes.json()]);

  const allThreads = [
    ...(activeData.threads   ?? []),
    ...(archivedData.threads ?? []),
  ];

  // タグIDでスレッドを集計
  let unresolvedCount = 0;
  let progressCount   = 0;

  for (const thread of allThreads) {
    const tags = thread.applied_tags ?? [];
    if (tags.includes(env.TAG_UNRESOLVED)) unresolvedCount++;
    if (tags.includes(env.TAG_PROGRESS))   progressCount++;
  }

  // 日本時間の今日の日付
  const dateLabel = new Intl.DateTimeFormat('ja-JP', {
    timeZone: 'Asia/Tokyo',
    month:    'long',
    day:      'numeric',
    weekday:  'short'
  }).format(new Date());

  const lines = [
    `📊 **${dateLabel}の対応状況サマリー**`,
    `🔴 未対応: **${unresolvedCount}件**`,
    `🟡 対応中: **${progressCount}件**`,
  ];

  if (unresolvedCount === 0 && progressCount === 0) {
    lines.push('✅ 未対応・対応中のメールはありません！');
  } else if (unresolvedCount >= 5) {
    lines.push('⚠️ 未対応が5件以上あります。確認をお願いします。');
  }

  // BotとしてリマインドチャンネルへPOST
  const postRes = await discordFetch(
    `/channels/${env.SUMMARY_CHANNEL_ID}/messages`,
    'POST',
    env.BOT_TOKEN,
    { content: lines.join('\n') }
  );

  return postRes.ok
    ? new Response('OK', { status: 200 })
    : new Response('Failed to post summary', { status: 500 });
}

// -------------------------------------------------------
// メイン
// -------------------------------------------------------
export default {
  async fetch(request, env, ctx) {

    const bodyText         = await request.text();
    const discordSignature = request.headers.get('X-Signature-Ed25519');
    const discordTimestamp = request.headers.get('X-Signature-Timestamp');
    const gasSignature     = request.headers.get('X-GAS-Signature');
    const gasTimestamp     = request.headers.get('X-GAS-Timestamp');
    const gasAction        = request.headers.get('X-GAS-Action');

    // ── GASからのリクエスト（メール通知 / 日次サマリー）──────────
    if (request.method === 'POST' && !discordSignature) {

      const authError = await verifyGasRequest(bodyText, gasSignature, gasTimestamp, env.GAS_SECRET);
      if (authError) return authError;

      // X-GAS-Action: summary の場合はサマリー処理へ
      if (gasAction === 'summary') {
        return handleSummary(env);
      }

      // 通常のメール通知処理
      let data;
      try {
        data = JSON.parse(bodyText);
      } catch {
        return new Response('Bad Request', { status: 400 });
      }

      const rawTitle   = data?.embeds?.[0]?.title ?? '(件名なし)';
      const threadName = rawTitle.replace('【未対応】', '').trim().slice(0, 100) || '(件名なし)';

      const res = await discordFetch(
        `/channels/${env.CHANNEL_ID}/threads`,
        'POST',
        env.BOT_TOKEN,
        {
          name:         threadName,
          applied_tags: [env.TAG_UNRESOLVED],
          message: {
            content:    data.content,
            embeds:     data.embeds,
            components: data.components
          }
        }
      );

      return res.ok
        ? new Response('OK', { status: 200 })
        : new Response('Internal Server Error', { status: 500 });
    }

    // ── Discord Interaction（署名検証あり）────────────────────
    if (!discordSignature || !discordTimestamp) {
      return new Response('Unauthorized', { status: 401 });
    }

    const isValid = await verifyDiscordSignature(
      bodyText, discordSignature, discordTimestamp, env.PUBLIC_KEY
    );
    if (!isValid) return new Response('Unauthorized', { status: 401 });

    let interaction;
    try {
      interaction = JSON.parse(bodyText);
    } catch {
      return new Response('Bad Request', { status: 400 });
    }

    // PING
    if (interaction.type === 1) {
      return Response.json({ type: 1 });
    }

    // ボタン操作
    if (interaction.type === 3) {
      const customId      = interaction.data.custom_id;
      const threadId      = interaction.channel_id;
      const originalEmbed = interaction.message.embeds[0];

      const state = STATE[customId];
      if (!state) return Response.json({ type: 1 });

      ctx.waitUntil(
        new Promise(resolve => setTimeout(resolve, 500)).then(() =>
          discordFetch(
            `/channels/${threadId}`,
            'PATCH',
            env.BOT_TOKEN,
            {
              applied_tags: env[state.targetTag] ? [env[state.targetTag]] : [],
              archived:     state.shouldArchive
            }
          )
        )
      );

      const cleanTitle = (originalEmbed?.title ?? '')
        .replace(/【未対応】|【対応中】|【完了】/g, '')
        .trim();

      return Response.json({
        type: 7,
        data: {
          embeds:     [{ ...originalEmbed, title: `${state.titleStatus} ${cleanTitle}`, color: state.color }],
          components: state.components
        }
      });
    }

    return new Response('Unauthorized', { status: 401 });
  }
};
