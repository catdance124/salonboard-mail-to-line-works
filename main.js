// ============================================================
// 設定（スクリプトプロパティから取得）
// ============================================================
const CONFIG = (() => {
  const p = PropertiesService.getScriptProperties();
  return {
    // Gmail フィルター条件（公開情報はそのまま）
    SENDER_EMAIL:     "yoyaku_system@salonboard.com",
    PROCESSED_LABEL:  "LW転送済み",

    // LINE WORKS OAuth 2.0 認証情報（スクリプトプロパティから取得）
    CLIENT_ID:        p.getProperty("CLIENT_ID"),
    CLIENT_SECRET:    p.getProperty("CLIENT_SECRET"),
    SERVICE_ACCOUNT:  p.getProperty("SERVICE_ACCOUNT"),
    PRIVATE_KEY:      (() => {
      const raw = p.getProperty("PRIVATE_KEY");
      if (raw.includes('\n')) return raw;
      if (raw.includes('\\n')) return raw.replace(/\\n/g, '\n');
      const b64 = raw.replace(/-----[A-Z ]+-----/g, '').replace(/\s+/g, '');
      return `-----BEGIN PRIVATE KEY-----\n${b64.match(/.{1,64}/g).join('\n')}\n-----END PRIVATE KEY-----`;
    })(),

    // Bot & 送信先（スクリプトプロパティから取得）
    BOT_ID:           p.getProperty("BOT_ID"),
    CHANNEL_ID:       p.getProperty("CHANNEL_ID"),
  };
})();

// ============================================================
// メイン処理
// ============================================================
function forwardGmailToLineWorks() {
  const label = getOrCreateLabel(CONFIG.PROCESSED_LABEL);
  const query = `from:${CONFIG.SENDER_EMAIL} -label:${CONFIG.PROCESSED_LABEL}`;
  const threads = GmailApp.search(query);

  if (threads.length === 0) {
    Logger.log("転送対象メールなし");
    return;
  }

  const token = getAccessToken();
  if (!token) {
    Logger.log("アクセストークン取得失敗。処理中断。");
    return;
  }

  threads.forEach((thread) => {
    thread.getMessages().forEach((message) => {
      const subject = message.getSubject();
      const body = message.getPlainBody();

      let text;
      if (subject.includes("予約連絡")) {
        text = formatReservation(body);
      } else if (subject.includes("キャンセル連絡")) {
        text = formatCancellation(body);
      } else {
        Logger.log(`未対応の件名のためスキップ: ${subject}`);
        return;
      }

      const success = sendToLineWorks(token, text);
      if (success) {
        thread.addLabel(label);
        Logger.log(`転送成功: ${subject}`);
      } else {
        Logger.log(`転送失敗: ${subject}`);
      }
    });
  });
}

// ============================================================
// 予約連絡のフォーマット
// ============================================================
function formatReservation(body) {
  const extract = (pattern) => {
    const m = body.match(pattern);
    return m ? m[1].trim() : "不明";
  };

  const reservationNo = extract(/■予約番号\s*\n\s*(.+)/);
  const name          = extract(/■氏名\s*\n\s*(.+)/);
  const datetime      = extract(/■来店日時\s*\n\s*(.+)/);
  const staff         = extract(/■指名スタッフ\s*\n\s*(.+)/);
  const amount        = extract(/予約時合計金額\s*([0-9,]+円)/);

  const menuMatch = body.match(/■メニュー\s*\n([\s\S]*?)■ご利用クーポン/);
  const menus = menuMatch
    ? menuMatch[1].split("\n").map(l => l.trim()).filter(l => l).join("／")
    : "不明";

  const couponMatch = body.match(/■ご利用クーポン\s*\n([\s\S]*?)(?:■合計金額|PC版SALON)/);
  const couponLines = couponMatch
    ? couponMatch[1].split("\n").map(l => l.trim()).filter(l => l && !l.match(/^\[.+\]$/))
    : [];
  const coupon = couponLines.length > 0 ? couponLines.join("／") : "なし";

  return (
    `🆕 予約が入りました\n` +
    `───────\n` +
    `📅 ${datetime}\n` +
    `👤 ${name}\n` +
    `\n` +
    `💅 ${menus}\n` +
    `🎫 ${coupon}\n` +
    `💰 ${amount}\n` +
    `👩‍🎨 指名：${staff}\n` +
    `📌 予約番号：${reservationNo}`
  );
}

// ============================================================
// キャンセル連絡のフォーマット
// ============================================================
function formatCancellation(body) {
  const extract = (pattern) => {
    const m = body.match(pattern);
    return m ? m[1].trim() : "不明";
  };

  const reservationNo = extract(/■予約番号\s*\n\s*(.+)/);
  const name          = extract(/■氏名\s*\n\s*(.+)/);
  const datetime      = extract(/■来店日時\s*\n\s*(.+)/);
  const staff         = extract(/■指名スタッフ\s*\n\s*(.+)/);

  const menuMatch = body.match(/■メニュー\s*\n([\s\S]*?)(?:■ご利用クーポン|PC版SALON)/);
  const menus = menuMatch
    ? menuMatch[1].split("\n").map(l => l.trim()).filter(l => l).join("／")
    : "不明";

  return (
    `❌ キャンセルがありました\n` +
    `───────\n` +
    `📅 ${datetime}\n` +
    `👤 ${name}\n` +
    `\n` +
    `💅 ${menus}\n` +
    `👩‍🎨 指名：${staff}\n` +
    `📌 予約番号：${reservationNo}`
  );
}

// ============================================================
// JWT生成 → アクセストークン取得
// ============================================================
function getAccessToken() {
  try {
    const now = Math.floor(Date.now() / 1000);
    const header  = { alg: "RS256", typ: "JWT" };
    const payload = {
      iss: CONFIG.CLIENT_ID,
      sub: CONFIG.SERVICE_ACCOUNT,
      iat: now,
      exp: now + 3600,
    };

    const encHeader    = base64UrlEncode(JSON.stringify(header));
    const encPayload   = base64UrlEncode(JSON.stringify(payload));
    const signingInput = `${encHeader}.${encPayload}`;
    const signature    = Utilities.computeRsaSha256Signature(signingInput, CONFIG.PRIVATE_KEY);
    const encSignature = Utilities.base64EncodeWebSafe(signature).replace(/=+$/, "");
    const jwt = `${signingInput}.${encSignature}`;

    const response = UrlFetchApp.fetch("https://auth.worksmobile.com/oauth2/v2.0/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      payload: {
        assertion:     jwt,
        grant_type:    "urn:ietf:params:oauth:grant-type:jwt-bearer",
        client_id:     CONFIG.CLIENT_ID,
        client_secret: CONFIG.CLIENT_SECRET,
        scope:         "bot.message",
      },
      muteHttpExceptions: true,
    });

    const code = response.getResponseCode();
    const json = JSON.parse(response.getContentText());
    if (code === 200 && json.access_token) {
      Logger.log("アクセストークン取得成功");
      return json.access_token;
    } else {
      Logger.log(`トークン取得失敗: ${code} - ${JSON.stringify(json)}`);
      return null;
    }
  } catch (e) {
    Logger.log(`JWT生成エラー: ${e.message}`);
    return null;
  }
}

// ============================================================
// LINE WORKS APIでメッセージ送信
// ============================================================
function sendToLineWorks(token, text) {
  const url = `https://www.worksapis.com/v1.0/bots/${CONFIG.BOT_ID}/channels/${CONFIG.CHANNEL_ID}/messages`;

  const response = UrlFetchApp.fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    payload: JSON.stringify({
      content: { type: "text", text: text },
    }),
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  if (code === 200 || code === 201) return true;
  Logger.log(`送信APIエラー: ${code} - ${response.getContentText()}`);
  return false;
}

// ============================================================
// ユーティリティ
// ============================================================
function base64UrlEncode(str) {
  return Utilities.base64EncodeWebSafe(str).replace(/=+$/, "");
}

function getOrCreateLabel(labelName) {
  return GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
}

// ============================================================
// テスト：フォーマット確認（APIは呼ばない）
// ============================================================
// テストデータ
const TEST_RESERVATION = `■予約番号
　BE00000001
■氏名
　山田 太郎（ヤマダ タロウ）
■来店日時
　2026年05月22日（金）12:30
■指名スタッフ
　指名なし
■メニュー
　ジェル
　オフメニュー
　（所要時間目安：1時間30分）
■ご利用クーポン
　[全員]
　【テスト】ダミークーポン
■合計金額
　予約時合計金額　6,800円
　今回の利用ギフト券　利用なし
　今回の利用ポイント　利用なし
　お支払い予定金額　6,800円`;

const TEST_CANCELLATION = `■予約番号
　BE00000002
■氏名
　鈴木 花子（スズキ ハナコ）
■来店日時
　2026年05月13日（水）13:30
■指名スタッフ
　テストスタッフ
■メニュー
　ジェル
　（所要時間目安：2時間15分）
■ご利用クーポン
　【テスト】ダミークーポン`;

// [全員]なしの形式のテスト
const TEST_RESERVATION_NO_BRACKET = `■予約番号
　BE92221561
■氏名
　ダミー（ダミー）
■来店日時
　2026年05月10日（日）10:30
■指名スタッフ
　指名なし
■メニュー
　ジェル
　オフメニュー
　オプションメニュー
　（所要時間目安：3時間10分）
■ご利用クーポン
　【テスト】ダミークーポン
■合計金額
　予約時合計金額　12,150円
　今回の利用ギフト券　利用なし
　今回の利用ポイント　利用なし
　お支払い予定金額　12,150円`;

// ============================================================
// テスト：フォーマット確認（ログのみ・API呼ばない）
// ============================================================
function testFormat() {
  Logger.log("=== 予約連絡 ===");
  Logger.log(formatReservation(TEST_RESERVATION));
  Logger.log("=== キャンセル連絡 ===");
  Logger.log(formatCancellation(TEST_CANCELLATION));
  Logger.log("=== 予約連絡（[全員]なし形式） ===");
  Logger.log(formatReservation(TEST_RESERVATION_NO_BRACKET));
}

// ============================================================
// テスト：フォーマット確認 + LINE WORKSへ実際に送信
// ============================================================
function testSend() {
  const token = getAccessToken();
  if (!token) {
    Logger.log("トークン取得失敗。認証情報を確認してください。");
    return;
  }

  const resText = formatReservation(TEST_RESERVATION);
  const canText = formatCancellation(TEST_CANCELLATION);

  Logger.log("=== 予約連絡 送信 ===");
  Logger.log(sendToLineWorks(token, resText) ? "成功" : "失敗");

  Logger.log("=== キャンセル連絡 送信 ===");
  Logger.log(sendToLineWorks(token, canText) ? "成功" : "失敗");
}
