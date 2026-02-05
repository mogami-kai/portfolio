const CFG = {
  WORK_SHEET: "WORK_DAILY",
  RAW_SHEET: "RAW_LINE",
  LOG_SHEET: "T_LINE_MESSAGE_LOG",

  // 未来日補正：今日から何日以上未来なら「前年」とみなす（年跨ぎ対策）
  FUTURE_DAY_LIMIT: 30,

  // 重複チェック（ログの後ろ何行を見るか）
  DUP_CHECK_ROWS: 2000,

  // 送信用
  CHANNEL_ACCESS_TOKEN: ""
};

// WORK_DAILYの列（固定）
const WORK_HEADERS = [
  "管理ID",        // A: uniqueKey
  "日付",          // B: yyyy/MM/dd
  "年月",          // C: yyyy-MM
  "取引先",        // D
  "契約種別",      // E: 常用/請負
  "現場",          // F
  "職人名",        // G
  "人工",          // H: 1.0 / 0.5 etc
  "残業H",         // I: 0 / 1 / 1.5 etc
  "元メッセージID",// J
  "登録日時"       // K: yyyy/MM/dd HH:mm:ss
];

// ----------------------------------------------------
// 初期化（手動実行OK）: シートとヘッダーを揃える
// ----------------------------------------------------
function initOpsDb() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // WORK_DAILY
  let w = ss.getSheetByName(CFG.WORK_SHEET);
  if (!w) w = ss.insertSheet(CFG.WORK_SHEET);
  ensureHeaderRow_(w, WORK_HEADERS);

  // RAW_LINE
  let r = ss.getSheetByName(CFG.RAW_SHEET);
  if (!r) r = ss.insertSheet(CFG.RAW_SHEET);
  ensureHeaderRow_(r, ["Timestamp", "GroupID", "UserID", "Text"]);

  // LOG
  let l = ss.getSheetByName(CFG.LOG_SHEET);
  if (!l) l = ss.insertSheet(CFG.LOG_SHEET);
  ensureHeaderRow_(l, ["MessageID", "Timestamp", "GroupID", "UserID", "Text", "Status", "Detail"]);
}

// ----------------------------------------------------
// 1) LINE Webhook エントリーポイント
// ----------------------------------------------------
function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) return ContentService.createTextOutput("OK");

  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput("JSON Parse Error");
  }

  const event = payload.events && payload.events[0];
  if (!event) return ContentService.createTextOutput("No Event");

  const ts = new Date(event.timestamp || Date.now());
  const groupId = (event.source && event.source.groupId) ? event.source.groupId : "";
  const userId  = (event.source && event.source.userId)  ? event.source.userId  : "";

  // A) 送信取消（Unsend）
  if (event.type === "unsend") {
    const unsentMessageId = event.unsend && event.unsend.messageId;
    if (unsentMessageId) {
      const count = deleteByMessageId_(unsentMessageId);
      appendLog_(ts, unsentMessageId, groupId, userId, "[UNSEND]", "DELETED", `削除件数: ${count}`);
    }
    return ContentService.createTextOutput("OK");
  }

  // B) テキストメッセージ受信
  if (event.type === "message" && event.message && event.message.type === "text") {
    const messageId = String(event.message.id || "");
    const messageText = String(event.message.text || "");

    // 生ログ
    appendRawLineSafe_(ts, groupId, userId, messageText);

    // 重複チェック
    if (isDuplicateMessageId_(messageId)) {
      appendLog_(ts, messageId, groupId, userId, messageText, "DUPLICATE", "処理スキップ");
      return ContentService.createTextOutput("OK");
    }

    // 登録
    try {
      const result = processLineMessage_(messageText, messageId, ts);
      appendLog_(ts, messageId, groupId, userId, messageText, "SUCCESS", `登録: ${result.rows}件`);
    } catch (err) {
      appendLog_(ts, messageId, groupId, userId, messageText, "ERROR", err && err.stack ? err.stack : String(err));
    }
  }

  return ContentService.createTextOutput("OK");
}

// ----------------------------------------------------
// 2) メイン処理：テキスト解析 → WORK_DAILYへ追記
// ----------------------------------------------------
function processLineMessage_(text, messageId, receivedAt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.WORK_SHEET);
  if (!sheet) throw new Error(`シートが見つかりません: ${CFG.WORK_SHEET}`);

  // ヘッダーがズレてると Middle_db 側も崩れるので、最低限揃える
  ensureHeaderRow_(sheet, WORK_HEADERS);

  const lines = String(text)
    .split(/\r\n|\r|\n/)
    .map(s => String(s).trim())
    .filter(Boolean);

  let currentDate = null;
  let currentClient = "";
  let currentWorkType = ""; // 常用/請負
  let currentSite = "";

  let nextIsClient = false;
  let nextIsSite = false;

  // ブロック共通指定（後付け上書き）
  let blockDefaultQty = null;
  let blockDefaultOt  = null;
  let blockStartIndex = 0;

  const newRows = [];
  const tz = Session.getScriptTimeZone();
  const nowStr = Utilities.formatDate(new Date(), tz, "yyyy/MM/dd HH:mm:ss");

  for (const rawLine of lines) {

    // A) 日付判定
    const dateObj = parseDateLineFlex_(rawLine, receivedAt);
    if (dateObj) {
      currentDate = dateObj;

      // ブロック初期化
      currentClient = "";
      currentWorkType = "";
      currentSite = "";
      blockDefaultQty = null;
      blockDefaultOt = null;

      nextIsClient = true;
      nextIsSite = false;
      continue;
    }
    if (!currentDate) continue;

    // B) 取引先 + 契約種別 + (混在なら現場も)
    if (nextIsClient) {
      const parsed = parseClientWorkTypeAndMaybeSite_(rawLine);

      currentClient = parsed.clientName;
      currentWorkType = parsed.workType;

      if (parsed.siteName) {
        currentSite = parsed.siteName;
        nextIsClient = false;
        nextIsSite = false;

        blockStartIndex = newRows.length;
        blockDefaultQty = null;
        blockDefaultOt = null;
      } else {
        nextIsClient = false;
        nextIsSite = true;
      }
      continue;
    }

    // C) 現場
    if (nextIsSite) {
      currentSite = rawLine;
      nextIsSite = false;

      blockStartIndex = newRows.length;
      blockDefaultQty = null;
      blockDefaultOt = null;
      continue;
    }

    // C-2) 共通指定（人工のみ）
    const qtyOnly = parseQtyOnlyLine_(rawLine);
    if (qtyOnly !== null) {
      blockDefaultQty = qtyOnly;
      for (let i = blockStartIndex; i < newRows.length; i++) newRows[i][7] = blockDefaultQty; // H:人工
      continue;
    }

    // C-3) 共通指定（残業のみ）
    const otOnly = parseOtOnlyLine_(rawLine);
    if (otOnly !== null) {
      blockDefaultOt = otOnly;
      for (let i = blockStartIndex; i < newRows.length; i++) newRows[i][8] = blockDefaultOt; // I:残業H
      continue;
    }

    // D) 職人行（カオス対応）
    const workers = parseWorkerLineChaos_(rawLine);
    if (workers.length > 0) {
      const ym = Utilities.formatDate(currentDate, tz, "yyyy-MM");
      const dateFormatted = Utilities.formatDate(currentDate, tz, "yyyy/MM/dd");

      for (let idx = 0; idx < workers.length; idx++) {
        const w = workers[idx];

        // ✅ Middle_db 側でUPSERTしやすいように messageId + idx
        const uniqueKey = `${messageId}_${idx}`;

        // 個別指定 > 共通指定 > デフォルト
        const qtyFinal = w.hasQty ? w.qty : (blockDefaultQty !== null ? blockDefaultQty : 1.0);
        const otFinal  = w.hasOt  ? w.ot  : (blockDefaultOt  !== null ? blockDefaultOt  : 0);

        newRows.push([
          uniqueKey,        // A 管理ID
          dateFormatted,    // B 日付
          ym,               // C 年月
          currentClient,    // D 取引先
          currentWorkType,  // E 契約種別
          currentSite,      // F 現場
          w.name,           // G 職人名
          qtyFinal,         // H 人工
          otFinal,          // I 残業H
          messageId,        // J 元メッセージID
          nowStr            // K 登録日時
        ]);
      }
    }
  }

  if (newRows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newRows.length, WORK_HEADERS.length).setValues(newRows);

    // 書式（最低限）
    sheet.getRange(lastRow + 1, 2, newRows.length, 1).setNumberFormat("yyyy/mm/dd"); // 日付
    sheet.getRange(lastRow + 1, 8, newRows.length, 2).setNumberFormat("0.00");       // 人工・残業H
  }

  return { rows: newRows.length };
}

// ----------------------------------------------------
// 3) パース関数群
// ----------------------------------------------------

// 日付解析（曜日/英語曜日/括弧などがあっても数字だけ拾う）
function parseDateLineFlex_(text, baseDate) {
  // ✅ / を潰さない正規化
  const s = normalizeText_(text).replace(/\s/g, "");

  // 2026/01/16, 2026-01-16
  let m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  // 1/16(火) / 01/16 / 1-16 / 1/16Tue
  m = s.match(/(\d{1,2})[\/\-](\d{1,2})/);
  if (m) return fixFutureDate_(baseDate, Number(m[1]), Number(m[2]));

  // 1月16日
  m = s.match(/(\d{1,2})月(\d{1,2})/);
  if (m) return fixFutureDate_(baseDate, Number(m[1]), Number(m[2]));

  return null;
}

// 未来日付補正
function fixFutureDate_(baseDate, mm, dd) {
  const y = baseDate.getFullYear();
  let d = new Date(y, mm - 1, dd);

  const diff = (d - baseDate) / (1000 * 60 * 60 * 24);

  // 未来すぎる → 去年に戻す
  if (diff > CFG.FUTURE_DAY_LIMIT) return new Date(y - 1, mm - 1, dd);

  // 過去すぎる（年跨ぎ直後のメッセージなど） → 来年に寄せる
  if (diff < -300) return new Date(y + 1, mm - 1, dd);

  return d;
}

// 取引先行：常用/請負抽出 + 1行に現場が混ざる場合も拾う
function parseClientWorkTypeAndMaybeSite_(line) {
  // ✅ / | を保持したまま統一
  let normalized = normalizeText_(line)
    .replace(/[　]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  let clientPart = normalized;
  let sitePart = "";

  // 例: "恵興業 常用 / 追浜造船所"
  // 例: "恵興業 常用｜追浜造船所"
  if (/[\/|]/.test(normalized)) {
    const parts = normalized.split(/[\/|]/).map(s => s.trim()).filter(Boolean);
    if (parts.length >= 2) {
      clientPart = parts[0];
      sitePart = parts.slice(1).join(" / ");
    }
  }

  const hasJoyo = /常用/.test(clientPart);
  const hasUkeoi = /請負/.test(clientPart);

  let workType = "";
  if (hasJoyo) workType = "常用";
  if (hasUkeoi) workType = "請負";

  let clientName = clientPart
    .replace(/常用/g, "")
    .replace(/請負/g, "")
    .replace(/\s+/g, " ")
    .trim();

  return { clientName, workType, siteName: sitePart };
}

// 共通指定：人工だけの行
function parseQtyOnlyLine_(line) {
  const raw = normalizeText_(line).toLowerCase().trim();
  const s = raw.replace(/\s+/g, "");

  // 半日系
  if (["半日", "半", "半勤", "午前", "午後", "午前のみ", "午後のみ", "am", "pm", "半日勤務"].includes(raw)) return 0.5;

  // 1日系
  if (["一日", "1日", "１日", "全日"].includes(raw)) return 1.0;

  // 0.5 / 0.5日 / 0.5day
  let m = s.match(/^([0-9]+(\.[0-9]+)?)\s*(日|day)?$/);
  if (m) return parseFloat(m[1]);

  // 人工0.5 / 稼働0.5
  m = s.match(/^(人工|稼働)\s*([0-9]+(\.[0-9]+)?)$/);
  if (m) return parseFloat(m[2]);

  return null;
}

// 共通指定：残業だけの行
function parseOtOnlyLine_(line) {
  const raw = normalizeText_(line).toLowerCase().trim();
  const s = raw.replace(/\s+/g, "");

  // 残業1 / 残業1h / 残業1時間
  let m = s.match(/^残業([0-9]+(\.[0-9]+)?)((h|時間)?)$/);
  if (m) return parseFloat(m[1]);

  // 残1
  m = s.match(/^残([0-9]+(\.[0-9]+)?)$/);
  if (m) return parseFloat(m[1]);

  // 1h / +1h
  m = s.match(/^\+?([0-9]+(\.[0-9]+)?)h$/);
  if (m) return parseFloat(m[1]);

  // OT1 / OT1h / overtime1h
  m = s.match(/^ot([0-9]+(\.[0-9]+)?)h?$/);
  if (m) return parseFloat(m[1]);

  m = s.match(/^overtime([0-9]+(\.[0-9]+)?)h?$/);
  if (m) return parseFloat(m[1]);

  return null;
}

// カオス混在：名前と状態が同一行で順不同でもOK
function parseWorkerLineChaos_(line) {
  const tokens = tokenizeChaos_(line);
  if (tokens.length === 0) return [];

  // 名前判定（状態語や英数は除外）
  const isName = (t) => {
    if (!t) return false;
    if (isStatusToken_(t)) return false;
    return /^[一-龥々ぁ-んァ-ヶ]{1,12}$/.test(t);
  };

  const workers = [];
  let current = null;

  // 状態が先に来たとき用
  let pending = { qty: null, ot: null };

  for (const t of tokens) {
    // 状態
    const st = parseStatusToken_(t);
    if (st) {
      if (current) {
        if (st.qty !== null) { current.qty = st.qty; current.hasQty = true; }
        if (st.ot  !== null) { current.ot  = st.ot;  current.hasOt  = true; }
      } else {
        if (st.qty !== null) pending.qty = st.qty;
        if (st.ot  !== null) pending.ot  = st.ot;
      }
      continue;
    }

    // 名前
    if (isName(t)) {
      if (current) workers.push(current);
      current = { name: t, qty: 1.0, ot: 0, hasQty: false, hasOt: false };

      // pending 適用
      if (pending.qty !== null) { current.qty = pending.qty; current.hasQty = true; }
      if (pending.ot  !== null) { current.ot  = pending.ot;  current.hasOt  = true; }
      pending = { qty: null, ot: null };
      continue;
    }

    // 数字だけ → 人工（個別指定）
    const num = parseFloatSafe_(t);
    if (num !== null) {
      if (current) {
        current.qty = num;
        current.hasQty = true;
      } else {
        pending.qty = num;
      }
      continue;
    }
  }

  if (current) workers.push(current);

  // 0人なら誤爆なので空
  if (workers.length === 0) return [];

  return workers;
}

// ----------------------------------------------------
// 4) 正規化/トークン化/状態判定
// ----------------------------------------------------

// ✅ 改訂：/ と | を潰さない（＝日付/取引先分割が死なない）
function normalizeText_(text) {
  let s = String(text || "");
  if (s.normalize) s = s.normalize("NFKC");

  s = s
    .replace(/[　]+/g, " ")
    .replace(/[、，,]/g, " ")
    .replace(/[。．]/g, ".")
    .replace(/[＋+]/g, "+");

  // 全角を半角に寄せるが、記号自体は保持
  s = s
    .replace(/／/g, "/")
    .replace(/｜/g, "|");

  return s;
}

// ✅ 職人行専用：ここでだけ / | を空白にする（名前抽出の邪魔を消す）
function tokenizeChaos_(line) {
  let s = normalizeText_(line);

  // かっこ類は空白化
  s = s.replace(/[（）\(\)\[\]【】]/g, " ");

  // 職人行では区切り記号は邪魔なので空白化
  s = s.replace(/[\/|]/g, " ");

  // 「残業 1 h」→「1h」
  s = s.replace(/(\d)\s*h\b/gi, "$1h");

  // 連続スペース整理
  s = s.replace(/\s+/g, " ").trim();
  if (!s) return [];

  return s.split(" ").filter(Boolean);
}

function parseFloatSafe_(t) {
  const s = String(t || "").trim();
  if (!s) return null;
  if (!/^([0-9]+(\.[0-9]+)?)$/.test(s)) return null;
  const v = parseFloat(s);
  return isNaN(v) ? null : v;
}

function isStatusToken_(t) {
  return !!parseStatusToken_(t);
}

function parseStatusToken_(token) {
  const raw = normalizeText_(token).toLowerCase().trim();
  const s = raw.replace(/\s+/g, "");

  // 半日系（人工=0.5）
  if (["半日", "半", "半勤", "午前", "午後", "午前のみ", "午後のみ", "am", "pm", "半日勤務"].includes(raw)) {
    return { qty: 0.5, ot: null };
  }

  // 0.5日 / 0.5day（数字単独は人工と衝突するため「日/day付きのみ」）
  let m = s.match(/^([0-9]+(\.[0-9]+)?)((日|day))$/);
  if (m) return { qty: parseFloat(m[1]), ot: null };

  // 残業1 / 残業1h / 残業1時間
  m = s.match(/^残業([0-9]+(\.[0-9]+)?)((h|時間)?)$/);
  if (m) return { qty: null, ot: parseFloat(m[1]) };

  // 残1
  m = s.match(/^残([0-9]+(\.[0-9]+)?)$/);
  if (m) return { qty: null, ot: parseFloat(m[1]) };

  // 1h / +1h
  m = s.match(/^\+?([0-9]+(\.[0-9]+)?)h$/);
  if (m) return { qty: null, ot: parseFloat(m[1]) };

  // OT1 / OT1h
  m = s.match(/^ot([0-9]+(\.[0-9]+)?)h?$/);
  if (m) return { qty: null, ot: parseFloat(m[1]) };

  // overtime1h
  m = s.match(/^overtime([0-9]+(\.[0-9]+)?)h?$/);
  if (m) return { qty: null, ot: parseFloat(m[1]) };

  return null;
}

// ----------------------------------------------------
// 5) ログ・削除・ユーティリティ
// ----------------------------------------------------
function ensureHeaderRow_(sheet, headers) {
  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  const row1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 1行目が空っぽに近いなら書く
  const hasAny = row1.some(v => String(v || "").trim() !== "");
  if (!hasAny) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  // すでに何かある場合：足りない分だけ右に足す（既存列を壊さない）
  for (let i = 0; i < headers.length; i++) {
    const cur = String(row1[i] || "").trim();
    if (!cur) sheet.getRange(1, i + 1).setValue(headers[i]);
  }
}

function appendLog_(ts, msgId, groupId, userId, text, status, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CFG.LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CFG.LOG_SHEET);
    sheet.getRange(1, 1, 1, 7).setValues([["MessageID", "Timestamp", "GroupID", "UserID", "Text", "Status", "Detail"]]);
  }
  sheet.appendRow([
    String(msgId || ""),
    Utilities.formatDate(ts, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss"),
    String(groupId || ""),
    String(userId || ""),
    String(text || ""),
    String(status || ""),
    String(detail || "")
  ]);
}

function isDuplicateMessageId_(msgId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.LOG_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return false;

  const lastRow = sheet.getLastRow();
  const startRow = Math.max(2, lastRow - CFG.DUP_CHECK_ROWS + 1);
  const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 1).getValues();

  const ids = data.flat().map(String);
  return ids.includes(String(msgId));
}

function appendRawLineSafe_(ts, groupId, userId, text) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CFG.RAW_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CFG.RAW_SHEET);
    sheet.getRange(1, 1, 1, 4).setValues([["Timestamp", "GroupID", "UserID", "Text"]]);
  }

  sheet.appendRow([
    Utilities.formatDate(ts, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss"),
    String(groupId || ""),
    String(userId || ""),
    String(text || "")
  ]);
}

function deleteByMessageId_(msgId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.WORK_SHEET);
  if (!sheet) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  // J列（10列目）が「元メッセージID」
  const ids = sheet.getRange(2, 10, lastRow - 1, 1).getValues().flat().map(String);
  const targetId = String(msgId);

  const rowsToDelete = [];
  for (let i = 0; i < ids.length; i++) {
    if (ids[i] === targetId) rowsToDelete.push(i + 2);
  }

  rowsToDelete.reverse().forEach(r => sheet.deleteRow(r));
  return rowsToDelete.length;
}
