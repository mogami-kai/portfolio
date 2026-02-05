const MIDDLE_SS_ID = "1u7nJyh7wQDWN13nwmPfV0DI4ysOqsM1ef1UtIbA2ZgE";
const MIDDLE_SHEET_NAME = "APPROVED_DAILY";

const ADMIN_SHEET_NAME = "ADMIN_DAILY";
const WORKER_SHEET_NAME = "SECRET_M_WORKER";

// ✅ 1行目ヘッダーの辞書（列ズレに強い）
const JP_HEADER_ALIASES = {
  key:      ["管理ID", "uniqueKey", "ID"],
  ym:       ["年月", "YM", "請求月"],
  date:     ["日付", "Date", "作業日"],

  site:     ["現場名（元請/現場）", "現場名", "現場", "場所", "Site"],
  workType: ["契約種別（常用/請負）", "契約種別", "常用/請負"],  // ★追加
  worker:   ["職人名", "氏名", "名前", "Worker"],
  qty:      ["人工", "稼働", "Qty"],
  ot:       ["残業H", "残業(H)", "残業", "OT"],
  status:   ["状態", "ステータス", "Status"],                 // ★追加

  // マスタ・計算列
  base:   ["(マスタ)基本日当", "基本日当"],
  rate:   ["(マスタ)残業倍率", "残業倍率"],
  manual: ["【手動】日当変更", "手動"],
  fixed:  ["確定日当", "確定"],
  total:  ["最終支給額", "支給額（給与）", "支給額", "【手動】支給額上書き"]
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("管理メニュー")
    .addItem("① 承認済データを取り込む（同期）", "syncFromMiddleJP")
    .addToUi();
}

function syncFromMiddleJP() {
  const ui = SpreadsheetApp.getUi();
  const adminSS = SpreadsheetApp.getActiveSpreadsheet();
  const adminSheet = adminSS.getSheetByName(ADMIN_SHEET_NAME);
  const workerSheet = adminSS.getSheetByName(WORKER_SHEET_NAME);

  if (!adminSheet) throw new Error(`Sheet not found: ${ADMIN_SHEET_NAME}`);
  if (!workerSheet) throw new Error(`Sheet not found: ${WORKER_SHEET_NAME}`);

  // 1) Adminヘッダー解析
  const col = getHeaderMapJP_(adminSheet);

  // ★必須項目に「契約種別」追加
  const required = ["key", "ym", "date", "site", "workType", "worker", "qty", "ot", "status"];
  const missing = required.filter(k => !col[k]);

  if (missing.length > 0) {
    ui.alert("ヘッダーが見つかりません: " + missing.join(", "));
    return;
  }

  // 2) Middle(承認済)取得
  let middleSheet;
  try {
    middleSheet = SpreadsheetApp.openById(MIDDLE_SS_ID).getSheetByName(MIDDLE_SHEET_NAME);
  } catch(e) {
    ui.alert("Middleファイルにアクセスできません。IDを確認してください。");
    return;
  }
  if (!middleSheet) {
    ui.alert("Middleシートが見つかりません");
    return;
  }

  const lastRow = middleSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("Middle側に承認済みデータがありません");
    return;
  }

  // ✅ Middle APPROVED_DAILY は 14列（A〜N）
  // 0 管理ID
  // 1 日付
  // 2 年月
  // 3 取引先名
  // 4 契約種別（常用/請負）
  // 5 現場名
  // 6 職人名
  // 7 人工
  // 8 残業H
  // 9 元メッセージID
  // 10 登録日時
  // 11 状態（APPROVED）
  // 12 承認者
  // 13 承認時刻
  const data = middleSheet.getRange(2, 1, lastRow - 1, 14).getValues();

  // 3) マスタ & 既存データ確認
  const validWorkers = loadValidWorkersJP_(workerSheet);
  const existingKeys = loadExistingKeys_(adminSheet, col.key);

  const newObjs = [];
  let skippedDup = 0;
  const missingNames = new Set();

  for (const r of data) {
    const key = String(r[0] || "").trim();
    if (!key) continue;

    // 既にAdminにあるならスキップ
    if (existingKeys.has(key)) { skippedDup++; continue; }

    const workerName = String(r[6] || "").trim();
    if (!validWorkers.has(workerName)) {
      if (workerName) missingNames.add(workerName);
      continue;
    }

    const client = String(r[3] || "").trim();
    const workType = String(r[4] || "").trim();
    const site = String(r[5] || "").trim();

    // ✅ 現場名は「取引先 / 現場」に統一（見やすい）
    const siteMerged = (client && site) ? `${client} / ${site}` : (site || client);

    newObjs.push({
      key: key,
      ym: r[2],
      date: r[1],
      site: siteMerged,
      workType: workType,    // ★ここに別列で入れる
      worker: workerName,
      qty: r[7],
      ot: r[8],
      status: String(r[11] || "APPROVED") // ★状態列に入れる
    });
  }

  if (newObjs.length === 0) {
    let msg = `新規取込なし\n(重複スキップ: ${skippedDup}件)`;
    if (missingNames.size > 0) {
      msg += `\n\n⚠️以下の職人名がマスタ「SECRET_M_WORKER」に無いため、取り込めませんでした：\n[ ${Array.from(missingNames).join(", ")} ]\n\n↑この名前をA列に追加してください。`;
    }
    ui.alert(msg);
    return;
  }

  // 4) Adminへ書き込み
  const startRow = adminSheet.getLastRow() + 1;
  const lastCol = adminSheet.getLastColumn();
  const out = Array.from({ length: newObjs.length }, () => Array(lastCol).fill(""));

  for (let i = 0; i < newObjs.length; i++) {
    const o = newObjs[i];
    setCell_(out[i], col.key, o.key);
    setCell_(out[i], col.ym, o.ym);
    setCell_(out[i], col.date, o.date);
    setCell_(out[i], col.site, o.site);
    setCell_(out[i], col.workType, o.workType); // ★追加
    setCell_(out[i], col.worker, o.worker);
    setCell_(out[i], col.qty, o.qty);
    setCell_(out[i], col.ot, o.ot);
    setCell_(out[i], col.status, o.status);     // ★追加
  }

  adminSheet.getRange(startRow, 1, out.length, lastCol).setValues(out);

  // 5) 計算式と書式
  applyFormulasJP_(adminSheet, col, startRow, out.length);
  applyFormatsJP_(adminSheet, col, startRow, out.length);

  ui.alert(`同期完了: ${newObjs.length}件取り込みました`);
}

// ---------------- Helper ----------------

function getHeaderMapJP_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const map = {};
  for (const k in JP_HEADER_ALIASES) {
    const idx = header.findIndex(h => JP_HEADER_ALIASES[k].includes(h));
    if (idx >= 0) map[k] = idx + 1;
  }
  return map;
}

function loadValidWorkersJP_(s) {
  const last = s.getLastRow();
  if (last < 2) return new Set();
  const d = s.getRange(2, 1, last - 1, 3).getValues();
  const set = new Set();
  d.forEach(r => { if (r[0]) set.add(String(r[0]).trim()); });
  return set;
}

function loadExistingKeys_(s, k) {
  const last = s.getLastRow();
  if (last < 2) return new Set();
  const vals = s.getRange(2, k, last - 1, 1).getValues().flat();
  return new Set(vals.map(String));
}

function setCell_(row, c, v) { if (c) row[c - 1] = v; }
function colToA1_(c) {
  let s = "";
  while (c > 0) { s = String.fromCharCode(65 + (c - 1) % 26) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

// ▼▼▼ 残業倍率を「1.25固定」にする ▼▼▼
function applyFormulasJP_(s, c, r, n) {
  if (!c.base || !c.total) return;

  const w = colToA1_(c.worker), q = colToA1_(c.qty), o = colToA1_(c.ot);
  const b = colToA1_(c.base), rt = colToA1_(c.rate), m = colToA1_(c.manual), f = colToA1_(c.fixed);

  const formulas = Array.from({ length: n }, (_, i) => {
    const idx = r + i;
    return [
      `=IFERROR(VLOOKUP(${w}${idx}, ${WORKER_SHEET_NAME}!A:C, 2, 0), 0)`, // 基本日当
      "1.25", // 🔴 残業倍率固定
      ``,
      `=IF(${m}${idx}<>"", ${m}${idx}, ${b}${idx})`, // 確定日当
      `=ROUNDDOWN(${f}${idx}*${q}${idx} + (${f}${idx}/8*${rt}${idx})*${o}${idx})` // 最終計算
    ];
  });

  if (c.base + 4 === c.total) {
    s.getRange(r, c.base, n, 5).setFormulas(formulas);
  }
}

function applyFormatsJP_(s, c, r, n) {
  if (c.date) s.getRange(r, c.date, n, 1).setNumberFormat("yyyy/mm/dd");
  if (c.qty)  s.getRange(r, c.qty,  n, 1).setNumberFormat("0.00");
  if (c.ot)   s.getRange(r, c.ot,   n, 1).setNumberFormat("0.00");

  // 🔴 1.25を「1」表示にしない
  if (c.rate) s.getRange(r, c.rate, n, 1).setNumberFormat("0.00");

  if (c.total) s.getRange(r, c.total, n, 1).setNumberFormat("#,##0");
}
