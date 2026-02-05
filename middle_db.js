const OPS_SS_ID = "1iBY2DnltjGbAltS0OtQer1Wn9vaZf2G2vFlaPPAk3RQ";
const OPS_SHEET_NAME = "WORK_DAILY";

const REVIEW_SHEET   = "REVIEW_DAILY";
const APPROVED_SHEET = "APPROVED_DAILY";
const REJECTED_SHEET = "REJECTED_LOG";

const UPSERT_OPEN_ROWS = true;

// ============================
// メニュー（4つ）
// ============================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("管理メニュー")
    .addItem("① 現場データを取り込む（同期）", "syncOpsToReview")
    .addItem("② OPENを承認して確定する", "approveOpenToApproved")
    .addItem("③ OPENを却下してREJECTへ", "rejectOpenToRejected")
    .addSeparator()
    .addItem("⚠️ 全データを初期化（リセット）", "emergencyReset")
    .addToUi();
}

// ============================
// 必要シートを自動作成 + ヘッダー強制
// ============================
function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let review    = ss.getSheetByName(REVIEW_SHEET);
  let approved  = ss.getSheetByName(APPROVED_SHEET);
  let rejected  = ss.getSheetByName(REJECTED_SHEET);

  if (!review)   review   = ss.insertSheet(REVIEW_SHEET);
  if (!approved) approved = ss.insertSheet(APPROVED_SHEET);
  if (!rejected) rejected = ss.insertSheet(REJECTED_SHEET);

  // REVIEWヘッダー（A〜Z / 26列）
  const reviewHeader = [
    "管理ID","日付","年月","取引先名","契約種別（常用/請負）","現場名","職人名","人工","残業H","元メッセージID","登録日時",
    "状態","承認者","承認時刻",
    "【手動】取引先名","【手動】契約種別","【手動】現場名","【手動】職人名","【手動】人工","【手動】残業H",
    "最終_取引先名","最終_契約種別","最終_現場名","最終_職人名","最終_人工","最終_残業H"
  ];

  // APPROVEDヘッダー（A〜N / 14列）
  const approvedHeader = [
    "管理ID","日付","年月","取引先名","契約種別（常用/請負）","現場名","職人名","人工","残業H",
    "元メッセージID","登録日時","状態","承認者","承認時刻"
  ];

  // REJECTEDヘッダー（A〜O / 15列）
  const rejectedHeader = [
    "管理ID","日付","年月","取引先名","契約種別（常用/請負）","現場名","職人名","人工","残業H",
    "元メッセージID","登録日時","状態","却下者","却下時刻","却下理由"
  ];

  const ensureHeader = (sheet, headerArr) => {
    const lastCol = Math.max(sheet.getLastColumn(), headerArr.length);
    const row1 = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const current = row1.slice(0, headerArr.length).map(v => String(v || "").trim());
    const expected = headerArr.map(v => String(v).trim());
    const same = JSON.stringify(current) === JSON.stringify(expected);

    if (!same) {
      sheet.getRange(1, 1, 1, headerArr.length).setValues([headerArr]);
    }
  };

  ensureHeader(review, reviewHeader);
  ensureHeader(approved, approvedHeader);
  ensureHeader(rejected, rejectedHeader);

  // REVIEWの数式(U〜Z)は常に維持
  setFormulasToReview_(review);

  review.setFrozenRows(1);
  approved.setFrozenRows(1);
  rejected.setFrozenRows(1);

  return { review, approved, rejected };
}

// ============================
// ① 同期：Ops → REVIEW
// ============================
function syncOpsToReview() {
  const { review, approved, rejected } = ensureSheets_();

  const ops = getOpsSheet_();
  const opsLastRow = getRealLastRow_(ops);
  if (opsLastRow < 2) {
    SpreadsheetApp.getUi().alert("Ops側にデータがありません");
    return;
  }

  const idx = buildHeaderIndexMap_(ops);

  const opsLastCol = ops.getLastColumn();
  const opsData = ops.getRange(2, 1, opsLastRow - 1, opsLastCol).getValues();

  // APPROVEDにあるIDは触らない
  const approvedIds = new Set();
  const approvedLast = getRealLastRow_(approved);
  if (approvedLast >= 2) {
    approved.getRange(2, 1, approvedLast - 1, 1).getValues().flat()
      .forEach(id => approvedIds.add(String(id)));
  }

  // REJECTEDにあるIDも基本スキップ（却下済を復活させたくない前提）
  const rejectedIds = new Set();
  const rejectedLast = getRealLastRow_(rejected);
  if (rejectedLast >= 2) {
    rejected.getRange(2, 1, rejectedLast - 1, 1).getValues().flat()
      .forEach(id => rejectedIds.add(String(id)));
  }

  // REVIEW内のID→行番号（UPSERT用）
  const reviewIdToRow = new Map();
  const reviewLast = getRealLastRow_(review);
  if (reviewLast >= 2) {
    const ids = review.getRange(2, 1, reviewLast - 1, 1).getValues().flat();
    ids.forEach((id, i) => {
      if (id) reviewIdToRow.set(String(id), i + 2);
    });
  }

  let appended = 0;
  let updated = 0;
  const newRows = [];

  for (const r of opsData) {
    const key = String(r[idx.管理ID] ?? "").trim();
    if (!key) continue;

    if (approvedIds.has(key)) continue;
    if (rejectedIds.has(key)) continue;

    // REVIEWにあるなら更新（OPENのみ）
    if (UPSERT_OPEN_ROWS && reviewIdToRow.has(key)) {
      const rowNum = reviewIdToRow.get(key);
      const status = String(review.getRange(rowNum, 12).getValue()); // L=状態
      if (status === "OPEN") {
        review.getRange(rowNum, 1, 1, 11).setValues([extractOpsRow11_(r, idx)]);
        updated++;
      }
      continue;
    }

    // REVIEWに無い → 新規追加
    const base11 = extractOpsRow11_(r, idx);
    newRows.push([
      ...base11,
      "OPEN", "", "",
      "", "", "", "", "", ""
    ]);
  }

  // 新規追加
  if (newRows.length > 0) {
    const startRow = Math.max(2, getRealLastRow_(review) + 1);
    review.getRange(startRow, 1, newRows.length, 20).setValues(newRows);
    appended = newRows.length;
  }

  setFormulasToReview_(review);

  SpreadsheetApp.getUi().alert(`同期完了 ✅\n追加: ${appended}件\n更新: ${updated}件`);
}

// ============================
// ② 承認：REVIEW(OPEN) → APPROVED
// ============================
function approveOpenToApproved() {
  const { review, approved } = ensureSheets_();

  const me = Session.getActiveUser().getEmail();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

  const last = getRealLastRow_(review);
  if (last < 2) {
    SpreadsheetApp.getUi().alert("データがありません");
    return;
  }

  const data = review.getRange(2, 1, last - 1, 26).getValues();

  const push = [];
  const deleteRows = [];

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][11]) !== "OPEN") continue;

    // 最終値（U〜Z）
    const finalClient   = data[i][20];
    const finalWorkType = data[i][21];
    const finalSite     = data[i][22];
    const finalWorker   = data[i][23];
    const finalQty      = data[i][24];
    const finalOt       = data[i][25];

    push.push([
      data[i][0], data[i][1], data[i][2],
      finalClient, finalWorkType, finalSite,
      finalWorker, finalQty, finalOt,
      data[i][9], data[i][10],
      "APPROVED", me, now
    ]);

    deleteRows.push(i + 2);
  }

  if (push.length === 0) {
    SpreadsheetApp.getUi().alert("承認対象（OPEN）がありません");
    return;
  }

  const start = Math.max(2, getRealLastRow_(approved) + 1);
  approved.getRange(start, 1, push.length, 14).setValues(push);

  // REVIEWから削除
  deleteRows.sort((a, b) => b - a);
  deleteRows.forEach(r => {
    if (r === 2) {
      review.getRange(2, 1, 1, 20).clearContent();
    } else {
      review.deleteRow(r);
    }
  });

  setFormulasToReview_(review);

  SpreadsheetApp.getUi().alert(`${push.length}件 承認しました ✅`);
}

// ============================
// ③ 却下：REVIEW(OPEN) → REJECTED_LOG（理由付き）
// ============================
function rejectOpenToRejected() {
  const { review, rejected } = ensureSheets_();
  const ui = SpreadsheetApp.getUi();

  const last = getRealLastRow_(review);
  if (last < 2) {
    ui.alert("データがありません");
    return;
  }

  // 却下理由を入力（全件共通）
  const res = ui.prompt("却下理由", "例）名前誤り / 現場違い / 二重送信 など", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const reason = String(res.getResponseText() || "").trim();
  if (!reason) {
    ui.alert("却下理由が空です。キャンセルしました。");
    return;
  }

  const me = Session.getActiveUser().getEmail();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

  // A〜Z（26列）
  const data = review.getRange(2, 1, last - 1, 26).getValues();

  const push = [];
  const deleteRows = [];

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][11]) !== "OPEN") continue;

    // 最終値（U〜Z）
    const finalClient   = data[i][20];
    const finalWorkType = data[i][21];
    const finalSite     = data[i][22];
    const finalWorker   = data[i][23];
    const finalQty      = data[i][24];
    const finalOt       = data[i][25];

    push.push([
      data[i][0], data[i][1], data[i][2],
      finalClient, finalWorkType, finalSite,
      finalWorker, finalQty, finalOt,
      data[i][9], data[i][10],
      "REJECTED", me, now, reason
    ]);

    deleteRows.push(i + 2);
  }

  if (push.length === 0) {
    ui.alert("却下対象（OPEN）がありません");
    return;
  }

  const start = Math.max(2, getRealLastRow_(rejected) + 1);
  rejected.getRange(start, 1, push.length, 15).setValues(push);

  // REVIEWから削除
  deleteRows.sort((a, b) => b - a);
  deleteRows.forEach(r => {
    if (r === 2) {
      review.getRange(2, 1, 1, 20).clearContent();
    } else {
      review.deleteRow(r);
    }
  });

  setFormulasToReview_(review);

  ui.alert(`${push.length}件 却下しました ✅\n理由：${reason}`);
}

// ============================
// REVIEW 数式注入（U〜Z）
// ============================
function setFormulasToReview_(sheet) {
  const formulas = [[
    '=ARRAYFORMULA(IF(A2:A="","",IF(O2:O<>"",O2:O,D2:D)))', // U 最終_取引先名
    '=ARRAYFORMULA(IF(A2:A="","",IF(P2:P<>"",P2:P,E2:E)))', // V 最終_契約種別
    '=ARRAYFORMULA(IF(A2:A="","",IF(Q2:Q<>"",Q2:Q,F2:F)))', // W 最終_現場名
    '=ARRAYFORMULA(IF(A2:A="","",IF(R2:R<>"",R2:R,G2:G)))', // X 最終_職人名
    '=ARRAYFORMULA(IF(A2:A="","",IF(S2:S<>"",S2:S,H2:H)))', // Y 最終_人工
    '=ARRAYFORMULA(IF(A2:A="","",IF(T2:T<>"",T2:T,I2:I)))'  // Z 最終_残業H
  ]];

  sheet.getRange("U2:Z2").setFormulas(formulas);
}

// ============================
// Ops取得
// ============================
function getOpsSheet_() {
  const opsSS = SpreadsheetApp.openById(OPS_SS_ID);
  const ops = opsSS.getSheetByName(OPS_SHEET_NAME);
  if (!ops) throw new Error(`Opsにシート「${OPS_SHEET_NAME}」がありません`);
  return ops;
}

// ============================
// Opsヘッダー自動判定（列ズレ耐性）
// ============================
function buildHeaderIndexMap_(opsSheet) {
  const header = opsSheet.getRange(1, 1, 1, opsSheet.getLastColumn()).getValues()[0]
    .map(v => String(v || "").trim());

  const find = (candidates) => {
    for (let i = 0; i < header.length; i++) {
      const h = header[i];
      if (!h) continue;
      if (candidates.includes(h)) return i;
    }
    return -1;
  };

  const map = {
    管理ID: find(["管理ID", "ID"]),
    日付: find(["日付"]),
    年月: find(["年月"]),
    取引先名: find(["取引先名", "元請", "元請名"]),
    契約種別: find(["契約種別（常用/請負）", "契約種別", "常用/請負"]),
    現場名: find(["現場名", "現場"]),
    職人名: find(["職人名", "職人"]),
    人工: find(["人工"]),
    残業H: find(["残業H", "残業"]),
    元メッセージID: find(["元メッセージID", "messageId", "messageid"]),
    登録日時: find(["登録日時", "受信時刻", "受信日時"])
  };

  if (map.管理ID === -1) throw new Error("Opsのヘッダーに「管理ID」が見つかりません。");
  if (map.日付 === -1) throw new Error("Opsのヘッダーに「日付」が見つかりません。");
  if (map.年月 === -1) throw new Error("Opsのヘッダーに「年月」が見つかりません。");

  return map;
}

// ============================
// Ops1行 → REVIEW A〜K(11列)
// ============================
function extractOpsRow11_(r, idx) {
  const safe = (i) => (i === -1 ? "" : r[i]);
  return [
    safe(idx.管理ID),
    safe(idx.日付),
    safe(idx.年月),
    safe(idx.取引先名),
    safe(idx.契約種別),
    safe(idx.現場名),
    safe(idx.職人名),
    safe(idx.人工),
    safe(idx.残業H),
    safe(idx.元メッセージID),
    safe(idx.登録日時)
  ];
}

// ============================
// 最終行（A列基準）
// ============================
function getRealLastRow_(sheet) {
  const vals = sheet.getRange("A:A").getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== "" && vals[i][0] != null) return i + 1;
  }
  return 0;
}

// ============================
// ⚠️ 緊急リセット（REVIEW/APPROVED/REJECTED全消し）
// ============================
function emergencyReset() {
  const { review, approved, rejected } = ensureSheets_();
  const ui = SpreadsheetApp.getUi();

  const ok = ui.alert("⚠️ 警告", "REVIEW / APPROVED / REJECTED のデータを全消去しますか？", ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  // REVIEW
  if (review.getFilter()) review.getFilter().remove();
  let maxRow = review.getLastRow();
  if (maxRow >= 3) review.deleteRows(3, maxRow - 2);
  if (maxRow >= 2) review.getRange(2, 1, 1, 20).clearContent();
  setFormulasToReview_(review);

  // APPROVED
  if (approved.getFilter()) approved.getFilter().remove();
  maxRow = approved.getLastRow();
  if (maxRow >= 2) approved.getRange(2, 1, maxRow - 1, approved.getLastColumn()).clearContent();

  // REJECTED
  if (rejected.getFilter()) rejected.getFilter().remove();
  maxRow = rejected.getLastRow();
  if (maxRow >= 2) rejected.getRange(2, 1, maxRow - 1, rejected.getLastColumn()).clearContent();

  ui.alert("🗑️ リセット完了（REVIEW数式も再注入済み）");
}
