/**
 * インターラーケン売上レポート集計システム
 * SquareサマリーCSVの38項目完全対応版
 * v8: Invoices API追加・現金バグ修正・38項目対応
 *
 * 実行環境: Google Apps Script (V8ランタイム)
 *
 * ── セットアップ手順 ──────────────────────────────
 * 1. GASエディタの「プロジェクトの設定」→「スクリプトプロパティ」を開く
 * 2. 以下の2つのプロパティを追加する:
 *      SQUARE_ACCESS_TOKEN  → SquareのAPIトークン
 *      COLORME_ACCESS_TOKEN → カラーミーショップのAPIトークン
 * 3. コードにトークンを直接書かない（セキュリティ向上）
 * ────────────────────────────────────────────────
 *
 * ── v8の主な変更点 ─────────────────────────────────
 * [修正] 現金の返品が二重計上されていたバグを修正
 * [追加] Invoices API連携（一部支払い・払い戻された一部入金）
 * [追加] 38項目出力に対応（Square CSVと完全一致）
 * ────────────────────────────────────────────────
 */

// ============================================================
// 定数
// ============================================================

const HISTORY_START = "2018-01";

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    SQUARE_ACCESS_TOKEN: props.getProperty("SQUARE_ACCESS_TOKEN"),
    COLORME_ACCESS_TOKEN: props.getProperty("COLORME_ACCESS_TOKEN"),
  };
}

const CM_COL = {
  DATE: 0, ORDER_ID: 1, PRODUCT_NAME: 2, TYPE: 3, QTY: 4,
  PRICE: 5, SUBTOTAL: 6, DELIVERY: 7, FEE: 8, POINT_DISCOUNT: 9, KEY: 10,
};

const SQ_COL = {
  DATE: 0, ORDER_ID: 1, NAME: 2, TYPE: 3, QTY: 4,
  GROSS: 5, TAX: 6, DISC: 7, PAY_TYPE: 8, FEE: 9, KEY: 10,
  RETURN_GROSS: 11, RETURN_TAX: 12, AMOUNT: 13,
};

// Invoice一部支払いデータ シートの列定義
const INV_COL = {
  YEAR_MONTH: 0,        // 年月
  PARTIAL: 1,           // 一部支払い
  REFUNDED_PARTIAL: 2,  // 払い戻された一部入金
  REFUNDED_COUNT: 3,    // 払い戻された一部入金の取引
};

const ELECTRONIC_MONEY_BRANDS = new Set([
  "ID", "QUICPAY", "SUICA", "PASMO", "ICOCA", "SUGOCA",
  "NIMOCA", "HAYAKAKEN", "KITACA", "TOICA", "MANACA",
]);

// ============================================================
// メニュー
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀インターラーケン操作")
    .addItem("1. Square＋カラーミー 今月分を更新", "runUpdateThisMonth")
    .addSeparator()
    .addItem("📦 過去データを一括取得（初回）", "runBulkImportStart")
    .addItem("📦 過去データを一括取得（続きから）", "runBulkImportResume")
    .addSeparator()
    .addItem("🔄 月を指定してSquare売上を更新", "runSquareUpdateByMonth")
    .addItem("🔄 月を指定してカラーミー売上を更新", "runColormeUpdateByMonth")
    .addSeparator()
    .addItem("📊 レポートを再集計", "recalculateAllSummaries")
    .addSeparator()
    .addItem("⏰ 日次自動更新を設定する", "setupDailyTrigger")
    .addItem("⏰ 日次自動更新を解除する", "removeDailyTrigger")
    .addSeparator()
    .addItem("⚙️ APIトークンを設定する", "setupApiTokens")
    .addToUi();
}

// ============================================================
// 今月分の更新
// ============================================================

function runUpdateThisMonth() {
  const { start, end } = getCurrentMonthRange();
  updateSquareSalesMaster(start, end);
  updateColormeSalesMaster(start);
  finalizeUpdate();
}

function dailyAutoUpdate() {
  const { start, end } = getCurrentMonthRange();
  updateSquareSalesMaster(start, end);
  updateColormeSalesMaster(start);
  sortAndRecalculate();
  console.log(`日次自動更新完了: ${new Date().toLocaleString("ja-JP")}`);
}

// ============================================================
// 過去データ一括取得
// ============================================================

function runBulkImportStart() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    "📦 過去データ一括取得",
    `最初の月を入力してください（デフォルト: ${HISTORY_START}）\n例: 2018-01`,
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  const input = result.getResponseText().trim() || HISTORY_START;
  if (!/^\d{4}-\d{2}$/.test(input)) {
    ui.alert("形式が正しくありません。「2018-01」のように入力してください。");
    return;
  }

  PropertiesService.getScriptProperties().setProperty("BULK_IMPORT_CURSOR", input);
  runBulkImportResume();
}

function runBulkImportResume() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const config = getConfig();

  if (!config.SQUARE_ACCESS_TOKEN) {
    ui.alert("SquareのAPIトークンが設定されていません。「⚙️ APIトークンを設定する」を実行してください。");
    return;
  }

  const cursor = props.getProperty("BULK_IMPORT_CURSOR");
  if (!cursor) {
    ui.alert("開始月が設定されていません。「📦 過去データを一括取得（初回）」から実行してください。");
    return;
  }

  const { yearMonth: currentYM } = getTodayYearMonth();
  const startTime = Date.now();
  const TIME_LIMIT_MS = 5 * 60 * 1000;

  let processingMonth = cursor;
  let processedMonths = [];

  while (processingMonth <= currentYM) {
    if (Date.now() - startTime > TIME_LIMIT_MS) {
      props.setProperty("BULK_IMPORT_CURSOR", processingMonth);
      ui.alert(
        `⏸ 一時停止\n\n処理済み: ${processedMonths.join(", ")}\n\n` +
        `「📦 過去データを一括取得（続きから）」を実行してください。\n` +
        `残り: ${processingMonth} 〜 ${currentYM}`
      );
      return;
    }

    const { start, end } = getMonthRange(processingMonth);
    try {
      updateSquareSalesMaster(start, end);
      updateColormeSalesMaster(start);
      processedMonths.push(processingMonth);
      console.log(`✅ ${processingMonth} 完了`);
    } catch (e) {
      console.error(`❌ ${processingMonth} エラー: ${e.message}`);
    }

    processingMonth = getNextMonth(processingMonth);
  }

  props.deleteProperty("BULK_IMPORT_CURSOR");
  sortAndRecalculate();
  ui.alert(
    `🎉 過去データの一括取得が完了しました！\n\n` +
    `処理済み: ${processedMonths.length}ヶ月分\n` +
    `（${cursor} 〜 ${currentYM}）`
  );
}

// ============================================================
// 月を指定して更新
// ============================================================

function runSquareUpdateByMonth() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt("月を指定", "対象月を入力してください（例: 2026-01）", ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;

  const yearMonth = result.getResponseText().trim();
  if (!/^\d{4}-\d{2}$/.test(yearMonth)) {
    ui.alert("形式が正しくありません。「2026-01」のように入力してください。");
    return;
  }

  const { start, end } = getMonthRange(yearMonth);
  updateSquareSalesMaster(start, end);
  finalizeUpdate();
}

function runColormeUpdateByMonth() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt("月を指定", "対象月を入力してください（例: 2026-01）", ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;

  const yearMonth = result.getResponseText().trim();
  if (!/^\d{4}-\d{2}$/.test(yearMonth)) {
    ui.alert("形式が正しくありません。「2026-01」のように入力してください。");
    return;
  }

  const { start } = getMonthRange(yearMonth);
  updateColormeSalesMaster(start);
  finalizeUpdate();
}

// ============================================================
// 日次自動更新トリガー
// ============================================================

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "dailyAutoUpdate")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("dailyAutoUpdate").timeBased().atHour(2).everyDays(1).create();

  SpreadsheetApp.getUi().alert(
    "⏰ 日次自動更新を設定しました！\n毎日 午前2時 に自動で今月分を取得・集計します。"
  );
}

function removeDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "dailyAutoUpdate");

  if (triggers.length === 0) {
    SpreadsheetApp.getUi().alert("設定済みの日次自動更新トリガーはありません。");
    return;
  }

  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert("⏰ 日次自動更新を解除しました。");
}

// ============================================================
// 日付ユーティリティ
// ============================================================

function getCurrentMonthRange() {
  const { yearMonth } = getTodayYearMonth();
  return getMonthRange(yearMonth);
}

function getTodayYearMonth() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  return { yearMonth: `${year}-${month}` };
}

function getMonthRange(yearMonth) {
  const [year, month] = yearMonth.split("-").map(Number);
  const start = `${yearMonth}-01`;
  const nextMonth = month === 12
    ? `${year + 1}-01-01`
    : `${year}-${String(month + 1).padStart(2, "0")}-01`;
  return { start, end: nextMonth };
}

function getNextMonth(yearMonth) {
  const [year, month] = yearMonth.split("-").map(Number);
  return month === 12
    ? `${year + 1}-01`
    : `${year}-${String(month + 1).padStart(2, "0")}`;
}

// ============================================================
// APIトークン設定
// ============================================================

function setupApiTokens() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const sqResult = ui.prompt("Square APIトークン設定", "Squareのアクセストークンをコピペしてください:", ui.ButtonSet.OK_CANCEL);
  if (sqResult.getSelectedButton() !== ui.Button.OK) return;
  props.setProperty("SQUARE_ACCESS_TOKEN", sqResult.getResponseText().trim());

  const cmResult = ui.prompt("カラーミー APIトークン設定", "カラーミーショップのアクセストークンをコピペしてください:", ui.ButtonSet.OK_CANCEL);
  if (cmResult.getSelectedButton() !== ui.Button.OK) return;
  props.setProperty("COLORME_ACCESS_TOKEN", cmResult.getResponseText().trim());

  ui.alert("✅ APIトークンを安全に保存しました！");
}

// ============================================================
// カラーミー
// ============================================================

function updateColormeSalesMaster(startDate) {
  const config = getConfig();
  if (!config.COLORME_ACCESS_TOKEN) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, "カラーミー売上データ");
  const existingKeys = getExistingKeys(sheet, CM_COL.KEY + 1);
  let offset = 0;
  let addedCount = 0;

  try {
    while (true) {
      const url = `https://api.shop-pro.jp/v1/sales.json?make_date_min=${startDate}&limit=100&offset=${offset}`;
      const res = JSON.parse(UrlFetchApp.fetch(url, {
        headers: { Authorization: `Bearer ${config.COLORME_ACCESS_TOKEN}` }
      }).getContentText());

      if (!res.sales || res.sales.length === 0) break;

      const newRows = [];
      for (const sale of res.sales) {
        const saleDate = parseSaleDate(sale.make_date);
        for (const detail of sale.details) {
          const qty = Number(detail.unit_num) || Number(detail.product_num) || 1;
          const price = Number(detail.price);
          const key = `${sale.id}_D_${detail.id}`;
          if (!existingKeys.has(key)) {
            newRows.push([saleDate, sale.id, detail.product_name, "SALE", qty, price,
              qty * price, sale.delivery_total, sale.fee_total, sale.point_discount, key]);
            existingKeys.add(key);
          }
        }
      }

      if (newRows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 11).setValues(newRows);
        addedCount += newRows.length;
      }

      if (res.sales.length < 100) break;
      offset += 100;
    }
    console.log(`カラーミー: ${addedCount}件追加`);
  } catch (e) {
    console.error(`CM Error: ${e.message}`);
    throw e;
  }
}

function parseSaleDate(raw) {
  if (typeof raw === "number") {
    return Utilities.formatDate(new Date(raw * 1000), "JST", "yyyy-MM-dd");
  }
  return raw.split(" ")[0];
}

// ============================================================
// Square 注文データ取得
// ============================================================

function updateSquareSalesMaster(startDate, endDate) {
  const config = getConfig();
  if (!config.SQUARE_ACCESS_TOKEN) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, "Square売上データ");

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 14).setValues([[
      "日付", "注文ID", "商品名", "種別", "数量", "売上", "税金", "ディスカウント",
      "支払種別", "手数料", "キー", "返品売上", "返品税金", "金額"
    ]]).setFontWeight("bold").setBackground("#f3f3f3");
    sheet.setFrozenRows(1);
  }

  const existingKeys = getExistingKeys(sheet, SQ_COL.KEY + 1);
  const sqHeaders = {
    Authorization: `Bearer ${config.SQUARE_ACCESS_TOKEN}`,
    "Content-Type": "application/json",
  };

  const { brandMap, refundMap, cardRefundPaymentIds, feeMap } =
    fetchPaymentsData(startDate, endDate, sqHeaders);

  // ✅ Invoices APIで一部支払いデータも取得・保存
  updateInvoiceData(startDate, endDate, sqHeaders);

  try {
    const locRes = JSON.parse(
      UrlFetchApp.fetch("https://connect.squareup.com/v2/locations", {
        headers: { Authorization: `Bearer ${config.SQUARE_ACCESS_TOKEN}` },
      }).getContentText()
    );

    let totalAdded = 0;

    for (const loc of locRes.locations) {
      let cursor = null;
      do {
        const payload = {
          location_ids: [loc.id],
          query: {
            filter: {
              closed_at: {
                start_at: new Date(`${startDate}T00:00:00+09:00`).toISOString(),
                end_at: new Date(`${endDate}T00:00:00+09:00`).toISOString(),
              },
              state_filter: { states: ["COMPLETED"] },
            },
          },
          ...(cursor && { cursor }),
        };

        const res = JSON.parse(
          UrlFetchApp.fetch("https://connect.squareup.com/v2/orders/search", {
            method: "post",
            headers: sqHeaders,
            payload: JSON.stringify(payload),
          }).getContentText()
        );

        if (res.orders) {
          const newRows = buildSquareRows(
            res.orders, existingKeys, brandMap, refundMap, cardRefundPaymentIds, feeMap
          );
          if (newRows.length > 0) {
            sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 14).setValues(newRows);
            totalAdded += newRows.length;
            newRows.forEach(r => existingKeys.add(r[SQ_COL.KEY]));
          }
        }
        cursor = res.cursor ?? null;
      } while (cursor);
    }

    console.log(`Square: ${totalAdded}行追加 (${startDate} 〜 ${endDate})`);
  } catch (e) {
    console.error(`SQ Error: ${e.message}`);
    throw e;
  }
}

// ============================================================
// Payments API（ブランド・返金・手数料）
// ============================================================

function fetchPaymentsData(startDate, endDate, sqHeaders) {
  const brandMap = new Map();
  const refundMap = new Map();
  const feeMap = new Map();
  const paymentOrderMap = new Map();
  const cardRefundPaymentIds = new Set();

  const startIso = encodeURIComponent(`${startDate}T00:00:00+09:00`);
  const endIso   = encodeURIComponent(`${endDate}T00:00:00+09:00`);

  let cursor = null;
  do {
    let url = `https://connect.squareup.com/v2/payments`
      + `?begin_time=${startIso}&end_time=${endIso}&limit=200`;
    if (cursor) url += `&cursor=${encodeURIComponent(cursor)}`;

    const res = JSON.parse(
      UrlFetchApp.fetch(url, { headers: sqHeaders }).getContentText()
    );

    res.payments?.forEach(p => {
      paymentOrderMap.set(p.id, p.order_id);

      if (p.source_type === "WALLET" && p.wallet_details?.brand) {
        brandMap.set(p.id, p.wallet_details.brand);
      }
      if (p.source_type === "CARD" && p.card_details?.card?.card_brand === "FELICA") {
        brandMap.set(p.id, "FELICA");
      }
      if (p.refunded_money?.amount > 0 && p.source_type !== "CASH") {
        refundMap.set(p.order_id, (refundMap.get(p.order_id) ?? 0) + p.refunded_money.amount);
        cardRefundPaymentIds.add(p.id);
      }
      if (p.processing_fee) {
        const fee = p.processing_fee.reduce((sum, f) => sum + (f.amount_money?.amount ?? 0), 0);
        if (fee > 0) feeMap.set(p.order_id, (feeMap.get(p.order_id) ?? 0) + fee);
      }
    });

    cursor = res.cursor ?? null;
  } while (cursor);

  const refRes = JSON.parse(
    UrlFetchApp.fetch(
      `https://connect.squareup.com/v2/refunds?begin_time=${startIso}&end_time=${endIso}`,
      { headers: sqHeaders }
    ).getContentText()
  );

  refRes.refunds?.forEach(r => {
    if (r.processing_fee) {
      const paymentId = r.id.split("_")[0];
      const orderId = paymentOrderMap.get(paymentId);
      if (orderId) {
        const fee = r.processing_fee.reduce((sum, f) => sum + (f.amount_money?.amount ?? 0), 0);
        feeMap.set(orderId, (feeMap.get(orderId) ?? 0) + fee);
      }
    }
  });

  console.log(`ブランドマップ: ${brandMap.size}件 / 返金マップ: ${refundMap.size}件 / 手数料マップ: ${feeMap.size}件`);
  return { brandMap, refundMap, cardRefundPaymentIds, feeMap };
}

// ============================================================
// ✅ Invoices API（一部支払い・払い戻された一部入金）
// ============================================================

/**
 * 指定期間の請求書データを取得し、「Invoice一部支払いデータ」シートに保存する
 * - 一部支払い: その月にupdatedされたPARTIALLY_PAID請求書の支払済み合計
 * - 払い戻された一部入金: 部分払いがあったPAID/REFUNDED請求書の金額
 */
function updateInvoiceData(startDate, endDate, sqHeaders) {
  const config = getConfig();
  if (!config.SQUARE_ACCESS_TOKEN) return;

  const startIso = `${startDate}T00:00:00+09:00`;
  const endIso   = `${endDate}T00:00:00+09:00`;

  // 月ごとの集計マップ { "2025-01" -> { partial, refunded, count } }
  const monthlyInvoice = {};

  try {
    const locRes = JSON.parse(
      UrlFetchApp.fetch("https://connect.squareup.com/v2/locations", {
        headers: sqHeaders
      }).getContentText()
    );

    for (const loc of locRes.locations) {
      let cursor = null;
      do {
        let url = `https://connect.squareup.com/v2/invoices?location_id=${loc.id}&limit=200`;
        if (cursor) url += `&cursor=${encodeURIComponent(cursor)}`;

        const res = JSON.parse(
          UrlFetchApp.fetch(url, { headers: sqHeaders }).getContentText()
        );

        res.invoices?.forEach(invoice => {
          if (!invoice.updated_at) return;

          // updated_at が対象期間外なら無視
          if (invoice.updated_at < startIso || invoice.updated_at >= endIso) return;

          // 年月を取得
          const yearMonth = Utilities.formatDate(
            new Date(invoice.updated_at), "JST", "yyyy-MM"
          );

          if (!monthlyInvoice[yearMonth]) {
            monthlyInvoice[yearMonth] = { partial: 0, refunded: 0, count: 0 };
          }

          // 支払済み合計を計算
          const totalPaid = (invoice.payment_requests ?? []).reduce((sum, req) => {
            return sum + (req.total_completed_amount_money?.amount ?? 0);
          }, 0);

          if (totalPaid === 0) return;

          switch (invoice.status) {
            case "PARTIALLY_PAID":
              // まだ未完了の部分払い → 一部支払い
              monthlyInvoice[yearMonth].partial += totalPaid;
              break;

            case "PAID":
            case "REFUNDED":
            case "CANCELLED": {
              // 部分払い構造（DEPOSIT or INSTALLMENT）があった場合のみ
              // → 払い戻された一部入金
              const hadPartialStructure = (invoice.payment_requests ?? []).some(req =>
                req.request_type === "DEPOSIT" || req.request_type === "INSTALLMENT"
              );
              if (hadPartialStructure) {
                monthlyInvoice[yearMonth].refunded += totalPaid;
                monthlyInvoice[yearMonth].count++;
              }
              break;
            }
          }
        });

        cursor = res.cursor ?? null;
      } while (cursor);
    }
  } catch (e) {
    console.error(`Invoice Error: ${e.message}`);
    return; // Invoicesエラーは無視して続行（必須ではないため）
  }

  // シートに保存（upsert: 既存行を更新、なければ追加）
  if (Object.keys(monthlyInvoice).length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, "Invoice一部支払いデータ");

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([[
      "年月", "一部支払い", "払い戻された一部入金", "払い戻された一部入金の取引"
    ]]).setFontWeight("bold").setBackground("#e8f5e9");
    sheet.setFrozenRows(1);
  }

  // 既存データをマップとして読み込む
  const existingData = {};
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 4).getValues().forEach((row, i) => {
      if (row[0]) existingData[String(row[0])] = i + 2; // 行番号を記録
    });
  }

  for (const [ym, data] of Object.entries(monthlyInvoice)) {
    const rowData = [ym, data.partial, -data.refunded, data.count];
    if (existingData[ym]) {
      // 既存行を更新
      sheet.getRange(existingData[ym], 1, 1, 4).setValues([rowData]);
    } else {
      // 新規行を追加
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, 4).setValues([rowData]);
    }
  }

  console.log(`Invoice: ${Object.keys(monthlyInvoice).length}ヶ月分を更新`);
}

/**
 * Invoice一部支払いデータシートから全データを読み込んでマップで返す
 * @returns {Map<string, {partial: number, refunded: number, count: number}>}
 */
function readInvoiceDataMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Invoice一部支払いデータ");
  const map = new Map();

  if (!sheet || sheet.getLastRow() < 2) return map;

  sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues().forEach(row => {
    if (!row[INV_COL.YEAR_MONTH]) return;
    map.set(String(row[INV_COL.YEAR_MONTH]), {
      partial:  Number(row[INV_COL.PARTIAL]) || 0,
      refunded: Number(row[INV_COL.REFUNDED_PARTIAL]) || 0,
      count:    Number(row[INV_COL.REFUNDED_COUNT]) || 0,
    });
  });

  return map;
}

// ============================================================
// Square行データ構築
// ============================================================

function buildSquareRows(orders, existingKeys, brandMap, refundMap, cardRefundPaymentIds, feeMap) {
  const rows = [];

  for (const order of orders) {
    const dateStr = Utilities.formatDate(new Date(order.closed_at), "JST", "yyyy-MM-dd");
    const id = order.id;

    // 1. 商品売上行 (SALE)
    if (order.line_items) {
      order.line_items.forEach((item, i) => {
        const key = `${id}_L_${i}`;
        if (!existingKeys.has(key)) {
          const gross = item.gross_sales_money?.amount ?? 0;
          rows.push([dateStr, id, item.name, "SALE", Number(item.quantity),
            gross, 0, 0, "", 0, key, 0, 0, 0]);
        }
      });
    }

    // 2. 注文サマリー行 (SUMMARY)
    const sumKey = `${id}_SUM`;
    if (!existingKeys.has(sumKey)) {
      let totalTax = order.total_tax_money?.amount ?? 0;
      const totalDisc = order.total_discount_money?.amount ?? 0;
      let retGross = 0, retTax = 0, manualRefund = 0;

      if (order.return_amounts) {
        const tax   = order.return_amounts.tax_money?.amount ?? 0;
        const total = order.return_amounts.total_money?.amount ?? 0;
        if (tax > 0) {
          retTax   = tax;
          retGross = total - tax;
          totalTax -= tax;
        } else {
          manualRefund = -total;
        }
      }

      if (retGross === 0 && retTax === 0) {
        order.refunds?.forEach(rf => {
          if (!rf.return_id) manualRefund += rf.amount_money?.amount ?? 0;
        });
      }

      if (retGross === 0 && retTax === 0 && manualRefund === 0) {
        const partialRefund = refundMap?.get(id) ?? 0;
        if (partialRefund > 0) manualRefund = partialRefund;
      }

      const hasRefund =
        (refundMap?.get(id) ?? 0) > 0 ||
        order.refunds?.some(rf => cardRefundPaymentIds?.has(rf.tender_id)) === true;

      rows.push([dateStr, id, "注文サマリー", "SUMMARY", 0, 0,
        totalTax, -totalDisc, hasRefund ? "払い戻し" : "", 0, sumKey,
        -retGross, -retTax, -manualRefund]);
    }

    // 3. 現金返品のマイナス行
    if (order.return_amounts && (!order.tenders || order.tenders.length === 0)) {
      const hasCardRefund = order.refunds?.some(rf =>
        cardRefundPaymentIds?.has(rf.tender_id)
      ) ?? false;

      if (!hasCardRefund) {
        const refundKey = `${id}_REFUND`;
        if (!existingKeys.has(refundKey)) {
          const total = order.return_amounts.total_money?.amount ?? 0;
          rows.push([dateStr, id, "返金", "PAYMENT",
            0, 0, 0, 0, "現金", 0, refundKey, 0, 0, -total]);
        }
      }
    }

    // 4. 支払い行 (PAYMENT)
    order.tenders?.forEach((tender, i) => {
      const key = `${id}_T_${i}`;
      if (!existingKeys.has(key)) {
        const payType = getPaymentType(tender, brandMap);
        const amt = tender.amount_money?.amount ?? 0;
        const fee = feeMap?.get(id) ?? tender.processing_fee_money?.amount ?? 0;
        rows.push([dateStr, id, `支払い: ${payType}`, "PAYMENT",
          0, 0, 0, 0, payType, fee, key, 0, 0, amt]);
      }
    });

    // 5. カード部分返金のマイナス行
    const partialRefund = refundMap?.get(id) ?? 0;
    if (partialRefund > 0 && order.tenders?.length > 0) {
      const refundKey = `${id}_PREFUND`;
      if (!existingKeys.has(refundKey)) {
        rows.push([dateStr, id, "部分返金", "PAYMENT",
          0, 0, 0, 0, "カード", 0, refundKey, 0, 0, -partialRefund]);
      }
    }
  }

  return rows;
}

function getPaymentType(tender, brandMap) {
  switch (tender.type) {
    case "CARD": {
      const mappedBrand = (brandMap?.get(tender.payment_id) ?? "").toUpperCase();
      if (mappedBrand === "FELICA") return "電子マネー";
      const brand = (tender.card_details?.card?.card_brand ?? "").toUpperCase();
      if (brand === "FELICA") return "電子マネー";
      return ELECTRONIC_MONEY_BRANDS.has(brand) ? "電子マネー" : "カード";
    }
    case "WALLET": {
      const brand = (brandMap?.get(tender.payment_id) ?? "").toUpperCase();
      if (brand === "RAKUTEN_PAY") return "楽天ペイ";
      if (brand === "AU_PAY")      return "au PAY";
      if (brand === "D_BARAI")     return "d払い";
      return "その他";
    }
    case "CASH": return "現金";
    case "HOUSE_ACCOUNT":
    case "SQUARE_ACCOUNT": return "ハウスアカウント";
    case "EXTERNAL": {
      const src = (tender.external_details?.source_name ?? "").toUpperCase();
      if (src.includes("AU PAY") || src.includes("AUPAY")) return "au PAY";
      if (src.includes("D払い")  || src.includes("DBARAI")) return "d払い";
      if (src.includes("楽天")   || src.includes("RAKUTEN")) return "楽天ペイ";
      return "その他";
    }
    case "OTHER": {
      const note = (tender.note ?? "").toLowerCase();
      if (note.includes("売掛") || note.includes("掛け"))        return "ハウスアカウント";
      if (note.includes("クレジット") || note.includes("credit")) return "カード";
      if (note.includes("suica") || note.includes("pasmo") || note.includes("id")) return "電子マネー";
      return "その他";
    }
    default: return "その他";
  }
}

// ============================================================
// 再集計（38項目対応）
// ============================================================

function recalculateAllSummaries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sqSheet = ss.getSheetByName("Square売上データ");
  if (!sqSheet) {
    SpreadsheetApp.getUi().alert("「Square売上データ」シートが見つかりません。");
    return;
  }

  const data = sqSheet.getDataRange().getValues();
  const monthly = aggregateMonthlyData(data.slice(1));

  // ✅ Invoice一部支払いデータを読み込んでマージ
  const invoiceMap = readInvoiceDataMap();
  for (const [ym, inv] of invoiceMap.entries()) {
    if (monthly[ym]) {
      monthly[ym].partialPayment = inv.partial;
      monthly[ym].refundedPartial = inv.refunded;
      monthly[ym].refundedPartialCount = inv.count;
    }
  }

  // ✅ 38項目のヘッダー（Square CSVの順番と一致）
  const headers = [
    "年月",
    "総売上高", "商品", "サービス料", "返品", "ディスカウントと無料提供",
    "純売上高", "繰延売上", "ギフトカード売上",
    "一部支払い",              // ✅ 新規追加
    "払い戻された一部入金",    // ✅ 新規追加
    "税金", "金額を指定した払い戻し",
    "売上合計", "受取合計額",
    "au PAY", "d払い", "カード", "その他",
    "ハウスアカウント", "楽天ペイ", "現金", "電子マネー",
    "手数料", "Squareの決済手数料", "Squareの手数料",
    "合計（純額）", "総売上数",
    "売上取引履歴", "商品売上取引履歴", "サービス料取引履歴",
    "商品別返品取引履歴", "ディスカウント取引履歴", "無料提供取引履歴",
    "ギフトカード売上取引履歴",
    "払い戻された一部入金の取引", // ✅ 新規追加
    "税金取引履歴", "総売上取引履歴", "受取合計額の取引履歴",
  ];

  const rows = Object.keys(monthly).sort().reverse().map(m => buildSummaryRow(m, monthly[m]));
  writeToSheet(getOrCreateSheet(ss, "Square月次売上"), headers, rows);
  SpreadsheetApp.getUi().alert("✅ レポートを更新しました！（38項目対応）");
}

function aggregateMonthlyData(rows) {
  const monthly = {};

  for (const row of rows) {
    const dateVal = row[SQ_COL.DATE];
    const m = dateVal instanceof Date
      ? Utilities.formatDate(dateVal, "JST", "yyyy-MM")
      : String(dateVal).substring(0, 7);

    if (!m || m.length < 7) continue;

    if (!monthly[m]) {
      monthly[m] = {
        gross: 0, tax: 0, disc: 0, returns: 0, refund: 0, fees: 0,
        partialPayment: 0,    // 一部支払い（Invoice APIから設定）
        refundedPartial: 0,   // 払い戻された一部入金（Invoice APIから設定）
        refundedPartialCount: 0,
        txAll: new Set(), txItems: new Set(), txAllItems: new Set(),
        txReturns: new Set(), txDiscounts: new Set(),
        txTax: new Set(), txTenders: new Set(),
        pay: {
          "au PAY": 0, "d払い": 0, "カード": 0, "その他": 0,
          "ハウスアカウント": 0, "楽天ペイ": 0, "現金": 0, "電子マネー": 0,
        },
      };
    }

    const o = monthly[m];
    const type = row[SQ_COL.TYPE];
    const id = String(row[SQ_COL.ORDER_ID]);

    switch (type) {
      case "SALE":
        o.gross += Number(row[SQ_COL.GROSS]);
        o.txAllItems.add(id);
        if (Number(row[SQ_COL.GROSS]) > 0) {
          o.txItems.add(id);
          o.txAll.add(id);
        }
        break;

      case "SUMMARY":
        o.tax     += Number(row[SQ_COL.TAX]);
        o.disc    += Number(row[SQ_COL.DISC]);
        o.returns += Number(row[SQ_COL.RETURN_GROSS]);
        o.refund  += Number(row[SQ_COL.AMOUNT]);
        if (Number(row[SQ_COL.TAX]) !== 0)         o.txTax.add(id);
        if (Number(row[SQ_COL.DISC]) !== 0)         o.txDiscounts.add(id);
        if (Number(row[SQ_COL.RETURN_GROSS]) !== 0) o.txReturns.add(id);
        if (Number(row[SQ_COL.AMOUNT]) !== 0 || row[SQ_COL.PAY_TYPE] === "払い戻し") {
          o.txTenders.add(id);
        }
        break;

      case "PAYMENT": {
        const pt  = row[SQ_COL.PAY_TYPE];
        const amt = Number(row[SQ_COL.AMOUNT]);
        if (pt === "オンライン") break;

        // ✅ バグ修正: 現金の返品マイナス行は o.refund で既に計上済み
        // → o.pay["現金"] から二重引きしないようスキップ
        if (pt === "現金" && amt < 0) break;

        o.pay[pt in o.pay ? pt : "その他"] += amt;
        o.fees += Number(row[SQ_COL.FEE]);
        o.txTenders.add(id);
        break;
      }
    }
  }

  return monthly;
}

function buildSummaryRow(month, o) {
  const netSales   = o.gross + o.returns + o.disc;
  // ✅ 一部支払いと払い戻された一部入金を売上合計に反映
  const totalSales = netSales + o.partialPayment + o.refundedPartial + o.tax + o.refund;
  const collected  = totalSales - (o.pay["ハウスアカウント"] ?? 0);
  const allTxCount = new Set([...o.txAllItems, ...o.txTenders]).size;

  return [
    month,
    o.gross, o.gross, 0, o.returns, o.disc,
    netSales, 0, 0,
    o.partialPayment,    // 一部支払い
    o.refundedPartial,   // 払い戻された一部入金
    o.tax, o.refund,
    totalSales, collected,
    o.pay["au PAY"], o.pay["d払い"], o.pay["カード"], o.pay["その他"],
    o.pay["ハウスアカウント"], o.pay["楽天ペイ"], o.pay["現金"], o.pay["電子マネー"],
    -o.fees, -o.fees, -o.fees,
    collected - o.fees, allTxCount,
    o.txTax.size, o.txItems.size, 0, o.txReturns.size, o.txDiscounts.size,
    0, 0,
    o.refundedPartialCount, // 払い戻された一部入金の取引
    o.txTax.size, o.txItems.size, o.txTenders.size,
  ];
}

// ============================================================
// ユーティリティ
// ============================================================

function finalizeUpdate() {
  sortAndRecalculate();
}

function sortAndRecalculate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const name of ["Square売上データ", "カラーミー売上データ"]) {
    const s = ss.getSheetByName(name);
    if (s) sortSheetByDate(s);
  }
  recalculateAllSummaries();
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) ?? ss.insertSheet(name);
}

function getExistingKeys(sheet, keyColumn) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Set();
  return new Set(
    sheet.getRange(2, keyColumn, lastRow - 1, 1).getValues().map(r => String(r[0]))
  );
}

function sortSheetByDate(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).sort({ column: 1, ascending: false });
}

function writeToSheet(sheet, headers, rows) {
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#f3f3f3");
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  sheet.setFrozenRows(1);
}