/**
 * インターラーケン売上レポート集計システム
 * SquareサマリーCSVの38項目に完全準拠
 * v7.5: 年月列テキスト固定版 🔥
 *
 * 実行環境: Google Apps Script (V8ランタイム)
 *
 * 使い方:
 *   「Square設定」シートのA2（開始日）・B2（終了日）に期間を入力して
 *   メニュー「Square売上を更新」を押すだけ！
 *
 * 対応済み項目（38項目）:
 * - WALLET支払い種別（au PAY / d払い / 楽天ペイ）
 * - FELICA（電子マネー）
 * - 現金返品のマイナス処理
 * - カード部分返金のマイナス処理
 * - 返金手数料の戻り処理（期間外の元支払いも個別API取得）
 * - 期間外の元支払いのsource_type取得
 * - CANCELEDの注文除外
 * - EXTERNAL返金のその他分類
 * - ゼロ円取引の取引履歴カウント
 * - シート新規作成時のヘッダー行保護
 * - 一部支払い / 払い戻された一部入金 / 払い戻された一部入金の取引（NEW）
 */

// ============================================================
// 定数
// ============================================================

const CONFIG = {
  get SQUARE_ACCESS_TOKEN() {
    return PropertiesService.getScriptProperties().getProperty(
      "SQUARE_ACCESS_TOKEN",
    );
  },
  get COLORME_ACCESS_TOKEN() {
    return PropertiesService.getScriptProperties().getProperty(
      "COLORME_ACCESS_TOKEN",
    );
  },
};

const CM_COL = {
  DATE: 0,
  ORDER_ID: 1,
  PRODUCT_NAME: 2,
  TYPE: 3,
  QTY: 4,
  PRICE: 5,
  SUBTOTAL: 6,
  DELIVERY: 7,
  FEE: 8,
  POINT_DISCOUNT: 9,
  KEY: 10,
};

const SQ_COL = {
  DATE: 0,
  ORDER_ID: 1,
  NAME: 2,
  TYPE: 3,
  QTY: 4,
  GROSS: 5,
  TAX: 6,
  DISC: 7,
  PAY_TYPE: 8,
  FEE: 9,
  KEY: 10,
  RETURN_GROSS: 11,
  RETURN_TAX: 12,
  AMOUNT: 13,
};

const SQ_HEADERS_KEY = [
  "日付",
  "注文ID",
  "商品名",
  "種別",
  "数量",
  "売上",
  "税金",
  "割引",
  "支払種別",
  "手数料",
  "キー",
  "返品売上",
  "返品税金",
  "金額",
];

const ELECTRONIC_MONEY_BRANDS = new Set([
  "ID",
  "QUICPAY",
  "SUICA",
  "PASMO",
  "ICOCA",
  "SUGOCA",
  "NIMOCA",
  "HAYAKAKEN",
  "KITACA",
  "TOICA",
  "MANACA",
]);

// ============================================================
// メニュー
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀インターラーケン操作")
    .addItem("1. カラーミー売上を更新", "runColormeUpdate")
    .addItem("2. Square売上を更新", "runSquareUpdate")
    .addSeparator()
    .addItem("3. レポートを再集計", "recalculateAllSummaries")
    .addSeparator()
    .addItem("⚠️ Squareデータをクリア", "clearSquareData")
    .addToUi();
}

// ============================================================
// 設定シート
// ============================================================

/**
 * 「Square設定」シートからstart/endDateを取得
 * シートがなければ自動作成してサンプルを入力する
 */
function getSettingsRange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Square設定");

  if (!sheet) {
    sheet = ss.insertSheet("Square設定");
    sheet
      .getRange("A1:C1")
      .setValues([["開始日", "終了日", "取得済み月"]])
      .setFontWeight("bold")
      .setBackground("#f3f3f3");
    // デフォルト: 今月1日〜今日
    const now = new Date();
    const firstDay = Utilities.formatDate(
      new Date(now.getFullYear(), now.getMonth(), 1),
      "JST",
      "yyyy-MM-dd",
    );
    const today = Utilities.formatDate(now, "JST", "yyyy-MM-dd");
    sheet.getRange("A2:B2").setValues([[firstDay, today]]);
    sheet.setColumnWidth(1, 130);
    sheet.setColumnWidth(2, 130);
    SpreadsheetApp.getUi().alert(
      "「Square設定」シートを作成しました。\nA2に開始日、B2に終了日を入力してから再実行してください。",
    );
    return null;
  }

  const vals = sheet.getRange("A2:B2").getValues()[0];
  const startDate =
    vals[0] instanceof Date
      ? Utilities.formatDate(vals[0], "JST", "yyyy-MM-dd")
      : String(vals[0]).trim();
  const endDate =
    vals[1] instanceof Date
      ? Utilities.formatDate(vals[1], "JST", "yyyy-MM-dd")
      : String(vals[1]).trim();

  if (!startDate || !endDate || startDate.length < 10 || endDate.length < 10) {
    SpreadsheetApp.getUi().alert(
      "「Square設定」シートのA2に開始日、B2に終了日を\nyyyy-MM-dd形式で入力してください。",
    );
    return null;
  }

  return { startDate, endDate };
}

// ============================================================
// エントリーポイント
// ============================================================

function runColormeUpdate() {
  const settings = getSettingsRange();
  if (!settings) return;
  updateColormeSalesMaster(settings.startDate);
  finalizeUpdate();
}

function runSquareUpdate() {
  const settings = getSettingsRange();
  if (!settings) return;
  updateSquareRange(settings.startDate, settings.endDate);
  finalizeUpdate();
}

// ============================================================
// カラーミー
// ============================================================

function updateColormeSalesMaster(startDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, "カラーミー売上データ");
  const existingKeys = getExistingKeys(sheet, CM_COL.KEY + 1);
  let offset = 0;

  try {
    while (true) {
      const url = `https://api.shop-pro.jp/v1/sales.json?make_date_min=${startDate}&limit=100&offset=${offset}`;
      const res = JSON.parse(
        UrlFetchApp.fetch(url, {
          headers: { Authorization: `Bearer ${CONFIG.COLORME_ACCESS_TOKEN}` },
        }).getContentText(),
      );

      if (!res.sales || res.sales.length === 0) break;

      const newRows = [];
      for (const sale of res.sales) {
        const saleDate = parseSaleDate(sale.make_date);
        for (const detail of sale.details) {
          const qty =
            Number(detail.unit_num) || Number(detail.product_num) || 1;
          const price = Number(detail.price);
          const key = `${sale.id}_D_${detail.id}`;
          if (!existingKeys.has(key)) {
            newRows.push([
              saleDate,
              sale.id,
              detail.product_name,
              "SALE",
              qty,
              price,
              qty * price,
              sale.delivery_total,
              sale.fee_total,
              sale.point_discount,
              key,
            ]);
          }
        }
      }
      if (newRows.length > 0) {
        sheet
          .getRange(sheet.getLastRow() + 1, 1, newRows.length, 11)
          .setValues(newRows);
      }
      if (res.sales.length < 100) break;
      offset += 100;
    }
  } catch (e) {
    console.error(`CM Error: ${e.message}`);
  }
}

function parseSaleDate(raw) {
  if (typeof raw === "number") {
    return Utilities.formatDate(new Date(raw * 1000), "JST", "yyyy-MM-dd");
  }
  return raw.split(" ")[0];
}

// ============================================================
// Square: 期間指定取得
// ============================================================

/**
 * 指定期間を月ごとに分割してSquareデータを取得する
 * 例: startDate="2025-03-01", endDate="2026-02-28"
 *     → 2025-03, 2025-04, ... 2026-02 の12ヶ月を順番に取得
 */
function updateSquareRange(startDate, endDate) {
  const months = getMonthsInRange(startDate, endDate);
  console.log(
    `取得対象: ${months.length}ヶ月 (${months[0]} 〜 ${months[months.length - 1]})`,
  );

  // 設定シートから再開位置を取得（タイムアウト時の続きから再開）
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = ss.getSheetByName("Square設定");
  const lastDone = settingSheet
    ? String(settingSheet.getRange("C2").getValue()).trim()
    : "";
  let resume = !lastDone || lastDone === "完了";

  for (const month of months) {
    if (!resume) {
      if (month === lastDone) resume = true;
      continue;
    }
    const monthStart = month + "-01";
    const nextMonth = getNextMonthStart(month);
    const monthEnd = nextMonth > endDate + "X" ? addOneDay(endDate) : nextMonth;
    console.log(`=== ${month} 取得開始 ===`);
    updateSquareMonth(monthStart, monthEnd, month);

    // 完了した月を記録（途中でタイムアウトしても次回ここから再開）
    if (settingSheet) {
      settingSheet.getRange("C2").setValue(month);
      SpreadsheetApp.flush();
    }
  }

  // 全月完了したらリセット
  if (settingSheet) {
    settingSheet.getRange("C2").setValue("完了");
    SpreadsheetApp.getUi().alert("✅ Square売上データの取得が完了しました！");
  }
}

/** "yyyy-MM" の配列を返す */
function getMonthsInRange(startDate, endDate) {
  const months = [];
  const start = new Date(startDate + "T00:00:00+09:00");
  const end = new Date(endDate + "T00:00:00+09:00");
  let cur = new Date(start.getFullYear(), start.getMonth(), 1);
  while (cur <= end) {
    months.push(Utilities.formatDate(cur, "JST", "yyyy-MM"));
    cur.setMonth(cur.getMonth() + 1);
  }
  return months;
}

/** "yyyy-MM" → 翌月の "yyyy-MM-dd" */
function getNextMonthStart(month) {
  const [y, m] = month.split("-").map(Number);
  const d = new Date(y, m, 1); // JavaScriptのDateは0始まり月なので m=次月
  return Utilities.formatDate(d, "JST", "yyyy-MM-dd");
}

/** "yyyy-MM-dd" の翌日を返す */
function addOneDay(dateStr) {
  const d = new Date(dateStr + "T00:00:00+09:00");
  d.setDate(d.getDate() + 1);
  return Utilities.formatDate(d, "JST", "yyyy-MM-dd");
}

// ============================================================
// Square: 支払いデータ取得
// ============================================================

function fetchPaymentsData(startDate, endDate, sqHeaders) {
  const brandMap = new Map();
  const refundMap = new Map();
  const feeMap = new Map();
  const partialMap = new Map(); // NEW: order_id → 一部支払い合計
  const paymentOrderMap = new Map();
  const sourceTypeMap = new Map();
  const cardRefundPaymentIds = new Set();
  const externalRefundPaymentIds = new Set();
  let cursor = null;

  do {
    let url =
      `https://connect.squareup.com/v2/payments` +
      `?begin_time=${encodeURIComponent(startDate + "T00:00:00+09:00")}` +
      `&end_time=${encodeURIComponent(endDate + "T00:00:00+09:00")}` +
      `&limit=200`;
    if (cursor) url += `&cursor=${encodeURIComponent(cursor)}`;

    const res = JSON.parse(
      UrlFetchApp.fetch(url, { headers: sqHeaders }).getContentText(),
    );

    res.payments?.forEach((p) => {
      paymentOrderMap.set(p.id, p.order_id);
      sourceTypeMap.set(p.id, p.source_type);

      if (p.source_type === "WALLET" && p.wallet_details?.brand) {
        brandMap.set(p.id, p.wallet_details.brand);
      }
      if (
        p.source_type === "CARD" &&
        p.card_details?.card?.card_brand === "FELICA"
      ) {
        brandMap.set(p.id, "FELICA");
      }
      if (p.processing_fee) {
        const fee = p.processing_fee.reduce(
          (sum, f) => sum + (f.amount_money?.amount ?? 0),
          0,
        );
        if (fee > 0)
          feeMap.set(p.order_id, (feeMap.get(p.order_id) ?? 0) + fee);
      }

      // 一部支払い: approved_money < total_money のとき差額が「一部支払い」
      // Square仕様: total_money = 実際の注文合計, approved_money = 実際に受け取った金額
      if (p.approved_money && p.total_money) {
        const diff =
          (p.total_money.amount ?? 0) - (p.approved_money.amount ?? 0);
        if (diff > 0 && p.order_id) {
          partialMap.set(p.order_id, (partialMap.get(p.order_id) ?? 0) + diff);
        }
      }
    });

    cursor = res.cursor ?? null;
  } while (cursor);

  const refRes = JSON.parse(
    UrlFetchApp.fetch(
      `https://connect.squareup.com/v2/refunds` +
        `?begin_time=${encodeURIComponent(startDate + "T00:00:00+09:00")}` +
        `&end_time=${encodeURIComponent(endDate + "T00:00:00+09:00")}`,
      { headers: sqHeaders },
    ).getContentText(),
  );

  refRes.refunds?.forEach((r) => {
    // 返金手数料（期間外の元支払いは個別APIで取得）
    if (r.processing_fee) {
      let orderId = paymentOrderMap.get(r.payment_id);
      if (!orderId) {
        try {
          const pRes = JSON.parse(
            UrlFetchApp.fetch(
              `https://connect.squareup.com/v2/payments/${r.payment_id}`,
              { headers: sqHeaders, muteHttpExceptions: true },
            ).getContentText(),
          );
          orderId = pRes.payment?.order_id;
        } catch (e) {}
      }
      if (orderId) {
        const fee = r.processing_fee.reduce(
          (sum, f) => sum + (f.amount_money?.amount ?? 0),
          0,
        );
        feeMap.set(orderId, (feeMap.get(orderId) ?? 0) + fee);
      }
    }

    // 返金のsource_type取得（期間外の元支払いは個別APIで取得）
    if (r.amount_money?.amount > 0) {
      let sourceType = sourceTypeMap.get(r.payment_id);
      if (!sourceType) {
        try {
          const pRes = JSON.parse(
            UrlFetchApp.fetch(
              `https://connect.squareup.com/v2/payments/${r.payment_id}`,
              { headers: sqHeaders, muteHttpExceptions: true },
            ).getContentText(),
          );
          sourceType = pRes.payment?.source_type;
        } catch (e) {}
      }
      if (sourceType === "CARD") {
        refundMap.set(
          r.order_id,
          (refundMap.get(r.order_id) ?? 0) + r.amount_money.amount,
        );
        cardRefundPaymentIds.add(r.payment_id);
      } else if (sourceType === "EXTERNAL") {
        externalRefundPaymentIds.add(r.payment_id);
      }
    }
  });

  console.log(
    `ブランドマップ: ${brandMap.size}件 / 部分返金: ${refundMap.size}件 / 手数料: ${feeMap.size}件 / 一部支払い: ${partialMap.size}件`,
  );
  return {
    brandMap,
    refundMap,
    cardRefundPaymentIds,
    externalRefundPaymentIds,
    feeMap,
    partialMap,
  };
}

// ============================================================
// Square: 行データ生成
// ============================================================

function buildSquareRows(
  orders,
  existingKeys,
  brandMap,
  refundMap,
  cardRefundPaymentIds,
  externalRefundPaymentIds,
  feeMap,
  partialMap,
) {
  const rows = [];

  for (const order of orders) {
    const dateStr = Utilities.formatDate(
      new Date(order.closed_at),
      "JST",
      "yyyy-MM-dd",
    );
    const id = order.id;

    // 1. 商品売上行 (SALE)
    if (order.line_items) {
      order.line_items.forEach((item, i) => {
        const key = `${id}_L_${i}`;
        if (!existingKeys.has(key)) {
          const gross = item.gross_sales_money?.amount ?? 0;
          rows.push([
            dateStr,
            id,
            item.name,
            "SALE",
            Number(item.quantity),
            gross,
            0,
            0,
            "",
            0,
            key,
            0,
            0,
            0,
          ]);
        }
      });
    }

    // 2. 注文サマリー行 (SUMMARY)
    const sumKey = `${id}_SUM`;
    if (!existingKeys.has(sumKey)) {
      let totalTax = order.total_tax_money?.amount ?? 0;
      const totalDisc = order.total_discount_money?.amount ?? 0;
      let retGross = 0,
        retTax = 0,
        manualRefund = 0;

      if (order.return_amounts) {
        const tax = order.return_amounts.tax_money?.amount ?? 0;
        const total = order.return_amounts.total_money?.amount ?? 0;
        if (tax > 0) {
          retTax = tax;
          retGross = total - tax;
          totalTax -= tax;
        } else {
          const hasCatalogItem = order.returns?.some((r) =>
            r.return_line_items?.some((i) => i.catalog_object_id),
          );
          retGross = hasCatalogItem ? total : 0;
          manualRefund = hasCatalogItem ? 0 : total;
        }
      }

      // Payments APIの部分返金を反映（商品返品がない場合のみ）
      if (retGross === 0 && retTax === 0 && manualRefund === 0) {
        const partialRefund = refundMap?.get(id) ?? 0;
        if (partialRefund > 0) manualRefund = partialRefund;
      }

      // 払い戻しフラグ（受取合計額の取引履歴カウント用）
      const hasRefund =
        (refundMap?.get(id) ?? 0) > 0 ||
        order.refunds?.some((rf) => cardRefundPaymentIds?.has(rf.tender_id)) ===
          true;

      rows.push([
        dateStr,
        id,
        "注文サマリー",
        "SUMMARY",
        0,
        0,
        totalTax,
        -totalDisc,
        hasRefund ? "払い戻し" : "",
        0,
        sumKey,
        -retGross,
        -retTax,
        -manualRefund,
      ]);
    }

    // 3. 返品のマイナス行（tenderなし）
    if (
      order.return_amounts &&
      (!order.tenders || order.tenders.length === 0)
    ) {
      const refundKey = `${id}_REFUND`;
      if (!existingKeys.has(refundKey)) {
        const total = order.return_amounts.total_money?.amount ?? 0;
        const hasCardRefund =
          order.refunds?.some((rf) =>
            cardRefundPaymentIds?.has(rf.tender_id),
          ) ?? false;
        const hasExternalRefund =
          order.refunds?.some((rf) =>
            externalRefundPaymentIds?.has(rf.tender_id),
          ) ?? false;
        const payType = hasCardRefund
          ? "カード"
          : hasExternalRefund
            ? "その他"
            : "現金";
        rows.push([
          dateStr,
          id,
          "返金",
          "PAYMENT",
          0,
          0,
          0,
          0,
          payType,
          0,
          refundKey,
          0,
          0,
          -total,
        ]);
      }
    }

    // 4. 支払い行 (PAYMENT)
    order.tenders?.forEach((tender, i) => {
      const key = `${id}_T_${i}`;
      if (!existingKeys.has(key)) {
        const payType = getPaymentType(tender, brandMap);
        const amt = tender.amount_money?.amount ?? 0;
        const fee = feeMap?.get(id) ?? tender.processing_fee_money?.amount ?? 0;
        rows.push([
          dateStr,
          id,
          `支払い: ${payType}`,
          "PAYMENT",
          0,
          0,
          0,
          0,
          payType,
          fee,
          key,
          0,
          0,
          amt,
        ]);
      }
    });

    // 5. カード部分返金のマイナス行
    const partialRefund = refundMap?.get(id) ?? 0;
    if (partialRefund > 0 && order.tenders?.length > 0) {
      const refundKey = `${id}_PREFUND`;
      if (!existingKeys.has(refundKey)) {
        rows.push([
          dateStr,
          id,
          "部分返金",
          "PAYMENT",
          0,
          0,
          0,
          0,
          "カード",
          0,
          refundKey,
          0,
          0,
          -partialRefund,
        ]);
      }
    }

    // 6. 一部支払い行 (PARTIAL) ← NEW
    const partial = partialMap?.get(id) ?? 0;
    if (partial > 0) {
      const partialKey = `${id}_PARTIAL`;
      if (!existingKeys.has(partialKey)) {
        rows.push([
          dateStr,
          id,
          "一部支払い",
          "PARTIAL",
          0,
          0,
          0,
          0,
          "",
          0,
          partialKey,
          0,
          0,
          partial,
        ]);
      }
    }
  }

  return rows;
}

function getPaymentType(tender, brandMap) {
  switch (tender.type) {
    case "CARD": {
      // brandMapのキーはPayments APIのpayment_id = tender.payment_id または tender.id
      const tenderId = tender.payment_id ?? tender.id;
      const mappedBrand = (brandMap?.get(tenderId) ?? "").toUpperCase();
      if (mappedBrand === "FELICA") return "電子マネー";
      const brand = (tender.card_details?.card_brand ?? "").toUpperCase();
      if (brand === "FELICA") return "電子マネー";
      return ELECTRONIC_MONEY_BRANDS.has(brand) ? "電子マネー" : "カード";
    }
    case "WALLET": {
      const tenderId = tender.payment_id ?? tender.id;
      const brand = (brandMap?.get(tenderId) ?? "").toUpperCase();
      if (brand === "RAKUTEN_PAY") return "楽天ペイ";
      if (brand === "AU_PAY") return "au PAY";
      if (brand === "D_BARAI") return "d払い";
      return "その他";
    }
    case "CASH":
      return "現金";
    case "HOUSE_ACCOUNT":
    case "SQUARE_ACCOUNT":
      return "ハウスアカウント";
    case "EXTERNAL": {
      const src = (tender.external_details?.source_name ?? "").toUpperCase();
      if (src.includes("AU PAY") || src.includes("AUPAY")) return "au PAY";
      if (src.includes("D払い") || src.includes("DBARAI")) return "d払い";
      if (src.includes("楽天") || src.includes("RAKUTEN")) return "楽天ペイ";
      return "その他";
    }
    case "OTHER": {
      const note = (tender.note ?? "").toLowerCase();
      if (note.includes("オンライン") || note.includes("online"))
        return "その他";
      if (note.includes("代引")) return "その他";
      if (note.includes("pay pay") || note.includes("paypay")) return "その他";
      if (note.includes("売掛") || note.includes("掛け"))
        return "ハウスアカウント";
      if (note.includes("クレジット") || note.includes("credit"))
        return "カード";
      if (
        note.includes("suica") ||
        note.includes("pasmo") ||
        note.includes("id")
      )
        return "電子マネー";
      return "その他";
    }
    default:
      return "その他";
  }
}

// ============================================================
// 再集計
// ============================================================

function recalculateAllSummaries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sqSheet = ss.getSheetByName("Square売上データ");
  if (!sqSheet) return;

  const data = sqSheet.getDataRange().getValues();
  const monthly = aggregateMonthlyData(data.slice(1));

  const headers = [
    "年月",
    "総売上高",
    "商品",
    "サービス料",
    "返品",
    "ディスカウントと無料提供",
    "純売上高",
    "繰延売上",
    "ギフトカード売上",
    "一部支払い",
    "払い戻された一部入金", // NEW (38項目)
    "税金",
    "金額を指定した払い戻し",
    "売上合計",
    "受取合計額",
    "au PAY",
    "d払い",
    "カード",
    "その他",
    "ハウスアカウント",
    "楽天ペイ",
    "現金",
    "電子マネー",
    "手数料",
    "Squareの決済手数料",
    "Squareの手数料",
    "合計（純額）",
    "総売上数",
    "売上取引履歴",
    "商品売上取引履歴",
    "サービス料取引履歴",
    "商品別返品取引履歴",
    "ディスカウント取引履歴",
    "無料提供取引履歴",
    "ギフトカード売上取引履歴",
    "払い戻された一部入金の取引", // NEW (38項目)
    "税金取引履歴",
    "総売上取引履歴",
    "受取合計額の取引履歴",
  ];

  const rows = Object.keys(monthly)
    .sort()
    .reverse()
    .map((m) => buildSummaryRow(m, monthly[m]));
  const targetSheet = getOrCreateSheet(ss, "Square月次売上");
  writeToSheet(targetSheet, headers, rows);

  // 年月列をテキスト形式で上書き（スプレッドシートが日付型に変換するのを防ぐ）
  const lastRow = targetSheet.getLastRow();
  if (lastRow >= 2) {
    const monthValues = rows.map((r) => {
      const raw = String(r[0]);
      const parts = raw.split("-");
      return [
        parts.length >= 2 ? parts[0] + "-" + parts[1].padStart(2, "0") : raw,
      ];
    });
    targetSheet
      .getRange(2, 1, monthValues.length, 1)
      .setNumberFormat("@")
      .setValues(monthValues);
  }
}

function aggregateMonthlyData(rows) {
  const monthly = {};

  for (const row of rows) {
    const dateVal = row[SQ_COL.DATE];
    let m;
    if (dateVal instanceof Date) {
      m = Utilities.formatDate(dateVal, "JST", "yyyy-MM");
    } else {
      // 文字列「yyyy-MM-dd」または「yyyy-M-dd」を正規化して「yyyy-MM」に統一
      const parts = String(dateVal).split("-");
      if (parts.length >= 2) {
        m = parts[0] + "-" + parts[1].padStart(2, "0");
      } else {
        m = String(dateVal).substring(0, 7);
      }
    }

    if (!m || m.length < 7) continue;

    if (!monthly[m]) {
      monthly[m] = {
        gross: 0,
        tax: 0,
        disc: 0,
        returns: 0,
        refund: 0,
        fees: 0,
        partial: 0,
        txAll: new Set(),
        txItems: new Set(),
        txAllItems: new Set(),
        txReturns: new Set(),
        txDiscounts: new Set(),
        txTax: new Set(),
        txTenders: new Set(),
        txPartial: new Set(),
        pay: {
          "au PAY": 0,
          d払い: 0,
          カード: 0,
          その他: 0,
          ハウスアカウント: 0,
          楽天ペイ: 0,
          現金: 0,
          電子マネー: 0,
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
        o.tax += Number(row[SQ_COL.TAX]);
        o.disc += Number(row[SQ_COL.DISC]);
        o.returns += Number(row[SQ_COL.RETURN_GROSS]);
        o.refund += Number(row[SQ_COL.AMOUNT]);
        if (Number(row[SQ_COL.TAX]) !== 0) o.txTax.add(id);
        if (Number(row[SQ_COL.DISC]) !== 0) o.txDiscounts.add(id);
        if (Number(row[SQ_COL.RETURN_GROSS]) !== 0) o.txReturns.add(id);
        if (
          Number(row[SQ_COL.AMOUNT]) !== 0 ||
          row[SQ_COL.PAY_TYPE] === "払い戻し"
        ) {
          o.txTenders.add(id);
        }
        break;

      case "PAYMENT": {
        const pt = row[SQ_COL.PAY_TYPE];
        const amt = Number(row[SQ_COL.AMOUNT]);
        if (pt === "オンライン") break;
        o.pay[pt in o.pay ? pt : "その他"] += amt;
        o.fees += Number(row[SQ_COL.FEE]);
        o.txTenders.add(id);
        break;
      }

      case "PARTIAL": // NEW: 一部支払い
        o.partial += Number(row[SQ_COL.AMOUNT]);
        o.txPartial.add(id);
        break;
    }
  }

  return monthly;
}

function buildSummaryRow(month, o) {
  const netSales = o.gross + o.returns + o.disc;
  const totalSales = netSales + o.tax + o.refund;
  const collected = totalSales - (o.pay["ハウスアカウント"] ?? 0);
  const allTxCount = new Set([...o.txAllItems, ...o.txTenders]).size;

  return [
    month,
    o.gross,
    o.gross,
    0,
    o.returns,
    o.disc,
    netSales,
    0,
    0,
    o.partial,
    -o.partial, // 一部支払い / 払い戻された一部入金
    o.tax,
    o.refund,
    totalSales,
    collected,
    o.pay["au PAY"],
    o.pay["d払い"],
    o.pay["カード"],
    o.pay["その他"],
    o.pay["ハウスアカウント"],
    o.pay["楽天ペイ"],
    o.pay["現金"],
    o.pay["電子マネー"],
    -o.fees,
    -o.fees,
    -o.fees,
    collected - o.fees,
    allTxCount,
    o.txTax.size,
    o.txItems.size,
    0,
    o.txReturns.size,
    o.txDiscounts.size,
    0,
    0,
    o.txPartial.size, // 払い戻された一部入金の取引
    o.txTax.size,
    o.txItems.size,
    o.txTenders.size,
  ];
}

// ============================================================
// ユーティリティ
// ============================================================

function finalizeUpdate() {
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
    sheet
      .getRange(2, keyColumn, lastRow - 1, 1)
      .getValues()
      .map((r) => String(r[0])),
  );
}

function sortSheetByDate(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .sort({ column: 1, ascending: false });
}

function writeToSheet(sheet, headers, rows) {
  sheet.clear();
  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#f3f3f3");
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  sheet.setFrozenRows(1);
}

function clearSquareData() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    "確認",
    "Square売上データをすべて削除します。よろしいですか？",
    ui.ButtonSet.YES_NO,
  );
  if (res !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName("Square売上データ");
  if (existing) ss.deleteSheet(existing);
}

// ============================================================
// Square: 月単位取得（内部関数）
// ============================================================

function updateSquareMonth(startDate, endDate, targetMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, "Square売上データ");

  // ヘッダー行がない場合は追加（データが1行目に書き込まれるのを防ぐ）
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SQ_HEADERS_KEY.length).setValues([SQ_HEADERS_KEY]);
  }

  const sqHeaders = {
    Authorization: `Bearer ${CONFIG.SQUARE_ACCESS_TOKEN}`,
    "Content-Type": "application/json",
  };

  const {
    brandMap,
    refundMap,
    cardRefundPaymentIds,
    externalRefundPaymentIds,
    feeMap,
    partialMap,
  } = fetchPaymentsData(startDate, endDate, sqHeaders);

  // Refunds APIから返品注文IDを収集
  const refRes = JSON.parse(
    UrlFetchApp.fetch(
      "https://connect.squareup.com/v2/refunds" +
        "?begin_time=" +
        encodeURIComponent(startDate + "T00:00:00+09:00") +
        "&end_time=" +
        encodeURIComponent(endDate + "T00:00:00+09:00"),
      { headers: sqHeaders },
    ).getContentText(),
  );

  const returnOrderIds = new Set();
  refRes.refunds?.forEach((r) => {
    if (r.order_id) returnOrderIds.add(r.order_id);
  });
  console.log(`返品注文ID収集: ${returnOrderIds.size}件`);

  const locRes = JSON.parse(
    UrlFetchApp.fetch("https://connect.squareup.com/v2/locations", {
      headers: { Authorization: `Bearer ${CONFIG.SQUARE_ACCESS_TOKEN}` },
    }).getContentText(),
  );

  const existingKeys = getExistingKeys(sheet, SQ_COL.KEY + 1);

  for (const loc of locRes.locations) {
    let cursor = null;
    do {
      const payload = {
        location_ids: [loc.id],
        query: {
          filter: {
            closed_at: {
              start_at: startDate + "T00:00:00+09:00",
              end_at: endDate + "T00:00:00+09:00",
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
        }).getContentText(),
      );

      if (res.orders) {
        const monthOrders = res.orders.filter(
          (o) =>
            Utilities.formatDate(new Date(o.closed_at), "JST", "yyyy-MM") ===
            targetMonth,
        );

        const ordersToProcess = monthOrders
          .filter((o) => !returnOrderIds.has(o.id))
          .concat(
            monthOrders
              .filter((o) => returnOrderIds.has(o.id))
              .map((o) => {
                const detail = JSON.parse(
                  UrlFetchApp.fetch(
                    `https://connect.squareup.com/v2/orders/${o.id}`,
                    { headers: sqHeaders },
                  ).getContentText(),
                );
                return detail.order ?? o;
              }),
          );

        const newRows = buildSquareRows(
          ordersToProcess,
          existingKeys,
          brandMap,
          refundMap,
          cardRefundPaymentIds,
          externalRefundPaymentIds,
          feeMap,
          partialMap,
        );
        if (newRows.length > 0) {
          sheet
            .getRange(sheet.getLastRow() + 1, 1, newRows.length, 14)
            .setValues(newRows);
          newRows.forEach((row) => existingKeys.add(String(row[SQ_COL.KEY])));
        }
      }

      cursor = res.cursor ?? null;
    } while (cursor);
  }

  // Payments APIで取得したorder_idのうちシートにないものを個別取得
  const allPaymentOrderIds = new Set();
  let pCursor = null;
  do {
    let url =
      "https://connect.squareup.com/v2/payments" +
      "?begin_time=" +
      encodeURIComponent(startDate + "T00:00:00+09:00") +
      "&end_time=" +
      encodeURIComponent(endDate + "T00:00:00+09:00") +
      "&limit=200";
    if (pCursor) url += "&cursor=" + encodeURIComponent(pCursor);
    const pRes = JSON.parse(
      UrlFetchApp.fetch(url, { headers: sqHeaders }).getContentText(),
    );
    pRes.payments?.forEach((p) => {
      if (p.order_id) allPaymentOrderIds.add(p.order_id);
    });
    pCursor = pRes.cursor ?? null;
  } while (pCursor);

  const missingIds = [...allPaymentOrderIds].filter(
    (id) => !existingKeys.has(id + "_SUM"),
  );
  console.log(`未取得注文: ${missingIds.length}件`);

  if (missingIds.length > 0) {
    const missingOrders = missingIds
      .map((id) => {
        const detail = JSON.parse(
          UrlFetchApp.fetch(`https://connect.squareup.com/v2/orders/${id}`, {
            headers: sqHeaders,
          }).getContentText(),
        );
        return detail.order;
      })
      .filter((o) => o && o.state !== "CANCELED");

    const missingRows = buildSquareRows(
      missingOrders,
      existingKeys,
      brandMap,
      refundMap,
      cardRefundPaymentIds,
      externalRefundPaymentIds,
      feeMap,
      partialMap,
    );
    if (missingRows.length > 0) {
      sheet
        .getRange(sheet.getLastRow() + 1, 1, missingRows.length, 14)
        .setValues(missingRows);
    }
  }
}
