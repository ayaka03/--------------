/**
 * インターラーケン売上レポート集計システム
 * SquareサマリーCSVの35項目に完全準拠
 * v6: 多月対応・セキュリティ改善・コード整理版
 *
 * 実行環境: Google Apps Script (V8ランタイム)
 */

// ============================================================
// 定数
// ============================================================

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    SQUARE_ACCESS_TOKEN: props.getProperty("SQUARE_ACCESS_TOKEN"),
    COLORME_ACCESS_TOKEN: props.getProperty("COLORME_ACCESS_TOKEN"),
  };
}

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
    .addItem("1. カラーミー売上を更新（今月）", "runColormeUpdateThisMonth")
    .addItem("2. Square売上を更新（今月）", "runSquareUpdateThisMonth")
    .addSeparator()
    .addItem("3. 月を指定してSquare売上を更新", "runSquareUpdateByMonth")
    .addItem("4. 月を指定してカラーミー売上を更新", "runColormeUpdateByMonth")
    .addSeparator()
    .addItem("5. レポートを再集計", "recalculateAllSummaries")
    .addSeparator()
    .addItem("⚙️ APIトークンを設定する", "setupApiTokens")
    .addToUi();
}

// ── 今月のデータを更新 ──

function runColormeUpdateThisMonth() {
  const { start } = getCurrentMonthRange();
  updateColormeSalesMaster(start);
  finalizeUpdate();
}

function runSquareUpdateThisMonth() {
  const { start, end } = getCurrentMonthRange();
  updateSquareSalesMaster(start, end);
  finalizeUpdate();
}

// ── 月を指定して更新（多月対応） ──

function runSquareUpdateByMonth() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    "月を指定",
    "対象月を入力してください（例: 2026-01）",
    ui.ButtonSet.OK_CANCEL,
  );
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
  const result = ui.prompt(
    "月を指定",
    "対象月を入力してください（例: 2026-01）",
    ui.ButtonSet.OK_CANCEL,
  );
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
// 日付ユーティリティ
// ============================================================

/**
 * 今月の開始日・終了日を返す
 * @returns {{ start: string, end: string }}
 */
function getCurrentMonthRange() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  return getMonthRange(`${year}-${month}`);
}

/**
 * 指定した年月の開始日・翌月1日を返す
 * @param {string} yearMonth - "2026-02" 形式
 * @returns {{ start: string, end: string }}
 */
function getMonthRange(yearMonth) {
  const [year, month] = yearMonth.split("-").map(Number);
  const start = `${yearMonth}-01`;
  const nextMonth =
    month === 12
      ? `${year + 1}-01-01`
      : `${year}-${String(month + 1).padStart(2, "0")}-01`;
  return { start, end: nextMonth };
}

// ============================================================
// APIトークン設定ヘルパー（初期設定用）
// ============================================================

/**
 * 初回セットアップ用: ダイアログでAPIトークンを設定する
 * コードにトークンを書く代わりに、この関数を一度実行すればOK
 */
function setupApiTokens() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const sqResult = ui.prompt(
    "Square APIトークン設定",
    "Squareのアクセストークンをコピペしてください:",
    ui.ButtonSet.OK_CANCEL,
  );
  if (sqResult.getSelectedButton() !== ui.Button.OK) return;
  props.setProperty("SQUARE_ACCESS_TOKEN", sqResult.getResponseText().trim());

  const cmResult = ui.prompt(
    "カラーミー APIトークン設定",
    "カラーミーショップのアクセストークンをコピペしてください:",
    ui.ButtonSet.OK_CANCEL,
  );
  if (cmResult.getSelectedButton() !== ui.Button.OK) return;
  props.setProperty("COLORME_ACCESS_TOKEN", cmResult.getResponseText().trim());

  ui.alert("✅ APIトークンを安全に保存しました！");
}

// ============================================================
// カラーミー
// ============================================================

function updateColormeSalesMaster(startDate) {
  const config = getConfig();
  if (!config.COLORME_ACCESS_TOKEN) {
    SpreadsheetApp.getUi().alert(
      "カラーミーのAPIトークンが設定されていません。「⚙️ APIトークンを設定する」を実行してください。",
    );
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, "カラーミー売上データ");
  const existingKeys = getExistingKeys(sheet, CM_COL.KEY + 1);
  let offset = 0;
  let addedCount = 0;

  try {
    while (true) {
      const url = `https://api.shop-pro.jp/v1/sales.json?make_date_min=${startDate}&limit=100&offset=${offset}`;
      const res = JSON.parse(
        UrlFetchApp.fetch(url, {
          headers: { Authorization: `Bearer ${config.COLORME_ACCESS_TOKEN}` },
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
            existingKeys.add(key); // 重複防止のため追加済みキーに追加
          }
        }
      }

      if (newRows.length > 0) {
        sheet
          .getRange(sheet.getLastRow() + 1, 1, newRows.length, 11)
          .setValues(newRows);
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
// Square
// ============================================================

function updateSquareSalesMaster(startDate, endDate) {
  const config = getConfig();
  if (!config.SQUARE_ACCESS_TOKEN) {
    SpreadsheetApp.getUi().alert(
      "SquareのAPIトークンが設定されていません。「⚙️ APIトークンを設定する」を実行してください。",
    );
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, "Square売上データ");

  // ヘッダー行がなければ作成
  if (sheet.getLastRow() === 0) {
    sheet
      .getRange(1, 1, 1, 14)
      .setValues([
        [
          "日付",
          "注文ID",
          "商品名",
          "種別",
          "数量",
          "売上",
          "税金",
          "ディスカウント",
          "支払種別",
          "手数料",
          "キー",
          "返品売上",
          "返品税金",
          "金額",
        ],
      ])
      .setFontWeight("bold")
      .setBackground("#f3f3f3");
    sheet.setFrozenRows(1);
  }

  const existingKeys = getExistingKeys(sheet, SQ_COL.KEY + 1);
  const sqHeaders = {
    Authorization: `Bearer ${config.SQUARE_ACCESS_TOKEN}`,
    "Content-Type": "application/json",
  };

  // ✅ 修正: endDateを引数から受け取る（ハードコードをなくす）
  const { brandMap, refundMap, cardRefundPaymentIds, feeMap } =
    fetchPaymentsData(startDate, endDate, sqHeaders);

  try {
    const locRes = JSON.parse(
      UrlFetchApp.fetch("https://connect.squareup.com/v2/locations", {
        headers: { Authorization: `Bearer ${config.SQUARE_ACCESS_TOKEN}` },
      }).getContentText(),
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
          }).getContentText(),
        );

        if (res.orders) {
          const newRows = buildSquareRows(
            res.orders,
            existingKeys,
            brandMap,
            refundMap,
            cardRefundPaymentIds,
            feeMap,
          );
          if (newRows.length > 0) {
            sheet
              .getRange(sheet.getLastRow() + 1, 1, newRows.length, 14)
              .setValues(newRows);
            totalAdded += newRows.length;
            // 追加したキーをexistingKeysに反映（同一実行内での重複を防ぐ）
            newRows.forEach((r) => existingKeys.add(r[SQ_COL.KEY]));
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
// Payments API（ブランド・返金・手数料の取得）
// ============================================================

/**
 * ✅ 修正: endDateを引数で受け取れるようにした（ハードコードを廃止）
 * @param {string} startDate - "2026-02-01" 形式
 * @param {string} endDate   - "2026-03-01" 形式（翌月1日）
 * @param {Object} sqHeaders - Authorization ヘッダー
 */
function fetchPaymentsData(startDate, endDate, sqHeaders) {
  const brandMap = new Map();
  const refundMap = new Map();
  const feeMap = new Map();
  const paymentOrderMap = new Map();
  const cardRefundPaymentIds = new Set();

  const startIso = encodeURIComponent(`${startDate}T00:00:00+09:00`);
  const endIso = encodeURIComponent(`${endDate}T00:00:00+09:00`);

  let cursor = null;
  do {
    let url =
      `https://connect.squareup.com/v2/payments` +
      `?begin_time=${startIso}` +
      `&end_time=${endIso}` +
      `&limit=200`;
    if (cursor) url += `&cursor=${encodeURIComponent(cursor)}`;

    const res = JSON.parse(
      UrlFetchApp.fetch(url, { headers: sqHeaders }).getContentText(),
    );

    res.payments?.forEach((p) => {
      paymentOrderMap.set(p.id, p.order_id);

      if (p.source_type === "WALLET" && p.wallet_details?.brand) {
        brandMap.set(p.id, p.wallet_details.brand);
      }
      if (
        p.source_type === "CARD" &&
        p.card_details?.card?.card_brand === "FELICA"
      ) {
        brandMap.set(p.id, "FELICA");
      }
      if (p.refunded_money?.amount > 0 && p.source_type !== "CASH") {
        refundMap.set(
          p.order_id,
          (refundMap.get(p.order_id) ?? 0) + p.refunded_money.amount,
        );
        cardRefundPaymentIds.add(p.id);
      }
      if (p.processing_fee) {
        const fee = p.processing_fee.reduce(
          (sum, f) => sum + (f.amount_money?.amount ?? 0),
          0,
        );
        if (fee > 0)
          feeMap.set(p.order_id, (feeMap.get(p.order_id) ?? 0) + fee);
      }
    });

    cursor = res.cursor ?? null;
  } while (cursor);

  // Refunds APIで手数料戻りを取得
  const refRes = JSON.parse(
    UrlFetchApp.fetch(
      `https://connect.squareup.com/v2/refunds?begin_time=${startIso}&end_time=${endIso}`,
      { headers: sqHeaders },
    ).getContentText(),
  );

  refRes.refunds?.forEach((r) => {
    if (r.processing_fee) {
      const paymentId = r.id.split("_")[0];
      const orderId = paymentOrderMap.get(paymentId);
      if (orderId) {
        const fee = r.processing_fee.reduce(
          (sum, f) => sum + (f.amount_money?.amount ?? 0),
          0,
        );
        feeMap.set(orderId, (feeMap.get(orderId) ?? 0) + fee);
      }
    }
  });

  console.log(
    `ブランドマップ: ${brandMap.size}件 / 返金マップ: ${refundMap.size}件 / 手数料マップ: ${feeMap.size}件`,
  );
  return { brandMap, refundMap, cardRefundPaymentIds, feeMap };
}

// ============================================================
// Square行データ構築
// ============================================================

function buildSquareRows(
  orders,
  existingKeys,
  brandMap,
  refundMap,
  cardRefundPaymentIds,
  feeMap,
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
          manualRefund = -total;
        }
      }

      if (retGross === 0 && retTax === 0) {
        order.refunds?.forEach((rf) => {
          if (!rf.return_id) manualRefund += rf.amount_money?.amount ?? 0;
        });
      }

      if (retGross === 0 && retTax === 0 && manualRefund === 0) {
        const partialRefund = refundMap?.get(id) ?? 0;
        if (partialRefund > 0) manualRefund = partialRefund;
      }

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

    // 3. 現金返品のマイナス行
    if (
      order.return_amounts &&
      (!order.tenders || order.tenders.length === 0)
    ) {
      const hasCardRefund =
        order.refunds?.some((rf) => cardRefundPaymentIds?.has(rf.tender_id)) ??
        false;

      if (!hasCardRefund) {
        const refundKey = `${id}_REFUND`;
        if (!existingKeys.has(refundKey)) {
          const total = order.return_amounts.total_money?.amount ?? 0;
          rows.push([
            dateStr,
            id,
            "返金",
            "PAYMENT",
            0,
            0,
            0,
            0,
            "現金",
            0,
            refundKey,
            0,
            0,
            -total,
          ]);
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
  }

  return rows;
}

function getPaymentType(tender, brandMap) {
  switch (tender.type) {
    case "CARD": {
      const mappedBrand = (
        brandMap?.get(tender.payment_id) ?? ""
      ).toUpperCase();
      if (mappedBrand === "FELICA") return "電子マネー";
      const brand = (tender.card_details?.card?.card_brand ?? "").toUpperCase();
      if (brand === "FELICA") return "電子マネー";
      return ELECTRONIC_MONEY_BRANDS.has(brand) ? "電子マネー" : "カード";
    }
    case "WALLET": {
      const brand = (brandMap?.get(tender.payment_id) ?? "").toUpperCase();
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
  if (!sqSheet) {
    SpreadsheetApp.getUi().alert(
      "「Square売上データ」シートが見つかりません。",
    );
    return;
  }

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
    "税金取引履歴",
    "総売上取引履歴",
    "受取合計額の取引履歴",
  ];

  const rows = Object.keys(monthly)
    .sort()
    .reverse()
    .map((m) => buildSummaryRow(m, monthly[m]));
  writeToSheet(getOrCreateSheet(ss, "Square月次売上"), headers, rows);
  SpreadsheetApp.getUi().alert("✅ すべてのレポートを更新しました！");
}

function aggregateMonthlyData(rows) {
  const monthly = {};

  for (const row of rows) {
    const dateVal = row[SQ_COL.DATE];
    const m =
      dateVal instanceof Date
        ? Utilities.formatDate(dateVal, "JST", "yyyy-MM")
        : String(dateVal).substring(0, 7);

    if (!m || m.length < 7) continue;

    if (!monthly[m]) {
      monthly[m] = {
        gross: 0,
        tax: 0,
        disc: 0,
        returns: 0,
        refund: 0,
        fees: 0,
        txAll: new Set(),
        txItems: new Set(),
        txAllItems: new Set(),
        txReturns: new Set(),
        txDiscounts: new Set(),
        txTax: new Set(),
        txTenders: new Set(),
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

/*
 * ⚠️ 削除した関数について
 *
 * updateSquare2026Feb() → runSquareUpdateByMonth() で代替
 *   「3. 月を指定してSquare売上を更新」から「2026-02」と入力すれば同じ動作
 *
 * clearSquareData() → 削除（確認なしのシート削除は危険なため廃止）
 *   必要な場合はスプレッドシートから手動でシートを削除してください
 */
