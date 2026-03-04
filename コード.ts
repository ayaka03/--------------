/**
 * インターラーケン売上レポート集計システム
 * SquareサマリーCSVの35項目に完全準拠
 * v3: 現金返品・カード返金・部分返金・FELICA・WALLET対応
 */

const CONFIG = {
  SQUARE_ACCESS_TOKEN:
    "EAAAl3VlBqnOihdeDGqTuOyfuE8juXQrSNR6cgpX-RDtVxxFyr4d7daw5jil-oow",
  COLORME_ACCESS_TOKEN:
    "4fd03a83f636c4517b72bf23cde52b797fff500263e34e3f26ac2c26f3c10ee7",
  START_DATE: "2026-02-01",
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

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀インターラーケン操作")
    .addItem("1. カラーミー売上を更新", "runColormeUpdate")
    .addItem("2. Square売上を更新", "runSquareUpdate")
    .addSeparator()
    .addItem("3. レポートを再集計", "recalculateAllSummaries")
    .addToUi();
}

function runColormeUpdate() {
  updateColormeSalesMaster(CONFIG.START_DATE);
  finalizeUpdate();
}

function runSquareUpdate() {
  updateSquareSalesMaster(CONFIG.START_DATE);
  finalizeUpdate();
}

// ============================================================
// カラーミー関連
// ============================================================

function updateColormeSalesMaster(startDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, "カラーミー売上データ");
  const existingKeys = getExistingKeys(sheet, CM_COL.KEY + 1);
  let offset = 0;

  try {
    while (true) {
      const url = `https://api.shop-pro.jp/v1/sales.json?make_date_min=${startDate}&limit=100&offset=${offset}`;
      const options = {
        headers: { Authorization: `Bearer ${CONFIG.COLORME_ACCESS_TOKEN}` },
      };
      const res = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
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
// Square関連
// ============================================================

function updateSquareSalesMaster(startDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, "Square売上データ");

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
  }

  const existingKeys = getExistingKeys(sheet, SQ_COL.KEY + 1);
  const startAt = new Date(`${startDate}T00:00:00+09:00`).toISOString();
  const sqHeaders = {
    Authorization: `Bearer ${CONFIG.SQUARE_ACCESS_TOKEN}`,
    "Content-Type": "application/json",
  };

  const { brandMap, refundMap, cardRefundPaymentIds } = fetchPaymentsData(
    startDate,
    sqHeaders,
  );

  try {
    const locRes = JSON.parse(
      UrlFetchApp.fetch("https://connect.squareup.com/v2/locations", {
        headers: { Authorization: `Bearer ${CONFIG.SQUARE_ACCESS_TOKEN}` },
      }).getContentText(),
    );

    for (const loc of locRes.locations) {
      let cursor = null;
      do {
        const payload = {
          location_ids: [loc.id],
          query: {
            filter: {
              closed_at: { start_at: startAt },
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
          );
          if (newRows.length > 0) {
            sheet
              .getRange(sheet.getLastRow() + 1, 1, newRows.length, 14)
              .setValues(newRows);
          }
        }

        cursor = res.cursor ?? null;
      } while (cursor);
    }
  } catch (e) {
    console.error(`SQ Error: ${e.message}`);
  }
}

// Payments APIからWALLETブランド・FELICA・返金情報を一括取得
function fetchPaymentsData(startDate, sqHeaders) {
  const brandMap = new Map();
  const refundMap = new Map();
  const cardRefundPaymentIds = new Set(); // カード返金のpayment_id
  let cursor = null;

  do {
    let url = `https://connect.squareup.com/v2/payments?begin_time=${encodeURIComponent(startDate + "T00:00:00+09:00")}&end_time=${encodeURIComponent("2026-03-01T00:00:00+09:00")}&limit=200`;
    if (cursor) url += `&cursor=${encodeURIComponent(cursor)}`;

    const res = JSON.parse(
      UrlFetchApp.fetch(url, { headers: sqHeaders }).getContentText(),
    );

    res.payments?.forEach((p) => {
      // WALLETのブランド判別
      if (p.source_type === "WALLET" && p.wallet_details?.brand) {
        brandMap.set(p.id, p.wallet_details.brand);
      }
      // FELICAの判別
      if (
        p.source_type === "CARD" &&
        p.card_details?.card?.card_brand === "FELICA"
      ) {
        brandMap.set(p.id, "FELICA");
      }
      // 現金以外の返金を記録
      if (p.refunded_money?.amount > 0 && p.source_type !== "CASH") {
        refundMap.set(
          p.order_id,
          (refundMap.get(p.order_id) ?? 0) + p.refunded_money.amount,
        );
        cardRefundPaymentIds.add(p.id); // payment_idを記録
      }
    });

    cursor = res.cursor ?? null;
  } while (cursor);

  console.log(`ブランドマップ取得完了: ${brandMap.size}件`);
  console.log(`部分返金マップ取得完了: ${refundMap.size}件`);
  console.log(`カード返金ID取得完了: ${cardRefundPaymentIds.size}件`);
  return { brandMap, refundMap, cardRefundPaymentIds };
}

function buildSquareRows(
  orders,
  existingKeys,
  brandMap,
  refundMap,
  cardRefundPaymentIds,
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

      // return_amountsがない場合のみrefundsを処理
      if (retGross === 0 && retTax === 0) {
        order.refunds?.forEach((rf) => {
          if (!rf.return_id) manualRefund += rf.amount_money?.amount ?? 0;
        });
      }

      // Payments APIの部分返金を反映（商品返品がない場合のみ）
      if (retGross === 0 && retTax === 0 && manualRefund === 0) {
        const partialRefund = refundMap?.get(id) ?? 0;
        if (partialRefund > 0) {
          manualRefund = partialRefund;
        }
      }

      rows.push([
        dateStr,
        id,
        "注文サマリー",
        "SUMMARY",
        0,
        0,
        totalTax,
        -totalDisc,
        "",
        0,
        sumKey,
        -retGross,
        -retTax,
        -manualRefund,
      ]);
    }

    // 3. 現金返品のマイナス行（tenderがなく、カード返金でない場合）
    if (
      order.return_amounts &&
      (!order.tenders || order.tenders.length === 0)
    ) {
      // refundsのtender_idがカード返金かどうか確認
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
        const fee = tender.processing_fee_money?.amount ?? 0;
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
      const brand = (tender.card_details?.card_brand ?? "").toUpperCase();
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
      if (
        src.includes("D払い") ||
        src.includes("DBARAI") ||
        src.includes("D-BARAI")
      )
        return "d払い";
      if (src.includes("楽天") || src.includes("RAKUTEN")) return "楽天ペイ";
      return "その他";
    }
    case "OTHER": {
      const note = (tender.note ?? "").toLowerCase();
      if (note.includes("オンライン") || note.includes("online"))
        return "その他";
      if (note.includes("代引")) return "その他";
      if (note.includes("pay pay") || note.includes("paypay")) return "その他";
      if (
        note.includes("売掛") ||
        note.includes("掛け") ||
        note.includes("aoi 売掛")
      )
        return "ハウスアカウント";
      if (note.includes("クレジット") || note.includes("credit"))
        return "カード";
      if (
        note.includes("id") ||
        note.includes("マナカ") ||
        note.includes("manaca") ||
        note.includes("イオン") ||
        note.includes("suica") ||
        note.includes("pasmo")
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
  SpreadsheetApp.getUi().alert("すべてのレポートを更新しました！😍");
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
        items: 0,
        txAll: new Set(),
        txItems: new Set(),
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
        o.items += Number(row[SQ_COL.QTY]);
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
        break;

      case "PAYMENT": {
        const pt = row[SQ_COL.PAY_TYPE];
        const amt = Number(row[SQ_COL.AMOUNT]);
        if (pt === "オンライン") break;
        if (pt in o.pay) {
          o.pay[pt] += amt;
        } else {
          o.pay["その他"] += amt;
        }
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
    o.items,
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

function clearSquareData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ss.getSheetByName("Square売上データ");
  if (existing) ss.deleteSheet(existing);
  SpreadsheetApp.getUi().alert(
    "シートを削除しました！次は updateSquare2026Feb() を実行してください",
  );
}

// ============================================================
// テスト用：2月データ再取得
// ============================================================

function updateSquare2026Feb() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const existing = ss.getSheetByName("Square売上データ");
  if (existing) ss.deleteSheet(existing);
  const sheet = ss.insertSheet("Square売上データ");

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

  const sqHeaders = {
    Authorization: `Bearer ${CONFIG.SQUARE_ACCESS_TOKEN}`,
    "Content-Type": "application/json",
  };

  const { brandMap, refundMap, cardRefundPaymentIds } = fetchPaymentsData(
    "2026-02-01",
    sqHeaders,
  );

  const locRes = JSON.parse(
    UrlFetchApp.fetch("https://connect.squareup.com/v2/locations", {
      headers: { Authorization: `Bearer ${CONFIG.SQUARE_ACCESS_TOKEN}` },
    }).getContentText(),
  );

  for (const loc of locRes.locations) {
    let cursor = null;
    do {
      const payload = {
        location_ids: [loc.id],
        query: {
          filter: {
            closed_at: {
              start_at: "2026-02-01T00:00:00+09:00",
              end_at: "2026-03-01T00:00:00+09:00",
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
        const febOrders = res.orders.filter((o) => {
          const m = Utilities.formatDate(
            new Date(o.closed_at),
            "JST",
            "yyyy-MM",
          );
          return m === "2026-02";
        });
        const newRows = buildSquareRows(
          febOrders,
          new Set(),
          brandMap,
          refundMap,
          cardRefundPaymentIds,
        );
        if (newRows.length > 0) {
          sheet
            .getRange(sheet.getLastRow() + 1, 1, newRows.length, 14)
            .setValues(newRows);
        }
      }

      cursor = res.cursor ?? null;
    } while (cursor);
  }

  SpreadsheetApp.getUi().alert(
    "取得完了！次は recalculateAllSummaries() を実行してください",
  );
}
