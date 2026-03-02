/**
 * インターラーケン売上レポート集計システム
 * SquareサマリーCSVの35項目に完全準拠
 *
 * 実行環境: Google Apps Script (V8ランタイム)
 * VS Code補完: JSDoc + @types/google-apps-script
 */

// ============================================================
// JSDoc 型定義 (VS Codeの補完・型チェックに使用)
// ============================================================

/**
 * @typedef {Object} Config
 * @property {string} SQUARE_ACCESS_TOKEN
 * @property {string} COLORME_ACCESS_TOKEN
 * @property {string} START_DATE
 */

/**
 * @typedef {Object} ColormeDetail
 * @property {string|number} id
 * @property {string} product_name
 * @property {string|number} [unit_num]
 * @property {string|number} [product_num]
 * @property {string|number} price
 */

/**
 * @typedef {Object} ColormeSale
 * @property {string|number} id
 * @property {string|number} make_date
 * @property {ColormeDetail[]} details
 * @property {number} delivery_total
 * @property {number} fee_total
 * @property {number} point_discount
 */

/**
 * @typedef {Object} ColormeResponse
 * @property {ColormeSale[]} [sales]
 */

/**
 * @typedef {Object} SquareMoney
 * @property {number} [amount]
 * @property {string} [currency]
 */

/**
 * @typedef {Object} SquareLineItem
 * @property {string} name
 * @property {string} quantity
 * @property {SquareMoney} [gross_sales_money]
 */

/**
 * @typedef {Object} SquareReturnAmounts
 * @property {SquareMoney} [gross_return_money]
 * @property {SquareMoney} [tax_money]
 */

/**
 * @typedef {Object} SquareReturn
 * @property {SquareReturnAmounts} [return_amounts]
 */

/**
 * @typedef {Object} SquareRefund
 * @property {string} [return_id]
 * @property {SquareMoney} [amount_money]
 */

/**
 * @typedef {Object} SquareCardDetails
 * @property {string} [card_brand]
 */

/**
 * @typedef {Object} SquareExternalDetails
 * @property {string} [source_name]
 */

/**
 * @typedef {Object} SquareTender
 * @property {string} type
 * @property {SquareMoney} [amount_money]
 * @property {SquareMoney} [processing_fee_money]
 * @property {SquareCardDetails} [card_details]
 * @property {SquareExternalDetails} [external_details]
 */

/**
 * @typedef {Object} SquareOrder
 * @property {string} id
 * @property {string} closed_at
 * @property {SquareLineItem[]} [line_items]
 * @property {SquareReturn[]} [returns]
 * @property {SquareRefund[]} [refunds]
 * @property {SquareTender[]} [tenders]
 * @property {SquareMoney} [total_tax_money]
 * @property {SquareMoney} [total_discount_money]
 */

/**
 * @typedef {Object} SquareOrdersResponse
 * @property {SquareOrder[]} [orders]
 * @property {string} [cursor]
 */

/**
 * @typedef {Object} SquareLocation
 * @property {string} id
 * @property {string} [name]
 */

/**
 * @typedef {Object} SquareLocationsResponse
 * @property {SquareLocation[]} locations
 */

/**
 * @typedef {Object} MonthlyPayments
 * @property {number} auPay
 * @property {number} dBarai
 * @property {number} card
 * @property {number} other
 * @property {number} houseAccount
 * @property {number} rakuten
 * @property {number} cash
 * @property {number} eMoney
 */

/**
 * @typedef {Object} MonthlyData
 * @property {number} gross
 * @property {number} tax
 * @property {number} disc
 * @property {number} returns
 * @property {number} refund
 * @property {number} fees
 * @property {number} items
 * @property {Set<string>} txAll
 * @property {Set<string>} txItems
 * @property {Set<string>} txReturns
 * @property {Set<string>} txDiscounts
 * @property {Set<string>} txTax
 * @property {Set<string>} txTenders
 * @property {Record<string, number>} pay
 */

// ============================================================
// 定数 (Constants)
// ============================================================

/** @type {Config} */
const CONFIG = {
  SQUARE_ACCESS_TOKEN:
    "EAAAl3VlBqnOihdeDGqTuOyfuE8juXQrSNR6cgpX-RDtVxxFyr4d7daw5jil-oow",
  COLORME_ACCESS_TOKEN:
    "4fd03a83f636c4517b72bf23cde52b797fff500263e34e3f26ac2c26f3c10ee7",
  START_DATE: "2026-02-01",
};

/** カラーミー売上データシートの列インデックス (0始まり) */
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

/** Square売上データシートの列インデックス (0始まり) */
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

/** 電子マネーとして扱うカードブランドのSet */
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
// エントリーポイント (Entry Points)
// ============================================================

/** スプレッドシートを開いたときにカスタムメニューを追加する */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚀インターラーケン操作")
    .addItem("1. カラーミー売上を更新", "runColormeUpdate")
    .addItem("2. Square売上を更新", "runSquareUpdate")
    .addSeparator()
    .addItem("3. レポートを再集計", "recalculateAllSummaries")
    .addToUi();
}

/** カラーミー売上を更新してレポートを最終化する */
function runColormeUpdate() {
  updateColormeSalesMaster(CONFIG.START_DATE);
  finalizeUpdate();
}

/** Square売上を更新してレポートを最終化する */
function runSquareUpdate() {
  updateSquareSalesMaster(CONFIG.START_DATE);
  finalizeUpdate();
}

// ============================================================
// 1. カラーミー関連 (Colorme)
// ============================================================

/**
 * カラーミー売上データマスタを更新する
 * @param {string} startDate - 取得開始日 (例: "2026-02-01")
 */
function updateColormeSalesMaster(startDate) {
  /** @type {GoogleAppsScript.Spreadsheet.Spreadsheet} */
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  /** @type {GoogleAppsScript.Spreadsheet.Sheet} */
  const sheet = getOrCreateSheet(ss, "カラーミー売上データ");

  const existingKeys = getExistingKeys(sheet, CM_COL.KEY + 1);
  let offset = 0;

  try {
    while (true) {
      const url = `https://api.shop-pro.jp/v1/sales.json?make_date_min=${startDate}&limit=100&offset=${offset}`;

      /** @type {GoogleAppsScript.URL_Fetch.URLFetchRequestOptions} */
      const options = {
        headers: { Authorization: `Bearer ${CONFIG.COLORME_ACCESS_TOKEN}` },
      };

      /** @type {ColormeResponse} */
      const res = JSON.parse(UrlFetchApp.fetch(url, options).getContentText());

      if (!res.sales || res.sales.length === 0) break;

      /** @type {Array<Array<string|number>>} */
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

/**
 * カラーミーの make_date を "yyyy-MM-dd" 形式に変換する
 * @param {string|number} raw - UNIXタイムスタンプ(秒)または "yyyy-MM-dd HH:mm:ss" 文字列
 * @returns {string}
 */
function parseSaleDate(raw) {
  if (typeof raw === "number") {
    return Utilities.formatDate(new Date(raw * 1000), "JST", "yyyy-MM-dd");
  }
  return raw.split(" ")[0];
}

// ============================================================
// 2. Square関連 (Square)
// ============================================================

/**
 * Square売上データマスタを更新する
 * @param {string} startDate - 取得開始日 (例: "2026-02-01")
 */
function updateSquareSalesMaster(startDate) {
  /** @type {GoogleAppsScript.Spreadsheet.Spreadsheet} */
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  /** @type {GoogleAppsScript.Spreadsheet.Sheet} */
  const sheet = getOrCreateSheet(ss, "Square売上データ");

  // ヘッダーがなければ自動で追加
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
  const startAt = new Date(`${startDate}T00:00:00+09:00`).toISOString();

  /** @type {GoogleAppsScript.URL_Fetch.HttpHeaders} */
  const sqHeaders = {
    Authorization: `Bearer ${CONFIG.SQUARE_ACCESS_TOKEN}`,
    "Content-Type": "application/json",
  };

  try {
    /** @type {SquareLocationsResponse} */
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

        /** @type {SquareOrdersResponse} */
        const res = JSON.parse(
          UrlFetchApp.fetch("https://connect.squareup.com/v2/orders/search", {
            method: "post",
            headers: sqHeaders,
            payload: JSON.stringify(payload),
          }).getContentText(),
        );

        if (res.orders) {
          const newRows = buildSquareRows(res.orders, existingKeys);
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

/**
 * SquareのOrderリストからシートに書き込む行データを生成する
 * @param {SquareOrder[]} orders
 * @param {Set<string>} existingKeys
 * @returns {Array<Array<string|number>>}
 */
function buildSquareRows(orders, existingKeys) {
  /** @type {Array<Array<string|number>>} */
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
      const totalTax = order.total_tax_money?.amount ?? 0;
      const totalDisc = order.total_discount_money?.amount ?? 0;

      let retGross = 0;
      let retTax = 0;
      let manualRefund = 0;

      if (order.return_amounts) {
        const tax = order.return_amounts.tax_money?.amount ?? 0;
        const total = order.return_amounts.total_money?.amount ?? 0;
        if (tax > 0) {
          retTax = tax;
          retGross = total - tax;
        } else {
          manualRefund = -total;
        }
      }

      order.refunds?.forEach((rf) => {
        if (!rf.return_id) manualRefund += rf.amount_money?.amount ?? 0;
      });

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

    // 3. 支払い行 (PAYMENT)
    order.tenders?.forEach((tender, i) => {
      const key = `${id}_T_${i}`;
      if (!existingKeys.has(key)) {
        const payType = getPaymentType(tender);
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
  }

  return rows;
}

/**
 * SquareのTenderオブジェクトから支払い種別名を返す
 * @param {SquareTender} tender
 * @returns {string}
 */
function getPaymentType(tender) {
  switch (tender.type) {
    case "CARD": {
      const brand = (tender.card_details?.card_brand ?? "").toUpperCase();
      return ELECTRONIC_MONEY_BRANDS.has(brand) ? "電子マネー" : "カード";
    }
    case "CASH":
      return "現金";
    case "HOUSE_ACCOUNT":
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
      return "電子マネー";
    }
    case "OTHER": {
      const note = (tender.note ?? "").toLowerCase();
      // オンライン注文（カラーミーと重複するため除外用）
      if (note.includes("オンライン") || note.includes("online"))
        return "オンライン";
      // 売掛（別タブ管理用）
      if (
        note.includes("売掛") ||
        note.includes("掛け") ||
        note.includes("かけ")
      )
        return "売掛";
      // クレジット系（テラス・エポス等）→ カード扱い
      if (note.includes("クレジット") || note.includes("credit"))
        return "カード";
      // 電子マネー系
      if (
        note.includes("id") ||
        note.includes("マナカ") ||
        note.includes("manaca") ||
        note.includes("イオン") ||
        note.includes("kuikku") ||
        note.includes("suica") ||
        note.includes("pasmo")
      )
        return "電子マネー";
      // 金シャチ系
      if (note.includes("金シャチ") || note.includes("きんしゃち"))
        return "その他";
      // PayPay系（デフォルト）
      return "その他";
    }
    default:
      return "その他";
  }
}

// ============================================================
// 3. 再集計 (Recalculate - CSV 35項目完全再現)
// ============================================================

/** すべての月次集計を再計算してSheetに書き出す */
function recalculateAllSummaries() {
  /** @type {GoogleAppsScript.Spreadsheet.Spreadsheet} */
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sqSheet = ss.getSheetByName("Square売上データ");
  if (!sqSheet) return;

  const data = sqSheet.getDataRange().getValues();
  const monthly = aggregateMonthlyData(data.slice(1)); // ヘッダー行をスキップ

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

/**
 * Square売上データシートの行データから月次集計マップを構築する
 * @param {Array<Array<any>>} rows
 * @returns {Record<string, MonthlyData>}
 */
function aggregateMonthlyData(rows) {
  /** @type {Record<string, MonthlyData>} */
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
        // オンライン・売掛は集計から除外
        if (pt === "オンライン" || pt === "売掛") break;
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

/**
 * 月次集計データから35列のサマリー行を構築する
 * @param {string} month
 * @param {MonthlyData} o
 * @returns {Array<string|number>}
 */
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
// 4. 共通ユーティリティ (Utilities)
// ============================================================

/** 全シートのデータをソートして月次集計を再実行する */
function finalizeUpdate() {
  /** @type {GoogleAppsScript.Spreadsheet.Spreadsheet} */
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (const name of ["Square売上データ", "カラーミー売上データ"]) {
    const s = ss.getSheetByName(name);
    if (s) sortSheetByDate(s);
  }

  recalculateAllSummaries();
}

/**
 * シートを名前で取得し、存在しない場合は新規作成する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) ?? ss.insertSheet(name);
}

/**
 * シートの指定列から既存のキーをSetで返す
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} keyColumn - キーが入っている列番号 (1始まり)
 * @returns {Set<string>}
 */
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

/**
 * シートのデータ行を日付の降順でソートする (2行目以降)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function sortSheetByDate(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .sort({ column: 1, ascending: false });
}

/**
 * シートをクリアしてヘッダーとデータを書き込む
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} headers
 * @param {Array<Array<string|number>>} rows
 */
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

function checkTenderNames() {
  const sqHeaders = {
    Authorization: `Bearer ${CONFIG.SQUARE_ACCESS_TOKEN}`,
    "Content-Type": "application/json",
  };

  const locRes = JSON.parse(
    UrlFetchApp.fetch("https://connect.squareup.com/v2/locations", {
      headers: { Authorization: `Bearer ${CONFIG.SQUARE_ACCESS_TOKEN}` },
    }).getContentText(),
  );

  const payload = {
    location_ids: [locRes.locations[0].id],
    query: {
      filter: {
        closed_at: { start_at: "2026-02-01T00:00:00+09:00" },
        state_filter: { states: ["COMPLETED"] },
      },
    },
    limit: 10,
  };

  const res = JSON.parse(
    UrlFetchApp.fetch("https://connect.squareup.com/v2/orders/search", {
      method: "post",
      headers: sqHeaders,
      payload: JSON.stringify(payload),
    }).getContentText(),
  );

  res.orders?.forEach((order) => {
    order.tenders?.forEach((tender) => {
      console.log(
        JSON.stringify({
          type: tender.type,
          card_brand: tender.card_details?.card_brand,
          source_name: tender.external_details?.source_name,
          amount: tender.amount_money?.amount,
        }),
      );
    });
  });
}
