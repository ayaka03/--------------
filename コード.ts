/**
 * インターラーケン売上レポート集計システム - 完璧版
 * SquareサマリーCSVの35項目に完全準拠
 */

// --- 設定情報 ---
const SQUARE_ACCESS_TOKEN = 'EAAAl3VlBqnOihdeDGqTuOyfuE8juXQrSNR6cgpX-RDtVxxFyr4d7daw5jil-oow';
const COLORME_ACCESS_TOKEN = '4fd03a83f636c4517b72bf23cde52b797fff500263e34e3f26ac2c26f3c10ee7';
const START_DATE = "2026-02-01"; 

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀インターラーケン操作')
    .addItem('1. カラーミー売上を更新', 'runColormeUpdate')
    .addItem('2. Square売上を更新', 'runSquareUpdate')
    .addSeparator()
    .addItem('3. レポートを再集計', 'recalculateAllSummaries')
    .addToUi();
}

// ============================================================
// 1. カラーミー関連
// ============================================================
function runColormeUpdate() { updateColormeSalesMaster(START_DATE); finalizeUpdate(); }

function updateColormeSalesMaster(startDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, 'カラーミー売上データ');
  const existingKeys = getExistingKeys(sheet, 11);
  let offset = 0;
  try {
    while (true) {
      const url = `https://api.shop-pro.jp/v1/sales.json?make_date_min=${startDate}&limit=100&offset=${offset}`;
      const res = JSON.parse(UrlFetchApp.fetch(url, { headers: { "Authorization": "Bearer " + COLORME_ACCESS_TOKEN } }).getContentText());
      if (!res.sales || res.sales.length === 0) break;
      let rows = [];
      res.sales.forEach(sale => {
        let sdRaw = sale.make_date;
        let saleDate = (typeof sdRaw === 'number') ? Utilities.formatDate(new Date(sdRaw * 1000), "JST", "yyyy-MM-dd") : sdRaw.split(" ")[0];
        sale.details.forEach(detail => {
          let qty = Number(detail.unit_num) || Number(detail.product_num) || 1;
          const key = String(sale.id) + "_D_" + String(detail.id);
          if (!existingKeys.has(key)) {
            rows.push([saleDate, sale.id, detail.product_name, "SALE", qty, Number(detail.price), qty * Number(detail.price), sale.delivery_total, sale.fee_total, sale.point_discount, key]);
          }
        });
      });
      if (rows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 11).setValues(rows);
      if (res.sales.length < 100) break;
      offset += 100;
    }
  } catch(e) { console.error("CM Error: " + e.message); }
}

// ============================================================
// 2. スクエア関連（精密抽出ロジック）
// ============================================================
function runSquareUpdate() { updateSquareSalesMaster(START_DATE); finalizeUpdate(); }

function updateSquareSalesMaster(startDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(ss, 'Square売上データ');
  const existingKeys = getExistingKeys(sheet, 11); 
  const startAt = new Date(startDate + "T00:00:00+09:00").toISOString();

  try {
    const locRes = JSON.parse(UrlFetchApp.fetch('https://connect.squareup.com/v2/locations', { 
      headers: { "Authorization": "Bearer " + SQUARE_ACCESS_TOKEN } 
    }));
    
    for (const loc of locRes.locations) {
      let cursor = null;
      do {
        const payload = { 
          "location_ids": [loc.id], 
          "query": { "filter": { "closed_at": { "start_at": startAt }, "state_filter": { "states": ["COMPLETED"] } } }, 
          "cursor": cursor 
        };
        const res = JSON.parse(UrlFetchApp.fetch('https://connect.squareup.com/v2/orders/search', { 
          method: "post", 
          headers: { "Authorization": "Bearer " + SQUARE_ACCESS_TOKEN, "Content-Type": "application/json" }, 
          payload: JSON.stringify(payload) 
        }));

        if (res.orders) {
          let rows = [];
          res.orders.forEach(order => {
            const dateStr = Utilities.formatDate(new Date(order.closed_at), "JST", "yyyy-MM-dd");
            const id = order.id;

            // 1. 商品売上行
            if (order.line_items) {
              order.line_items.forEach((item, i) => {
                const key = id + "_L_" + i;
                if (!existingKeys.has(key)) {
                  const gross = Number(item.gross_sales_money?.amount || 0);
                  rows.push([dateStr, id, item.name, "SALE", Number(item.quantity), gross, 0, 0, "", 0, key, 0, 0, 0]);
                }
              });
            }

            // 2. 注文サマリー行（税金・割引・返品・払い戻しを一括計上）
            const sumKey = id + "_SUM";
            if (!existingKeys.has(sumKey)) {
              const totalTax = Number(order.total_tax_money?.amount || 0);
              const totalDisc = Number(order.total_discount_money?.amount || 0);
              
              let retGross = 0, retTax = 0;
              if (order.returns) {
                order.returns.forEach(r => {
                  retGross += Number(r.return_amounts?.gross_return_money?.amount || 0);
                  retTax += Number(r.return_amounts?.tax_money?.amount || 0);
                });
              }
              
              let manualRefund = 0;
              if (order.refunds) {
                order.refunds.forEach(rf => {
                  if (!rf.return_id) manualRefund += Number(rf.amount_money?.amount || 0);
                });
              }
              rows.push([dateStr, id, "注文サマリー", "SUMMARY", 0, 0, totalTax, -totalDisc, "", 0, sumKey, -retGross, -retTax, -manualRefund]);
            }

            // 3. 支払い（決済種別）行
            if (order.tenders) {
              order.tenders.forEach((tender, i) => {
                const key = id + "_T_" + i;
                if (!existingKeys.has(key)) {
                  const payType = getPaymentType(tender);
                  const amt = Number(tender.amount_money?.amount || 0);
                  const fee = Number(tender.processing_fee_money?.amount || 0);
                  rows.push([dateStr, id, "支払い: " + payType, "PAYMENT", 0, 0, 0, 0, payType, fee, key, 0, 0, amt]);
                }
              });
            }
          });
          if (rows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 14).setValues(rows);
        }
        cursor = res.cursor;
      } while (cursor);
    }
  } catch(e) { console.error("SQ Error: " + e.message); }
}

function getPaymentType(tender) {
  if (!tender) return "その他";
  const type = tender.type;
  if (type === "CARD") {
    const brand = (tender.card_details?.card_brand || "").toUpperCase();
    if (["ID", "QUICPAY", "SUICA", "PASMO", "ICOCA", "SUGOCA", "NIMOCA", "HAYAKAKEN", "KITACA", "TOICA", "MANACA"].includes(brand)) return "電子マネー";
    return "カード";
  }
  if (type === "CASH") return "現金";
  if (type === "HOUSE_ACCOUNT") return "ハウスアカウント";
  if (type === "EXTERNAL") {
    const src = (tender.external_details?.source_name || "").toUpperCase();
    if (src.includes("AU PAY") || src.includes("AUPAY")) return "au PAY";
    if (src.includes("D払い") || src.includes("DBARAI") || src.includes("D-BARAI") || src.includes("D BARAI")) return "d払い";
    if (src.includes("楽天") || src.includes("RAKUTEN")) return "楽天ペイ";
    if (src.includes("PAYPAY")) return "その他";
    return "電子マネー";
  }
  return "その他";
}

// ============================================================
// 3. 再集計（CSV 35項目完全再現）
// ============================================================
function recalculateAllSummaries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sqSheet = ss.getSheetByName('Square売上データ');
  if (!sqSheet) return;

  const data = sqSheet.getDataRange().getValues();
  let monthly = {};

  data.slice(1).forEach(row => {
    const m = row[0] instanceof Date ? Utilities.formatDate(row[0], "JST", "yyyy-MM") : String(row[0]).substring(0, 7);
    if (!m || m.length < 7) return;

    if (!monthly[m]) {
      monthly[m] = {
        gross:0, tax:0, disc:0, returns:0, refund:0, fees:0, items:0,
        txAll: new Set(), txItems: new Set(), txReturns: new Set(), txDiscounts: new Set(), txTax: new Set(), txTenders: new Set(),
        pay: {"au PAY":0, "d払い":0, "カード":0, "その他":0, "ハウスアカウント":0, "楽天ペイ":0, "現金":0, "電子マネー":0}
      };
    }
    const o = monthly[m];
    const type = row[3];
    const id = String(row[1]);

    if (type === "SALE") {
      o.gross += Number(row[5]);
      o.items += Number(row[4]);
      if (Number(row[5]) > 0) { o.txItems.add(id); o.txAll.add(id); }
    } else if (type === "SUMMARY") {
      o.tax += Number(row[6]);
      o.disc += Number(row[7]);
      o.returns += Number(row[11]);
      o.refund += Number(row[13]);
      if (Number(row[6]) !== 0) o.txTax.add(id);
      if (Number(row[7]) !== 0) o.txDiscounts.add(id);
      if (Number(row[11]) !== 0) o.txReturns.add(id);
    } else if (type === "PAYMENT") {
      const pt = row[8];
      const amt = Number(row[13]);
      if (o.pay.hasOwnProperty(pt)) o.pay[pt] += amt; else o.pay["その他"] += amt;
      o.fees += Number(row[9]);
      o.txTenders.add(id);
    }
  });

  const headers = ["年月", "総売上高", "商品", "サービス料", "返品", "ディスカウントと無料提供", "純売上高", "繰延売上", "ギフトカード売上", "税金", "金額を指定した払い戻し", "売上合計", "受取合計額", "au PAY", "d払い", "カード", "その他", "ハウスアカウント", "楽天ペイ", "現金", "電子マネー", "手数料", "Squareの決済手数料", "Squareの手数料", "合計（純額）", "総売上数", "売上取引履歴", "商品売上取引履歴", "サービス料取引履歴", "商品別返品取引履歴", "ディスカウント取引履歴", "無料提供取引履歴", "ギフトカード売上取引履歴", "税金取引履歴", "総売上取引履歴", "受取合計額の取引履歴"];
  
  const rows = Object.keys(monthly).sort().reverse().map(m => {
    const o = monthly[m];
    const netSales = o.gross + o.returns + o.disc;
    const totalSales = netSales + o.tax + o.refund;
    const collected = totalSales - (o.pay["ハウスアカウント"] || 0);
    
    return [
      m, o.gross, o.gross, 0, o.returns, o.disc, netSales, 0, 0, o.tax, o.refund, totalSales, collected,
      o.pay["au PAY"], o.pay["d払い"], o.pay["カード"], o.pay["その他"], o.pay["ハウスアカウント"], o.pay["楽天ペイ"], o.pay["現金"], o.pay["電子マネー"],
      -o.fees, -o.fees, -o.fees, collected - o.fees,
      o.items, o.txTax.size, o.txItems.size, 0, o.txReturns.size, o.txDiscounts.size, 0, 0, o.txTax.size, o.txItems.size, o.txTenders.size
    ];
  });

  writeToSheet(getOrCreateSheet(ss, 'Square月次売上'), headers, rows);
  SpreadsheetApp.getUi().alert("すべてのレポートを更新しました！😍");
}

function finalizeUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Square売上データ', 'カラーミー売上データ'].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s) sortSheetByDate(s);
  });
  recalculateAllSummaries();
}
function getOrCreateSheet(ss, name) { let s = ss.getSheetByName(name); if (!s) s = ss.insertSheet(name); return s; }
function getExistingKeys(sheet, keyColumnIndex) { 
  let lr = sheet.getLastRow(); 
  if (lr < 2) return new Set(); 
  return new Set(sheet.getRange(2, keyColumnIndex, lr - 1, 1).getValues().map(r => String(r[0]))); 
}
function sortSheetByDate(sheet) { 
  let lr = sheet.getLastRow(); 
  if (lr < 2) return; 
  sheet.getRange(2, 1, lr - 1, sheet.getLastColumn()).sort({column: 1, ascending: false}); 
}
function writeToSheet(sheet, h, r) { 
  sheet.clear(); 
  sheet.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight("bold").setBackground("#f3f3f3"); 
  if (r.length > 0) sheet.getRange(2, 1, r.length, r[0].length).setValues(r); 
  sheet.setFrozenRows(1); 
}









function testAuto() {
  SpreadsheetApp.getActiveSpreadsheet().
}