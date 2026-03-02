/**
 * インターラーケン売上レポート集計システム
 * SquareサマリーCSVの35項目に完全準拠
 *
 * 前提: @types/google-apps-script がインストール済み
 *   npm install --save-dev @types/google-apps-script
 */

// ============================================================
// 型定義 (Type Definitions)
// ============================================================

/** アクセストークンなどの設定情報 */
interface Config {
  readonly SQUARE_ACCESS_TOKEN: string;
  readonly COLORME_ACCESS_TOKEN: string;
  readonly START_DATE: string;
}

/** カラーミーAPIの注文明細 */
interface ColormeDetail {
  id: string | number;
  product_name: string;
  unit_num?: string | number;
  product_num?: string | number;
  price: string | number;
}

/** カラーミーAPIの注文 */
interface ColormeSale {
  id: string | number;
  make_date: string | number;
  details: ColormeDetail[];
  delivery_total: number;
  fee_total: number;
  point_discount: number;
}

/** カラーミーAPIのレスポンス */
interface ColormeResponse {
  sales?: ColormeSale[];
}

/** Square APIの金額オブジェクト */
interface SquareMoney {
  amount?: number;
  currency?: string;
}

/** Square APIの注文明細行 */
interface SquareLineItem {
  name: string;
  quantity: string;
  gross_sales_money?: SquareMoney;
}

/** Square APIの返品 */
interface SquareReturn {
  return_amounts?: {
    gross_return_money?: SquareMoney;
    tax_money?: SquareMoney;
  };
}

/** Square APIの払い戻し */
interface SquareRefund {
  return_id?: string;
  amount_money?: SquareMoney;
}

/** Square APIのカード情報 */
interface SquareCardDetails {
  card_brand?: string;
}

/** Square APIの外部決済情報 */
interface SquareExternalDetails {
  source_name?: string;
}

/** Square APIの支払い(Tender) */
interface SquareTender {
  type: string;
  amount_money?: SquareMoney;
  processing_fee_money?: SquareMoney;
  card_details?: SquareCardDetails;
  external_details?: SquareExternalDetails;
}

/** Square APIの注文 */
interface SquareOrder {
  id: string;
  closed_at: string;
  line_items?: SquareLineItem[];
  returns?: SquareReturn[];
  refunds?: SquareRefund[];
  tenders?: SquareTender[];
  total_tax_money?: SquareMoney;
  total_discount_money?: SquareMoney;
}

/** Square APIの注文検索レスポンス */
interface SquareOrdersResponse {
  orders?: SquareOrder[];
  cursor?: string;
}

/** Square APIのロケーション */
interface SquareLocation {
  id: string;
  name?: string;
}

/** Square APIのロケーション一覧レスポンス */
interface SquareLocationsResponse {
  locations: SquareLocation[];
}

/** 月次集計の支払い種別ごとの金額 */
interface MonthlyPayments {
  "au PAY": number;
  "d払い": number;
  "カード": number;
  "その他": number;
  "ハウスアカウント": number;
  "楽天ペイ": number;
  "現金": number;
  "電子マネー": number;
}

/** 月次集計データ */
interface MonthlyData {
  gross: number;
  tax: number;
  disc: number;
  returns: number;
  refund: number;
  fees: number;
  items: number;
  txAll: Set<string>;
  txItems: Set<string>;
  txReturns: Set<string>;
  txDiscounts: Set<string>;
  txTax: Set<string>;
  txTenders: Set<string>;
  pay: MonthlyPayments;
}

/** 月次集計マップ */
type MonthlyMap = Record<string, MonthlyData>;

/** Square売上データシートの行タイプ */
type RowType = "SALE" | "SUMMARY" | "PAYMENT";

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
} as const;

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
} as const;

// ============================================================
// 設定 (Config)
// ============================================================

const CONFIG: Config = {
  SQUARE_ACCESS_TOKEN: 'EAAAl3VlBqnOihdeDGqTuOyfuE8juXQrSNR6cgpX-RDtVxxFyr4d7daw5jil-oow',
  COLORME_ACCESS_TOKEN: '4fd03a83f636c4517b72bf23cde52b797fff500263e34e3f26ac2c26f3c10ee7',
  START_DATE: '2026-02-01',
};

// ============================================================
// エントリーポイント (Entry Points)
// ============================================================

/** スプレッドシートを開いたときにカスタムメニューを追加する */
function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu('🚀インターラーケン操作')
    .addItem('1. カラーミー売上を更新', 'runColormeUpdate')
    .addItem('2. Square売上を更新', 'runSquareUpdate')
    .addSeparator()
    .addItem('3. レポートを再集計', 'recalculateAllSummaries')
    .addToUi();
}

/** カラーミー売上を更新してレポートを最終化する */
function runColormeUpdate(): void {
  updateColormeSalesMaster(CONFIG.START_DATE);
  finalizeUpdate();
}

/** Square売上を更新してレポートを最終化する */
function runSquareUpdate(): void {
  updateSquareSalesMaster(CONFIG.START_DATE);
  finalizeUpdate();
}

// ============================================================
// 1. カラーミー関連 (Colorme)
// ============================================================

/**
 * カラーミー売上データマスタを更新する
 * @param startDate - 取得開始日 (例: "2026-02-01")
 */
function updateColormeSalesMaster(startDate: string): void {
  const ss: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet: GoogleAppsScript.Spreadsheet.Sheet = getOrCreateSheet(ss, 'カラーミー売上データ');
  const existingKeys: Set<string> = getExistingKeys(sheet, CM_COL.KEY + 1); // 1始まり列番号
  let offset = 0;

  try {
    while (true) {
      const url = `https://api.shop-pro.jp/v1/sales.json?make_date_min=${startDate}&limit=100&offset=${offset}`;
      const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        headers: { Authorization: `Bearer ${CONFIG.COLORME_ACCESS_TOKEN}` },
      };
      const res = JSON.parse(
        UrlFetchApp.fetch(url, options).getContentText()
      ) as ColormeResponse;

      if (!res.sales || res.sales.length === 0) break;

      const newRows: (string | number)[][] = [];

      for (const sale of res.sales) {
        const saleDate: string = parseSaleDate(sale.make_date);

        for (const detail of sale.details) {
          const qty: number = Number(detail.unit_num) || Number(detail.product_num) || 1;
          const price: number = Number(detail.price);
          const key = `${sale.id}_D_${detail.id}`;

          if (!existingKeys.has(key)) {
            newRows.push([
              saleDate,
              sale.id,
              detail.product_name,
              'SALE',
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
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 11).setValues(newRows);
      }

      if (res.sales.length < 100) break;
      offset += 100;
    }
  } catch (e) {
    console.error(`CM Error: ${(e as Error).message}`);
  }
}

/**
 * カラーミーの sale.make_date を "yyyy-MM-dd" 形式に変換する
 * @param raw - UNIX秒タイムスタンプ(数値) または "yyyy-MM-dd HH:mm:ss" 文字列
 */
function parseSaleDate(raw: string | number): string {
  if (typeof raw === 'number') {
    return Utilities.formatDate(new Date(raw * 1000), 'JST', 'yyyy-MM-dd');
  }
  return raw.split(' ')[0];
}

// ============================================================
// 2. Square関連 (Square)
// ============================================================

/**
 * Square売上データマスタを更新する
 * @param startDate - 取得開始日 (例: "2026-02-01")
 */
function updateSquareSalesMaster(startDate: string): void {
  const ss: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet: GoogleAppsScript.Spreadsheet.Sheet = getOrCreateSheet(ss, 'Square売上データ');
  const existingKeys: Set<string> = getExistingKeys(sheet, SQ_COL.KEY + 1);
  const startAt: string = new Date(`${startDate}T00:00:00+09:00`).toISOString();

  const sqHeaders: GoogleAppsScript.URL_Fetch.HttpHeaders = {
    Authorization: `Bearer ${CONFIG.SQUARE_ACCESS_TOKEN}`,
    'Content-Type': 'application/json',
  };

  try {
    const locRes = JSON.parse(
      UrlFetchApp.fetch('https://connect.squareup.com/v2/locations', {
        headers: { Authorization: `Bearer ${CONFIG.SQUARE_ACCESS_TOKEN}` },
      }).getContentText()
    ) as SquareLocationsResponse;

    for (const loc of locRes.locations) {
      let cursor: string | null | undefined = null;

      do {
        const payload = {
          location_ids: [loc.id],
          query: {
            filter: {
              closed_at: { start_at: startAt },
              state_filter: { states: ['COMPLETED'] },
            },
          },
          cursor: cursor ?? undefined,
        };

        const res = JSON.parse(
          UrlFetchApp.fetch('https://connect.squareup.com/v2/orders/search', {
            method: 'post',
            headers: sqHeaders,
            payload: JSON.stringify(payload),
          }).getContentText()
        ) as SquareOrdersResponse;

        if (res.orders) {
          const newRows: (string | number)[][] = buildSquareRows(res.orders, existingKeys);
          if (newRows.length > 0) {
            sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 14).setValues(newRows);
          }
        }

        cursor = res.cursor;
      } while (cursor);
    }
  } catch (e) {
    console.error(`SQ Error: ${(e as Error).message}`);
  }
}

/**
 * SquareのOrderリストからシートに書き込む行データを生成する
 */
function buildSquareRows(
  orders: SquareOrder[],
  existingKeys: Set<string>
): (string | number)[][] {
  const rows: (string | number)[][] = [];

  for (const order of orders) {
    const dateStr: string = Utilities.formatDate(new Date(order.closed_at), 'JST', 'yyyy-MM-dd');
    const id: string = order.id;

    // 1. 商品売上行 (SALE)
    if (order.line_items) {
      order.line_items.forEach((item: SquareLineItem, i: number) => {
        const key = `${id}_L_${i}`;
        if (!existingKeys.has(key)) {
          const gross: number = item.gross_sales_money?.amount ?? 0;
          rows.push([dateStr, id, item.name, 'SALE', Number(item.quantity), gross, 0, 0, '', 0, key, 0, 0, 0]);
        }
      });
    }

    // 2. 注文サマリー行 (SUMMARY)
    const sumKey = `${id}_SUM`;
    if (!existingKeys.has(sumKey)) {
      const totalTax: number = order.total_tax_money?.amount ?? 0;
      const totalDisc: number = order.total_discount_money?.amount ?? 0;

      let retGross = 0;
      let retTax = 0;
      order.returns?.forEach((r: SquareReturn) => {
        retGross += r.return_amounts?.gross_return_money?.amount ?? 0;
        retTax += r.return_amounts?.tax_money?.amount ?? 0;
      });

      let manualRefund = 0;
      order.refunds?.forEach((rf: SquareRefund) => {
        if (!rf.return_id) manualRefund += rf.amount_money?.amount ?? 0;
      });

      rows.push([dateStr, id, '注文サマリー', 'SUMMARY', 0, 0, totalTax, -totalDisc, '', 0, sumKey, -retGross, -retTax, -manualRefund]);
    }

    // 3. 支払い行 (PAYMENT)
    order.tenders?.forEach((tender: SquareTender, i: number) => {
      const key = `${id}_T_${i}`;
      if (!existingKeys.has(key)) {
        const payType: string = getPaymentType(tender);
        const amt: number = tender.amount_money?.amount ?? 0;
        const fee: number = tender.processing_fee_money?.amount ?? 0;
        rows.push([dateStr, id, `支払い: ${payType}`, 'PAYMENT', 0, 0, 0, 0, payType, fee, key, 0, 0, amt]);
      }
    });
  }

  return rows;
}

/**
 * SquareのTenderオブジェクトから支払い種別名を返す
 * @param tender - Square Tender オブジェクト
 * @returns 支払い種別の日本語名
 */
function getPaymentType(tender: SquareTender): string {
  const ELECTRONIC_MONEY_BRANDS = new Set([
    'ID', 'QUICPAY', 'SUICA', 'PASMO', 'ICOCA',
    'SUGOCA', 'NIMOCA', 'HAYAKAKEN', 'KITACA', 'TOICA', 'MANACA',
  ]);

  switch (tender.type) {
    case 'CARD': {
      const brand = (tender.card_details?.card_brand ?? '').toUpperCase();
      return ELECTRONIC_MONEY_BRANDS.has(brand) ? '電子マネー' : 'カード';
    }
    case 'CASH':
      return '現金';
    case 'HOUSE_ACCOUNT':
      return 'ハウスアカウント';
    case 'EXTERNAL': {
      const src = (tender.external_details?.source_name ?? '').toUpperCase();
      if (src.includes('AU PAY') || src.includes('AUPAY')) return 'au PAY';
      if (src.includes('D払い') || src.includes('DBARAI') || src.includes('D-BARAI') || src.includes('D BARAI')) return 'd払い';
      if (src.includes('楽天') || src.includes('RAKUTEN')) return '楽天ペイ';
      if (src.includes('PAYPAY')) return 'その他';
      return '電子マネー';
    }
    default:
      return 'その他';
  }
}

// ============================================================
// 3. 再集計 (Recalculate - CSV 35項目完全再現)
// ============================================================

/** すべての月次集計を再計算してSheetに書き出す */
function recalculateAllSummaries(): void {
  const ss: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sqSheet: GoogleAppsScript.Spreadsheet.Sheet | null = ss.getSheetByName('Square売上データ');
  if (!sqSheet) return;

  const data: unknown[][] = sqSheet.getDataRange().getValues();
  const monthly: MonthlyMap = aggregateMonthlyData(data.slice(1)); // ヘッダー行をスキップ

  const headers: string[] = [
    '年月', '総売上高', '商品', 'サービス料', '返品', 'ディスカウントと無料提供',
    '純売上高', '繰延売上', 'ギフトカード売上', '税金', '金額を指定した払い戻し',
    '売上合計', '受取合計額',
    'au PAY', 'd払い', 'カード', 'その他', 'ハウスアカウント', '楽天ペイ', '現金', '電子マネー',
    '手数料', 'Squareの決済手数料', 'Squareの手数料', '合計（純額）',
    '総売上数', '売上取引履歴', '商品売上取引履歴', 'サービス料取引履歴',
    '商品別返品取引履歴', 'ディスカウント取引履歴', '無料提供取引履歴',
    'ギフトカード売上取引履歴', '税金取引履歴', '総売上取引履歴', '受取合計額の取引履歴',
  ];

  const rows: (string | number)[][] = Object.keys(monthly)
    .sort()
    .reverse()
    .map((m: string) => buildSummaryRow(m, monthly[m]));

  writeToSheet(getOrCreateSheet(ss, 'Square月次売上'), headers, rows);
  SpreadsheetApp.getUi().alert('すべてのレポートを更新しました！😍');
}

/**
 * Square売上データシートの行データから月次集計マップを構築する
 */
function aggregateMonthlyData(rows: unknown[][]): MonthlyMap {
  const monthly: MonthlyMap = {};

  for (const row of rows) {
    const dateVal = row[SQ_COL.DATE];
    const m: string = dateVal instanceof Date
      ? Utilities.formatDate(dateVal, 'JST', 'yyyy-MM')
      : String(dateVal).substring(0, 7);

    if (!m || m.length < 7) continue;

    if (!monthly[m]) {
      monthly[m] = {
        gross: 0, tax: 0, disc: 0, returns: 0, refund: 0, fees: 0, items: 0,
        txAll: new Set(), txItems: new Set(), txReturns: new Set(),
        txDiscounts: new Set(), txTax: new Set(), txTenders: new Set(),
        pay: { 'au PAY': 0, 'd払い': 0, 'カード': 0, 'その他': 0, 'ハウスアカウント': 0, '楽天ペイ': 0, '現金': 0, '電子マネー': 0 },
      };
    }

    const o: MonthlyData = monthly[m];
    const type = row[SQ_COL.TYPE] as RowType;
    const id = String(row[SQ_COL.ORDER_ID]);

    switch (type) {
      case 'SALE':
        o.gross += Number(row[SQ_COL.GROSS]);
        o.items += Number(row[SQ_COL.QTY]);
        if (Number(row[SQ_COL.GROSS]) > 0) {
          o.txItems.add(id);
          o.txAll.add(id);
        }
        break;

      case 'SUMMARY':
        o.tax += Number(row[SQ_COL.TAX]);
        o.disc += Number(row[SQ_COL.DISC]);
        o.returns += Number(row[SQ_COL.RETURN_GROSS]);
        o.refund += Number(row[SQ_COL.AMOUNT]);
        if (Number(row[SQ_COL.TAX]) !== 0) o.txTax.add(id);
        if (Number(row[SQ_COL.DISC]) !== 0) o.txDiscounts.add(id);
        if (Number(row[SQ_COL.RETURN_GROSS]) !== 0) o.txReturns.add(id);
        break;

      case 'PAYMENT': {
        const pt = row[SQ_COL.PAY_TYPE] as keyof MonthlyPayments;
        const amt = Number(row[SQ_COL.AMOUNT]);
        if (pt in o.pay) {
          o.pay[pt] += amt;
        } else {
          o.pay['その他'] += amt;
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
 */
function buildSummaryRow(month: string, o: MonthlyData): (string | number)[] {
  const netSales: number = o.gross + o.returns + o.disc;
  const totalSales: number = netSales + o.tax + o.refund;
  const collected: number = totalSales - (o.pay['ハウスアカウント'] ?? 0);

  return [
    month, o.gross, o.gross, 0, o.returns, o.disc, netSales, 0, 0, o.tax, o.refund,
    totalSales, collected,
    o.pay['au PAY'], o.pay['d払い'], o.pay['カード'], o.pay['その他'],
    o.pay['ハウスアカウント'], o.pay['楽天ペイ'], o.pay['現金'], o.pay['電子マネー'],
    -o.fees, -o.fees, -o.fees, collected - o.fees,
    o.items, o.txTax.size, o.txItems.size, 0, o.txReturns.size, o.txDiscounts.size, 0, 0,
    o.txTax.size, o.txItems.size, o.txTenders.size,
  ];
}

// ============================================================
// 4. 共通ユーティリティ (Utilities)
// ============================================================

/**
 * 全シートのデータをソートして月次集計を再実行する
 */
function finalizeUpdate(): void {
  const ss: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  for (const name of ['Square売上データ', 'カラーミー売上データ']) {
    const s: GoogleAppsScript.Spreadsheet.Sheet | null = ss.getSheetByName(name);
    if (s) sortSheetByDate(s);
  }

  recalculateAllSummaries();
}

/**
 * シートを名前で取得し、存在しない場合は新規作成する
 */
function getOrCreateSheet(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  name: string
): GoogleAppsScript.Spreadsheet.Sheet {
  return ss.getSheetByName(name) ?? ss.insertSheet(name);
}

/**
 * シートの指定列から既存のキーをSetで返す
 * @param sheet - 対象シート
 * @param keyColumn - キーが入っている列番号 (1始まり)
 */
function getExistingKeys(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  keyColumn: number
): Set<string> {
  const lastRow: number = sheet.getLastRow();
  if (lastRow < 2) return new Set<string>();
  return new Set<string>(
    sheet.getRange(2, keyColumn, lastRow - 1, 1).getValues().map((r: unknown[]) => String(r[0]))
  );
}

/**
 * シートのデータ行を日付の降順でソートする (2行目以降)
 */
function sortSheetByDate(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  const lastRow: number = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .sort({ column: 1, ascending: false });
}

/**
 * シートをクリアしてヘッダーとデータを書き込む
 */
function writeToSheet(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  headers: string[],
  rows: (string | number)[][]
): void {
  sheet.clear();
  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#f3f3f3');
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
  sheet.setFrozenRows(1);
}