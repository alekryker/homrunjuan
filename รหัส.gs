/**
 * homrunjuan – Inventory Web App (Apps Script)
 * Backend APIs + Sheet bootstrap + Document numbering
 */

const SHEETS = { 
  PRODUCTS: 'Products',
  MOVES: 'StockMovements',
};
const TZ = Session.getScriptTimeZone() || 'Asia/Bangkok';

/** Create sheets & headers if not exists */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SHEETS.PRODUCTS, [
    'productCode','name','unit','sellPrice','buyPrice','minQty','createdAt','updatedAt','active'
  ]);
  ensureSheet_(ss, SHEETS.MOVES, [
    'dateISO','docType','docNo','productCode','qtyChange','unit','price','note'
  ]);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const range = sh.getRange(1,1,1,headers.length);
  const values = range.getValues()[0];
  let needs = false;
  for (let i=0;i<headers.length;i++) if (values[i] !== headers[i]) { needs = true; break; }
  if (needs) {
    range.setValues([headers]);
    sh.setFrozenRows(1);
  }
}

/** Serve UI */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('homrunjuan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Utilities */
function nowISO_() {
  return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function getNextDocNo_(prefix) { // S = Sale, P = Purchase
  const d = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd');
  const key = `${prefix}_${d}_SEQ`;
  const props = PropertiesService.getDocumentProperties();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    let seq = Number(props.getProperty(key) || '0');
    seq++;
    props.setProperty(key, String(seq));
    return `${prefix}-${d}-${('000' + seq).slice(-3)}`;
  } finally {
    lock.releaseLock();
  }
}

/** PRODUCTS */
function listProducts() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.PRODUCTS);
  const rng = sh.getRange(2,1,Math.max(sh.getLastRow()-1,0),9);
  const rows = rng.getValues().filter(r => r[0]);
  return rows.map(r => ({
    productCode: String(r[0]),
    name:        r[1],
    category:    r[2],
    sellPrice:   Number(r[3]||0),
    buyPrice:    Number(r[4]||0),
    minQty:      Number(r[5]||0),
    createdAt:   r[6],
    updatedAt:   r[7],
    active:      r[8] === '' ? true : !!r[8],
  }));
}


function saveProduct(p) {
  if (!p || !p.productCode) throw new Error('กรุณาระบุรหัสสินค้า');
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.PRODUCTS);
  const data = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),9).getValues();
  const idx = data.findIndex(r => String(r[0]) === String(p.productCode));
  const t = nowISO_();
  const row = [
    p.productCode,
    p.name||'',
    p.unit||'',
    Number(p.sellPrice||0),
    Number(p.buyPrice||0),
    Number(p.minQty||0),
    idx === -1 ? t : (data[idx][6] || t),
    t,
    p.active===false?false:true,
  ];
  if (idx === -1) {
    sh.appendRow(row);
  } else {
    sh.getRange(idx+2,1,1,9).setValues([row]);
  }
  return { ok:true };
}

/** STOCK MOVEMENTS helpers */
function appendMovements_(items, docType, docNo) {
  if (!Array.isArray(items) || items.length===0) throw new Error('No items');
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.MOVES);
  const t = nowISO_();
  const rows = items.map(it => [
    t,
    docType,
    docNo,
    it.productCode,
    docType==='SALE' ? -Math.abs(Number(it.qty||0)) : Math.abs(Number(it.qty||0)),
    it.unit||'',
    Number(it.price||0),
    it.note||''
  ]);
  sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);
}

/** Create SALE or PURCHASE document */
function createSale(doc) {
  // doc: { items:[{productCode, price, qty, note}], note }
  if (!doc || !Array.isArray(doc.items) || doc.items.length===0) throw new Error('ไม่มีรายการขาย');
  const products = listProducts();
  const map = new Map(products.map(p => [String(p.productCode), p]));
  const cleaned = doc.items.map(x => {
    const p = map.get(String(x.productCode));
    if (!p) throw new Error(`ไม่พบสินค้า: ${x.productCode}`);
    return {
      productCode: String(x.productCode),
      unit: p.unit||'',
      price: Number(x.price||p.sellPrice||0),
      qty: Number(x.qty||0),
      note: x.note||''
    };
  });
  const docNo = getNextDocNo_('S');
  appendMovements_(cleaned, 'SALE', docNo);
  const total = cleaned.reduce((s,it)=> s + it.price*it.qty, 0);
  return { ok:true, docNo, total };
}

function createPurchase(doc) {
  if (!doc || !Array.isArray(doc.items) || doc.items.length===0) throw new Error('ไม่มีรายการซื้อ');
  const products = listProducts();
  const map = new Map(products.map(p => [String(p.productCode), p]));
  const cleaned = doc.items.map(x => {
    const p = map.get(String(x.productCode));
    if (!p) throw new Error(`ไม่พบสินค้า: ${x.productCode}`);
    return {
      productCode: String(x.productCode),
      unit: p.unit||'',
      price: Number(x.price||p.buyPrice||0),
      qty: Number(x.qty||0),
      note: x.note||''
    };
  });
  const docNo = getNextDocNo_('P');
  appendMovements_(cleaned, 'PURCHASE', docNo);
  const total = cleaned.reduce((s,it)=> s + it.price*it.qty, 0);
  return { ok:true, docNo, total };
}

/** Reports */
function getStockReport(opts) {
  opts = opts || {}; // { lowOnly:boolean, sort:'desc'|'asc' }
  const products = listProducts();
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.MOVES);
  const rng = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),8);
  const rows = rng.getValues();
  const sum = new Map();
  for (const r of rows) {
    const code = String(r[3]);
    const qty = Number(r[4]||0);
    if (!code) continue;
    sum.set(code, (sum.get(code)||0) + qty);
  }
  const out = products.map(p => {
    const qty = Number(sum.get(String(p.productCode))||0);
    const low = qty <= Number(p.minQty||0);
    return {
      productCode: p.productCode,
      name: p.name,
      category: p.category,
      stock: qty,
      minQty: Number(p.minQty||0),
      low,
    };
  }).filter(x => opts.lowOnly ? x.low : true);

  out.sort((a,b) => (opts.sort==='asc' ? a.stock-b.stock : b.stock-a.stock));
  return out;
}

function getSalesReport(month) {
  // month: 'YYYY-MM' (optional). If missing, return current month
  const d = new Date();
  const current = Utilities.formatDate(d, TZ, 'yyyy-MM');
  const target = month || current;
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.MOVES);
  const rng = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),8).getValues();
  const daily = new Map(); // key = yyyy-MM-dd -> {qty, amount}
  for (const r of rng) {
    const dateISO = r[0];
    const type = r[1];
    if (type !== 'SALE') continue;
    const dateStr = String(dateISO||'').slice(0,10);
    if (!dateStr.startsWith(target)) continue;
    const qty = Math.abs(Number(r[4]||0));
    const price = Number(r[6]||0);
    const amt = qty * price;
    const cur = daily.get(dateStr) || { qty:0, amount:0 };
    cur.qty += qty;
    cur.amount += amt;
    daily.set(dateStr, cur);
  }
  const out = Array.from(daily.entries()).map(([date, v]) => ({date, qty:v.qty, amount:v.amount}));
  out.sort((a,b)=> a.date.localeCompare(b.date));
  return { month: target, days: out };
}

function getPurchaseReport(month) {
  const d = new Date();
  const current = Utilities.formatDate(d, TZ, 'yyyy-MM');
  const target = month || current;
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.MOVES);
  const rng = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),8).getValues();
  const daily = new Map();
  for (const r of rng) {
    const dateISO = r[0];
    const type = r[1];
    if (type !== 'PURCHASE') continue;
    const dateStr = String(dateISO||'').slice(0,10);
    if (!dateStr.startsWith(target)) continue;
    const qty = Math.abs(Number(r[4]||0));
    const price = Number(r[6]||0);
    const amt = qty * price;
    const cur = daily.get(dateStr) || { qty:0, amount:0 };
    cur.qty += qty;
    cur.amount += amt;
    daily.set(dateStr, cur);
  }
  const out = Array.from(daily.entries()).map(([date, v]) => ({date, qty:v.qty, amount:v.amount}));
  out.sort((a,b)=> a.date.localeCompare(b.date));
  return { month: target, days: out };
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SHEETS.PRODUCTS, [
    'productCode','name','category','sellPrice','buyPrice','minQty','createdAt','updatedAt','active'
  ]);
  ensureSheet_(ss, SHEETS.MOVES, [
    'dateISO','docType','docNo','productCode','qtyChange','price','note'
  ]);
}

function listProducts() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.PRODUCTS);
  const rng = sh.getRange(2,1,Math.max(sh.getLastRow()-1,0),9);
  const rows = rng.getValues().filter(r => r[0]);
  return rows.map(r => ({
    productCode: String(r[0]),
    name: r[1],
    category: r[2],
    sellPrice: Number(r[3]||0),
    buyPrice: Number(r[4]||0),
    minQty: Number(r[5]||0),
    createdAt: r[6],
    updatedAt: r[7],
    active: r[8] === '' ? true : !!r[8],
  }));
}

function saveProduct(p) {
  if (!p || !p.productCode) throw new Error('กรุณาระบุรหัสสินค้า');
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.PRODUCTS);
  const data = sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),9).getValues();
  const idx = data.findIndex(r => String(r[0]) === String(p.productCode));
  const t = nowISO_();
  const row = [
    p.productCode,
    p.name||'',
    p.category||'',
    Number(p.sellPrice||0),
    Number(p.buyPrice||0),
    Number(p.minQty||0),
    idx === -1 ? t : (data[idx][6] || t),
    t,
    p.active===false?false:true,
  ];
  if (idx === -1) {
    sh.appendRow(row);
  } else {
    sh.getRange(idx+2,1,1,9).setValues([row]);
  }
  return { ok:true };
}

function normalizeCategories(){
  const sh = SpreadsheetApp.getActive().getSheetByName('Products');
  if (!sh) throw new Error('ไม่พบชีต Products');
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const catIdx = hdr.indexOf('category');
  if (catIdx === -1) throw new Error('ไม่พบคอลัมน์ category');

  const last = sh.getLastRow();
  if (last < 2) return;
  const rng = sh.getRange(2,1,last-1,sh.getLastColumn());
  const data = rng.getValues();

  // map เพื่อรวมหมวดที่ต่างกันแค่ช่องว่าง/ตัวพิมพ์ ให้เป็นรูปแบบแรกที่พบ
  const seen = {};
  for (let i=0;i<data.length;i++){
    let c = String(data[i][catIdx]||'').trim();
    const key = c.toLowerCase();
    if (!c) continue;
    if (!seen[key]) seen[key] = c;      // จดรูปแบบแรกไว้
    data[i][catIdx] = seen[key];        // เขียนกลับเป็นรูปเดียวกัน
  }
  rng.setValues(data);
}



/** Include for HTML templating */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
