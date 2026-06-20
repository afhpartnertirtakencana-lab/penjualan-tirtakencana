// ══════════════════════════════════════════════════════════════
//  TIRTA KENCANA — Google Apps Script Backend (v8.3 Optimized)
//  - Performance optimizations: batch reads/writes, reduced API calls
// ══════════════════════════════════════════════════════════════
var SPREADSHEET_ID = '12q8SwBtoww9Y9c6EZ46-SNa1q5TKIXnf9g_3wLbsNK4';
var SS;
var SHEET_CACHE = {};
var HEADER_CACHE = {};

function getSpreadsheet() {
  if (!SS) {
    try {
      if (SPREADSHEET_ID) SS = SpreadsheetApp.openById(SPREADSHEET_ID);
      else SS = SpreadsheetApp.getActiveSpreadsheet();
    } catch (e) {
      throw new Error('Gagal membuka Spreadsheet: ' + e.message);
    }
  }
  return SS;
}

var SHEET_NAMES = {
  TRX       : ['Transaksi',    'Transactions'],
  DETAIL    : ['DetailTrx',    'TrxDetail'],
  PRODUCTS  : ['Produk',       'Products'],
  CUSTOMERS : ['Pelanggan',    'Customers'],
  SETTINGS  : ['Pengaturan',   'Settings'],
  BARANG    : ['InputBarang',  'BarangMasuk'],
  SETORAN   : ['Setoran',      'Setoran'],
  USERS     : ['Users',        'User'],
  MASTERSTOCK : ['MasterStock', 'MasterStock'],
  DRIVERS   : ['Driver',       'Drivers']
};

function sheetName(candidates) {
  if (!candidates || !Array.isArray(candidates)) return 'Sheet1';
  for (var i = 0; i < candidates.length; i++) {
    try { if (SS.getSheetByName(candidates[i])) return candidates[i]; } catch (e) {}
  }
  return candidates[0];
}

var SHEET = {
  TRX      : sheetName(SHEET_NAMES.TRX),
  DETAIL   : sheetName(SHEET_NAMES.DETAIL),
  PRODUCTS : sheetName(SHEET_NAMES.PRODUCTS),
  CUSTOMERS: sheetName(SHEET_NAMES.CUSTOMERS),
  SETTINGS : sheetName(SHEET_NAMES.SETTINGS),
  BARANG   : sheetName(SHEET_NAMES.BARANG),
  SETORAN  : sheetName(SHEET_NAMES.SETORAN),
  USERS    : sheetName(SHEET_NAMES.USERS),
  MASTERSTOCK : sheetName(SHEET_NAMES.MASTERSTOCK),
  DRIVERS  : sheetName(SHEET_NAMES.DRIVERS)
};

// ══════════════════════════════════════════════════════════════
//  doGet — JSONP entry point
// ══════════════════════════════════════════════════════════════
function doGet(e) {
  if (!e || !e.parameter) {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, data: 'Tirta Kencana GAS v8.2' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var cb   = e.parameter.callback || '';
  var fn   = e.parameter.fn       || '';
  var args = [];
  try { args = JSON.parse(decodeURIComponent(e.parameter.args || '[]')); } catch(x) {
    try { args = JSON.parse(e.parameter.args || '[]'); } catch(x2) { args = []; }
  }
  var result, isOk = true, errMsg = '';
  try {
    switch (fn) {
      case 'getTrxList'           : result = getTrxList();                        break;
      case 'saveTrx'              : result = saveTrx(args[0]);                    break;
      case 'updateStatus'         : result = updateStatus(args[0], args[1]);      break;
      case 'deleteTrx'            : result = deleteTrx(args[0]);                  break;
      case 'getProducts'          : result = getProducts();                        break;
      case 'saveProducts'         : result = saveProducts(args[0]);               break;
      case 'getCustomers'         : result = getCustomers();                       break;
      case 'addCustomer'          : result = addCustomer(args[0]);                break;
      case 'saveCustomers'        : result = saveCustomers(args[0]);              break;
      case 'getSettings'          : result = getSettings();                        break;
      case 'saveSettings'         : result = saveSettings(args[0]);               break;
      case 'saveInputBarang'      : result = saveInputBarang(args[0]);            break;
      case 'getInputBarangHistory': result = getInputBarangHistory();             break;
      case 'saveSetoran'          : result = saveSetoran(args[0]);                break;
      case 'getSetoranHistory'    : result = getSetoranHistory();                 break;
      case 'uploadFoto'           : result = uploadFoto(args[0]);                 break;
      case 'getFotoList'          : result = getFotoList(args[0]);                break;
      case 'saveMasterStock'      : result = saveMasterStock(args[0]);            break;
      case 'getMasterStockByDate' : result = getMasterStockByDate(args[0]);       break;
      case 'getStockBarang'       : result = getStockBarang();                    break;
      case 'getUsers'             : result = getUsers();                           break;
      case 'saveUser'             : result = saveUser(args[0]);                   break;
      case 'deleteUser'           : result = deleteUser(args[0]);                 break;
      case 'getDrivers'           : result = getDrivers();                         break;
      case 'saveDrivers'          : result = saveDrivers(args[0]);                 break;
      case 'updateInputBarang'    : result = updateInputBarang(args[0]);          break;
      case 'deleteInputBarang'    : result = deleteInputBarang(args[0]);          break;
      case 'getTrxDetail'         : result = getTrxDetail(args[0]);               break;
      case 'updateStockAwal'      : result = updateStockAwal(args[0]);            break;
      default: throw new Error('Fungsi tidak dikenal: ' + fn);
    }
  } catch(err) { isOk = false; errMsg = err.message; }
  var payload = isOk ? JSON.stringify({ ok: true, data: result }) : JSON.stringify({ ok: false, error: errMsg });
  if (cb) return ContentService.createTextOutput(cb + '(' + payload + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════════
//  HELPER - OPTIMIZED WITH CACHING
// ══════════════════════════════════════════════════════════════
function getSheet(name) { 
  if (!SHEET_CACHE[name]) {
    try {
      var sh = SS.getSheetByName(name); 
      if (!sh) sh = SS.insertSheet(name); 
      SHEET_CACHE[name] = sh;
    } catch(e) {
      throw new Error('Gagal mengakses sheet "' + name + '": ' + e.message);
    }
  }
  return SHEET_CACHE[name];
}

function getSheetData(name) { 
  var sheet = getSheet(name);
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) return [];
  return sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
}

function getHeaders(name) {
  if (HEADER_CACHE[name]) return HEADER_CACHE[name];
  var sheet = getSheet(name);
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) return [];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  HEADER_CACHE[name] = headers.map(function(c) { return String(c||'').toLowerCase().trim(); });
  return HEADER_CACHE[name];
}

function clearHeaderCache() {
  HEADER_CACHE = {};
}
function formatDateCell(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Jakarta', 'yyyy-MM-dd');
  var s = String(val).trim(); if (s.length >= 10) return s.substring(0,10); return s;
}

// ══════════════════════════════════════════════════════════════
//  DRIVERS
// ══════════════════════════════════════════════════════════════
function getDrivers() {
  try {
    var sheet = getSheet(SHEET.DRIVERS);
    var data = sheet.getDataRange().getValues();
    if (data.length === 0 || String(data[0][0] || '').toLowerCase() !== 'nama driver') {
      sheet.clearContents();
      sheet.appendRow(['Nama Driver', 'Status', 'Tanggal Dibuat']);
      ['oji', 'padong', 'said', 'dedi', 'zehpudin'].forEach(function(d) {
        sheet.appendRow([d, 'aktif', Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd')]);
      });
      data = sheet.getDataRange().getValues();
    }
    var drivers = [];
    for (var i = 1; i < data.length; i++) {
      var nama = String(data[i][0] || '').trim();
      if (nama) drivers.push(nama);
    }
    return drivers;
  } catch(e) {
    throw new Error('getDrivers gagal: ' + e.message);
  }
}

function saveDrivers(driverList) {
  try {
    // Sama persis pola saveCustomers — args[0] bisa array atau string JSON
    if (typeof driverList === 'string') {
      try { driverList = JSON.parse(driverList); } catch(e) { driverList = [driverList]; }
    }
    if (!Array.isArray(driverList)) driverList = driverList ? [String(driverList)] : [];
    var sheet = getSheet(SHEET.DRIVERS);
    sheet.clearContents();
    sheet.appendRow(['Nama Driver', 'Status', 'Tanggal Dibuat']);
    var now = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
    driverList.forEach(function(nama) {
      if (String(nama || '').trim()) sheet.appendRow([String(nama).trim(), 'aktif', now]);
    });
    return { ok: true, count: driverList.length };
  } catch(e) {
    throw new Error('saveDrivers gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  SETTINGS
// ══════════════════════════════════════════════════════════════
function getSettings() {
  try {
    var data = getSheetData(SHEET.SETTINGS);
    var cfg  = {};
    for (var i = 0; i < data.length; i++) {
      var key = String(data[i][0]||'').trim();
      var val = String(data[i][1]||'').trim();
      if (!key) continue;
      if (['bank1','bank2','salesList'].indexOf(key) >= 0) {
        try { cfg[key] = JSON.parse(val); } catch(e) { cfg[key] = (key==='salesList') ? [] : {}; }
      } else { cfg[key] = val; }
    }
    if (!cfg.storeName) cfg.storeName = 'Tirta Kencana';
    if (!cfg.tagline)   cfg.tagline   = 'Distributor Air Minum Terpercaya';
    if (!cfg.bank1)     cfg.bank1     = { nama:'BCA', norek:'6930099099', penerima:'Hendri' };
    if (!cfg.bank2)     cfg.bank2     = { nama:'bluBCA', norek:'002283588888', penerima:'Hendri' };
    if (!cfg.salesList) cfg.salesList = [];
    return cfg;
  } catch(e) {
    throw new Error('getSettings gagal: ' + e.message);
  }
}

function saveSettings(cfg) {
  if (!cfg || typeof cfg !== 'object') throw new Error('Data settings tidak valid');
  var sh = getSheet(SHEET.SETTINGS); sh.clearContents();
  sh.getRange(1,1,8,2).setValues([
    ['storeName', String(cfg.storeName||'')],
    ['tagline',   String(cfg.tagline||'')],
    ['address',   String(cfg.address||'')],
    ['phone',     String(cfg.phone||'')],
    ['footer',    String(cfg.footer||'')],
    ['salesList', JSON.stringify(cfg.salesList || [])],
    ['bank1',     JSON.stringify(cfg.bank1 || {})],
    ['bank2',     JSON.stringify(cfg.bank2 || {})]
  ]);
  return 'ok';
}

// ══════════════════════════════════════════════════════════════
//  USERS
// ══════════════════════════════════════════════════════════════
function getUsers() {
  try {
    var sheet = getSheet(SHEET.USERS);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      var defaultUsers = [
        ['admin','admin','020730'],
        ['hasan','sales','1234'],
        ['ujang','sales','1234'],
        ['oji','driver','1234'],
        ['padong','driver','1234'],
        ['said','driver','1234'],
        ['dedi','driver','1234'],
        ['zehpudin','driver','1234']
      ];
      if (data.length === 0) sheet.appendRow(['Nama','Role','Password']);
      defaultUsers.forEach(function(u) { sheet.appendRow(u); });
      data = sheet.getDataRange().getValues();
    }
    var users = [];
    for (var i = 1; i < data.length; i++) {
      users.push({ name: String(data[i][0]||''), role: String(data[i][1]||'sales'), password: String(data[i][2]||'') });
    }
    return users;
  } catch(e) {
    throw new Error('getUsers gagal: ' + e.message);
  }
}

function saveUser(user) {
  if (!user || !user.name || !user.password) throw new Error('Data user tidak valid');
  var users = getUsers();
  var idx = -1;
  for (var i = 0; i < users.length; i++) { if (users[i].name === user.name) { idx = i; break; } }
  if (idx >= 0) users[idx] = { name: user.name, role: user.role || 'sales', password: user.password };
  else users.push({ name: user.name, role: user.role || 'sales', password: user.password });
  var sh = getSheet(SHEET.USERS); sh.clearContents();
  sh.appendRow(['Nama', 'Role', 'Password']);
  users.forEach(function(u) { sh.appendRow([u.name, u.role, u.password]); });
  return { ok: true, name: user.name };
}

function deleteUser(name) {
  if (!name) throw new Error('Nama user diperlukan');
  var users = getUsers().filter(function(u) { return u.name !== name; });
  var sh = getSheet(SHEET.USERS); sh.clearContents();
  sh.appendRow(['Nama', 'Role', 'Password']);
  users.forEach(function(u) { sh.appendRow([u.name, u.role, u.password]); });
  return { ok: true, deleted: name };
}

// ══════════════════════════════════════════════════════════════
//  PRODUK
// ══════════════════════════════════════════════════════════════
function parseHarga(val) { if (typeof val === 'number') return val; var s = String(val||'').replace(/[^0-9]/g,''); return parseInt(s)||0; }

function ensureProductHeaders() {
  try {
    var sh = getSheet(SHEET.PRODUCTS);
    var data = sh.getDataRange().getValues();
    if (data.length === 0 || String(data[0][0]||'').toLowerCase() !== 'sku') {
      sh.clearContents();
      sh.appendRow(['SKU','Barcode','Nama','Harga','Modal','Satuan','StokAwal','HasBarcode']);
    }
  } catch(e) { throw new Error('Gagal inisialisasi produk: ' + e.message); }
}

function getProducts() {
  try {
    ensureProductHeaders();
    var data = getSheetData(SHEET.PRODUCTS); if (data.length <= 1) return [];
    var headers = getHeaders(SHEET.PRODUCTS);
    var col = {
      sku: headers.indexOf('sku'), barcode: headers.indexOf('barcode'), nama: headers.indexOf('nama'),
      harga: headers.indexOf('harga'), modal: headers.indexOf('modal'), satuan: headers.indexOf('satuan'),
      stokAwal: headers.indexOf('stokawal'), hasBarcode: headers.indexOf('hasbarcode')
    };
    var products = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i]; var sku = String(r[col.sku]||'').trim(); if (!sku) continue;
      products.push({
        sku: sku, barcode: col.barcode >= 0 ? String(r[col.barcode]||'') : '',
        nama: col.nama >= 0 ? String(r[col.nama]||'') : '',
        harga: col.harga >= 0 ? parseHarga(r[col.harga]) : 0,
        modal: col.modal >= 0 ? parseHarga(r[col.modal]) : 0,
        satuan: col.satuan >= 0 ? String(r[col.satuan]||'Pcs') : 'Pcs',
        stokAwal: col.stokAwal >= 0 ? parseInt(r[col.stokAwal])||0 : 0,
        hasBarcode: col.hasBarcode >= 0 ? String(r[col.hasBarcode]||'')==='true' : false
      });
    }
    return products;
  } catch(e) {
    throw new Error('getProducts gagal: ' + e.message);
  }
}

function saveProducts(list) {
  if (!list) return 'error: data kosong';
  if (!Array.isArray(list)) { try { list = JSON.parse(list); } catch(e) { return 'error: format tidak valid'; } }
  var sh = getSheet(SHEET.PRODUCTS); sh.clearContents();
  var headerRow = ['SKU','Barcode','Nama','Harga','Modal','Satuan','StokAwal','HasBarcode'];
  var rows = [headerRow];
  for (var i = 0; i < list.length; i++) {
    var p = list[i];
    rows.push([String(p.sku||''), String(p.barcode||''), String(p.nama||''), Number(p.harga)||0, Number(p.modal)||0, String(p.satuan||'Pcs'), parseInt(p.stokAwal)||0, p.hasBarcode?'true':'false']);
  }
  if (rows.length > 0) sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  return 'ok';
}

// ══════════════════════════════════════════════════════════════
//  CUSTOMERS
// ══════════════════════════════════════════════════════════════
function getCustomers() {
  try {
    var data = getSheetData(SHEET.CUSTOMERS);
    var list = [];
    for (var i = 0; i < data.length; i++) {
      var v = String(data[i][0]||'').trim();
      if (v && v.toLowerCase() !== 'nama' && v.toLowerCase() !== 'pelanggan') list.push(v);
    }
    return list;
  } catch(e) { throw new Error('getCustomers gagal: ' + e.message); }
}
function addCustomer(name) { if (!name) throw new Error('Nama kosong'); if (getCustomers().indexOf(name) >= 0) return 'exists'; getSheet(SHEET.CUSTOMERS).appendRow([name]); return 'ok'; }
function saveCustomers(list) { 
  if (!Array.isArray(list)) throw new Error('Harus array'); 
  var sh = getSheet(SHEET.CUSTOMERS); 
  sh.clearContents(); 
  var rows = [];
  for (var i = 0; i < list.length; i++) { if (list[i]) rows.push([list[i]]); }
  if (rows.length > 0) sh.getRange(1, 1, rows.length, 1).setValues(rows);
  return 'ok'; 
}

// ══════════════════════════════════════════════════════════════
//  TRANSAKSI
// ══════════════════════════════════════════════════════════════
function ensureTrxHeaders() {
  try {
    var sh = getSheet(SHEET.TRX);
    var data = sh.getDataRange().getValues();
    if (!data.length || String(data[0][0]||'').trim().toLowerCase() !== 'id') {
      sh.insertRowBefore(1);
      sh.getRange(1,1,1,9).setValues([['ID','Tanggal','Customer','Sales','Gross','Diskon','Nett','Status','CreatedAt']]);
    }
    var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(c) { return String(c||'').toLowerCase(); });
    if (headers.indexOf('totalmodal') === -1) {
      sh.insertColumnAfter(9); sh.getRange(1,10).setValue('TotalModal');
      sh.insertColumnAfter(10); sh.getRange(1,11).setValue('TotalProfit');
    }
  } catch(e) { throw new Error('Gagal inisialisasi header transaksi: ' + e.message); }
}

function ensureDetailHeaders() {
  try {
    var sh = getSheet(SHEET.DETAIL);
    var data = sh.getDataRange().getValues();
    if (!data.length) {
      sh.appendRow(['TrxID','Tanggal','Customer','Sales','Status','SKU','Barcode','Nama','Qty','Harga','Modal','DiscRpPer','NettPer','Subtotal','Profit','CreatedAt']);
      return;
    }
    var hdr = data[0].map(function(c) { return String(c||'').toLowerCase().trim(); });
    if (hdr.indexOf('trxid') === -1) {
      sh.insertRowBefore(1);
      sh.getRange(1,1,1,16).setValues([['TrxID','Tanggal','Customer','Sales','Status','SKU','Barcode','Nama','Qty','Harga','Modal','DiscRpPer','NettPer','Subtotal','Profit','CreatedAt']]);
    }
  } catch(e) { throw new Error('Gagal inisialisasi header detail: ' + e.message); }
}

function saveTrx(trx) {
  try {
    if (!trx || !trx.id) throw new Error('Data transaksi tidak valid: id kosong');
    ensureTrxHeaders(); ensureDetailHeaders();
    var now = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var items = Array.isArray(trx.items) ? trx.items : [];
    var totalModal = 0, totalProfit = 0;
    
    // Pre-calculate item values to avoid repeated computation
    var processedItems = [];
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var modal = Number(it.modal)||0, qty = Number(it.qty)||1, harga = Number(it.harga)||0, disc = Number(it.discRpPer)||0;
      var nettPer = harga - disc, subtotal = nettPer * qty, profit = (nettPer - modal) * qty;
      totalModal += modal * qty;
      totalProfit += profit;
      processedItems.push({ modal: modal, qty: qty, harga: harga, disc: disc, nettPer: nettPer, subtotal: subtotal, profit: profit, sku: it.sku||'', barcode: it.barcode||'', nama: it.nama||'' });
    }
    
    var trxSheet = getSheet(SHEET.TRX);
    var trxHeaders = getHeaders(SHEET.TRX);
    var newTrxRow = new Array(trxHeaders.length).fill('');
    for (var j = 0; j < trxHeaders.length; j++) {
      var key = trxHeaders[j];
      if (key === 'id' || key === 'trxid') newTrxRow[j] = trx.id;
      else if (key === 'tanggal' || key === 'tgl') newTrxRow[j] = trx.tgl || '';
      else if (key === 'customer' || key === 'pelanggan') newTrxRow[j] = trx.customer || '';
      else if (key === 'sales') newTrxRow[j] = trx.sales || '';
      else if (key === 'gross') newTrxRow[j] = Number(trx.gross) || 0;
      else if (key === 'diskon' || key === 'discount') newTrxRow[j] = Number(trx.diskon) || 0;
      else if (key === 'nett' || key === 'net') newTrxRow[j] = Number(trx.nett) || 0;
      else if (key === 'status' || key === 'pembayaran') newTrxRow[j] = trx.status || 'belumTransfer';
      else if (key === 'createdat') newTrxRow[j] = now;
      else if (key === 'totalmodal') newTrxRow[j] = totalModal;
      else if (key === 'totalprofit') newTrxRow[j] = totalProfit;
    }
    trxSheet.appendRow(newTrxRow);
    
    // Batch insert detail rows
    var detSheet = getSheet(SHEET.DETAIL);
    var detHeaders = getHeaders(SHEET.DETAIL);
    var rowsToAppend = [];
    for (var k = 0; k < processedItems.length; k++) {
      var it = processedItems[k];
      var newDetRow = new Array(detHeaders.length).fill('');
      for (var m = 0; m < detHeaders.length; m++) {
        var key = detHeaders[m];
        if (key === 'trxid') newDetRow[m] = trx.id;
        else if (key === 'tanggal' || key === 'tgl') newDetRow[m] = trx.tgl || '';
        else if (key === 'customer' || key === 'pelanggan') newDetRow[m] = trx.customer || '';
        else if (key === 'sales') newDetRow[m] = trx.sales || '';
        else if (key === 'status' || key === 'pembayaran') newDetRow[m] = trx.status || '';
        else if (key === 'sku') newDetRow[m] = it.sku;
        else if (key === 'barcode') newDetRow[m] = it.barcode;
        else if (key === 'nama') newDetRow[m] = it.nama;
        else if (key === 'qty') newDetRow[m] = it.qty;
        else if (key === 'harga') newDetRow[m] = it.harga;
        else if (key === 'modal') newDetRow[m] = it.modal;
        else if (key === 'discrpper' || key === 'disc') newDetRow[m] = it.disc;
        else if (key === 'nettper' || key === 'nett') newDetRow[m] = it.nettPer;
        else if (key === 'subtotal') newDetRow[m] = it.subtotal;
        else if (key === 'profit') newDetRow[m] = it.profit;
        else if (key === 'createdat') newDetRow[m] = now;
      }
      rowsToAppend.push(newDetRow);
    }
    if (rowsToAppend.length > 0) {
      detSheet.getRange(detSheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    }
    
    return { ok: true, id: trx.id, totalModal: totalModal, totalProfit: totalProfit };
  } catch(e) {
    throw new Error('saveTrx gagal: ' + e.message);
  }
}

function getTrxList() {
  try {
    ensureTrxHeaders();
    var data = getSheetData(SHEET.TRX); if (data.length <= 1) return [];
    var headers = data[0].map(function(c) { return String(c||'').trim().toLowerCase(); });
    var col = {
      id: headers.indexOf('id') >= 0 ? headers.indexOf('id') : 0,
      tgl: headers.indexOf('tanggal') >= 0 ? headers.indexOf('tanggal') : (headers.indexOf('tgl') >= 0 ? headers.indexOf('tgl') : 1),
      customer: headers.indexOf('customer') >= 0 ? headers.indexOf('customer') : (headers.indexOf('pelanggan') >= 0 ? headers.indexOf('pelanggan') : 2),
      sales: headers.indexOf('sales') >= 0 ? headers.indexOf('sales') : 3,
      gross: headers.indexOf('gross') >= 0 ? headers.indexOf('gross') : 4,
      diskon: headers.indexOf('diskon') >= 0 ? headers.indexOf('diskon') : (headers.indexOf('discount') >= 0 ? headers.indexOf('discount') : 5),
      nett: headers.indexOf('nett') >= 0 ? headers.indexOf('nett') : (headers.indexOf('net') >= 0 ? headers.indexOf('net') : 6),
      status: headers.indexOf('status') >= 0 ? headers.indexOf('status') : (headers.indexOf('pembayaran') >= 0 ? headers.indexOf('pembayaran') : 7)
    };
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i]; var id = String(r[col.id]||'').trim(); if (!id) continue;
      list.push({
        id: id, tgl: formatDateCell(r[col.tgl]), customer: String(r[col.customer]||''),
        sales: String(r[col.sales]||''), gross: Number(r[col.gross])||0, diskon: Number(r[col.diskon])||0,
        nett: Number(r[col.nett])||0, status: String(r[col.status]||'belumTransfer')
      });
    }
    return list;
  } catch(e) {
    throw new Error('getTrxList gagal: ' + e.message);
  }
}

function updateStatus(id, newStatus) {
  try {
    if (!id || !newStatus) throw new Error('ID dan status diperlukan');
    var trxSheet = getSheet(SHEET.TRX);
    var trxData = trxSheet.getDataRange().getValues();
    var trxHeaders = getHeaders(SHEET.TRX);
    var idCol = trxHeaders.indexOf('id'); if (idCol < 0) idCol = 0;
    var statusCol = trxHeaders.indexOf('status'); if (statusCol < 0) statusCol = trxHeaders.indexOf('pembayaran'); if (statusCol < 0) statusCol = 7;
    for (var i = trxData.length - 1; i >= 1; i--) {
      if (String(trxData[i][idCol]||'').trim() === id) { trxSheet.getRange(i+1, statusCol+1).setValue(newStatus); break; }
    }
    var detSheet = getSheet(SHEET.DETAIL);
    var detData = detSheet.getDataRange().getValues();
    var detHeaders = getHeaders(SHEET.DETAIL);
    var trxIdCol = detHeaders.indexOf('trxid'); if (trxIdCol < 0) trxIdCol = 0;
    var detStatusCol = detHeaders.indexOf('status'); if (detStatusCol < 0) detStatusCol = detHeaders.indexOf('pembayaran'); if (detStatusCol < 0) detStatusCol = 4;
    for (var j = detData.length - 1; j >= 1; j--) {
      if (String(detData[j][trxIdCol]||'').trim() === id) { detSheet.getRange(j+1, detStatusCol+1).setValue(newStatus); }
    }
    return { ok: true, id: id, newStatus: newStatus };
  } catch(e) {
    throw new Error('updateStatus gagal: ' + e.message);
  }
}

function deleteTrx(id) {
  try {
    if (!id) throw new Error('ID transaksi diperlukan');
    var trxSheet = getSheet(SHEET.TRX);
    var trxData = trxSheet.getDataRange().getValues();
    var trxHeaders = getHeaders(SHEET.TRX);
    var idCol = trxHeaders.indexOf('id'); if (idCol < 0) idCol = 0;
    var rowToDelete = -1;
    for (var i = trxData.length - 1; i >= 1; i--) {
      if (String(trxData[i][idCol]||'').trim() === id) { rowToDelete = i + 1; break; }
    }
    if (rowToDelete > 0) trxSheet.deleteRow(rowToDelete);
    var detSheet = getSheet(SHEET.DETAIL);
    var detData = detSheet.getDataRange().getValues();
    var detHeaders = getHeaders(SHEET.DETAIL);
    var detIdCol = detHeaders.indexOf('trxid'); if (detIdCol < 0) detIdCol = 0;
    var rowsToDelete = [];
    for (var j = detData.length - 1; j >= 1; j--) {
      if (String(detData[j][detIdCol]||'').trim() === id) rowsToDelete.push(j + 1);
    }
    rowsToDelete.sort(function(a,b) { return b - a; });
    for (var k = 0; k < rowsToDelete.length; k++) { detSheet.deleteRow(rowsToDelete[k]); }
    return { ok: true, deleted: id };
  } catch(e) {
    throw new Error('deleteTrx gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  INPUT BARANG
// ══════════════════════════════════════════════════════════════
function saveInputBarang(entry) {
  try {
    if (!entry || !entry.id) throw new Error('Data input barang tidak valid');
    var sheet = getSheet(SHEET.BARANG);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['ID','GroupID','SKU','Nama','Qty','HargaModal','Disc','NetModal','Driver','Rit','Status','Tanggal','CreatedAt']);
    }
    sheet.appendRow([entry.id, entry.groupId || '', entry.sku, entry.nama, entry.qty, entry.hargaModal, entry.disc, entry.netModal, entry.driver, entry.rit, entry.status, entry.date || '', entry.createdAt || '']);
    return 'ok';
  } catch(e) {
    throw new Error('saveInputBarang gagal: ' + e.message);
  }
}

function getInputBarangHistory() {
  try {
    var sheet = getSheet(SHEET.BARANG);
    if (sheet.getLastRow() <= 1) return [];
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(c) { return String(c||'').trim().toLowerCase(); });
    var col = {
      id: headers.indexOf('id'), groupId: headers.indexOf('groupid'), sku: headers.indexOf('sku'),
      nama: headers.indexOf('nama'), qty: headers.indexOf('qty'), hargaModal: headers.indexOf('hargamodal'),
      disc: headers.indexOf('disc'), netModal: headers.indexOf('netmodal'), driver: headers.indexOf('driver'),
      rit: headers.indexOf('rit'), status: headers.indexOf('status'),
      date: headers.indexOf('tanggal') >= 0 ? headers.indexOf('tanggal') : headers.indexOf('date'),
      createdAt: headers.indexOf('createdat')
    };
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      list.push({
        id: col.id >= 0 ? r[col.id] : '', groupId: col.groupId >= 0 ? r[col.groupId] : '',
        sku: col.sku >= 0 ? String(r[col.sku]) : '', nama: col.nama >= 0 ? String(r[col.nama]) : '',
        qty: col.qty >= 0 ? Number(r[col.qty])||0 : 0, hargaModal: col.hargaModal >= 0 ? Number(r[col.hargaModal])||0 : 0,
        disc: col.disc >= 0 ? Number(r[col.disc])||0 : 0, netModal: col.netModal >= 0 ? Number(r[col.netModal])||0 : 0,
        driver: col.driver >= 0 ? String(r[col.driver]) : '', rit: col.rit >= 0 ? String(r[col.rit]) : '',
        status: col.status >= 0 ? String(r[col.status]) : '',
        date: (col.date >= 0 && r[col.date]) ? formatDateCell(r[col.date]) : (col.createdAt >= 0 && r[col.createdAt] ? formatDateCell(r[col.createdAt]) : ''),
        createdAt: col.createdAt >= 0 ? formatDateCell(r[col.createdAt]) : ''
      });
    }
    return list;
  } catch(e) {
    throw new Error('getInputBarangHistory gagal: ' + e.message);
  }
}

function updateInputBarang(updated) {
  try {
    if (!updated || !updated.id) throw new Error('ID diperlukan');
    var sheet = getSheet(SHEET.BARANG);
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(c) { return String(c||'').toLowerCase(); });
    var idCol = headers.indexOf('id');
    var qtyCol = headers.indexOf('qty');
    var statusCol = headers.indexOf('status');
    var hargaModalCol = headers.indexOf('hargamodal');
    var discCol = headers.indexOf('disc');
    var netModalCol = headers.indexOf('netmodal');
    if (idCol < 0 || qtyCol < 0 || statusCol < 0) throw new Error('Kolom tidak lengkap');
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idCol] || '').trim() === updated.id) {
        var row = i + 1;
        if (updated.qty !== undefined) sheet.getRange(row, qtyCol + 1).setValue(updated.qty);
        if (updated.status !== undefined) sheet.getRange(row, statusCol + 1).setValue(updated.status);
        if (hargaModalCol >= 0 && discCol >= 0 && netModalCol >= 0 && updated.qty !== undefined) {
          var harga = Number(data[i][hargaModalCol]) || 0;
          var disc = Number(data[i][discCol]) || 0;
          var netModal = (harga - disc) * updated.qty;
          sheet.getRange(row, netModalCol + 1).setValue(netModal);
        }
        return { ok: true, id: updated.id };
      }
    }
    throw new Error('Data tidak ditemukan');
  } catch(e) {
    throw new Error('updateInputBarang gagal: ' + e.message);
  }
}

function deleteInputBarang(id) {
  try {
    if (!id) throw new Error('ID diperlukan');
    var sheet = getSheet(SHEET.BARANG);
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(c) { return String(c||'').toLowerCase(); });
    var idCol = headers.indexOf('id');
    if (idCol < 0) throw new Error('Kolom ID tidak ditemukan');
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idCol] || '').trim() === id) { sheet.deleteRow(i + 1); return { ok: true, deleted: id }; }
    }
    throw new Error('Data tidak ditemukan');
  } catch(e) {
    throw new Error('deleteInputBarang gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  SETORAN
// ══════════════════════════════════════════════════════════════
function saveSetoran(data) {
  try {
    if (!data || !data.tgl) throw new Error('Data setoran tidak valid');
    var sheet = getSheet(SHEET.SETORAN);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Tanggal','Sales','GrandTotal','Makan','Tips','Parkir','Bensin','Flazz','Transfer','Cicilan','Tagihan','Ket1','Jml1','Ket2','Jml2','Total','Setor','Selisih','CreatedAt','FotoUrl']);
    }
    // Cek apakah kolom FotoUrl sudah ada
    var headers = sheet.getLastRow() > 0 ? sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(c){return String(c||'').toLowerCase();}) : [];
    if (headers.indexOf('fotourl') === -1 && sheet.getLastRow() > 0) {
      sheet.getRange(1, sheet.getLastColumn()+1).setValue('FotoUrl');
    }
    sheet.appendRow([data.tgl, data.sales, data.grandTotal, data.makan, data.tips, data.parkir, data.bensin, data.flazz, data.transfer, data.cicilan, data.tagihan, data.ket1, data.jml1, data.ket2, data.jml2, data.total, data.setor, data.selisih, Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss'), data.fotoUrl||'']);
    return 'ok';
  } catch(e) {
    throw new Error('saveSetoran gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  GET SETORAN HISTORY — Ambil semua riwayat setoran dari sheet
// ══════════════════════════════════════════════════════════════
function getSetoranHistory() {
  try {
    var sheet = getSheet(SHEET.SETORAN);
    if (sheet.getLastRow() === 0) return [];
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        var h = String(headers[j] || '').trim();
        var hl = h.toLowerCase();
        row[hl] = data[i][j];
        // Simpan juga dengan nama asli header untuk fallback
        row[h] = data[i][j];
      }
      result.push({
        tgl: row['tanggal'] || row['Tanggal'] || '',
        sales: row['sales'] || row['Sales'] || '',
        grandTotal: Number(row['grandtotal'] || row['grandTotal'] || row['grand total'] || row['GrandTot'] || 0),
        makan: Number(row['makan'] || row['uangmak'] || row['UangMak'] || row['UangMakan'] || 0),
        tips: Number(row['tips'] || row['Tips'] || 0),
        parkir: Number(row['parkir'] || row['Parkir'] || 0),
        bensin: Number(row['bensin'] || row['Bensin'] || 0),
        flazz: Number(row['flazz'] || row['Flazz'] || 0),
        transfer: Number(row['transfer'] || row['Transfer'] || 0),
        cicilan: Number(row['cicilan'] || row['cicilanhu'] || row['CicilanHu'] || row['Cicilan'] || row['cicilan hu'] || 0),
        tagihan: Number(row['tagihan'] || row['Tagihan'] || 0),
        ket1: row['ket1'] || row['ekstra1li'] || row['Ekstra1Li'] || row['ekstra1ket'] || row['Ekstra1Ket'] || row['ket 1'] || row['keterangan1'] || '',
        jml1: Number(row['jml1'] || row['ekstra1'] || row['Ekstra1'] || row['jumlah1'] || row['jml 1'] || 0),
        ket2: row['ket2'] || row['ekstra2li'] || row['Ekstra2Li'] || row['ekstra2ket'] || row['Ekstra2Ket'] || row['ket 2'] || row['keterangan2'] || '',
        jml2: Number(row['jml2'] || row['ekstra2'] || row['Ekstra2'] || row['jumlah2'] || row['jml 2'] || 0),
        total: Number(row['total'] || row['Total'] || 0),
        setor: Number(row['setor'] || row['Setor'] || 0),
        selisih: Number(row['selisih'] || row['Selisih'] || 0),
        createdAt: row['createdat'] || row['CreatedAt'] || '',
        fotoUrl: row['fotourl'] || row['FotoUrl'] || ''
      });
    }
    return result;
  } catch(e) {
    throw new Error('getSetoranHistory gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  FOTO BUKTI — Upload ke Google Drive (BARU)
//  Menerima: { base64, mimeType, nama, refId, jenis }
//  refId  = ID transaksi atau tanggal setoran
//  jenis  = 'transfer' | 'setoran' | 'pengiriman'
// ══════════════════════════════════════════════════════════════
function uploadFoto(payload) {
  try {
    if (!payload || !payload.base64) throw new Error('Data foto tidak valid');
    // Buat/ambil folder Tirta Kencana di Drive
    var folderName = 'Tirta Kencana - Bukti';
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    // Sub-folder per jenis
    var subName = payload.jenis || 'lainnya';
    var subFolders = folder.getFoldersByName(subName);
    var subFolder = subFolders.hasNext() ? subFolders.next() : folder.createFolder(subName);
    // Decode base64 dan simpan file
    var mimeType = payload.mimeType || 'image/jpeg';
    var ext = mimeType.indexOf('png') >= 0 ? '.png' : mimeType.indexOf('pdf') >= 0 ? '.pdf' : '.jpg';
    var now = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd-HHmmss');
    var fileName = (payload.nama || payload.refId || 'foto') + '-' + now + ext;
    var decoded = Utilities.newBlob(Utilities.base64Decode(payload.base64), mimeType, fileName);
    var file = subFolder.createFile(decoded);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileId = file.getId();
    var viewUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
    var thumbUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w400';
    // Simpan referensi ke sheet FotoBukti
    var fotoSheet = getSheet('FotoBukti');
    if (fotoSheet.getLastRow() === 0) {
      fotoSheet.appendRow(['Waktu','Jenis','RefID','NamaFile','ViewUrl','ThumbUrl','Uploader']);
    }
    fotoSheet.appendRow([now, subName, payload.refId||'', fileName, viewUrl, thumbUrl, payload.uploader||'']);
    return { ok: true, fileId: fileId, viewUrl: viewUrl, thumbUrl: thumbUrl, fileName: fileName };
  } catch(e) {
    throw new Error('uploadFoto gagal: ' + e.message);
  }
}

function getFotoList(refId) {
  try {
    var sheet = SS.getSheetByName('FotoBukti');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(c){ return String(c||'').toLowerCase(); });
    var col = {
      waktu: headers.indexOf('waktu'), jenis: headers.indexOf('jenis'),
      refId: headers.indexOf('refid'), nama: headers.indexOf('namafile'),
      viewUrl: headers.indexOf('viewurl'), thumbUrl: headers.indexOf('thumburl'),
      uploader: headers.indexOf('uploader')
    };
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (!refId || String(r[col.refId]||'').trim() === refId) {
        result.push({
          waktu: String(r[col.waktu]||''), jenis: String(r[col.jenis]||''),
          refId: String(r[col.refId]||''), nama: String(r[col.nama]||''),
          viewUrl: String(r[col.viewUrl]||''), thumbUrl: String(r[col.thumbUrl]||''),
          uploader: String(r[col.uploader]||'')
        });
      }
    }
    return result;
  } catch(e) {
    throw new Error('getFotoList gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  MASTER STOCK
// ══════════════════════════════════════════════════════════════
function saveMasterStock(payload) {
  try {
    if (!payload || !payload.date) return 'ok';
    var sheet = getSheet(SHEET.MASTERSTOCK);
    if (sheet.getLastRow() === 0) sheet.appendRow(['Tanggal', 'SKU', 'Nama', 'StokAwal', 'Masuk', 'Keluar', 'StokAkhir']);
    var stock = payload.stock || []; if (!stock.length) return 'ok';
    var data = sheet.getDataRange().getValues();
    var rowsToDelete = [];
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0] || '').trim() === payload.date) rowsToDelete.push(i + 1);
    }
    rowsToDelete.sort(function(a,b) { return b - a; });
    rowsToDelete.forEach(function(r) { sheet.deleteRow(r); });
    stock.forEach(function(s) {
      sheet.appendRow([payload.date, s.sku, s.nama, s.stokAwal || 0, s.masuk || 0, s.keluar || 0, s.stokAkhir || 0]);
    });
    return 'ok';
  } catch(e) {
    throw new Error('saveMasterStock gagal: ' + e.message);
  }
}

function getMasterStockByDate(date) {
  try {
    var sheet = SS.getSheetByName(SHEET.MASTERSTOCK); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); if (data.length <= 1) return [];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (String(r[0] || '').trim() === date) {
        result.push({ sku: String(r[1] || ''), nama: String(r[2] || ''), stokAwal: Number(r[3]) || 0, masuk: Number(r[4]) || 0, keluar: Number(r[5]) || 0, stokAkhir: Number(r[6]) || 0 });
      }
    }
    return result;
  } catch(e) {
    throw new Error('getMasterStockByDate gagal: ' + e.message);
  }
}

function getStockBarang() {
  try {
    var sheet = SS.getSheetByName(SHEET.MASTERSTOCK); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); if (data.length <= 1) return [];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var tgl = '';
      try { tgl = formatDateCell(r[0]); } catch(e) { tgl = String(r[0]||'').trim(); }
      if (!tgl) continue;
      result.push({ date: tgl, sku: String(r[1]||''), nama: String(r[2]||''), stokAwal: Number(r[3])||0, masuk: Number(r[4])||0, keluar: Number(r[5])||0, stokAkhir: Number(r[6])||0 });
    }
    return result;
  } catch(e) {
    throw new Error('getStockBarang gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  UPDATE STOK AWAL — untuk fitur Rollover Stok (BARU v8.2)
//  Menerima: { date, sku, nama, stokAwal, masuk, keluar }
//  Logika: jika baris date+sku sudah ada → update StokAwal saja
//          jika belum ada → append baris baru dengan masuk=0, keluar=0
// ══════════════════════════════════════════════════════════════
function updateStockAwal(payload) {
  try {
    if (!payload || !payload.date || !payload.sku) throw new Error('Data tidak valid: date dan sku diperlukan');
    var sheet = getSheet(SHEET.MASTERSTOCK);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Tanggal', 'SKU', 'Nama', 'StokAwal', 'Masuk', 'Keluar', 'StokAkhir']);
    }
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(c) { return String(c||'').trim().toLowerCase(); });
    var colTgl    = headers.indexOf('tanggal');  if (colTgl < 0)    colTgl = 0;
    var colSku    = headers.indexOf('sku');       if (colSku < 0)    colSku = 1;
    var colNama   = headers.indexOf('nama');      if (colNama < 0)   colNama = 2;
    var colAwal   = headers.indexOf('stokawal');  if (colAwal < 0)   colAwal = 3;
    var colMasuk  = headers.indexOf('masuk');     if (colMasuk < 0)  colMasuk = 4;
    var colKeluar = headers.indexOf('keluar');    if (colKeluar < 0) colKeluar = 5;
    var colAkhir  = headers.indexOf('stokakhir'); if (colAkhir < 0)  colAkhir = 6;

    var stokAwal   = Number(payload.stokAwal)  || 0;
    var masuk      = Number(payload.masuk)     || 0;
    var keluar     = Number(payload.keluar)    || 0;
    var stokAkhir  = stokAwal + masuk - keluar;

    // Cari baris yang sudah ada untuk date + sku yang sama
    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      var rowTgl = String(data[i][colTgl]||'').trim();
      var rowSku = String(data[i][colSku]||'').trim().toLowerCase();
      if (rowTgl === payload.date && rowSku === payload.sku.toLowerCase()) {
        foundRow = i + 1; // nomor baris di sheet (1-based)
        break;
      }
    }

    if (foundRow > 0) {
      // Update baris yang sudah ada — hanya update StokAwal, biarkan Masuk & Keluar tetap
      var existingMasuk  = Number(data[foundRow-1][colMasuk])  || 0;
      var existingKeluar = Number(data[foundRow-1][colKeluar]) || 0;
      var newAkhir = stokAwal + existingMasuk - existingKeluar;
      sheet.getRange(foundRow, colAwal  + 1).setValue(stokAwal);
      sheet.getRange(foundRow, colAkhir + 1).setValue(newAkhir);
    } else {
      // Append baris baru untuk hari ini
      var newRow = new Array(Math.max(colAkhir + 1, 7)).fill('');
      newRow[colTgl]    = payload.date;
      newRow[colSku]    = payload.sku;
      newRow[colNama]   = payload.nama || '';
      newRow[colAwal]   = stokAwal;
      newRow[colMasuk]  = masuk;
      newRow[colKeluar] = keluar;
      newRow[colAkhir]  = stokAkhir;
      sheet.appendRow(newRow);
    }

    return { ok: true, sku: payload.sku, date: payload.date, stokAwal: stokAwal };
  } catch(e) {
    throw new Error('updateStockAwal gagal: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════
//  TRX DETAIL
// ══════════════════════════════════════════════════════════════
function getTrxDetail(id) {
  try {
    var sheet = getSheet(SHEET.DETAIL);
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { id: id, items: [] };
    var headers = data[0].map(function(c) { return String(c||'').trim().toLowerCase(); });
    var col = {
      trxid: headers.indexOf('trxid'),
      sku: headers.indexOf('sku'),
      nama: headers.indexOf('nama'),
      qty: headers.indexOf('qty'),
      harga: headers.indexOf('harga'),
      modal: headers.indexOf('modal'),
      disc: headers.indexOf('discrpper'),
      nett: headers.indexOf('nettper'),
      subtotal: headers.indexOf('subtotal'),
      profit: headers.indexOf('profit')
    };
    var items = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (String(r[col.trxid]||'').trim() === id) {
        items.push({
          sku: col.sku>=0 ? String(r[col.sku]) : '',
          nama: col.nama>=0 ? String(r[col.nama]) : '',
          qty: col.qty>=0 ? Number(r[col.qty]) : 0,
          harga: col.harga>=0 ? Number(r[col.harga]) : 0,
          modal: col.modal>=0 ? Number(r[col.modal]) : 0,
          discRpPer: col.disc>=0 ? Number(r[col.disc]) : 0,
          nettPer: col.nett>=0 ? Number(r[col.nett]) : 0,
          subtotal: col.subtotal>=0 ? Number(r[col.subtotal]) : 0,
          profit: col.profit>=0 ? Number(r[col.profit]) : 0
        });
      }
    }
    return { id: id, items: items };
  } catch(e) {
    throw new Error('getTrxDetail gagal: ' + e.message);
  }
}
