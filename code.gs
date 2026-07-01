/**
 * TIRTA KENCANA POS - Google Apps Script
 * Backend untuk integrasi dengan Google Sheets
 * 
 * CARA PENGGUNAAN:
 * 1. Buka Google Sheets baru atau yang sudah ada
 * 2. Extensions > Apps Script
 * 3. Hapus semua kode yang ada, paste kode ini
 * 4. Save project (beri nama "Tirta Kencana POS")
 * 5. Deploy > New Deployment > Select type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy URL yang dihasilkan dan paste di pengaturan aplikasi
 */

// ════════════════════════════════════════════════════════
// KONFIGURASI
// ════════════════════════════════════════════════════════

const SHEET_NAMES = {
  SETTINGS: 'Settings',
  PRODUCTS: 'Products',
  CUSTOMERS: 'Customers',
  SALES_TRX: 'SalesTrx',
  PURCHASES: 'Purchases',
  STOCK: 'Stock',
  DEBT: 'Debt',
  SETORAN: 'Setoran',
  USERS: 'Users',
  BACKUP: 'Backup'
};

// ════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ════════════════════════════════════════════════════════

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function ensureHeaders(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
}

function getDataAsJSON(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0].map(h => String(h).toLowerCase().replace(/\s+/g, '_'));
  const rows = data.slice(1);
  
  return rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function timestamp() {
  return new Date().toISOString();
}

// ════════════════════════════════════════════════════════
// DO GET - HANDLER UTAMA
// ════════════════════════════════════════════════════════

function doGet(e) {
  const action = e.parameter.action;
  
  try {
    switch(action) {
      case 'getAllData':
        return handleGetAllData();
      case 'getSettings':
        return handleGetSettings();
      case 'getProducts':
        return handleGetProducts();
      case 'getCustomers':
        return handleGetCustomers();
      case 'getTrxList':
        return handleGetTrxList();
      case 'getPurchaseList':
        return handleGetPurchaseList();
      case 'getStockData':
        return handleGetStockData();
      case 'getDebtData':
        return handleGetDebtData();
      case 'getSetoranData':
        return handleGetSetoranData();
      case 'getUsers':
        return handleGetUsers();
      case 'getDetailByTrx':
        return handleGetDetailByTrx(e.parameter.id);
      default:
        return jsonResponse({ status: 'error', message: 'Action tidak dikenali' });
    }
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.toString() });
  }
}

// ════════════════════════════════════════════════════════
// DO POST - HANDLER UTAMA
// ════════════════════════════════════════════════════════

function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    switch(action) {
      case 'saveSale':
        return handleSaveSale(data);
      case 'savePurchase':
        return handleSavePurchase(data);
      case 'saveProduct':
        return handleSaveProduct(data);
      case 'saveProducts':
        return handleSaveProducts(data);
      case 'saveCustomer':
        return handleSaveCustomer(data);
      case 'addCustomer':
        return handleAddCustomer(data);
      case 'saveCustomers':
        return handleSaveCustomers(data);
      case 'saveSettings':
        return handleSaveSettings(data);
      case 'saveSetoran':
        return handleSaveSetoran(data);
      case 'saveDebt':
        return handleSaveDebt(data);
      case 'updateStock':
        return handleUpdateStock(data);
      case 'rolloverStock':
        return handleRolloverStock(data);
      case 'backupData':
        return handleBackupData(data);
      case 'restoreData':
        return handleRestoreData(data);
      case 'deleteMonth':
        return handleDeleteMonth(data);
      default:
        return jsonResponse({ status: 'error', message: 'Action POST tidak dikenali' });
    }
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.toString() });
  }
}

// ════════════════════════════════════════════════════════
// GET HANDLERS
// ════════════════════════════════════════════════════════

function handleGetAllData() {
  const data = {
    settings: getSettingsData(),
    products: getDataAsJSON(SHEET_NAMES.PRODUCTS),
    customers: getDataAsJSON(SHEET_NAMES.CUSTOMERS),
    sales: getDataAsJSON(SHEET_NAMES.SALES_TRX),
    purchases: getDataAsJSON(SHEET_NAMES.PURCHASES),
    stock: getDataAsJSON(SHEET_NAMES.STOCK),
    debt: getDataAsJSON(SHEET_NAMES.DEBT),
    setoran: getDataAsJSON(SHEET_NAMES.SETORAN),
    users: getDataAsJSON(SHEET_NAMES.USERS)
  };
  
  return jsonResponse({ status: 'success', data: data });
}

function getSettingsData() {
  const sheet = getSheet(SHEET_NAMES.SETTINGS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {};
  
  // Row 2 adalah data settings (row 1 adalah header)
  const headers = data[0];
  const values = data[1];
  
  const settings = {};
  headers.forEach((h, i) => {
    settings[String(h).toLowerCase()] = values[i];
  });
  
  // Parse salesList jika ada
  if (settings.sales_list) {
    try {
      settings.salesList = JSON.parse(settings.sales_list);
    } catch(e) {
      settings.salesList = [];
    }
  } else {
    settings.salesList = [];
  }
  
  return settings;
}

function handleGetSettings() {
  return jsonResponse(getSettingsData());
}

function handleGetProducts() {
  return jsonResponse(getDataAsJSON(SHEET_NAMES.PRODUCTS));
}

function handleGetCustomers() {
  const data = getDataAsJSON(SHEET_NAMES.CUSTOMERS);
  // Return hanya nama customer sebagai array simple
  return jsonResponse(data.map(c => c.nama || c.customer || ''));
}

function handleGetTrxList() {
  const data = getDataAsJSON(SHEET_NAMES.SALES_TRX);
  return jsonResponse(data);
}

function handleGetPurchaseList() {
  const data = getDataAsJSON(SHEET_NAMES.PURCHASES);
  return jsonResponse(data);
}

function handleGetStockData() {
  const data = getDataAsJSON(SHEET_NAMES.STOCK);
  return jsonResponse(data);
}

function handleGetDebtData() {
  const data = getDataAsJSON(SHEET_NAMES.DEBT);
  return jsonResponse(data);
}

function handleGetSetoranData() {
  const data = getDataAsJSON(SHEET_NAMES.SETORAN);
  return jsonResponse(data);
}

function handleGetUsers() {
  const data = getDataAsJSON(SHEET_NAMES.USERS);
  return jsonResponse(data);
}

function handleGetDetailByTrx(id) {
  if (!id) return jsonResponse([]);
  
  const data = getDataAsJSON(SHEET_NAMES.SALES_TRX);
  const trx = data.find(t => String(t.id) === String(id));
  
  if (!trx || !trx.items) return jsonResponse([]);
  
  try {
    const items = JSON.parse(trx.items);
    return jsonResponse(items);
  } catch(e) {
    return jsonResponse([]);
  }
}

// ════════════════════════════════════════════════════════
// POST HANDLERS
// ════════════════════════════════════════════════════════

function handleSaveSale(data) {
  const sheet = getSheet(SHEET_NAMES.SALES_TRX);
  const headers = ['id', 'tgl', 'customer', 'sales', 'status', 'items', 'gross', 'disc', 'biaya_tambahan', 'nett', 'foto_url', 'created_at', 'updated_at'];
  ensureHeaders(sheet, headers);
  
  const newRow = [
    data.id || Utilities.getUuid(),
    data.tgl || new Date().toISOString().slice(0, 10),
    data.customer || '-',
    data.sales || '-',
    data.status || 'cod',
    JSON.stringify(data.items || []),
    data.gross || 0,
    data.disc || 0,
    data.biaya_tambahan || 0,
    data.nett || 0,
    data.foto_url || '',
    timestamp(),
    timestamp()
  ];
  
  sheet.appendRow(newRow);
  
  // Update stock
  updateStockFromSale(data.items);
  
  return jsonResponse({ status: 'success', message: 'Penjualan berhasil disimpan' });
}

function handleSavePurchase(data) {
  const sheet = getSheet(SHEET_NAMES.PURCHASES);
  const headers = ['id', 'tgl', 'driver', 'rit', 'status', 'items', 'total_qty', 'gross_modal', 'total_disc', 'total_modal_bersih', 'foto_url', 'created_at', 'updated_at'];
  ensureHeaders(sheet, headers);
  
  const newRow = [
    data.id || Utilities.getUuid(),
    data.tgl || new Date().toISOString().slice(0, 10),
    data.driver || '-',
    data.rit || 1,
    data.status || 'kiriman',
    JSON.stringify(data.items || []),
    data.total_qty || 0,
    data.gross_modal || 0,
    data.total_disc || 0,
    data.total_modal_bersih || 0,
    data.foto_url || '',
    timestamp(),
    timestamp()
  ];
  
  sheet.appendRow(newRow);
  
  // Update stock
  updateStockFromPurchase(data.items);
  
  return jsonResponse({ status: 'success', message: 'Pembelian berhasil disimpan' });
}

function handleSaveProduct(data) {
  const sheet = getSheet(SHEET_NAMES.PRODUCTS);
  const headers = ['sku', 'barcode', 'nama', 'harga', 'satuan', 'has_barcode', 'updated_at'];
  ensureHeaders(sheet, headers);
  
  const sku = data.sku;
  const allData = sheet.getDataRange().getValues();
  let rowIndex = -1;
  
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][0]) === String(sku)) {
      rowIndex = i + 1;
      break;
    }
  }
  
  const rowData = [
    sku,
    data.barcode || '',
    data.nama || '',
    data.harga || 0,
    data.satuan || 'Pcs',
    data.hasBarcode ? 'TRUE' : 'FALSE',
    timestamp()
  ];
  
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  
  return jsonResponse({ status: 'success', message: 'Produk berhasil disimpan' });
}

function handleSaveProducts(data) {
  const products = data.products || data;
  if (!Array.isArray(products)) {
    return jsonResponse({ status: 'error', message: 'Data harus berupa array' });
  }
  
  const sheet = getSheet(SHEET_NAMES.PRODUCTS);
  const headers = ['sku', 'barcode', 'nama', 'harga', 'satuan', 'has_barcode', 'updated_at'];
  
  // Clear existing data (keep header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  
  ensureHeaders(sheet, headers);
  
  // Add all products
  products.forEach(p => {
    sheet.appendRow([
      p.sku || '',
      p.barcode || '',
      p.nama || '',
      p.harga || 0,
      p.satuan || 'Pcs',
      p.hasBarcode ? 'TRUE' : 'FALSE',
      timestamp()
    ]);
  });
  
  return jsonResponse({ status: 'success', message: `Berhasil menyimpan ${products.length} produk` });
}

function handleAddCustomer(data) {
  const name = data.name || data.nama;
  if (!name) {
    return jsonResponse({ status: 'error', message: 'Nama customer wajib diisi' });
  }
  
  const sheet = getSheet(SHEET_NAMES.CUSTOMERS);
  const headers = ['nama', 'alamat', 'telepon', 'created_at'];
  ensureHeaders(sheet, headers);
  
  sheet.appendRow([name, '', '', timestamp()]);
  
  return jsonResponse({ status: 'success', message: 'Customer berhasil ditambahkan' });
}

function handleSaveCustomer(data) {
  return handleAddCustomer(data);
}

function handleSaveCustomers(data) {
  const customers = data.customers || data;
  if (!Array.isArray(customers)) {
    return jsonResponse({ status: 'error', message: 'Data harus berupa array' });
  }
  
  const sheet = getSheet(SHEET_NAMES.CUSTOMERS);
  const headers = ['nama', 'alamat', 'telepon', 'created_at'];
  
  // Clear existing data (keep header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  
  ensureHeaders(sheet, headers);
  
  customers.forEach(c => {
    sheet.appendRow([
      typeof c === 'string' ? c : (c.nama || c.customer || ''),
      typeof c === 'object' ? (c.alamat || '') : '',
      typeof c === 'object' ? (c.telepon || '') : '',
      timestamp()
    ]);
  });
  
  return jsonResponse({ status: 'success', message: `Berhasil menyimpan ${customers.length} customer` });
}

function handleSaveSettings(data) {
  const sheet = getSheet(SHEET_NAMES.SETTINGS);
  const headers = [
    'store_name', 'tagline', 'address', 'phone', 'footer',
    'logo_url', 'qris_url',
    'bank1_nama', 'bank1_norek', 'bank1_penerima',
    'bank2_nama', 'bank2_norek', 'bank2_penerima',
    'sales_list', 'gas_url', 'sheets_id',
    'updated_at'
  ];
  
  ensureHeaders(sheet, headers);
  
  // Check if settings already exist
  const lastRow = sheet.getLastRow();
  
  const rowData = [
    data.storeName || 'TIRTA KENCANA',
    data.tagline || 'DO GOOD AND GOD WILL COME TO YOU',
    data.address || 'WTC MANGGA DUA LANTAIUG BLOK A/53',
    data.phone || '088211058000',
    data.footer || 'TERIMA KASIH ATAS PEMBELIAN ANDA',
    data.logo_url || '',
    data.qris_url || '',
    (data.bank1 && data.bank1.nama) || 'BCA',
    (data.bank1 && data.bank1.norek) || '6930099099',
    (data.bank1 && data.bank1.penerima) || 'HENDRI',
    (data.bank2 && data.bank2.nama) || 'bluBCA',
    (data.bank2 && data.bank2.norek) || '002283588888',
    (data.bank2 && data.bank2.penerima) || 'HENDRI',
    JSON.stringify(data.salesList || []),
    data.gas_url || '',
    data.sheets_id || '',
    timestamp()
  ];
  
  if (lastRow >= 2) {
    // Update existing
    sheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);
  } else {
    // Insert new
    sheet.appendRow(rowData);
  }
  
  return jsonResponse({ status: 'success', message: 'Pengaturan berhasil disimpan' });
}

function handleSaveSetoran(data) {
  const sheet = getSheet(SHEET_NAMES.SETORAN);
  const headers = [
    'id', 'tgl', 'sales', 'grand_total',
    'makan', 'tips', 'parkir', 'bensin', 'flazz', 'transfer', 'cicilan', 'ekstra1', 'ekstra2',
    'tagihan', 'total', 'setor', 'selisih',
    'created_at', 'updated_at'
  ];
  ensureHeaders(sheet, headers);
  
  const newRow = [
    data.id || Utilities.getUuid(),
    data.tgl || new Date().toISOString().slice(0, 10),
    data.sales || '-',
    data.grand_total || 0,
    data.makan || 0,
    data.tips || 0,
    data.parkir || 0,
    data.bensin || 0,
    data.flazz || 0,
    data.transfer || 0,
    data.cicilan || 0,
    data.ekstra1 || 0,
    data.ekstra2 || 0,
    data.tagihan || 0,
    data.total || 0,
    data.setor || 0,
    data.selisih || 0,
    timestamp(),
    timestamp()
  ];
  
  sheet.appendRow(newRow);
  
  return jsonResponse({ status: 'success', message: 'Setoran berhasil disimpan' });
}

function handleSaveDebt(data) {
  const sheet = getSheet(SHEET_NAMES.DEBT);
  const headers = ['id', 'tgl', 'customer', 'sales', 'nett', 'status', 'paid_date', 'payment_method', 'created_at', 'updated_at'];
  ensureHeaders(sheet, headers);
  
  const newRow = [
    data.id || Utilities.getUuid(),
    data.tgl || new Date().toISOString().slice(0, 10),
    data.customer || '-',
    data.sales || '-',
    data.nett || 0,
    data.status || 'belum_lunas',
    data.paid_date || '',
    data.payment_method || '',
    timestamp(),
    timestamp()
  ];
  
  sheet.appendRow(newRow);
  
  return jsonResponse({ status: 'success', message: 'Data hutang berhasil disimpan' });
}

function updateStockFromSale(items) {
  if (!items || !Array.isArray(items)) return;
  
  const stockSheet = getSheet(SHEET_NAMES.STOCK);
  const headers = ['tanggal', 'sku', 'nama', 'stock_awal', 'masuk', 'keluar', 'stock_akhir', 'keterangan', 'created_at'];
  ensureHeaders(stockSheet, headers);
  
  const today = new Date().toISOString().slice(0, 10);
  
  items.forEach(item => {
    const qty = Number(item.qty) || 0;
    if (qty <= 0) return;
    
    // Get current stock
    const currentStock = getCurrentStock(item.sku);
    
    stockSheet.appendRow([
      today,
      item.sku || '-',
      item.nama || '-',
      currentStock,
      0,
      qty,
      currentStock - qty,
      'Penjualan',
      timestamp()
    ]);
  });
}

function updateStockFromPurchase(items) {
  if (!items || !Array.isArray(items)) return;
  
  const stockSheet = getSheet(SHEET_NAMES.STOCK);
  const headers = ['tanggal', 'sku', 'nama', 'stock_awal', 'masuk', 'keluar', 'stock_akhir', 'keterangan', 'created_at'];
  ensureHeaders(stockSheet, headers);
  
  const today = new Date().toISOString().slice(0, 10);
  
  items.forEach(item => {
    const qty = Number(item.qty) || 0;
    if (qty <= 0) return;
    
    // Get current stock
    const currentStock = getCurrentStock(item.sku);
    
    stockSheet.appendRow([
      today,
      item.sku || '-',
      item.nama || '-',
      currentStock,
      qty,
      0,
      currentStock + qty,
      'Pembelian',
      timestamp()
    ]);
  });
}

function getCurrentStock(sku) {
  if (!sku) return 0;
  
  const stockData = getDataAsJSON(SHEET_NAMES.STOCK);
  if (stockData.length === 0) return 0;
  
  // Filter by SKU and get the latest entry
  const skuEntries = stockData.filter(s => String(s.sku) === String(sku));
  if (skuEntries.length === 0) return 0;
  
  // Sort by date descending and get the latest
  skuEntries.sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0));
  
  return Number(skuEntries[0].stock_akhir) || 0;
}

function handleUpdateStock(data) {
  const stockSheet = getSheet(SHEET_NAMES.STOCK);
  const headers = ['tanggal', 'sku', 'nama', 'stock_awal', 'masuk', 'keluar', 'stock_akhir', 'keterangan', 'created_at'];
  ensureHeaders(stockSheet, headers);
  
  const today = data.tanggal || new Date().toISOString().slice(0, 10);
  const currentStock = getCurrentStock(data.sku);
  const adjustment = Number(data.masuk) - Number(data.keluar);
  
  stockSheet.appendRow([
    today,
    data.sku || '-',
    data.nama || '-',
    currentStock,
    data.masuk || 0,
    data.keluar || 0,
    currentStock + adjustment,
    data.keterangan || 'Manual adjustment',
    timestamp()
  ]);
  
  return jsonResponse({ status: 'success', message: 'Stock berhasil diupdate' });
}

function handleRolloverStock(data) {
  const stockSheet = getSheet(SHEET_NAMES.STOCK);
  const headers = ['tanggal', 'sku', 'nama', 'stock_awal', 'masuk', 'keluar', 'stock_akhir', 'keterangan', 'created_at'];
  ensureHeaders(stockSheet, headers);
  
  const today = new Date().toISOString().slice(0, 10);
  const yesterday = new Date(Date.now() - 86400000).toISOString().slice(0, 10);
  
  // Get all unique SKUs from previous day
  const allStockData = getDataAsJSON(SHEET_NAMES.STOCK);
  const yesterdayData = allStockData.filter(s => String(s.tanggal) === yesterday);
  
  // Group by SKU and get final stock
  const skuMap = {};
  yesterdayData.forEach(s => {
    const sku = s.sku || '-';
    if (!skuMap[sku]) {
      skuMap[sku] = {
        sku: sku,
        nama: s.nama || '-',
        stock_akhir: Number(s.stock_akhir) || 0
      };
    }
  });
  
  // Create rollover entries
  Object.values(skuMap).forEach(item => {
    stockSheet.appendRow([
      today,
      item.sku,
      item.nama,
      item.stock_akhir, // Yesterday's final becomes today's initial
      0,
      0,
      item.stock_akhir,
      'Rollover otomatis',
      timestamp()
    ]);
  });
  
  return jsonResponse({ 
    status: 'success', 
    message: `Rollover berhasil untuk ${Object.keys(skuMap).length} SKU` 
  });
}

function handleBackupData(data) {
  const backupSheet = getSheet(SHEET_NAMES.BACKUP);
  const headers = ['backup_date', 'data_json', 'created_at'];
  ensureHeaders(backupSheet, headers);
  
  const allData = {
    settings: getSettingsData(),
    products: getDataAsJSON(SHEET_NAMES.PRODUCTS),
    customers: getDataAsJSON(SHEET_NAMES.CUSTOMERS),
    sales: getDataAsJSON(SHEET_NAMES.SALES_TRX),
    purchases: getDataAsJSON(SHEET_NAMES.PURCHASES),
    stock: getDataAsJSON(SHEET_NAMES.STOCK),
    debt: getDataAsJSON(SHEET_NAMES.DEBT),
    setoran: getDataAsJSON(SHEET_NAMES.SETORAN)
  };
  
  backupSheet.appendRow([
    data.backup_date || new Date().toISOString().slice(0, 10),
    JSON.stringify(allData),
    timestamp()
  ]);
  
  return jsonResponse({ status: 'success', message: 'Backup berhasil dibuat' });
}

function handleRestoreData(data) {
  const backupData = JSON.parse(data.data_json);
  
  if (backupData.products) handleSaveProducts(backupData.products);
  if (backupData.customers) handleSaveCustomers(backupData.customers);
  if (backupData.settings) handleSaveSettings(backupData.settings);
  
  // Note: Sales, purchases, stock data should be restored carefully
  // This is a simplified version
  
  return jsonResponse({ status: 'success', message: 'Restore berhasil' });
}

function handleDeleteMonth(data) {
  const month = data.month; // Format: YYYY-MM
  if (!month) {
    return jsonResponse({ status: 'error', message: 'Bulan harus ditentukan' });
  }
  
  const sheetsToClear = [
    SHEET_NAMES.SALES_TRX,
    SHEET_NAMES.PURCHASES,
    SHEET_NAMES.STOCK,
    SHEET_NAMES.DEBT,
    SHEET_NAMES.SETORAN
  ];
  
  let deletedCount = 0;
  
  sheetsToClear.forEach(sheetName => {
    const sheet = getSheet(sheetName);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    // Find rows to keep (not matching the month)
    const rowsToKeep = rows.filter(row => {
      const dateCol = row[1]; // Assuming column B is date
      if (!dateCol) return true;
      const dateStr = String(dateCol).slice(0, 7); // YYYY-MM
      return dateStr !== month;
    });
    
    // Clear sheet and rewrite
    if (rowsToKeep.length < rows.length) {
      deletedCount += rows.length - rowsToKeep.length;
      sheet.clear();
      sheet.appendRow(headers);
      rowsToKeep.forEach(row => sheet.appendRow(row));
    }
  });
  
  return jsonResponse({ 
    status: 'success', 
    message: `Berhasil menghapus ${deletedCount} data dari bulan ${month}` 
  });
}

// ════════════════════════════════════════════════════════
// INITIALIZATION FUNCTION (Run manually once)
// ════════════════════════════════════════════════════════

function initSheets() {
  const sheets = Object.values(SHEET_NAMES);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  sheets.forEach(name => {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
      Logger.log(`Created sheet: ${name}`);
    }
  });
  
  // Initialize Settings with defaults
  const settingsSheet = getSheet(SHEET_NAMES.SETTINGS);
  if (settingsSheet.getLastRow() === 0) {
    settingsSheet.appendRow([
      'store_name', 'tagline', 'address', 'phone', 'footer',
      'logo_url', 'qris_url',
      'bank1_nama', 'bank1_norek', 'bank1_penerima',
      'bank2_nama', 'bank2_norek', 'bank2_penerima',
      'sales_list', 'gas_url', 'sheets_id',
      'updated_at'
    ]);
    settingsSheet.appendRow([
      'TIRTA KENCANA',
      'DO GOOD AND GOD WILL COME TO YOU',
      'WTC MANGGA DUA LANTAIUG BLOK A/53',
      '088211058000',
      'TERIMA KASIH ATAS PEMBELIAN ANDA',
      'https://www.image2url.com/r2/default/images/1781919958885-4a0fd859-0e6f-4e9f-9370-702fbf72864f.jpg',
      'https://www.image2url.com/r2/default/images/1781920179544-5e92dd0f-d0fa-482c-afc2-88259a5d3000.jpg',
      'BCA', '6930099099', 'HENDRI',
      'bluBCA', '002283588888', 'HENDRI',
      '[]',
      '',
      SpreadsheetApp.getActiveSpreadsheet().getId(),
      timestamp()
    ]);
  }
  
  Logger.log('Initialization complete!');
}
